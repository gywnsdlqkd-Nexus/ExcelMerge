# -*- coding: utf-8 -*-
"""folder_compare 백엔드 단위 테스트 — 재귀 매칭 + same/modified/only_* 분류.

순수 로직(UI 비의존)이라 offscreen Qt 불필요.
실행: python -m pytest tests/test_folder_compare.py
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")   # 비동기 뷰 테스트용(Qt)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import excelmerge.folder_compare as fc
from excelmerge.folder_compare import (
    compare_folders, scan_folder, files_equal, summarize,
    SAME, MODIFIED, ONLY_A, ONLY_B,
)


def _write(path: str, data: bytes):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as f:
        f.write(data)


def _setup(tmp_path):
    a = os.path.join(tmp_path, "A")
    b = os.path.join(tmp_path, "B")
    # 공통(동일)
    _write(os.path.join(a, "same.xlsx"), b"XYZ")
    _write(os.path.join(b, "same.xlsx"), b"XYZ")
    # 공통(변경 — 내용 다름)
    _write(os.path.join(a, "changed.json"), b"11111")
    _write(os.path.join(b, "changed.json"), b"22222")
    # 공통(변경 — 크기 다름)
    _write(os.path.join(a, "sub/deep.xlsx"), b"AAA")
    _write(os.path.join(b, "sub/deep.xlsx"), b"AAAA")
    # A에만
    _write(os.path.join(a, "only_a.uasset"), b"a")
    # B에만
    _write(os.path.join(b, "sub/only_b.xlsx"), b"b")
    # 지원하지 않는 확장자 — 무시돼야 함
    _write(os.path.join(a, "note.txt"), b"ignore")
    _write(os.path.join(b, "note.txt"), b"ignore-diff")
    return a, b


def test_scan_folder_recursive_and_filter(tmp_path):
    a, _b = _setup(str(tmp_path))
    scanned = scan_folder(a)
    keys = set(scanned.keys())
    assert "same.xlsx" in keys
    assert "sub/deep.xlsx" in keys          # 재귀 하위폴더
    assert "only_a.uasset" in keys
    assert "note.txt" not in keys           # 미지원 확장자 제외
    # 값은 (표시경로, 절대경로)
    disp, absp = scanned["sub/deep.xlsx"]
    assert disp == "sub/deep.xlsx" and os.path.isabs(absp)


def test_compare_folders_classification(tmp_path):
    a, b = _setup(str(tmp_path))
    entries = compare_folders(a, b)
    by_rel = {e.rel_path: e.status for e in entries}

    assert by_rel["same.xlsx"] == SAME
    assert by_rel["changed.json"] == MODIFIED       # 같은 크기·다른 내용
    assert by_rel["sub/deep.xlsx"] == MODIFIED      # 다른 크기
    assert by_rel["only_a.uasset"] == ONLY_A
    assert by_rel["sub/only_b.xlsx"] == ONLY_B
    assert "note.txt" not in by_rel                 # 미지원 제외

    # 정렬 순서(rel_key 기준) 보장
    rels = [e.rel_path for e in entries]
    assert rels == sorted(rels, key=str.lower)

    counts = summarize(entries)
    assert counts[SAME] == 1
    assert counts[MODIFIED] == 2
    assert counts[ONLY_A] == 1
    assert counts[ONLY_B] == 1


def test_only_one_side(tmp_path):
    a, b = _setup(str(tmp_path))
    # B 폴더 미지정 → A 파일 전부 only_a
    entries = compare_folders(a, "")
    assert entries and all(e.status == ONLY_A for e in entries)
    assert all(e.path_b == "" for e in entries)


def test_files_equal(tmp_path):
    p1 = os.path.join(str(tmp_path), "x.bin")
    p2 = os.path.join(str(tmp_path), "y.bin")
    _write(p1, b"hello world" * 100000)   # 청크 경계 넘김
    _write(p2, b"hello world" * 100000)
    assert files_equal(p1, p2)
    _write(p2, b"hello world" * 100000 + b"!")
    assert not files_equal(p1, p2)
    assert not files_equal(p1, os.path.join(str(tmp_path), "missing.bin"))


def test_progress_callback(tmp_path):
    """진행 콜백이 (done,total)로 호출되고 단조 증가, 최종 total 도달 (Part B)."""
    a, b = _setup(str(tmp_path))
    calls = []
    compare_folders(a, b, progress=lambda d, t: calls.append((d, t)))
    assert calls, "progress 콜백 미호출"
    total = calls[-1][1]
    assert total == 3, f"매칭 쌍 3개여야: total={total}"   # same/changed/deep
    dones = [d for d, _ in calls]
    assert dones == sorted(dones), "done 단조 증가 아님"
    assert dones[-1] == total, "마지막 done이 total 미도달"


def test_pair_equal_cache(tmp_path):
    """(경로,mtime,size) 캐시 — 재스캔 시 안 바뀐 파일은 재비교 생략, 변경 파일만 재계산 (Part C)."""
    fc._pair_equal_cache.clear()
    a, b = _setup(str(tmp_path))

    real_files_equal = fc.files_equal
    count = {"n": 0}

    def _spy(pa, pb):
        count["n"] += 1
        return real_files_equal(pa, pb)

    fc.files_equal = _spy
    try:
        compare_folders(a, b)
        first = count["n"]
        assert first == 3, f"첫 스캔은 매칭 3쌍 비교: {first}"

        count["n"] = 0
        compare_folders(a, b)                 # 변경 없음 → 전부 캐시 히트
        assert count["n"] == 0, f"재스캔인데 재비교 발생: {count['n']}"

        # 한 파일 실제 변경(내용+크기 → mtime/size 달라짐) → 그 쌍만 재계산
        _write(os.path.join(b, "changed.json"), b"3333333333")
        count["n"] = 0
        compare_folders(a, b)
        assert count["n"] == 1, f"변경 1건만 재계산돼야: {count['n']}"
    finally:
        fc.files_equal = real_files_equal


def test_async_scan_view(tmp_path):
    """FolderCompareView.set_folders → 백그라운드 스캔 후 _entries가 동기 결과와 동일 (Part A)."""
    from PyQt5.QtWidgets import QApplication
    from PyQt5.QtCore import QEventLoop, QTimer
    from excelmerge.folder_view import FolderCompareView
    from excelmerge.statusbar import StatusBar

    app = QApplication.instance() or QApplication([])
    a, b = _setup(str(tmp_path))
    view = FolderCompareView(StatusBar())
    view.set_folders(a, b)

    loop = QEventLoop()

    def _chk():
        if view._entries:
            loop.quit()
        else:
            QTimer.singleShot(20, _chk)

    QTimer.singleShot(20, _chk)
    QTimer.singleShot(8000, loop.quit)   # 안전 타임아웃
    loop.exec_()

    by_rel = {e.rel_path: e.status for e in view._entries}
    assert by_rel.get("same.xlsx") == SAME
    assert by_rel.get("changed.json") == MODIFIED
    assert by_rel.get("only_a.uasset") == ONLY_A
    view.shutdown()


if __name__ == "__main__":
    import pytest
    pytest.main([__file__, "-q"])
