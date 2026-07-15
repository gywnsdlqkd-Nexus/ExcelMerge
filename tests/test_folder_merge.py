# -*- coding: utf-8 -*-
"""폴더 병합 순수 복사 로직(copy_pairs) 테스트 — 비동기 워커에서 분리된 부분."""
import os

from excelmerge.workers import copy_pairs


def test_copies_creating_dirs(tmp_path):
    src = tmp_path / "src.txt"
    src.write_text("new", encoding="utf-8")
    dst = tmp_path / "out" / "dst.txt"          # 아직 없는 대상(디렉터리도 없음)

    done, fails, merged = copy_pairs([(str(src), str(dst), "out/dst.txt")])

    assert done == 1 and not fails
    assert merged == ["out/dst.txt"]
    assert dst.read_text(encoding="utf-8") == "new"
    assert not os.path.exists(str(dst) + ".bak")   # 백업 비활성


def test_overwrite_no_bak(tmp_path):
    """기존 대상을 덮어쓸 때 .bak 백업을 만들지 않는다(사용자 요청으로 제거)."""
    src = tmp_path / "src.txt"; src.write_text("new", encoding="utf-8")
    dst = tmp_path / "dst.txt"; dst.write_text("old", encoding="utf-8")

    done, fails, merged = copy_pairs([(str(src), str(dst), "dst.txt")])

    assert done == 1 and not fails
    assert dst.read_text(encoding="utf-8") == "new"
    assert not os.path.exists(str(dst) + ".bak"), ".bak 를 만들면 안 됨(백업 비활성)"


def test_skips_empty_src_and_collects_fails(tmp_path):
    ok_src = tmp_path / "ok.txt"; ok_src.write_text("x", encoding="utf-8")
    pairs = [
        ("", str(tmp_path / "skip.txt"), "skip.txt"),           # 빈 src → 스킵
        (str(tmp_path / "missing.txt"), str(tmp_path / "d.txt"), "d.txt"),  # 없는 src → 실패
        (str(ok_src), str(tmp_path / "ok_copy.txt"), "ok_copy.txt"),
    ]
    done, fails, merged = copy_pairs(pairs)

    assert done == 1
    assert merged == ["ok_copy.txt"]
    assert len(fails) == 1 and fails[0].startswith("d.txt:")
    assert not os.path.exists(str(tmp_path / "skip.txt"))


def test_progress_callback(tmp_path):
    srcs = []
    for i in range(3):
        p = tmp_path / f"s{i}.txt"; p.write_text(str(i), encoding="utf-8"); srcs.append(p)
    pairs = [(str(p), str(tmp_path / f"o{i}.txt"), f"o{i}.txt") for i, p in enumerate(srcs)]
    seen = []
    copy_pairs(pairs, progress=lambda d, t: seen.append((d, t)))
    assert seen and seen[-1] == (3, 3)
