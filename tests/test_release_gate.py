# -*- coding: utf-8 -*-
"""릴리스가 실제 빌드 데이터로 전수 머지를 돌려 보는가 — 그리고 안 돌릴 땐 말하는가.

pytest 로는 저장 쪽 회귀를 다 못 잡는다. 합성 조작("셀 하나 고치기", "행 몇 개
지우기")은 실제 머지가 밟는 경로(열 매칭 · 빈 열/행 삭제 승격 · 행 삽입 · 서식 병합)를
밟지 않는다. v215 에서 터진 셋도, v222 의 '끝 행 삭제 + 삽입' 도 merge_check 에서만
나왔다.

문제는 그게 **사람이 기억해야 하는 규칙**이었다는 것이다. RELEASE.md 에 "저장 쪽을
고쳤다면 돌릴 것" 이라 적혀 있었는데 v219·v220·v221 이 그대로 나갔다. 그래서 릴리스
절차에 붙였다.

설정이 없으면 건너뛸 수밖에 없다(경로는 머신마다 다르다). 그때 **조용히 넘어가지
않는 것**이 이 파일이 지키는 핵심이다 — 조용하면 '돌렸겠거니' 가 되고, 그게 처음에
빠뜨린 이유였다.
"""
import json
import os
import subprocess

import pytest

import release

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _clear(monkeypatch):
    monkeypatch.delenv("EXCELMERGE_CHECK_A", raising=False)
    monkeypatch.delenv("EXCELMERGE_CHECK_B", raising=False)


# ── 설정 읽기 ────────────────────────────────────────────────────────────────

def test_no_config_is_not_an_error(monkeypatch):
    """경로는 머신마다 다르다 — 안 정해 뒀다고 릴리스를 막으면 안 된다."""
    _clear(monkeypatch)
    assert release.merge_check_dirs() is None


def test_only_one_side_configured_stops_the_release(monkeypatch, tmp_path):
    """한쪽만 설정한 건 설정 실수다 — 반쪽으로 돌리는 것보다 멈추는 게 낫다."""
    _clear(monkeypatch)
    monkeypatch.setenv("EXCELMERGE_CHECK_A", str(tmp_path))
    with pytest.raises(SystemExit):
        release.merge_check_dirs()


def test_a_missing_folder_stops_the_release(monkeypatch, tmp_path):
    _clear(monkeypatch)
    monkeypatch.setenv("EXCELMERGE_CHECK_A", str(tmp_path))
    monkeypatch.setenv("EXCELMERGE_CHECK_B", str(tmp_path / "없는폴더"))
    with pytest.raises(SystemExit):
        release.merge_check_dirs()


def test_quotes_around_a_path_are_tolerated(monkeypatch, tmp_path):
    """set VAR="D:\\..." 로 넣으면 따옴표가 값에 섞여 들어온다."""
    _clear(monkeypatch)
    a = tmp_path / "a"; b = tmp_path / "b"
    a.mkdir(); b.mkdir()
    monkeypatch.setenv("EXCELMERGE_CHECK_A", f'"{a}"')
    monkeypatch.setenv("EXCELMERGE_CHECK_B", f'"{b}"')
    assert release.merge_check_dirs() == (str(a), str(b))


# ── 건너뛸 때 반드시 말한다 ──────────────────────────────────────────────────

def test_skipping_without_config_is_announced(monkeypatch, capsys):
    """조용히 넘어가면 '돌렸겠거니' 가 된다 — 실제로 그래서 세 번 빠뜨렸다."""
    _clear(monkeypatch)
    release.merge_check_gate(False, os.environ.copy())
    said = capsys.readouterr().out
    assert "건너뜀" in said
    assert "EXCELMERGE_CHECK_A" in said, "어떻게 켜는지 알려 주지 않는다"


def test_explicit_skip_is_announced(monkeypatch, capsys):
    _clear(monkeypatch)
    release.merge_check_gate(True, os.environ.copy())
    said = capsys.readouterr().out
    assert "건너뜀" in said and "skip-merge-check" in said


def test_skipping_does_not_run_anything(monkeypatch):
    _clear(monkeypatch)
    calls = []
    monkeypatch.setattr(release, "run", lambda *a, **k: calls.append(a))
    release.merge_check_gate(False, os.environ.copy())
    assert calls == []


# ── 설정돼 있으면 실제로 돌린다 ──────────────────────────────────────────────

def test_it_runs_merge_check_with_the_excel_verdict(monkeypatch, tmp_path):
    """--excel 이 최종 판정이다. 구조 검사가 놓치는 것을 엑셀이 잡는다."""
    _clear(monkeypatch)
    a = tmp_path / "a"; b = tmp_path / "b"
    a.mkdir(); b.mkdir()
    monkeypatch.setenv("EXCELMERGE_CHECK_A", str(a))
    monkeypatch.setenv("EXCELMERGE_CHECK_B", str(b))
    seen = []
    monkeypatch.setattr(release, "run", lambda cmd, **k: seen.append(cmd))
    release.merge_check_gate(False, os.environ.copy())
    assert len(seen) == 1, seen
    cmd = seen[0]
    assert "merge_check.py" in cmd
    assert cmd[cmd.index("merge_check.py") + 1] == str(a)
    assert cmd[cmd.index("merge_check.py") + 2] == str(b)
    assert "--excel" in cmd, "엑셀 판정 없이 돌리면 절반만 보는 것이다"


def test_the_output_folder_is_temporary_and_cleaned(monkeypatch, tmp_path):
    """원본 옆에 결과를 흘리면 안 되고, 쌓여도 안 된다."""
    _clear(monkeypatch)
    a = tmp_path / "a"; b = tmp_path / "b"
    a.mkdir(); b.mkdir()
    monkeypatch.setenv("EXCELMERGE_CHECK_A", str(a))
    monkeypatch.setenv("EXCELMERGE_CHECK_B", str(b))
    seen = []
    monkeypatch.setattr(release, "run", lambda cmd, **k: seen.append(cmd))
    release.merge_check_gate(False, os.environ.copy())
    out_dir = seen[0][seen[0].index("merge_check.py") + 3]
    assert str(a) not in out_dir and str(b) not in out_dir
    assert not os.path.exists(out_dir), "임시 폴더를 치우지 않는다"


def test_a_failing_check_stops_the_release(monkeypatch, tmp_path):
    """merge_check 가 1 을 돌려주면 릴리스가 거기서 멈춰야 한다."""
    _clear(monkeypatch)
    a = tmp_path / "a"; b = tmp_path / "b"
    a.mkdir(); b.mkdir()
    monkeypatch.setenv("EXCELMERGE_CHECK_A", str(a))
    monkeypatch.setenv("EXCELMERGE_CHECK_B", str(b))

    def boom(cmd, **k):
        release.fail("merge_check 실패")

    monkeypatch.setattr(release, "run", boom)
    with pytest.raises(SystemExit):
        release.merge_check_gate(False, os.environ.copy())


# ── 절차 안에서의 자리 ───────────────────────────────────────────────────────

def test_the_gate_runs_before_the_commit():
    """실패하면 아무것도 커밋되지 않아야 한다 — 되돌릴 일이 없게."""
    src = open(os.path.join(ROOT, "release.py"), encoding="utf-8").read()
    body = src[src.index("def main("):]
    assert body.index("merge_check_gate(") < body.index('step("커밋")'), \
        "커밋 뒤에 검사한다 — 실패하면 커밋을 되돌려야 한다"


def test_the_gate_runs_after_the_version_bump():
    """pytest 와 같은 자리 — 버전을 올린 뒤에 돌려야 버전에 묶인 것도 함께 본다."""
    src = open(os.path.join(ROOT, "release.py"), encoding="utf-8").read()
    body = src[src.index("def main("):]
    assert body.index('step("버전 bump') < body.index("merge_check_gate(")


# ── 자체 점검(--selfcheck) 게이트 ───────────────────────────────────────────
#
# exe_smoke 는 창이 뜨는지만 본다. v221(저장 뒤 Ctrl+S 오류 창)은 그걸 통과하고 게시된
# 뒤 실기에서야 드러났다. 그 확인을 게시 앞으로 당긴다. 예전엔 uicheck.py 가 실제
# 키보드·마우스로 했는데, 사용자의 마우스를 가져갔고 같은 exe 가 통과했다 실패했다 했다.
# 이제 빌드된 exe 가 화면 없이 스스로 돌고 JSON 으로 답한다.

def _clear_sc(monkeypatch):
    monkeypatch.delenv("EXCELMERGE_SELFCHECK_A", raising=False)
    monkeypatch.delenv("EXCELMERGE_SELFCHECK_B", raising=False)


def _set_sc(monkeypatch, tmp_path):
    _clear_sc(monkeypatch)
    a = tmp_path / "a.xlsx"; b = tmp_path / "b.xlsx"
    a.write_text("x", encoding="utf-8"); b.write_text("y", encoding="utf-8")
    monkeypatch.setenv("EXCELMERGE_SELFCHECK_A", str(a))
    monkeypatch.setenv("EXCELMERGE_SELFCHECK_B", str(b))
    return str(a), str(b)


def _result(ok=True, version="225"):
    return {"version": version, "seconds": 4.0, "ok": ok,
            "checks": [{"name": "비교 완료", "ok": True, "detail": ""},
                       {"name": "B 가 저장됨", "ok": ok, "detail": ""}]}


def _fake_exe(monkeypatch, result, code=0, seen=None, raise_exc=None):
    """빌드된 exe 대신 — 마지막 인자(결과 경로)에 JSON 을 쓰고 종료 코드를 돌려준다."""
    def fake(cmd, **kw):
        if seen is not None:
            seen.append((cmd, kw))
        if raise_exc is not None:
            raise raise_exc
        if result is not None:
            with open(cmd[-1], "w", encoding="utf-8") as f:
                json.dump(result, f, ensure_ascii=False)
        return subprocess.CompletedProcess(cmd, code)
    monkeypatch.setattr(release.subprocess, "run", fake)


def test_sc_files_are_optional(monkeypatch):
    _clear_sc(monkeypatch)
    assert release.selfcheck_files() is None


def test_sc_only_one_side_stops_the_release(monkeypatch, tmp_path):
    _clear_sc(monkeypatch)
    f = tmp_path / "a.xlsx"; f.write_text("x", encoding="utf-8")
    monkeypatch.setenv("EXCELMERGE_SELFCHECK_A", str(f))
    with pytest.raises(SystemExit):
        release.selfcheck_files()


def test_sc_a_missing_file_stops_the_release(monkeypatch, tmp_path):
    _clear_sc(monkeypatch)
    f = tmp_path / "a.xlsx"; f.write_text("x", encoding="utf-8")
    monkeypatch.setenv("EXCELMERGE_SELFCHECK_A", str(f))
    monkeypatch.setenv("EXCELMERGE_SELFCHECK_B", str(tmp_path / "없음.xlsx"))
    with pytest.raises(SystemExit):
        release.selfcheck_files()


def test_sc_skipping_is_announced(monkeypatch, capsys):
    """조용히 넘어가면 '돌렸겠거니' 가 된다 — merge_check 에서 이미 겪었다."""
    _clear_sc(monkeypatch)
    release.selfcheck_gate("999", False, os.environ.copy())
    said = capsys.readouterr().out
    assert "건너뜀" in said and "EXCELMERGE_SELFCHECK_A" in said


def test_sc_explicit_skip_is_announced(monkeypatch, capsys):
    _clear_sc(monkeypatch)
    release.selfcheck_gate("999", True, os.environ.copy())
    said = capsys.readouterr().out
    assert "건너뜀" in said and "skip-selfcheck" in said


def test_sc_runs_the_built_exe_in_selfcheck_mode(monkeypatch, tmp_path):
    """설치본이 아니라 **방금 빌드한** exe 를 — 그게 나갈 물건이다."""
    a, b = _set_sc(monkeypatch, tmp_path)
    seen = []
    _fake_exe(monkeypatch, _result(), seen=seen)
    release.selfcheck_gate("225", False, os.environ.copy())
    (cmd, kw), = seen
    assert cmd[0].endswith(os.path.join("dist", "ExcelMerge", "ExcelMerge.exe"))
    assert cmd[1:4] == ["--selfcheck", a, b]
    assert kw.get("timeout"), "멈춘 exe 에 릴리스가 매달린다"


def test_sc_does_not_take_over_the_keyboard_and_mouse(monkeypatch, tmp_path, capsys):
    """사용자의 입력 장치를 쓰는 도구(uicheck.py)를 부르지 않는다."""
    _set_sc(monkeypatch, tmp_path)
    seen = []
    _fake_exe(monkeypatch, _result(), seen=seen)
    release.selfcheck_gate("225", False, os.environ.copy())
    assert not any("uicheck" in str(part) for cmd, _ in seen for part in cmd)
    assert "마우스 사용 없음" in capsys.readouterr().out


def test_sc_a_failing_check_stops_the_release(monkeypatch, tmp_path):
    _set_sc(monkeypatch, tmp_path)
    _fake_exe(monkeypatch, _result(ok=False), code=1)
    with pytest.raises(SystemExit):
        release.selfcheck_gate("225", False, os.environ.copy())


def test_sc_an_exit_code_alone_stops_the_release(monkeypatch, tmp_path):
    """JSON 이 ok 라고 해도 종료 코드가 0 이 아니면 믿지 않는다."""
    _set_sc(monkeypatch, tmp_path)
    _fake_exe(monkeypatch, _result(ok=True), code=1)
    with pytest.raises(SystemExit):
        release.selfcheck_gate("225", False, os.environ.copy())


def test_sc_a_missing_result_stops_the_release(monkeypatch, tmp_path):
    """exe 가 점검 모드로 못 뜨면(번들 누락 등) 결과 파일이 비어 있다 — 통과로 보면 안 된다."""
    _set_sc(monkeypatch, tmp_path)
    _fake_exe(monkeypatch, None, code=0)
    with pytest.raises(SystemExit):
        release.selfcheck_gate("225", False, os.environ.copy())


def test_sc_the_wrong_build_stops_the_release(monkeypatch, tmp_path):
    """dist 에 옛 빌드가 남아 있으면 그걸 점검하게 된다 — 버전으로 막는다."""
    _set_sc(monkeypatch, tmp_path)
    _fake_exe(monkeypatch, _result(version="224"))
    with pytest.raises(SystemExit):
        release.selfcheck_gate("225", False, os.environ.copy())


def test_sc_a_hung_exe_stops_the_release(monkeypatch, tmp_path):
    _set_sc(monkeypatch, tmp_path)
    _fake_exe(monkeypatch, None, raise_exc=subprocess.TimeoutExpired("exe", 300))
    with pytest.raises(SystemExit):
        release.selfcheck_gate("225", False, os.environ.copy())


def test_sc_the_result_file_is_cleaned(monkeypatch, tmp_path):
    _set_sc(monkeypatch, tmp_path)
    seen = []
    _fake_exe(monkeypatch, _result(), seen=seen)
    release.selfcheck_gate("225", False, os.environ.copy())
    assert not os.path.exists(seen[0][0][-1])


def test_sc_runs_after_the_smoke_and_before_publishing():
    """스모크가 통과한 exe 를 몰아 보고, 그 뒤에야 올린다."""
    src = open(os.path.join(ROOT, "release.py"), encoding="utf-8").read()
    body = src[src.index("def main("):]
    assert body.index("exe_smoke(args.version)") < body.index("selfcheck_gate("), \
        "스모크보다 먼저 몰아 본다"
    assert body.index("selfcheck_gate(") < body.index('step("푸시")'), \
        "게시한 뒤에 확인한다 — 늦다"
