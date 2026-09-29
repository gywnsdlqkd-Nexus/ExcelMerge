# -*- coding: utf-8 -*-
"""release.py 의 판단 로직 — 버전 bump · CHANGELOG 승격 · 릴리스 노트 추출.

손으로 돌리던 배포 절차를 한 명령으로 묶으면서, 실제로 났던 사고 두 가지를 스크립트가
막도록 했다. 여기서 고정하는 것은 그 '막는 규칙'이다.

  1. 버전을 올린 **뒤에** 테스트를 돌린다 — 새 버전 번호를 하드코딩한 테스트가 있으면
     옛 순서(bump 전 테스트)로는 못 잡는다. (순서는 main() 의 코드 순서로 고정.)
  2. [미배포] 가 비어 있으면 시작하지 않는다 — 본문이 빈 릴리스가 나가는 것을 막는다.
"""
import inspect

import pytest

import release


# ── 버전 ─────────────────────────────────────────────────────────────────────

SRC = '__version__ = "202"\n\nother = 1\n'


def test_bump_sets_the_new_version():
    assert release.bump_version(SRC, "203").startswith('__version__ = "203"')


def test_bump_refuses_to_go_backwards():
    with pytest.raises(ValueError, match="뒤로"):
        release.bump_version(SRC, "201")


def test_bump_refuses_the_same_version():
    with pytest.raises(ValueError, match="뒤로"):
        release.bump_version(SRC, "202")


def test_bump_needs_the_marker():
    with pytest.raises(ValueError):
        release.bump_version("version = 1\n", "203")


def test_current_version_reads_it():
    assert release.current_version(SRC) == "202"


# ── CHANGELOG ────────────────────────────────────────────────────────────────

CHG = """# Changelog

머리말

## [미배포]

### 수정

- 뭔가 고쳤다
- 또 고쳤다

## [202] - 2026-09-29

### 성능

- 예전 내용
"""

EMPTY = """# Changelog

## [미배포]

## [202] - 2026-09-29

- 예전 내용
"""


def test_unreleased_body_reads_only_that_section():
    body = release.unreleased_body(CHG)
    assert "뭔가 고쳤다" in body and "또 고쳤다" in body
    assert "예전 내용" not in body, "다음 절까지 먹었다"


def test_unreleased_body_is_empty_when_nothing_is_written():
    assert release.unreleased_body(EMPTY) == ""


def test_promote_creates_the_dated_section_and_keeps_unreleased():
    new = release.promote_changelog(CHG, "203", "2026-10-01")
    assert "## [203] - 2026-10-01" in new
    assert new.count("## [미배포]") == 1, "다음 릴리스를 적을 빈 절이 없다"
    # 승격 후에도 [미배포] 는 비어 있어야 한다(내용은 203 절로 갔다).
    assert release.unreleased_body(new) == ""
    assert "뭔가 고쳤다" in release.notes_for(new, "203")


def test_promote_refuses_an_empty_unreleased_section():
    with pytest.raises(ValueError, match="비어"):
        release.promote_changelog(EMPTY, "203", "2026-10-01")


def test_promote_needs_the_unreleased_marker():
    with pytest.raises(ValueError, match="미배포"):
        release.promote_changelog("# Changelog\n\n## [202] - x\n", "203", "d")


def test_notes_are_exactly_that_version_section():
    new = release.promote_changelog(CHG, "203", "2026-10-01")
    notes = release.notes_for(new, "203")
    assert notes.startswith("### 수정")
    assert "예전 내용" not in notes, "이전 버전 내용이 릴리스 노트에 섞였다"
    assert "## [" not in notes, "다음 절 머리말이 노트에 들어갔다"


def test_notes_roundtrip_matches_what_was_written():
    """승격 전 [미배포] 본문과 승격 후 노트가 같아야 한다 — 손실 없이 옮겨야 하므로."""
    before = release.unreleased_body(CHG)
    new = release.promote_changelog(CHG, "203", "2026-10-01")
    assert release.notes_for(new, "203") == before


def test_notes_for_unknown_version_raises():
    with pytest.raises(ValueError):
        release.notes_for(CHG, "999")


# ── 순서 보증 ────────────────────────────────────────────────────────────────

def test_tests_run_after_the_version_bump():
    """bump → 테스트 순서가 뒤집히면 버전에 묶인 테스트를 못 잡는다(실제로 겪은 사고)."""
    src = inspect.getsource(release.main)
    i_bump = src.index("버전 bump + CHANGELOG 승격")
    i_test = src.index("테스트 게이트")
    i_commit = src.index('step("커밋")')
    assert i_bump < i_test < i_commit, "bump → 테스트 → 커밋 순서가 아니다"


def test_publish_comes_after_build_and_smoke():
    """스모크를 통과하지 못한 빌드가 게시되면 안 된다."""
    src = inspect.getsource(release.main)
    assert src.index('step("빌드")') < src.index('step("exe 실행 스모크")')
    assert src.index('step("exe 실행 스모크")') < src.index('step("게시")')
    assert src.index('step("푸시")') < src.index('step("게시")')


# ── git status 파싱 ──────────────────────────────────────────────────────────
# `git status --porcelain` 은 줄 앞 두 칸이 상태 코드다(' M path' 처럼 앞이 공백일 수
# 있다). 출력을 통째로 strip 하면 첫 줄의 그 공백이 사라져 경로가 한 칸 밀리고, 허용
# 목록에 있는 파일이 '막는 변경'으로 잡힌다 — 실제로 그렇게 오작동했다.

ALLOWED = {"excelmerge/__init__.py", "CHANGELOG.md"}


def test_allowed_files_do_not_block():
    porcelain = " M CHANGELOG.md\nM  excelmerge/__init__.py\n"
    assert release.blocking_changes(porcelain, ALLOWED) == []


def test_other_changes_block():
    porcelain = " M CHANGELOG.md\n?? release.py\n M excelmerge/diff_view.py\n"
    blocked = release.blocking_changes(porcelain, ALLOWED)
    assert len(blocked) == 2
    assert any("release.py" in b for b in blocked)
    assert any("diff_view.py" in b for b in blocked)


def test_renamed_file_is_judged_by_its_new_path():
    assert release.blocking_changes("R  old.py -> CHANGELOG.md\n", ALLOWED) == []
    assert release.blocking_changes("R  CHANGELOG.md -> new.py\n", ALLOWED) != []


def test_quoted_path_with_spaces():
    assert release.blocking_changes('?? "some file.py"\n', ALLOWED) != []


def test_empty_status_means_clean():
    assert release.blocking_changes("", ALLOWED) == []
