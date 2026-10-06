# -*- coding: utf-8 -*-
"""실제 빌드 데이터로 머지를 전부 돌려 보고 결과가 성한지 확인한다.

**왜 있나.** 저장 쪽을 고칠 때마다 합성 조작으로 확인했다 — "셀 하나 고치기",
"행 몇 개 지우기". 그런데 실제 머지는 그보다 훨씬 많은 일을 한다(열 매칭, 빈 열·행
삭제 승격, 행 삽입, 서식 병합). v215 에서 터진 셋은 전부 **행 삽입** 경로였고, 합성
조작은 그 경로를 한 번도 밟지 않아 v212~v214 를 통과해 버렸다.

이 도구는 앱과 **같은 경로**를 탄다(StagedMergeWorker.run 과 같은 순서):

    비교(열 매칭) → 스테이징 → 방향별 패치 구성 → 빈 열·행 삭제 승격 → 저장

원본은 **읽기만** 한다. B 쪽을 출력 폴더로 복사해 거기에만 쓴다.

사용:
    python merge_check.py <A폴더> <B폴더> <출력폴더> [--excel] [--limit N]

      --excel   결과를 엑셀로 실제 열어 본다(Windows + Excel 필요). 이게 최종 판정이다 —
                구조 검사가 놓치는 것을 엑셀은 잡는다. 실제로 여덟 가지 원인 중 넷은
                구조 검사가 아니라 '엑셀이 열어 주는가' 로 처음 드러났다.
      --limit   앞에서부터 N 개만(빠른 확인용).

하나라도 어긋나면 종료 코드 1. 저장 쪽을 고쳤다면 릴리스 전에 한 번 돌릴 것.
"""
import argparse
import os
import shutil
import stat
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

for _stream in (sys.stdout, sys.stderr):      # cp949 콘솔에서 진행 문구로 죽지 않게
    try:
        _stream.reconfigure(errors="replace")
    except (AttributeError, OSError):
        pass

# 사용자 설정(키 기억·검사 제외 열)을 건드리지 않는다.
os.environ["APPDATA"] = tempfile.mkdtemp(prefix="merge_check_")

from excelmerge.constants import DIR_A2B                       # noqa: E402
from excelmerge.diff_engine import (compute_diff, match_columns,  # noqa: E402
                                    usable_col_meta)
from excelmerge.loaders import clear_values_cache, load_values_any  # noqa: E402
from excelmerge.merge_build import build_side_patches           # noqa: E402
from excelmerge.staging import stageable_cells                  # noqa: E402
from excelmerge.xlsx_writer import (_formula_mismatches,        # noqa: E402
                                    _package_mismatches,
                                    _promote_empty_cols_to_delete,
                                    _sheet_order_mismatches,
                                    _write_patches_to_file)

SHEETS = (".xlsx", ".xlsm")


def structural_problems(path) -> list:
    """저장이 끝난 파일을 구조로 훑는다 — 저장 검증이 보는 것과 같은 눈."""
    return (_package_mismatches(path) + _sheet_order_mismatches(path)
            + _formula_mismatches(path))


def merge_one(name, a_dir, b_dir, out_dir):
    """A→B 로 **전부** 병합해 out_dir 의 복사본에 쓴다.

    반환: None(바뀔 것 없음) 또는 {'patches', 'inserts', 'deleted_rows', 'deleted_cols'}
    """
    dst = os.path.join(out_dir, name)
    shutil.copyfile(os.path.join(b_dir, name), dst)
    os.chmod(dst, stat.S_IWRITE | stat.S_IREAD)   # P4V 읽기 전용 속성이 따라온다

    clear_values_cache()
    a = load_values_any(os.path.join(a_dir, name))
    clear_values_cache()
    b = load_values_any(os.path.join(b_dir, name))
    if not a or not b:
        os.remove(dst)
        return None

    col_meta = usable_col_meta(match_columns(a, b, 0), 0)
    matrix, row_meta = compute_diff(a, b, key_col=0, key_row=0, col_meta=col_meta)
    cells = stageable_cells(matrix, [(r, c) for r in range(len(matrix))
                                     for c in range(len(matrix[r]))])
    staged = {rc: DIR_A2B for rc in cells}
    patches, insert_rows, style_src = build_side_patches(
        DIR_A2B, matrix, row_meta, staged, col_meta)
    if not (patches or insert_rows):
        os.remove(dst)
        return None

    patches, del_rows, del_cols = _promote_empty_cols_to_delete(
        patches, set(), dst, None)
    style_src = {k: v for k, v in style_src.items() if k in patches}
    _write_patches_to_file(dst, patches, list(insert_rows.values()),
                           del_rows, del_cols, None,
                           os.path.join(a_dir, name), None, style_src)
    return {"patches": len(patches), "inserts": len(insert_rows),
            "deleted_rows": len(del_rows), "deleted_cols": sorted(del_cols)}


def open_in_excel(paths) -> list:
    """엑셀로 실제 열어 본다. 열리지 않은 파일 이름들을 돌려준다.

    구조 검사가 놓치는 것을 엑셀은 잡는다 — 이게 최종 판정이다.
    """
    if not paths:
        return []
    listing = tempfile.mktemp(suffix=".txt")
    with open(listing, "w", encoding="utf-8") as fh:
        fh.write("\n".join(paths))
    script = r"""
$ErrorActionPreference = 'Stop'
$paths = Get-Content -LiteralPath $env:MC_LIST -Encoding UTF8
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$xl.AutomationSecurity = 3     # 매크로 차단
$xl.AskToUpdateLinks = $false
foreach ($p in $paths) {
  if (-not $p) { continue }
  try { $wb = $xl.Workbooks.Open($p, 0); $wb.Close($false) }
  catch { Write-Output "FAIL`t$p" }
}
$xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
"""
    env = dict(os.environ, MC_LIST=listing)
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-NonInteractive",
                              "-Command", script],
                             capture_output=True, text=True, encoding="utf-8",
                             errors="replace", env=env, timeout=3600)
    except (OSError, subprocess.TimeoutExpired) as e:
        print(f"  (엑셀 확인을 건너뜁니다 — {type(e).__name__})")
        return []
    finally:
        try:
            os.remove(listing)
        except OSError:
            pass
    bad = [ln.split("\t", 1)[1] for ln in (out.stdout or "").splitlines()
           if ln.startswith("FAIL\t")]
    if out.returncode and not bad:
        print(f"  (엑셀 확인을 건너뜁니다 — powershell 종료 {out.returncode})")
        print("  ", (out.stderr or "").strip().splitlines()[:2])
        return []
    return bad


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("a_dir", help="A(원본) 폴더 — 읽기만 한다")
    ap.add_argument("b_dir", help="B(대상) 폴더 — 읽기만 한다")
    ap.add_argument("out_dir", help="결과를 쓸 폴더(비어 있어야 안전하다)")
    ap.add_argument("--excel", action="store_true",
                    help="결과를 엑셀로 열어 본다(최종 판정)")
    ap.add_argument("--limit", type=int, default=0, help="앞에서부터 N 개만")
    args = ap.parse_args()

    for d in (args.a_dir, args.b_dir):
        if not os.path.isdir(d):
            print(f"폴더가 없습니다: {d}")
            return 2
    os.makedirs(args.out_dir, exist_ok=True)

    names = sorted(n for n in os.listdir(args.b_dir)
                   if n.lower().endswith(SHEETS)
                   and os.path.exists(os.path.join(args.a_dir, n)))
    if not names:
        print("두 폴더에 공통으로 있는 시트 파일이 없습니다.")
        return 2

    merged, skipped = [], 0
    rejected, broken = [], []
    for i, name in enumerate(names, 1):
        if args.limit and len(merged) >= args.limit:
            break
        try:
            info = merge_one(name, args.a_dir, args.b_dir, args.out_dir)
        except Exception as e:
            lines = str(e).splitlines()
            rejected.append((name, lines[1] if len(lines) > 1 else lines[0]))
            continue
        if info is None:
            skipped += 1
            continue
        merged.append((name, info))
        bad = structural_problems(os.path.join(args.out_dir, name))
        if bad:
            broken.append((name, bad[:3]))
        if len(merged) % 25 == 0:
            print(f"  ... {len(merged)} 머지 ({i}/{len(names)})", flush=True)

    print()
    print(f"대상 {len(names)} 쌍 — 머지 {len(merged)} / "
          f"바뀔 것 없음 {skipped} / 저장 거부 {len(rejected)}")

    changed = [(n, i) for n, i in merged if i["deleted_rows"] or i["inserts"]
               or i["deleted_cols"]]
    if changed:
        print(f"\n행·열이 늘거나 준 파일 {len(changed)}:")
        for n, i in changed:
            bits = []
            if i["inserts"]:
                bits.append(f"삽입 {i['inserts']}행")
            if i["deleted_rows"]:
                bits.append(f"삭제 {i['deleted_rows']}행")
            if i["deleted_cols"]:
                cols = i["deleted_cols"]              # 수백 개가 나오는 파일이 있다
                shown = ", ".join(cols[:6])
                more = f" 외 {len(cols) - 6}개" if len(cols) > 6 else ""
                bits.append(f"삭제 {len(cols)}열({shown}{more})")
            print(f"   {n:<44} {' / '.join(bits)}")

    if rejected:
        print(f"\n저장 검증이 막은 파일 {len(rejected)}:")
        for n, why in rejected:
            print(f"   {n:<44} {why[:90]}")

    if broken:
        print(f"\n구조가 어긋난 결과 {len(broken)}:")
        for n, bad in broken:
            print(f"   {n}")
            for b in bad:
                print(f"      {b}")

    excel_bad = []
    if args.excel:
        paths = [os.path.join(args.out_dir, n) for n, _ in merged]
        print(f"\n엑셀로 {len(paths)}개를 열어 보는 중...", flush=True)
        excel_bad = open_in_excel(paths)
        print(f"엑셀이 거부: {len(excel_bad)} 개")
        for p in excel_bad:
            print(f"   {os.path.basename(p)}")

    ok = not (rejected or broken or excel_bad)
    print()
    print("모두 성합니다." if ok else "문제가 있습니다 — 위를 보세요.")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
