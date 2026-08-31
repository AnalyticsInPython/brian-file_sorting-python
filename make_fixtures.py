#!/usr/bin/env python3
"""
make_fixtures.py — sort_requests.py 검증용 request/ 폴더를 만든다.

계획서의 엣지 케이스 표를 그대로 파일로 구현한다. 각 파일 옆의 주석이
sort_requests.py 가 내려야 할 기대 판정이다.

  python3 make_fixtures.py            # request/ 생성 (이미 있으면 거부)
  python3 make_fixtures.py --force    # 기존 request/ 를 지우고 다시 생성
"""

import argparse
import io
import shutil
import sys
import zipfile
from pathlib import Path

# --------------------------------------------------------------- 내용 생성기

CSV_A = "company,metric,value\napple,revenue,1000\n"
CSV_B = "company,metric,value\napple,revenue,9999\n"


def ole2_bytes(payload: bytes = b"") -> bytes:
    """레거시 .xls 로 인식되도록 OLE2 복합문서 시그니처를 붙인 더미."""
    return b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 504 + payload


def xlsx_bytes(sheet_note: str) -> bytes:
    """openpyxl 없이 최소한의 ZIP 컨테이너(=xlsx로 감지)를 만든다."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0"?><Types '
                   'xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        z.writestr("xl/worksheets/sheet1.xml", f"<worksheet><!-- {sheet_note} --></worksheet>")
    return buf.getvalue()


HTML_TABLE = (
    "<html><head><meta charset='utf-8'></head><body>"
    "<table><tr><th>company</th><th>value</th></tr>"
    "<tr><td>apple</td><td>1000</td></tr></table></body></html>"
)

# --------------------------------------------------------------- 픽스처 정의
# (상대경로, 내용) — 내용이 str이면 utf-8로, bytes면 그대로 쓴다.

FIXTURES = [
    # --- 정상 (요청 예시 그대로) -------------------------------------
    ("apple_20260212.xls",            ole2_bytes()),          # placed 2026/02/12
    ("anthropic_20260213.csv",        CSV_A),                 # placed 2026/02/13
    ("microsoft_20260215.xlsx",       xlsx_bytes("msft")),    # placed 2026/02/15

    # --- 날짜 구분자 변형 --------------------------------------------
    ("nvidia_2026-02-16.csv",         CSV_A),                 # placed 2026/02/16
    ("tesla 2026.02.17.xlsx",         xlsx_bytes("tsla")),    # placed 2026/02/17
    ("meta_2026_02_18.csv",           CSV_A),                 # placed 2026/02/18

    # --- 연도 걸침 ----------------------------------------------------
    ("openai_20251231.csv",           CSV_A),                 # placed 2025/12/31

    # --- 대소문자 차이 + 확장자 대문자 --------------------------------
    # macOS는 대소문자를 구분하지 않아 같은 폴더에는 공존할 수 없다. 하위 폴더에
    # 두어 "목적지에서" apple_20260212.xls 와 충돌하게 만든다. 내용이 다르므로 __2.
    ("archive/APPLE_20260212.XLS",    ole2_bytes(b"different")),  # placed + NAME_COLLISION

    # --- 회사명 오타/표기 차이 (유사 이름 리포트 대상) -----------------
    ("Anthropic_20260214.csv",        CSV_A),                 # placed, 유사 이름 클러스터
    ("anthropicc_20260218.csv",       CSV_A),                 # placed, 편집거리 1

    # --- 무효한 날짜 --------------------------------------------------
    ("broken_20260230.csv",           CSV_A),                 # unresolved INVALID_DATE (2월 30일)
    ("broken_20261340.csv",           CSV_A),                 # unresolved INVALID_DATE (13월 40일)

    # --- 모호한 날짜 --------------------------------------------------
    ("amazon_02122026.csv",           CSV_A),                 # unresolved YEAR_OUT_OF_RANGE
    ("google_260212.csv",             CSV_A),                 # unresolved NO_DATE (2자리 연도)

    # --- 날짜 없음 ----------------------------------------------------
    ("readme.txt",                    "이 폴더 설명 파일입니다.\n"),  # unresolved NO_DATE

    # --- 날짜 두 개 ---------------------------------------------------
    ("report_20260212_to_20260215.csv", CSV_A),               # unresolved AMBIGUOUS_MULTI_DATE

    # --- 확장자 사기 --------------------------------------------------
    ("salesforce_20260219.xls",       HTML_TABLE),            # placed + EXT_MISMATCH(html)
    ("oracle_20260219.csv",           xlsx_bytes("oracle")),  # placed + EXT_MISMATCH(zip)
    ("adobe_20260220.xls",            xlsx_bytes("adobe")),   # placed + EXT_MISMATCH(실은 xlsx)

    # --- 0바이트 -------------------------------------------------------
    ("empty_20260219.csv",            b""),                   # unresolved EMPTY_FILE

    # --- 한글 파일명 (NFD/NFC 확인) -------------------------------------
    ("삼성전자_20260220.xlsx",         xlsx_bytes("005930")),  # placed 2026/02/20

    # --- CP949 인코딩 CSV (텍스트 판정 확인) -----------------------------
    ("네이버_20260221.csv",            "회사,값\n네이버,100\n".encode("cp949")),  # placed

    # --- 하위 폴더 ------------------------------------------------------
    ("archive/old_20260101.csv",      CSV_A),                 # placed 2026/01/01

    # --- 중복: 내용이 완전히 같은 동명 파일 ------------------------------
    ("archive/anthropic_20260213.csv", CSV_A),                # duplicate -> _duplicates/

    # --- 이름 충돌: 동명이지만 내용이 다름 --------------------------------
    ("archive/nvidia_2026-02-16.csv", CSV_B),                 # placed + NAME_COLLISION -> __2

    # --- OS 부산물 -------------------------------------------------------
    # 주의: macOS는 형식이 맞지 않는 .DS_Store 를 몇 초 안에 스스로 삭제한다.
    # 검증 시 사라져 있어도 정상이며, 같은 스킵 규칙은 아래 세 개로도 확인된다.
    (".DS_Store",                     b"\x00\x01bud1"),       # skipped SYSTEM_FILE
    ("._apple_20260212.xls",          b"\x00\x05\x16\x07"),   # skipped SYSTEM_FILE
    ("~$quarterly_20260222.xlsx",     b"\x00" * 16),          # skipped SYSTEM_FILE
    ("archive/Thumbs.db",             b"\x00" * 8),           # skipped SYSTEM_FILE
]

# 심볼릭 링크 (별도 처리)
SYMLINKS = [
    ("shortcut_20260212.xls", "apple_20260212.xls"),   # skipped SYMLINK
    # 회귀 방지: 깨진 심볼릭 링크가 대소문자 구분 프로브를 오판시켰던 버그.
    # exists() 는 링크를 따라가 False 를 돌려주므로 macOS 를 "구분함"으로 잘못
    # 판정했고, 대소문자만 다른 두 파일이 같은 목적지로 가서 하나가 사라졌다.
    ("Zdangling_20260101.xls", "./no_such_target"),    # skipped SYMLINK
]


def build(root: Path):
    root.mkdir(parents=True)
    for rel, content in FIXTURES:
        p = root / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        if isinstance(content, str):
            p.write_text(content, encoding="utf-8")
        else:
            p.write_bytes(content)

    for link_name, target in SYMLINKS:
        try:
            (root / link_name).symlink_to(target)   # skipped SYMLINK
        except OSError as e:
            print(f"  심볼릭 링크 생성 생략: {link_name} ({e})", file=sys.stderr)

    # 권한 없는 파일 (root로 실행하면 읽히므로 그 경우는 건너뛴다)
    locked = root / "locked_20260223.csv"
    locked.write_text(CSV_A, encoding="utf-8")
    locked.chmod(0o000)                              # unresolved UNREADABLE

    n = sum(1 for _ in root.rglob("*") if _.is_file() or _.is_symlink())
    print(f"{root}/ 생성 완료 — 파일 {n}개")
    print("다음: python3 sort_requests.py            (dry-run)")
    print("      python3 sort_requests.py --apply    (실제 이동)")


def main() -> int:
    ap = argparse.ArgumentParser(description="검증용 request/ 픽스처를 만든다.")
    ap.add_argument("root", nargs="?", default="request")
    ap.add_argument("--force", action="store_true", help="기존 폴더를 지우고 다시 만든다")
    args = ap.parse_args()

    root = Path(args.root)
    if root.exists():
        if not args.force:
            print(f"{root}/ 가 이미 있습니다. 지우고 다시 만들려면 --force 를 쓰세요.",
                  file=sys.stderr)
            return 2
        for p in root.rglob("*"):       # chmod 000 파일 때문에 rmtree 전에 권한 복구
            try:
                p.chmod(0o644 if p.is_file() else 0o755)
            except OSError:
                pass
        shutil.rmtree(root)

    build(root)
    return 0


if __name__ == "__main__":
    sys.exit(main())
