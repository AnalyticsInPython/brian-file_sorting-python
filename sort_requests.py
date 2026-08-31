#!/usr/bin/env python3
"""
sort_requests.py — request/ 안의 파일을 파일명에 들어 있는 날짜 기준으로
request_by_date/YYYY/MM/DD/ 트리에 재배치한다.

설계 원칙
---------
확실하게 판정되는 파일만 자동 배치하고, 애매한 것은 추측하지 않고
격리 폴더(_unresolved/)로 보낸 뒤 리포트한다. 잘못 정리된 파일은
잘못 정리됐다는 사실조차 드러나지 않기 때문에, "틀리게 정리"보다
"정리하지 않고 보고"가 항상 낫다.

기본 동작은 dry-run이다. --apply 없이는 디스크를 건드리지 않는다.
표준 라이브러리만 사용한다 (외부 패키지 불필요).

사용법
------
  python3 sort_requests.py                              # dry-run (기본)
  python3 sort_requests.py --apply                      # 실제 이동
  python3 sort_requests.py --apply --copy               # 원본 보존 복사
  python3 sort_requests.py --undo sort_manifest.jsonl   # 되돌리기
"""

import argparse
import codecs
import csv
import datetime as dt
import difflib
import hashlib
import json
import os
import re
import shutil
import signal
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

# ---------------------------------------------------------------- 설정 상수

DEFAULT_SRC = "request"
DEFAULT_DEST = "request_by_date"
DEFAULT_REPORT = "sort_report.csv"
DEFAULT_MANIFEST = "sort_manifest.jsonl"

UNRESOLVED_DIR = "_unresolved"
DUPLICATES_DIR = "_duplicates"

# 정리 대상이 아닌 OS/오피스 부산물
SKIP_EXACT = {".DS_Store", "Thumbs.db", "desktop.ini", ".localized", "Icon\r"}
SKIP_PREFIXES = ("._", "~$", ".~lock.")
SKIP_DIRS = {".Spotlight-V100", ".Trashes", ".fseventsd", ".TemporaryItems",
             "__MACOSX", ".git", ".svn"}

# 날짜로 인정하는 범위. 벗어나면 우연히 8자리가 맞은 숫자로 본다.
MIN_YEAR, MAX_YEAR = 1990, 2100

# YYYY[구분자]MM[구분자]DD 만 인정한다.
# 앞뒤 (?<!\d)(?!\d) 로 더 긴 숫자열의 일부가 잘려 매치되는 것을 막는다.
DATE_RE = re.compile(r"(?<!\d)(\d{4})[-_.\s]?(\d{2})[-_.\s]?(\d{2})(?!\d)")

# 유사 파일명 클러스터링 임계값 (difflib ratio)
SIMILARITY_THRESHOLD = 0.85

TEXTY = {"text/utf-8", "text/cp949", "text/other"}
EXT_EXPECT = {
    ".xls": {"ole2"},
    ".xlsx": {"zip"}, ".xlsm": {"zip"}, ".xlsb": {"zip"},
    ".doc": {"ole2"}, ".ppt": {"ole2"},
    ".docx": {"zip"}, ".pptx": {"zip"}, ".zip": {"zip"},
    ".pdf": {"pdf"},
    ".csv": TEXTY, ".tsv": TEXTY, ".txt": TEXTY, ".json": TEXTY,
    ".xml": TEXTY | {"html/xml"},
    ".html": {"html/xml"} | TEXTY, ".htm": {"html/xml"} | TEXTY,
}
# 확장자 불일치가 났을 때 사람에게 줄 힌트
FORMAT_HINT = {
    "zip": "실제로는 .xlsx/.docx 계열(ZIP 컨테이너)",
    "ole2": "실제로는 레거시 .xls/.doc 계열(OLE2)",
    "html/xml": "실제로는 HTML/XML — 엑셀 '웹페이지로 저장' 산출물일 가능성",
    "pdf": "실제로는 PDF",
    "binary": "알 수 없는 바이너리",
}

_interrupted = False

# ---------------------------------------------------------------- 데이터 모델


@dataclass
class Item:
    """파일 하나에 대한 수집 -> 판정 -> 배치 결과."""
    src: Path
    rel: str
    size: int = 0
    date: "dt.date | None" = None
    ext: str = ""
    detected: str = ""
    dest: "Path | None" = None
    status: str = ""          # placed | duplicate | unresolved | skipped | error
    reason: str = ""          # 대표 사유 코드
    flags: list = field(default_factory=list)   # 부가 경고 (배치는 진행)

    @property
    def name(self) -> str:
        return self.src.name


# ---------------------------------------------------------------- 유틸리티


def nfc(s: str) -> str:
    """macOS 파일시스템은 이름을 NFD로 돌려준다. 비교는 항상 NFC로."""
    return unicodedata.normalize("NFC", s)


def nearest_existing_dir(path: Path) -> Path:
    """아직 만들어지지 않은 경로에서 가장 가까운 상위 실존 디렉토리를 찾는다."""
    p = path if path.is_absolute() else Path.cwd() / path
    for cand in (p, *p.parents):
        if cand.is_dir():
            return cand
    return Path.cwd()


def fs_case_insensitive(probe_dir: Path, max_probes: int = 50) -> bool:
    """
    probe_dir이 놓인 파일시스템이 대소문자를 구분하는지 읽기 전용으로 판별한다.

    macOS APFS 기본값은 구분하지 않으므로, Python이 서로 다르게 보는 두 이름이
    OS 수준에서는 같은 파일로 덮어써질 수 있다. 구분하지 않는 볼륨을 구분한다고
    잘못 판단하면 파일이 경고 없이 사라지는 반면, 반대로 틀리면 __2 접미사가
    불필요하게 붙을 뿐이다. 그래서 확신이 없으면 항상 True(구분 안 함)로 답한다.

    exists()가 아니라 lexists()를 쓴다. exists()는 심볼릭 링크를 따라가므로
    깨진 링크가 프로브 대상이 되면 대소문자 무시 볼륨을 구분한다고 오판한다.
    한 항목만 보지 않고 여러 개를 확인해 하나라도 일치하면 '구분 안 함'으로 본다.
    """
    probed = 0
    try:
        for entry in probe_dir.iterdir():
            alt = entry.name.swapcase()
            if alt == entry.name:
                continue          # 대소문자가 없는 이름(숫자/한글 등)은 판별에 못 쓴다
            if os.path.lexists(probe_dir / alt):
                return True
            probed += 1
            if probed >= max_probes:
                break
    except OSError:
        return True
    return False if probed else True


def make_key_fn(case_insensitive: bool):
    """경로를 충돌 판정용 키로 바꾸는 함수를 만든다."""
    if case_insensitive:
        return lambda p: nfc(str(p)).casefold()
    return lambda p: nfc(str(p))


def file_hash(path: Path, chunk: int = 1 << 20) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            block = f.read(chunk)
            if not block:
                break
            h.update(block)
    return h.hexdigest()


def split_ext(name: str):
    stem, ext = os.path.splitext(name)
    return stem, ext


def human(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f}{unit}" if unit == "B" else f"{n:.1f}{unit}"
        n /= 1024.0
    return f"{n}B"


# ---------------------------------------------------------------- 1. 수집


def collect_files(src: Path, dest: Path):
    """src 하위를 재귀 탐색해 Item 목록을 만든다. 부산물/심볼릭 링크는 skipped."""
    items = []
    try:
        dest_res = dest.resolve()
    except OSError:
        dest_res = dest

    for dirpath, dirnames, filenames in os.walk(src):
        d = Path(dirpath)

        keep = []
        for name in sorted(dirnames):
            sub = d / name
            if name in SKIP_DIRS:
                continue
            if sub.is_symlink():
                items.append(Item(src=sub, rel=str(sub.relative_to(src)),
                                  status="skipped", reason="SYMLINK_DIR"))
                continue
            try:
                # dest가 src 안에 있으면 자기 산출물을 다시 먹지 않도록 잘라낸다
                if sub.resolve() == dest_res:
                    continue
            except OSError:
                pass
            keep.append(name)
        dirnames[:] = keep

        for name in sorted(filenames):
            p = d / name
            rel = str(p.relative_to(src))
            if name in SKIP_EXACT or name.startswith(SKIP_PREFIXES):
                items.append(Item(src=p, rel=rel, status="skipped",
                                  reason="SYSTEM_FILE"))
                continue
            if p.is_symlink():
                items.append(Item(src=p, rel=rel, status="skipped",
                                  reason="SYMLINK"))
                continue
            try:
                size = p.stat().st_size
            except OSError as e:
                items.append(Item(src=p, rel=rel, status="unresolved",
                                  reason=f"UNREADABLE: {e.strerror}"))
                continue
            items.append(Item(src=p, rel=rel, size=size,
                              ext=split_ext(name)[1].lower()))
    return items


# ---------------------------------------------------------------- 2. 날짜 파싱


def parse_date(stem: str):
    """
    파일명 stem에서 날짜를 뽑는다.
    반환: (date | None, reason). date가 None이면 격리 대상이다.
    """
    matches = list(DATE_RE.finditer(stem))
    if not matches:
        return None, "NO_DATE"

    good, bad = [], []
    for m in matches:
        y, mo, dd = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if not (MIN_YEAR <= y <= MAX_YEAR):
            bad.append("YEAR_OUT_OF_RANGE")
            continue
        try:
            good.append(dt.date(y, mo, dd))
        except ValueError:
            bad.append("INVALID_DATE")

    uniq = sorted(set(good))
    if not uniq:
        return None, bad[0] if bad else "NO_DATE"
    if len(uniq) > 1:
        return None, "AMBIGUOUS_MULTI_DATE"
    if bad:
        # 유효한 날짜 하나 + 깨진 숫자열이 함께 있는 경우: 배치하되 표시한다
        return uniq[0], "PARTIAL_INVALID_TOKEN"
    return uniq[0], ""


# ---------------------------------------------------------------- 3. 포맷 검증


def detect_format(path: Path):
    """앞 512바이트의 매직바이트로 실제 포맷을 판별한다."""
    try:
        with open(path, "rb") as f:
            head = f.read(512)
    except OSError as e:
        return "unreadable", e.strerror or str(e)

    if not head:
        return "empty", ""
    if head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return "ole2", ""
    if head.startswith(b"PK\x03\x04"):
        return "zip", ""
    if head.startswith(b"%PDF-"):
        return "pdf", ""

    probe = head.lstrip(codecs.BOM_UTF8).lstrip()
    low = probe[:64].lower()
    for sig in (b"<?xml", b"<html", b"<!doctype html", b"<table", b"<workbook"):
        if low.startswith(sig):
            return "html/xml", ""

    if b"\x00" in head:
        return "binary", ""
    # 512바이트 경계에서 멀티바이트 문자가 잘릴 수 있으므로 incremental decoder 사용
    for enc, label in (("utf-8-sig", "text/utf-8"), ("cp949", "text/cp949")):
        try:
            codecs.getincrementaldecoder(enc)().decode(head)
            return label, ""
        except UnicodeDecodeError:
            continue
    return "text/other", ""


def check_format(item: Item):
    """감지한 포맷과 확장자를 대조한다. 불일치는 flag만 남기고 배치는 진행."""
    detected, err = detect_format(item.src)
    item.detected = detected
    if detected == "unreadable":
        item.status = "unresolved"
        item.reason = f"UNREADABLE: {err}"
        return
    if detected == "empty":
        item.status = "unresolved"
        item.reason = "EMPTY_FILE"
        return
    expected = EXT_EXPECT.get(item.ext)
    if expected and detected not in expected:
        hint = FORMAT_HINT.get(detected, detected)
        item.flags.append(f"EXT_MISMATCH({item.ext} -> {hint})")


# ---------------------------------------------------------------- 4. 배치 계획


def resolve_dest(target: Path, item: Item, taken: dict, key_fn):
    """
    target 자리를 item에게 할당하고 (경로 | None, 사유)를 돌려준다.

      (path, "")                 깨끗하게 배치
      (path, "NAME_COLLISION")   같은 이름, 다른 내용 -> __2 접미사로 둘 다 보존
      (None, "ALREADY_PLACED")   디스크에 이미 같은 내용의 파일이 있음 (할 일 없음)
      (None, "DUPLICATE")        이번 실행의 다른 원본이 같은 내용으로 그 자리를 차지

    ALREADY_PLACED 를 DUPLICATE 와 구분하는 것이 재실행 멱등성의 핵심이다.
    구분하지 않으면 --copy 를 두 번 돌릴 때마다 _duplicates/ 가 계속 불어난다.
    """
    stem, ext = split_ext(target.name)
    cand, n = target, 1
    while True:
        k = key_fn(cand)
        claimed_by = taken.get(k)
        if claimed_by is None and not cand.exists():
            taken[k] = item.src
            return cand, ("NAME_COLLISION" if n > 1 else "")
        occupant = claimed_by if claimed_by is not None else cand
        try:
            same = file_hash(item.src) == file_hash(occupant)
        except OSError:
            same = False   # 해시할 수 없으면 다른 파일로 보고 접미사를 붙인다
        if same:
            return None, ("DUPLICATE" if claimed_by is not None else "ALREADY_PLACED")
        n += 1
        cand = target.with_name(f"{stem}__{n}{ext}")


def _note(item: Item, note: str):
    """이미 사유가 있으면 덮어쓰지 않고 flags 로 보낸다."""
    if not note:
        return
    if item.reason:
        item.flags.append(note)
    else:
        item.reason = note


def plan_moves(items, dest: Path, key_fn):
    """각 Item의 최종 목적지와 status를 확정한다."""
    taken: dict = {}
    for item in sorted(items, key=lambda i: nfc(i.rel)):
        if item.status:            # skipped / 수집 단계에서 이미 확정된 것
            continue

        date, reason = parse_date(split_ext(item.name)[0])
        if date is None:
            item.status = "unresolved"
            item.reason = reason
        else:
            item.date = date
            if reason:
                item.flags.append(reason)
            check_format(item)     # EMPTY_FILE / UNREADABLE 이면 여기서 unresolved

        quarantined = item.status == "unresolved"
        if quarantined:
            target = dest / UNRESOLVED_DIR / item.name
        else:
            target = dest / f"{date:%Y}" / f"{date:%m}" / f"{date:%d}" / item.name

        resolved, note = resolve_dest(target, item, taken, key_fn)

        if resolved is not None:
            item.dest = resolved
            if not quarantined:
                item.status = "placed"
            _note(item, note)
            continue

        if note == "ALREADY_PLACED":
            item.dest = None
            item.status = "already_placed"
            _note(item, "ALREADY_PLACED")
            continue

        # note == "DUPLICATE"
        if quarantined:
            item.dest = None
            item.status = "duplicate"
            _note(item, "DUPLICATE")
            continue

        # 날짜 폴더에는 하나만 두고 나머지는 _duplicates/ 에 보관한다 (삭제하지 않음)
        dup_target = dest / DUPLICATES_DIR / f"{date:%Y-%m-%d}" / item.name
        resolved, dup_note = resolve_dest(dup_target, item, taken, key_fn)
        item.dest = resolved
        if resolved is None:
            item.status = "already_placed"
            _note(item, "ALREADY_IN_DUPLICATES")
        else:
            item.status = "duplicate"
            _note(item, "DUPLICATE")
    return items


# ---------------------------------------------------------------- 5. 실행


def _on_sigint(signum, frame):
    global _interrupted
    _interrupted = True
    print("\n[중단 요청] 현재 파일 처리 후 안전하게 종료합니다...", file=sys.stderr)


def execute(items, dest: Path, copy: bool, manifest_path: Path):
    """실제 이동/복사. 각 파일 직전에 manifest를 flush+fsync 한다."""
    signal.signal(signal.SIGINT, _on_sigint)
    op = "copy" if copy else "move"
    done = 0

    dest.mkdir(parents=True, exist_ok=True)
    with open(manifest_path, "a", encoding="utf-8") as mf:
        mf.write(json.dumps({"op": "session", "mode": op,
                             "at": dt.datetime.now().isoformat(timespec="seconds"),
                             "dest": str(dest)}, ensure_ascii=False) + "\n")
        mf.flush()
        for item in items:
            if _interrupted:
                break
            if item.dest is None or item.status == "skipped":
                continue
            try:
                item.dest.parent.mkdir(parents=True, exist_ok=True)
                mf.write(json.dumps({"op": op, "src": str(item.src),
                                     "dest": str(item.dest)},
                                    ensure_ascii=False) + "\n")
                mf.flush()
                os.fsync(mf.fileno())
                if copy:
                    shutil.copy2(item.src, item.dest)
                else:
                    shutil.move(str(item.src), str(item.dest))
                done += 1
            except OSError as e:
                item.status = "error"
                item.reason = f"IO_ERROR: {e}"

    signal.signal(signal.SIGINT, signal.SIG_DFL)
    return done


def prune_empty_dirs(root: Path):
    """비게 된 하위 디렉토리를 정리한다 (root 자신은 남긴다)."""
    removed = 0
    for dirpath, dirnames, filenames in os.walk(root, topdown=False):
        d = Path(dirpath)
        if d == root:
            continue
        try:
            if not any(d.iterdir()):
                d.rmdir()
                removed += 1
        except OSError:
            pass
    return removed


# ---------------------------------------------------------------- 6. 리포트


NAME_TOKEN_RE = re.compile(r"[\W_]+", re.UNICODE)


def name_token(stem: str) -> str:
    """파일명에서 날짜/숫자/구분자를 걷어내고 '회사명'만 남긴다."""
    s = DATE_RE.sub(" ", stem)
    s = re.sub(r"\d+", " ", s)
    s = NAME_TOKEN_RE.sub(" ", s)
    return " ".join(s.split())


def find_similar_names(items):
    """오타/대소문자 차이로 갈라진 이름들을 묶어 보고한다. 이름은 바꾸지 않는다."""
    groups: dict = {}   # casefold 토큰 -> {원본 철자: [Item, ...]}
    for item in items:
        if item.status == "skipped":
            continue
        tok = name_token(split_ext(item.name)[0])
        if not tok:
            continue
        groups.setdefault(nfc(tok).casefold(), {}).setdefault(tok, []).append(item)

    keys = sorted(groups)
    parent = {k: k for k in keys}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for i, a in enumerate(keys):
        for b in keys[i + 1:]:
            if difflib.SequenceMatcher(None, a, b).ratio() >= SIMILARITY_THRESHOLD:
                ra, rb = find(a), find(b)
                if ra != rb:
                    parent[rb] = ra

    clusters: dict = {}
    for k in keys:
        clusters.setdefault(find(k), []).append(k)

    out = []
    for members in clusters.values():
        spellings: dict = {}
        for k in members:
            for spelling, its in groups[k].items():
                spellings.setdefault(spelling, []).extend(its)
        if len(spellings) > 1:
            out.append(sorted(spellings.items(), key=lambda kv: (-len(kv[1]), kv[0])))
    return sorted(out, key=lambda c: -sum(len(v) for _, v in c))


def write_report(items, path: Path, src: Path):
    """엑셀에서 바로 열리도록 utf-8-sig(BOM)로 저장한다."""
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["status", "reason", "flags", "parsed_date",
                    "source_path", "dest_path", "ext", "detected_format", "size_bytes"])
        for item in sorted(items, key=lambda i: (i.status, nfc(i.rel))):
            w.writerow([
                item.status,
                item.reason,
                "; ".join(item.flags),
                item.date.isoformat() if item.date else "",
                item.rel,
                str(item.dest) if item.dest else "",
                item.ext,
                item.detected,
                item.size,
            ])


def print_plan(items, dry_run: bool, limit: int = 12):
    by_status: dict = {}
    for item in items:
        by_status.setdefault(item.status, []).append(item)

    order = ["placed", "duplicate", "already_placed", "unresolved", "skipped", "error"]
    label = {"placed": "배치", "duplicate": "중복", "already_placed": "이미 정리됨",
             "unresolved": "격리", "skipped": "스킵", "error": "오류"}

    for status in order:
        group = by_status.get(status)
        if not group:
            continue
        print(f"\n[{label[status]}] {len(group)}건")
        for item in group[:limit]:
            note = f"  ({item.reason})" if item.reason else ""
            if status == "skipped" or item.dest is None:
                print(f"  - {item.rel}{note}")
                continue
            arrow = "->" if dry_run else "=>"
            print(f"  - {item.rel} {arrow} {item.dest}{note}")
        if len(group) > limit:
            print(f"  ... 외 {len(group) - limit}건 (리포트 CSV 참고)")

    mismatches = [i for i in items if any(f.startswith("EXT_MISMATCH") for f in i.flags)]
    if mismatches:
        print(f"\n[확장자 불일치] {len(mismatches)}건 — 배치는 정상 진행됨")
        for item in mismatches[:limit]:
            flag = next(f for f in item.flags if f.startswith("EXT_MISMATCH"))
            print(f"  - {item.rel}  {flag}")
        if len(mismatches) > limit:
            print(f"  ... 외 {len(mismatches) - limit}건")

    clusters = find_similar_names(items)
    if clusters:
        print(f"\n[유사 파일명] {len(clusters)}개 클러스터 — 같은 대상의 오타/표기 차이일 수 있습니다")
        print("  (파일명은 변경하지 않습니다. 확인 후 원본을 고치고 재실행하세요.)")
        for cluster in clusters:
            print("  ---")
            for spelling, its in cluster:
                dates = ", ".join(sorted({i.date.isoformat() for i in its if i.date}))
                print(f"    {spelling:<24} {len(its)}개  {dates}")

    counts = {s: len(by_status.get(s, [])) for s in order}
    print("\n" + "=" * 68)
    print(f"스캔 {len(items)}개 | 배치 {counts['placed']} | 격리 {counts['unresolved']} "
          f"| 중복 {counts['duplicate']} | 이미 정리됨 {counts['already_placed']} "
          f"| 스킵 {counts['skipped']} | 오류 {counts['error']}")
    collisions = sum(1 for i in items if i.reason == "NAME_COLLISION")
    if collisions:
        print(f"이름 충돌(내용 다름, 접미사 부여) {collisions}건")
    print("=" * 68)


# ---------------------------------------------------------------- 7. Undo


def undo(manifest_path: Path):
    if not manifest_path.exists():
        print(f"manifest를 찾을 수 없습니다: {manifest_path}", file=sys.stderr)
        return 2

    records = []
    roots = set()
    with open(manifest_path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
            except json.JSONDecodeError:
                continue
            if rec.get("op") == "session":
                # 빈 디렉토리 정리는 이 세션이 실제로 만든 출력 폴더로만 한정한다
                roots.add(Path(rec["dest"]))
            elif rec.get("op") in ("move", "copy"):
                records.append(rec)

    reverted = skipped = 0
    for rec in reversed(records):
        src, dest = Path(rec["src"]), Path(rec["dest"])
        try:
            if not dest.exists():
                skipped += 1
                continue
            if rec["op"] == "copy":
                # 복사본만 지운다. 원본이 없거나 내용이 다르면 손대지 않는다.
                if src.exists() and file_hash(src) == file_hash(dest):
                    dest.unlink()
                    reverted += 1
                else:
                    skipped += 1
                continue
            if src.exists():
                skipped += 1
                continue
            src.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(dest), str(src))
            reverted += 1
        except OSError as e:
            print(f"  되돌리기 실패: {dest} ({e})", file=sys.stderr)
            skipped += 1

    for root in roots:
        if root.exists():
            prune_empty_dirs(root)

    print(f"되돌림 {reverted}건 | 건너뜀 {skipped}건")
    if skipped:
        print("건너뛴 항목은 목적지가 이미 없거나 원본 자리에 다른 파일이 있는 경우입니다.")
    return 0


# ---------------------------------------------------------------- main


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="request/ 의 파일을 파일명 날짜 기준으로 정리한다.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__.split("사용법")[-1],
    )
    ap.add_argument("src", nargs="?", default=DEFAULT_SRC, help=f"원본 폴더 (기본: {DEFAULT_SRC})")
    ap.add_argument("-o", "--dest", default=DEFAULT_DEST, help=f"출력 폴더 (기본: {DEFAULT_DEST})")
    ap.add_argument("--apply", action="store_true", help="실제로 파일을 옮긴다 (없으면 dry-run)")
    ap.add_argument("--copy", action="store_true", help="이동 대신 복사한다 (원본 보존)")
    ap.add_argument("--report", default=DEFAULT_REPORT, help=f"리포트 CSV 경로 (기본: {DEFAULT_REPORT})")
    ap.add_argument("--manifest", default=DEFAULT_MANIFEST, help=f"manifest 경로 (기본: {DEFAULT_MANIFEST})")
    ap.add_argument("--undo", metavar="MANIFEST", help="manifest를 역순으로 되돌린다")
    args = ap.parse_args(argv)

    if args.undo:
        return undo(Path(args.undo))

    src, dest = Path(args.src), Path(args.dest)
    if not src.is_dir():
        print(f"원본 폴더가 없습니다: {src}", file=sys.stderr)
        return 2

    # key_fn은 목적지 경로 충돌 판정에만 쓰인다. src와 dest가 다른 볼륨일 수 있으므로
    # 반드시 dest 쪽(아직 없으면 가장 가까운 상위 폴더)을 프로브해야 한다.
    probe_dir = nearest_existing_dir(dest)
    ci = fs_case_insensitive(probe_dir)
    key_fn = make_key_fn(ci)

    items = collect_files(src, dest)
    if not items:
        print(f"{src} 안에 처리할 파일이 없습니다.")
        return 0

    plan_moves(items, dest, key_fn)

    mode = "이동" if not args.copy else "복사"
    print(f"원본: {src}  ->  출력: {dest}")
    print(f"모드: {'DRY-RUN (아무것도 변경하지 않음)' if not args.apply else mode}"
          f"   |   대소문자 구분: {'없음' if ci else '있음'} ({probe_dir} 기준)")

    print_plan(items, dry_run=not args.apply)

    if args.apply:
        done = execute(items, dest, args.copy, Path(args.manifest))
        if not args.copy:
            prune_empty_dirs(src)
        print(f"\n{mode} 완료: {done}건")
        failures = [i for i in items if i.status == "error"]
        if failures:
            print(f"실패 {len(failures)}건 — 원본은 그대로 남아 있습니다:")
            for item in failures:
                print(f"  - {item.rel}  ({item.reason})")
        print(f"되돌리기: python3 {Path(sys.argv[0]).name} --undo {args.manifest}")
        if _interrupted:
            print("중단되었습니다. manifest에는 완료된 항목만 기록되어 있어 그대로 되돌릴 수 있습니다.")
    else:
        print("\n실제로 옮기려면 --apply 를 붙여 다시 실행하세요.")

    write_report(items, Path(args.report), src)
    print(f"리포트: {args.report}")

    return 1 if any(i.status == "error" for i in items) else 0


if __name__ == "__main__":
    sys.exit(main())
