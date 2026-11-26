"""
한컴 공식 문서(ActionTable, ParameterSetTable PDF)에서 액션·매개변수 색인 JSON을 만듭니다.
hwp_action_help / hwp_run_action 도구가 이 JSON을 읽습니다.

사용:
    uv run --with pypdf python scripts/build_action_index.py <PDF 폴더> [출력 경로]
    예) uv run --with pypdf python scripts/build_action_index.py 참고자료/hancom-api
PDF 폴더에 ActionTable_*.pdf, ParameterSetTable_*.pdf가 있어야 합니다.
출력 기본값: src/tools/data/hwp_actions.json

PDF 표의 한 줄 = 한 항목. 설명이 여러 줄이면 다음 줄들이 이어 붙습니다(영문 식별자로 시작하지 않는 줄).
"""
import glob
import json
import os
import re
import sys

from pypdf import PdfReader

IDENT = r"[A-Za-z][A-Za-z0-9_]*"


def pdf_lines(path):
    for page in PdfReader(path).pages:
        for line in (page.extract_text() or "").splitlines():
            line = line.strip()
            if line:
                yield line


def _join(desc, more):
    more = more.strip()
    if not more:
        return desc
    return f"{desc} {more}".strip() if desc else more


# ---------------------------------------------------------------- 액션 표

ACTION_HEADER = "Action ID ParameterSet ID Description"
ACTION_START = re.compile(rf"^({IDENT})(?: (.*))?$")
PSET_TOKEN = re.compile(rf"^(?:-|\+|{IDENT}\*?)$")


def parse_actions(path):
    actions = {}
    cur = None
    started = False
    for line in pdf_lines(path):
        if line.startswith(ACTION_HEADER):
            started = True
            continue
        if not started or re.fullmatch(r"\d+", line) or line.startswith("최종 수정일"):
            continue
        m = ACTION_START.match(line)
        if m:
            name, rest = m.group(1), (m.group(2) or "")
            tokens = rest.split(" ", 1)
            pset, desc = None, rest
            if tokens and PSET_TOKEN.match(tokens[0]):
                pset, desc = tokens[0], (tokens[1] if len(tokens) > 1 else "")
            entry = {"pset": None, "desc": desc.strip()}
            if pset == "+":
                entry["note"] = "매개변수를 내부에서만 씀(외부에 노출 안 됨)"
            elif pset and pset != "-":
                if pset.endswith("*"):
                    pset = pset[:-1]
                    entry["needs_pset"] = True  # 매개변수를 직접 만들어 줘야 동작(Run만으로는 안 됨)
                entry["pset"] = pset
            cur = actions.setdefault(name, entry)
            continue
        if cur is not None:
            cur["desc"] = _join(cur["desc"], line)
    return actions


# ---------------------------------------------------------------- 매개변수 표

PSET_SECTION = re.compile(rf"^\d+\) ({IDENT}) ?: ?(.*)$")
ITEM_HEADER = re.compile(r"^Item (?:ID|Name) Type")
ITEM_START = re.compile(rf"^({IDENT}) (PIT_[A-Z0-9_]+)(?: (.*))?$")


def parse_psets(path):
    psets = {}
    cur_set = cur_item = None
    for line in pdf_lines(path):
        if re.fullmatch(r"\d+", line) or line.startswith("최종 수정일"):
            continue
        m = PSET_SECTION.match(line)
        if m:
            cur_set = psets.setdefault(m.group(1), {"desc": m.group(2).strip(), "items": []})
            cur_item = None
            continue
        if cur_set is None or ITEM_HEADER.match(line):
            continue
        m = ITEM_START.match(line)
        if m:
            item_id, ptype, rest = m.group(1), m.group(2), (m.group(3) or "")
            sub = None
            if ptype in ("PIT_SET", "PIT_ARRAY"):  # 하위 형식(묶음 이름 등)이 따라옴
                parts = rest.split(" ", 1)
                if parts and re.fullmatch(rf"{IDENT}|PIT_[A-Z0-9_]+", parts[0]):
                    sub, rest = parts[0], (parts[1] if len(parts) > 1 else "")
            cur_item = {"id": item_id, "type": ptype.replace("PIT_", "")}
            if sub:
                cur_item["sub"] = sub
            cur_item["desc"] = rest.strip()
            cur_set["items"].append(cur_item)
            continue
        if cur_item is not None:
            cur_item["desc"] = _join(cur_item["desc"], line)
        else:
            cur_set["desc"] = _join(cur_set["desc"], line)
    return psets


def parse_action_rows(path):
    """액션 표를 (이름, 매개변수 토큰, 설명) 순서 목록으로. 보정(verify) 전 단계."""
    rows = []
    started = False
    for line in pdf_lines(path):
        if line.startswith(ACTION_HEADER):
            started = True
            continue
        if not started or re.fullmatch(r"\d+", line) or line.startswith("최종 수정일"):
            continue
        m = ACTION_START.match(line)
        if m:
            name, rest = m.group(1), (m.group(2) or "")
            tokens = rest.split(" ", 1)
            pset, desc = None, rest
            if tokens and PSET_TOKEN.match(tokens[0]):
                pset, desc = tokens[0], (tokens[1] if len(tokens) > 1 else "")
            rows.append([name, pset, desc.strip()])
        elif rows:
            rows[-1][2] = _join(rows[-1][2], line)
    return rows


def _entry(pset, desc):
    entry = {"pset": None, "desc": desc}
    if pset == "+":
        entry["note"] = "매개변수를 내부에서만 씀(외부에 노출 안 됨)"
    elif pset and pset != "-":
        if pset.endswith("*"):
            pset = pset[:-1]
            entry["needs_pset"] = True  # 매개변수를 직접 만들어 줘야 동작(Run만으로는 안 됨)
        entry["pset"] = pset
    return entry


def verify_actions(rows, h):
    """
    실행 중인 한글(h)로 액션 이름을 검증·보정합니다.
    - PDF에서 두 줄로 잘린 긴 이름(예: ParagraphShapeAlignDistrib + ute)은 이어 붙여 보고 있으면 합침
    - 소문자로 시작하는 조각(예제 코드의 if/while 등)은 버림
    - 매개변수 묶음 이름은 한글이 알려 주는 실제 이름(SetID)으로 바꿈(원문 오타 보정)
    - 이 한글에 없는 이름은 verified=False로 남김(다른 버전에는 있을 수 있음)
    """
    def exists(name):
        try:
            return h.CreateAction(name) is not None
        except Exception:
            return False

    actions, stats = {}, {"merged": 0, "dropped": 0, "unverified": 0, "pset_fixed": 0}
    i = 0
    while i < len(rows):
        name, pset, desc = rows[i]
        ok = exists(name)
        if not ok and i + 1 < len(rows) and exists(name + rows[i + 1][0]):
            nxt = rows[i + 1]
            name, ok = name + nxt[0], True
            pset = pset or nxt[1]
            desc = _join(desc, nxt[2])
            stats["merged"] += 1
            i += 1
        i += 1
        if not ok and name[0].islower():
            stats["dropped"] += 1
            continue
        entry = _entry(pset, desc)
        if ok:
            try:
                s = h.CreateAction(name).CreateSet()
                sid = s.SetID if s is not None else ""
            except Exception:
                sid = ""
            if sid and sid != entry["pset"]:
                entry["pset"] = sid
                stats["pset_fixed"] += 1
        else:
            entry["verified"] = False
            stats["unverified"] += 1
        actions.setdefault(name, entry)
    return actions, stats


def _connect_hwp():
    """이미 켜진 한글에 붙습니다(없으면 None). 저장소 루트에서 실행한다고 가정."""
    try:
        sys.path.insert(0, os.getcwd())
        from src.tools.hwp_controller import HwpController
        c = HwpController()
        return c.hwp if c.connect() else None
    except Exception as e:
        print(f"한글 연결 실패(검증 생략): {e}")
        return None


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    src = args[0] if args else "참고자료/hancom-api"
    out = args[1] if len(args) > 1 else os.path.join("src", "tools", "data", "hwp_actions.json")
    action_pdf = sorted(glob.glob(os.path.join(src, "ActionTable_*.pdf")))[-1]
    pset_pdf = sorted(glob.glob(os.path.join(src, "ParameterSetTable_*.pdf")))[-1]
    h = None if "--no-verify" in sys.argv else _connect_hwp()
    if h is not None:
        actions, stats = verify_actions(parse_action_rows(action_pdf), h)
        print(f"한글로 검증: 잘린 이름 합침 {stats['merged']}, 코드 조각 버림 {stats['dropped']}, "
              f"이 한글에 없음 {stats['unverified']}, 매개변수 이름 보정 {stats['pset_fixed']}")
    else:
        actions = parse_actions(action_pdf)
        print("한글 없이 만듦(이름 검증·보정 안 함)")
    data = {
        "source": {"actions": os.path.basename(action_pdf), "psets": os.path.basename(pset_pdf),
                   "verified_with_hwp": h is not None},
        "actions": actions,
        "psets": parse_psets(pset_pdf),
    }
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=0, sort_keys=True)
    n_items = sum(len(p["items"]) for p in data["psets"].values())
    print(f"액션 {len(data['actions'])}개, 매개변수 묶음 {len(data['psets'])}개(항목 {n_items}개) → {out}")


if __name__ == "__main__":
    main()
