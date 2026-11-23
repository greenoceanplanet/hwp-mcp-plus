"""
서식 조건 찾기·스타일·메모 COM 헬퍼.

- 서식 조건 찾기: 한글 찾기(RepeatFind)의 글자모양 조건(FindCharShape)을 씀.
  실측: HParameterSet.HFindReplace는 이전 호출의 FindCharShape 값이 남으므로,
  매번 CreateAction("RepeatFind").CreateSet()으로 새 세트를 만들어야 조건이 섞이지 않음.
  글꼴 조건은 한글 찾기로 확인되지 않아, 문서를 HWPML(XML)로 읽어(GetTextFile "HWPML2X") 판정함.
- 스타일 목록·메모 목록도 HWPML로 읽음(문서를 저장하거나 바꾸지 않음).
- 스타일 적용: Style 액션의 Apply(스타일 번호).
- 메모 넣기: InsertFieldMemo → 글자 입력 → CloseEx.
"""

import logging
import re
import xml.etree.ElementTree as ET

from src.tools.hwp_edit_helpers import _insert_lines

logger = logging.getLogger("hwp-mcp-extended")

NAMED_COLORS = {
    "black": "#000000", "검정": "#000000",
    "red": "#FF0000", "빨강": "#FF0000",
    "blue": "#0000FF", "파랑": "#0000FF",
    "green": "#008000", "초록": "#008000",
    "yellow": "#FFFF00", "노랑": "#FFFF00",
    "gray": "#808080", "grey": "#808080", "회색": "#808080",
    "white": "#FFFFFF", "흰색": "#FFFFFF",
}


class _AutoOk:
    def __init__(self, h):
        self.h = h

    def __enter__(self):
        try:
            self.h.SetMessageBoxMode(0x00010000)
        except Exception:
            pass

    def __exit__(self, *_):
        try:
            self.h.SetMessageBoxMode(0x00000000)
        except Exception:
            pass
        return False


# ---------------------------------------------------------------- 조건 변환

def color_to_colorref(color) -> int:
    """'#RRGGBB'·'red'·정수 → 한글 COLORREF(0x00BBGGRR)."""
    if isinstance(color, int):
        return color
    s = str(color).strip()
    s = NAMED_COLORS.get(s.lower(), s)
    m = re.fullmatch(r"#?([0-9a-fA-F]{6})", s)
    if not m:
        raise ValueError(f"색은 '#RRGGBB' 또는 이름(red, blue 등)이어야 합니다: {color}")
    v = m.group(1)
    r, g, b = int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)
    return (b << 16) | (g << 8) | r


def colorref_to_hex(v: int) -> str:
    v = int(v)
    return f"#{v & 0xFF:02X}{(v >> 8) & 0xFF:02X}{(v >> 16) & 0xFF:02X}"


def build_charshape_conditions(bold=None, italic=None, underline=None, font_size=None, text_color=None) -> dict:
    """도구 인자 → FindCharShape 항목. None인 조건은 넣지 않음."""
    cond = {}
    if bold is not None:
        cond["Bold"] = 1 if bold else 0
    if italic is not None:
        cond["Italic"] = 1 if italic else 0
    if underline is not None:
        cond["UnderlineType"] = 1 if underline else 0
    if font_size is not None:
        size = float(font_size)
        if size <= 0:
            raise ValueError("font_size는 0보다 커야 합니다(pt).")
        cond["Height"] = int(round(size * 100))
    if text_color is not None and text_color != "":
        cond["TextColor"] = color_to_colorref(text_color)
    return cond


# ---------------------------------------------------------------- 서식 찾기(COM)

def _pos_key(h):
    try:
        return tuple(h.GetPos()[:3])
    except Exception:
        return None


def _current_page(h):
    try:
        ki = h.KeyIndicator()
        return int(ki[3]) if isinstance(ki, tuple) and len(ki) > 3 else None
    except Exception:
        return None


def _selected_text(h) -> str:
    try:
        return h.GetTextFile("UNICODE", "saveblock") or ""
    except Exception:
        return ""


def _find_next_with_shape(h, text: str, cond: dict) -> bool:
    act = h.CreateAction("RepeatFind")
    s = act.CreateSet()
    act.GetDefault(s)
    s.SetItem("FindString", text or "")
    s.SetItem("IgnoreFindString", 0 if text else 1)
    s.SetItem("IgnoreMessage", 1)
    s.SetItem("Direction", 0)
    s.SetItem("FindRegExp", 0)
    if cond:
        cs = s.CreateItemSet("FindCharShape", "CharShape")
        for k, v in cond.items():
            cs.SetItem(k, v)
    return bool(act.Execute(s))


def reset_find_conditions(h):
    """
    한글이 기억하는 '마지막 찾기 조건'에서 글자모양 조건을 지웁니다.
    실측: 글자+서식 조건으로 찾은 뒤에는 HFindReplace 기본값에 FindCharShape가 남아,
    이후의 일반 글자 찾기(find_text, hwp_patch 등)가 서식 조건까지 걸려 실패함.
    빈 FindCharShape로 한 번 찾기를 실행하면 기억된 조건이 비워짐. 커서는 원래 자리로.
    """
    saved = None
    try:
        saved = h.GetPos()
    except Exception:
        pass
    try:
        act = h.CreateAction("RepeatFind")
        s = act.CreateSet()
        act.GetDefault(s)
        s.SetItem("FindString", "​")  # 문서에 없을 글자
        s.SetItem("IgnoreFindString", 0)
        s.SetItem("IgnoreMessage", 1)
        s.SetItem("Direction", 0)
        s.CreateItemSet("FindCharShape", "CharShape")  # 빈 조건
        act.Execute(s)
    except Exception as e:
        logger.warning(f"찾기 조건 초기화 실패: {e}")
    finally:
        try:
            h.HAction.Run("Cancel")
            if saved:
                h.SetPos(*saved[:3])
        except Exception:
            pass


def find_by_charshape(h, text: str, cond: dict, max_results: int = 50) -> list:
    """
    글자모양 조건(cond)에 맞는 글자를 문서 처음부터 찾아 목록으로 돌려줍니다.
    text가 있으면 그 글자 중 조건에 맞는 것만. 커서는 원래 자리로 돌려놓습니다.
    Returns: [{"index", "text", "page"}]
    """
    saved = None
    try:
        saved = h.GetPos()
    except Exception:
        pass
    results = []
    try:
        with _AutoOk(h):
            h.HAction.Run("Cancel")
            h.HAction.Run("MoveDocBegin")
            last = None
            while len(results) < max_results:
                if not _find_next_with_shape(h, text, cond):
                    break
                cur = _pos_key(h)
                if last is not None and cur is not None and cur <= last:
                    break  # 문서 처음으로 되돌아감
                last = cur
                found = _selected_text(h).replace("\r\n", "\n").strip("\n")
                results.append({"index": len(results) + 1, "text": found, "page": _current_page(h)})
    finally:
        reset_find_conditions(h)
        try:
            h.HAction.Run("Cancel")
            if saved:
                h.SetPos(*saved[:3])
        except Exception:
            pass
    return results


# ---------------------------------------------------------------- HWPML 읽기

def read_hwpml(h) -> str:
    """문서 전체를 HWPML(XML) 글자로 읽습니다. 문서를 저장하거나 바꾸지 않습니다."""
    with _AutoOk(h):
        return h.GetTextFile("HWPML2X", "") or ""


def parse_hwpml(xml: str):
    xml = re.sub(r"^\s*<\?xml[^>]*\?>", "", xml or "", count=1)
    return ET.fromstring(xml)


# HWPML에는 글꼴 이름이 영문으로만 들어가는 경우가 있어(예: "HCR Dotum") 한글 이름과 짝지어 비교
FONT_ALIASES = {
    "hcrdotum": "함초롬돋움", "hcrbatang": "함초롬바탕",
    "malgungothic": "맑은 고딕", "gulim": "굴림", "gulimche": "굴림체",
    "dotum": "돋움", "dotumche": "돋움체", "batang": "바탕", "batangche": "바탕체",
    "gungsuh": "궁서", "gungsuhche": "궁서체",
    "nanumgothic": "나눔고딕", "nanummyeongjo": "나눔명조", "nanumbarungothic": "나눔바른고딕",
}
_ALIAS_REVERSE = {v.replace(" ", ""): k for k, v in FONT_ALIASES.items()}


def _with_aliases(names: list) -> list:
    out = list(names)
    for n in names:
        key = (n or "").replace(" ", "").lower()
        if key in FONT_ALIASES:
            out.append(FONT_ALIASES[key])
        if key in _ALIAS_REVERSE:
            out.append(_ALIAS_REVERSE[key])
    return out


def _font_names(root) -> dict:
    """{(lang, id): [이름들]}"""
    out = {}
    for face in root.iter("FONTFACE"):
        lang = face.get("Lang")
        for font in face.findall("FONT"):
            names = [font.get("Name") or ""]
            for sub in font.findall("SUBSTFONT"):
                names.append(sub.get("Name") or "")
            out[(lang, font.get("Id"))] = _with_aliases([n for n in names if n])
    return out


def _charshapes(root) -> dict:
    fonts = _font_names(root)
    shapes = {}
    for cs in root.iter("CHARSHAPE"):
        fid = cs.find("FONTID")
        hangul = fonts.get(("Hangul", fid.get("Hangul"))) if fid is not None else None
        latin = fonts.get(("Latin", fid.get("Latin"))) if fid is not None else None
        ul = cs.find("UNDERLINE")
        shapes[cs.get("Id")] = {
            "Height": int(cs.get("Height") or 0),
            "TextColor": int(cs.get("TextColor") or 0),
            "Bold": 1 if cs.find("BOLD") is not None else 0,
            "Italic": 1 if cs.find("ITALIC") is not None else 0,
            "UnderlineType": 1 if (ul is not None and (ul.get("Type") or "None") != "None") else 0,
            "fonts": (hangul or []) + (latin or []),
        }
    return shapes


def _shape_matches(shape: dict, cond: dict, font_name: str = None) -> bool:
    for k, v in cond.items():
        if shape.get(k) != v:
            return False
    if font_name:
        want = font_name.replace(" ", "").lower()
        if not any(want in (n or "").replace(" ", "").lower() for n in shape.get("fonts", [])):
            return False
    return True


def find_by_format_hwpml(xml: str, text: str, cond: dict, font_name: str = None, max_results: int = 50) -> list:
    """
    HWPML을 읽어 조건에 맞는 글자 덩어리를 돌려줍니다(쪽 번호는 알 수 없음).
    같은 문단에서 이어지는 조건 일치 글자는 하나로 합칩니다.
    Returns: [{"index", "text", "paragraph", "page": None}]
    """
    root = parse_hwpml(xml)
    shapes = _charshapes(root)
    results = []
    for para_idx, runs in _iter_paragraph_runs_flat(root):
        buf = ""
        for cs, chars in runs + [(None, None)]:
            ok = cs is not None and _shape_matches(shapes.get(cs, {}), cond, font_name)
            if ok:
                buf += chars
                continue
            if buf.strip():
                if not text or text in buf:
                    results.append({"index": len(results) + 1, "text": buf.strip() if not text else text,
                                    "context": buf.strip()[:200], "paragraph": para_idx, "page": None})
                    if len(results) >= max_results:
                        return results
            buf = ""
    return results


def _iter_paragraph_runs_flat(root):
    """모든 문단(P)을 문서 순서대로. 메모 PARALIST 안 문단은 제외."""
    memo_paras = set()
    for fb in root.iter("FIELDBEGIN"):
        if fb.get("Type") == "Memo":
            for p in fb.iter("P"):
                memo_paras.add(id(p))
    body = root.find("BODY")
    if body is None:
        return
    idx = 0
    for p in body.iter("P"):
        if id(p) in memo_paras:
            continue
        idx += 1
        runs = []
        for t in p.findall("TEXT"):
            cs = t.get("CharShape")
            for child in t.findall("CHAR"):
                runs.append((cs, child.text or ""))
        yield idx, runs


# ---------------------------------------------------------------- 스타일

def list_styles(xml: str) -> list:
    """문서 스타일 목록 [{"index", "name", "eng_name", "type"}]. index는 Style 액션 Apply 값."""
    root = parse_hwpml(xml)
    out = []
    for st in root.iter("STYLE"):
        try:
            idx = int(st.get("Id"))
        except (TypeError, ValueError):
            continue
        out.append({"index": idx, "name": st.get("Name") or "", "eng_name": st.get("EngName") or "",
                    "type": st.get("Type") or ""})
    return out


def resolve_style(styles: list, style):
    """이름·영문 이름·번호로 스타일을 찾습니다. 못 찾으면 None."""
    if isinstance(style, int) or (isinstance(style, str) and style.strip().isdigit()):
        idx = int(style)
        return next((s for s in styles if s["index"] == idx), None)
    key = str(style or "").strip()
    if not key:
        return None
    for s in styles:
        if key in (s["name"], s["eng_name"]):
            return s
    low = key.replace(" ", "").lower()
    for s in styles:
        if low in (s["name"].replace(" ", "").lower(), s["eng_name"].replace(" ", "").lower()):
            return s
    return None


def current_style_index(h):
    act = h.CreateAction("Style")
    s = act.CreateSet()
    act.GetDefault(s)
    return int(s.Item("Apply"))


def apply_style(h, index: int) -> bool:
    """현재 문단(또는 선택한 문단들)에 스타일 번호 index를 적용."""
    act = h.CreateAction("Style")
    s = act.CreateSet()
    act.GetDefault(s)
    s.SetItem("Apply", int(index))
    return bool(act.Execute(s))


# ---------------------------------------------------------------- 메모

def insert_memo(h, text: str) -> bool:
    """현재 선택(없으면 커서 위치)에 메모를 붙이고 text를 씁니다."""
    if not h.HAction.Run("InsertFieldMemo"):
        return False
    try:
        _insert_lines(h, text or "", "BreakPara")
    finally:
        h.HAction.Run("CloseEx")
    return True


def _para_text_excluding_memo(p) -> str:
    out = []
    for t in p.findall("TEXT"):
        for child in t:
            if child.tag == "CHAR":
                out.append(child.text or "")
    return "".join(out)


def list_memos(xml: str) -> list:
    """
    문서의 메모 목록.
    Returns: [{"index", "author", "content", "target", "paragraph_text"}]
      target: 메모가 붙은 본문 글자, paragraph_text: 그 문단 전체 글자(메모 내용 제외)
    """
    root = parse_hwpml(xml)
    body = root.find("BODY")
    if body is None:
        return []
    memos = []
    open_ids = {}

    # FIELDBEGIN → 들어 있는 문단 매핑
    para_map = {}
    for p in body.iter("P"):
        for t in p.findall("TEXT"):
            for child in t:
                if child.tag == "FIELDBEGIN":
                    para_map[id(child)] = p

    def walk(elem):
        for child in elem:
            tag = child.tag
            if tag == "FIELDBEGIN" and child.get("Type") == "Memo":
                content_parts = []
                for pl in child.findall("PARALIST"):
                    for mp in pl.iter("P"):
                        content_parts.append(_para_text_excluding_memo(mp))
                cmd = child.get("Command") or ""
                parts = cmd.split("/")
                author = parts[5] if len(parts) > 5 else ""
                p = para_map.get(id(child))
                rec = {"index": len(memos) + 1, "author": author, "content": "\n".join(content_parts).strip(),
                       "target": "", "paragraph_text": _para_text_excluding_memo(p).strip() if p is not None else ""}
                memos.append(rec)
                open_ids[child.get("FieldId")] = rec
                continue  # 메모 내용 안으로는 들어가지 않음
            if tag == "FIELDEND" and child.get("Type") == "Memo":
                open_ids.pop(child.get("FieldId"), None)
                continue
            if tag == "CHAR":
                for rec in open_ids.values():
                    rec["target"] += child.text or ""
            walk(child)

    walk(body)
    for m in memos:
        m["target"] = m["target"].strip()
    return memos
