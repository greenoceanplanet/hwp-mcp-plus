"""
검토·일괄 수정·문서 구조 읽기용 COM 헬퍼.

- 변경 추적 수락/거절 (TrackChangeApplyAll 등)
- 원문 기준 일괄 수정 (patch_document)
- 선택 영역 제자리 교체 (replace_selection)
- 누름틀 필드 읽기/쓰기 (GetFieldList / GetFieldText / PutFieldText)
- 문서 구조 요약 (scan_document)
- 쪽 단위 글자 읽기 (GetPageText)

hwp_edit_helpers와 마찬가지로 한글 COM 객체(HWPFrame.HwpObject)를 직접 받습니다.
"""

import logging

from src.tools.hwp_edit_helpers import (
    _insert_lines, _insert_plain, _is_picture, physical_page, printed_page, restore_pos, save_pos,
)
from src.tools.hwp_tracking import track_scope

logger = logging.getLogger("hwp-mcp-extended")

# GetFieldText/PutFieldText에서 여러 필드를 이을 때 쓰는 구분자
FIELD_SEP = "\x02"

TRACK_CHANGE_ACTIONS = {
    "apply_all": "TrackChangeApplyAll",
    "cancel_all": "TrackChangeCancelAll",
    "apply_next": "TrackChangeApplyNext",
    "cancel_next": "TrackChangeCancelNext",
    "apply": "TrackChangeApply",
    "cancel": "TrackChangeCancel",
}


class _MessageBoxAutoOk:
    """작업 동안 한글 대화창을 자동으로 확인(OK) 처리합니다."""

    def __init__(self, h):
        self.h = h

    def __enter__(self):
        try:
            self.h.SetMessageBoxMode(0x00010000)
        except Exception:
            pass
        return self

    def __exit__(self, *_):
        try:
            self.h.SetMessageBoxMode(0x00000000)
        except Exception:
            pass
        return False


# ---------------------------------------------------------------- 변경 추적

def resolve_track_changes(h, action: str):
    """변경 추적 내용을 수락(apply)하거나 거절(cancel)합니다."""
    act = TRACK_CHANGE_ACTIONS.get(action)
    if not act:
        return False, f"action은 {', '.join(TRACK_CHANGE_ACTIONS)} 중 하나여야 합니다: {action}"
    # 실측: 이 액션들은 실제로 반영/되돌림이 되어도 Run()이 False를 돌려줌 → 반환값으로 성패를 판단하지 않음
    with _MessageBoxAutoOk(h):
        h.HAction.Run(act)
    return True, f"{act} 실행 완료(처리할 변경 내용이 없었다면 문서는 그대로입니다)"


# ---------------------------------------------------------------- 찾기

def _pos_key(h):
    try:
        pos = h.GetPos()
        return tuple(pos[:3])
    except Exception:
        return None


def _find_next(h, text: str) -> bool:
    """현재 커서에서 앞쪽(문서 끝 방향)으로 text를 찾아 선택합니다."""
    pset = h.HParameterSet.HFindReplace
    h.HAction.GetDefault("RepeatFind", pset.HSet)
    pset.FindString = text
    pset.FindRegExp = 0
    pset.IgnoreMessage = 1
    pset.Direction = 0
    return bool(h.HAction.Execute("RepeatFind", pset.HSet))


def _find_nth(h, text: str, n: int) -> bool:
    """문서 처음부터 text의 n번째(1부터) 위치를 찾아 선택합니다. 문서 끝을 넘어 처음으로 돌아가면 실패.
    실패하면 커서를 원래 자리로 되돌립니다."""
    saved = save_pos(h)
    h.HAction.Run("Cancel")
    h.HAction.Run("MoveDocBegin")
    last = None
    for _ in range(n):
        if not _find_next(h, text):
            return restore_pos(h, saved)
        cur = _pos_key(h)
        if last is not None and cur is not None and cur <= last:
            return restore_pos(h, saved)  # 문서 처음으로 되돌아감
        last = cur
    return True


def _replace_selected(h, new: str):
    """선택된 글자를 new로 바꿉니다. new가 비면 지우기. 줄바꿈(\\n)은 문단 나눔."""
    if not new:
        h.HAction.Run("Delete")
        return
    if "\n" in new:
        h.HAction.Run("Delete")
        _insert_lines(h, new, "BreakPara")
    else:
        _insert_plain(h, new)  # InsertText는 선택 영역을 덮어씀


# ---------------------------------------------------------------- 일괄 수정

def patch_document(h, edits, text_reader=None, track: bool = False):
    """
    원문 일부(old)를 찾아 new로 바꾸는 작업을 건별로 실행하고 결과를 돌려줍니다.

    edits: [{"old": str, "new": str, "occurrence": int(선택), "all": bool(선택)}, ...]
      - old가 문서에 여러 번 있으면 occurrence(1부터)나 all=True가 없을 때 바꾸지 않고 "ambiguous".
    text_reader: 문서 전체 글자를 돌려주는 함수(일치 개수 세기용). None이면 개수를 세지 않음.
    track: True면 작업 동안 변경 추적을 켜고 끝나면 원래대로 돌림. False/None이면 설정을 건드리지 않음
      (MCP 도구에서는 tracked_edit 데코레이터가 변경 추적을 처리하므로 보통 None).

    Returns: [{"index", "old", "status", "count", "replaced", "message"}]
      status: ok / not_found / ambiguous / invalid / error
    """
    results = []
    if not isinstance(edits, list) or not edits:
        return [{"index": 0, "status": "invalid", "message": "edits는 비어 있지 않은 목록이어야 합니다."}]

    with track_scope(h, True if track else None):
        _patch_items(h, edits, text_reader, results)
    return results


def _patch_items(h, edits, text_reader, results):
    """patch_document의 항목별 처리. results에 결과를 채움."""
    for i, ed in enumerate(edits, 1):
        res = {"index": i, "old": None, "status": "error", "count": None, "replaced": 0, "message": ""}
        results.append(res)
        try:
            if not isinstance(ed, dict) or not ed.get("old"):
                res.update(status="invalid", message="각 항목에 비어 있지 않은 old가 필요합니다.")
                continue
            old = str(ed["old"])
            new = "" if ed.get("new") is None else str(ed.get("new"))
            res["old"] = old
            if "\n" in old or "\r" in old:
                res.update(status="invalid", message="old에 줄바꿈을 넣을 수 없습니다(한 문단 안의 글자만 찾을 수 있음).")
                continue
            occurrence = int(ed.get("occurrence") or 0)
            replace_all = bool(ed.get("all"))

            count = None
            if text_reader is not None:
                try:
                    count = (text_reader() or "").count(old)
                except Exception as e:
                    logger.debug(f"일치 개수 세기 실패(무시): {e}")
            res["count"] = count

            if count is not None and count > 1 and not occurrence and not replace_all:
                res.update(status="ambiguous",
                           message=f"{count}곳에서 일치합니다. old를 더 길게 쓰거나 occurrence/all을 지정하세요.")
                continue

            if replace_all:
                limit = count if count else 1000
                saved = save_pos(h)
                h.HAction.Run("Cancel")
                h.HAction.Run("MoveDocBegin")
                done, last = 0, None
                while done < limit and _find_next(h, old):
                    cur = _pos_key(h)
                    if last is not None and cur is not None and cur <= last:
                        h.HAction.Run("Cancel")
                        break  # 문서 처음으로 되돌아감
                    _replace_selected(h, new)
                    done += 1
                    last = _pos_key(h)
                res["replaced"] = done
                if done:
                    res.update(status="ok", message=f"{done}곳 바꿈")
                else:
                    restore_pos(h, saved)
                    res.update(status="not_found", message="찾지 못했습니다.")
                continue

            n = occurrence or 1
            if not _find_nth(h, old, n):
                res.update(status="not_found",
                           message="찾지 못했습니다." if n == 1 else f"{n}번째 일치를 찾지 못했습니다.")
                continue
            _replace_selected(h, new)
            res.update(status="ok", replaced=1, message="바꿈" if n == 1 else f"{n}번째 일치를 바꿈")
        except Exception as e:
            logger.error(f"patch 항목 {i} 실패: {e}", exc_info=True)
            res.update(status="error", message=str(e))


def summarize_patch(results) -> str:
    ok = sum(1 for r in results if r.get("status") == "ok")
    return f"{len(results)}건 중 {ok}건 성공, {len(results) - ok}건 실패"


# ---------------------------------------------------------------- 선택 영역 교체

def replace_selection(h, text: str, selected_text: str, preserve_linebreaks: bool = True):
    """선택된 영역을 text로 바꿉니다. selected_text는 호출 전에 읽어 둔 선택 글자."""
    if not selected_text:
        return False, "선택된 영역이 없습니다. 한글에서 바꿀 글자를 먼저 선택하세요."
    h.HAction.Run("Delete")
    if preserve_linebreaks:
        _insert_lines(h, text or "", "BreakPara")
    else:
        _insert_plain(h, (text or "").replace("\r\n", " ").replace("\n", " "))
    return True, f"선택 영역({len(selected_text)}자)을 {len(text or '')}자로 바꿨습니다."


# ---------------------------------------------------------------- 누름틀 필드

def _split_fields(s) -> list:
    return [x for x in str(s or "").split(FIELD_SEP) if x != ""]


def _field_base(name: str) -> str:
    """'이름{{0}}' → '이름'"""
    return name.split("{{", 1)[0]


def list_fields(h, number: int = 0, option: int = 0) -> list:
    """문서의 필드 이름 목록. number=1이면 같은 이름 필드에 {{n}} 번호를 붙임."""
    return _split_fields(h.GetFieldList(number, option))


def get_fields(h, names) -> dict:
    """필드 값 읽기. names가 비면 모든 필드."""
    if not names:
        names = list_fields(h, 1)
    if isinstance(names, str):
        names = [names]
    if not names:
        return {}
    texts = str(h.GetFieldText(FIELD_SEP.join(names)) or "").split(FIELD_SEP)
    texts += [""] * (len(names) - len(texts))
    return dict(zip(names, texts))


def put_fields(h, values: dict):
    """필드에 값 쓰기. 없는 필드 이름은 건너뛰고 알려줍니다."""
    if not values:
        return False, "values가 비어 있습니다.", []
    existing = set(_field_base(n) for n in list_fields(h, 0))
    names, texts, missing = [], [], []
    for name, value in values.items():
        name = str(name)
        if _field_base(name) not in existing:
            missing.append(name)
            continue
        names.append(name)
        texts.append("" if value is None else str(value))
    if names:
        h.PutFieldText(FIELD_SEP.join(names), FIELD_SEP.join(texts))
    msg = f"{len(names)}개 필드에 값을 넣었습니다."
    if missing:
        msg += f" 없는 필드: {', '.join(missing)}"
    return bool(names), msg, missing


# ---------------------------------------------------------------- 쪽 단위 읽기

def get_page_text(h, page: int) -> str:
    """page(1부터) 쪽의 글자를 돌려줍니다. 표·글상자 등 모든 글자 포함."""
    with _MessageBoxAutoOk(h):
        return h.GetPageText(page - 1, 0xFFFFFFFF) or ""


# ---------------------------------------------------------------- 문서 구조 요약

def _page_at_anchor(h, ctrl):
    """컨트롤이 놓인 쪽의 (순번, 인쇄 쪽 번호). 커서를 옮기므로 호출 전후로 위치를 저장/복원할 것."""
    try:
        h.SetPosBySet(ctrl.GetAnchorPos(0))
    except Exception:
        return None, None
    return physical_page(h), printed_page(h)


# 쪽 번호가 새로 시작될 수 있는 컨트롤: 구역 정의, 새 번호로 시작
_NUMBERING_CTRLS = ("secd", "nwno")


def _collect_controls(h, max_items: int = 200):
    """표·그림 목록과 그 밖의 컨트롤 개수, 쪽 번호가 바뀔 수 있는 지점 [(순번, 인쇄 번호)]를 모읍니다."""
    tables, pictures, others, numbering = [], [], {}, []
    ctrl = h.HeadCtrl
    while ctrl is not None:
        try:
            cid = ctrl.CtrlID
        except Exception:
            break
        if cid == "tbl" or _is_picture(ctrl):
            items = tables if cid == "tbl" else pictures
            if len(items) < max_items:
                page, printed = _page_at_anchor(h, ctrl)
                items.append({"index": len(items) + 1, "page": page, "printed_page": printed})
            else:
                items.append(None)
        else:
            others[cid] = others.get(cid, 0) + 1
            if cid in _NUMBERING_CTRLS:
                numbering.append(_page_at_anchor(h, ctrl))
        ctrl = ctrl.Next
    return tables, pictures, others, numbering


def _page_numbering(starts, total):
    """
    쪽 순번과 인쇄 쪽 번호의 대응을 구간으로 요약합니다.
    starts: 문서 순서대로 [(순번, 그 쪽의 인쇄 번호)] — 번호가 바뀔 수 있는 지점들.
    번호가 앞 구간에서 그대로 이어지는 지점은 합칩니다.
    Returns: [{"pages": "1-2", "printed": "1-2"}, {"pages": "3-20", "printed": "1-18"}]
    """
    by_page = {}
    for page, printed in starts:
        if page and printed is not None:
            by_page[page] = printed  # 같은 쪽에 여럿이면 뒤의 것(문서 순서상 나중)이 이김
    segs = []  # [순번 시작, 인쇄 번호 시작]
    for page in sorted(by_page):
        printed = by_page[page]
        if segs and printed == segs[-1][1] + (page - segs[-1][0]):
            continue  # 앞 구간에서 번호가 그대로 이어짐
        segs.append([page, printed])

    def span(a, b):
        return str(a) if a == b else f"{a}-{b}"

    out = []
    for i, (page, printed) in enumerate(segs):
        end = segs[i + 1][0] - 1 if i + 1 < len(segs) else (total or page)
        out.append({"pages": span(page, end), "printed": span(printed, printed + end - page)})
    return out


def _scan_get_text(h):
    """GetText(BSTR* text) 호출. 늦은 바인딩에서는 out 인자를 명시해야 해서 InvokeTypes를 씀."""
    ole = getattr(h, "_oleobj_", None)
    if ole is None:
        return h.GetText()
    import pythoncom
    dispid = ole.GetIDsOfNames("GetText")
    # 반환 long(VT_I4=3), 인자 BSTR*(VT_BSTR|VT_BYREF, PARAMFLAG_FOUT=2)
    return ole.InvokeTypes(dispid, 0, pythoncom.DISPATCH_METHOD, (3, 0), ((8 | 0x4000, 2),), "")


def _collect_headings(h, max_paras: int = 3000, max_headings: int = 200):
    """
    본문 문단을 훑어 개요(HeadingType=1) 문단을 모읍니다. 표 등 개체 안 문단은 건너뜀.
    InitScan → GetText 반복. 실측 반환: 2=같은 문단 글자, 3=다음 문단 시작(그 문단 글자 포함),
    4/5=개체 안으로 들어감/나옴, 0/1=끝. 문단이 시작될 때 MovePos(201: 스캔 위치로 커서 이동) 후
    ParaShape로 개요 여부를 확인합니다.
    """
    headings = []
    paras = 0
    h.InitScan(0x00, 0x0077, 0, 0, 0, 0)  # maskNormal, 문서 처음~끝 (늦은 바인딩은 인자 6개 모두 필요)
    try:
        depth = 0
        buf = ""
        heading_info = None
        started = False

        def flush():
            nonlocal buf, heading_info
            if heading_info is not None and buf.strip():
                level, number = heading_info
                headings.append({"level": level, "number": number, "text": buf.strip()[:120]})
            buf = ""
            heading_info = None

        def read_heading():
            try:
                h.MovePos(201, 0, 0)  # 늦은 바인딩은 인자 3개 모두 필요
                ps = h.ParaShape
                if int(ps.Item("HeadingType")) == 1:
                    number = ""
                    try:
                        number = str(h.GetHeadingString() or "")
                    except Exception:
                        pass
                    return int(ps.Item("Level")) + 1, number
            except Exception as e:
                logger.debug(f"문단 모양 읽기 실패(무시): {e}")
            return None

        while paras < max_paras and len(headings) < max_headings:
            ret = _scan_get_text(h)
            state, text = (ret[0], ret[1]) if isinstance(ret, tuple) else (ret, "")
            if state in (0, 1, 101, 102):
                break
            if state == 4:
                depth += 1
                continue
            if state == 5:
                depth = max(0, depth - 1)
                continue
            if depth or state not in (2, 3):
                continue
            if state == 3 or not started:
                flush()
                started = True
                paras += 1
                heading_info = read_heading()
            buf += text or ""
        flush()
    finally:
        h.ReleaseScan()
    return headings, paras


def scan_document(h, include_headings: bool = True) -> dict:
    """문서 구조 요약: 쪽 수, 표·그림 목록(쪽 위치), 개요 제목 목록."""
    saved = None
    try:
        saved = h.GetPos()
    except Exception:
        pass
    result = {"pages": None, "tables": [], "pictures": [], "other_controls": {}, "headings": []}
    try:
        try:
            result["pages"] = int(h.PageCount)
        except Exception:
            pass
        tables, pictures, others, numbering = _collect_controls(h)
        result["page_numbering"] = _page_numbering(numbering, result["pages"])
        result["table_count"] = len(tables)
        result["picture_count"] = len(pictures)
        result["tables"] = [t for t in tables if t]
        result["pictures"] = [p for p in pictures if p]
        result["other_controls"] = others
        if include_headings:
            try:
                result["headings"], result["paragraphs_scanned"] = _collect_headings(h)
            except Exception as e:
                logger.warning(f"개요 목록 읽기 실패: {e}")
                result["headings_error"] = str(e)
    finally:
        try:
            h.HAction.Run("Cancel")
            if saved:
                h.SetPos(*saved[:3])
        except Exception:
            pass
    return result
