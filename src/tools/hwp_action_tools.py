"""
한글 액션을 이름으로 찾아보고(hwp_action_help) 실행하는(hwp_run_action) 기능.

색인 data/hwp_actions.json은 한컴 공식 문서(ActionTable, ParameterSetTable)에서
scripts/build_action_index.py로 만든 것입니다(액션 이름·매개변수 묶음·항목·설명).
"""
import json
import logging
import os
from functools import lru_cache

try:
    from src.tools.hwp_format_helpers import _AutoOk
except ImportError:
    from hwp_format_helpers import _AutoOk

logger = logging.getLogger("hwp-action-tools")

INDEX_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "hwp_actions.json")

# 실행을 막는 액션. 파일 열기/닫기/저장·인쇄·창 조작·매크로/스크립트는 전용 도구를 쓰거나 위험함.
BLOCKED_PREFIXES = ("File", "Print", "Window", "Macro", "Script", "HwpCtrl", "Quit", "Exit")
# 이름에 이런 말이 들어가면 대화창을 여는 액션이라 한글이 대화창에서 멈춤
DIALOG_MARKERS = ("Dlg", "Dialog")


@lru_cache(maxsize=1)
def load_index() -> dict:
    with open(INDEX_PATH, encoding="utf-8") as f:
        return json.load(f)


def blocked_reason(name: str):
    """실행을 막을 이유(없으면 None)."""
    if name.startswith(BLOCKED_PREFIXES):
        return ("파일·인쇄·창·매크로·스크립트 계열 액션은 막혀 있습니다. "
                "열기/저장/닫기는 hwp_open·hwp_save·hwp_close_document를 쓰세요.")
    if any(m in name for m in DIALOG_MARKERS):
        return "대화창을 여는 액션은 한글이 대화창에서 멈추므로 막혀 있습니다. 매개변수(params)로 직접 지정하세요."
    return None


def _pset_detail(index: dict, pset_id: str):
    p = index["psets"].get(pset_id)
    if not p:
        return None
    return {"id": pset_id, "desc": p["desc"], "items": p["items"]}


def describe(name: str):
    """액션 하나의 자세한 설명. 없으면 None."""
    index = load_index()
    a = index["actions"].get(name)
    if a is None:
        return None
    out = {"action": name, **a}
    if a.get("pset"):
        out["pset_detail"] = _pset_detail(index, a["pset"]) or "(매개변수 표에 설명 없음)"
    blocked = blocked_reason(name)
    if blocked:
        out["blocked"] = blocked
    return out


def search(query: str, limit: int = 15) -> list:
    """이름·설명에 query의 모든 낱말이 들어간 액션 목록(대소문자 무시). 이름이 맞는 것을 앞에."""
    terms = [t.lower() for t in (query or "").split() if t]
    if not terms:
        return []
    hits = []
    for name, a in load_index()["actions"].items():
        hay = f"{name} {a.get('desc', '')}".lower()
        if all(t in hay for t in terms):
            in_name = all(t in name.lower() for t in terms)
            hits.append((0 if in_name else 1, len(name), name, a))
    hits.sort()
    return [{"action": n, "desc": a.get("desc", "")[:100], "pset": a.get("pset")}
            for _, _, n, a in hits[:limit]]


def _set_items(index: dict, pset_id: str, s, params: dict):
    """매개변수 묶음 s에 params를 넣습니다. 하위 묶음(PIT_SET)은 dict로 받음. 오류 메시지 또는 None."""
    known = {it["id"]: it for it in (index["psets"].get(pset_id) or {}).get("items", [])}
    for key, value in params.items():
        item = known.get(key)
        if known and item is None:
            return (f"'{pset_id}' 매개변수에 '{key}' 항목이 없습니다. "
                    f"쓸 수 있는 항목: {', '.join(known)}")
        if isinstance(value, dict):
            sub_id = (item or {}).get("sub")
            if not sub_id:
                return f"'{key}'는 하위 묶음이 아닙니다(값을 dict로 줄 수 없음)."
            sub = s.CreateItemSet(key, sub_id)
            err = _set_items(index, sub_id, sub, value)
            if err:
                return err
        elif isinstance(value, list):
            return f"'{key}': 배열 항목은 아직 지원하지 않습니다."
        else:
            s.SetItem(key, value)
    return None


def run_action(h, name: str, params: dict = None):
    """액션 실행. (성공, 메시지)

    - 매개변수 묶음이 있는 액션은 GetDefault → params 반영 → Execute(대화창 없이 실행).
      Run으로 실행하면 설정 대화창이 떠서 멈추므로 쓰지 않음.
    - 매개변수가 없는 액션은 Run.
    """
    name = (name or "").strip()
    if not name:
        return False, "action 이름이 필요합니다."
    blocked = blocked_reason(name)
    if blocked:
        return False, blocked
    try:
        act = h.CreateAction(name)
    except Exception:
        act = None
    if act is None:
        similar = [r["action"] for r in search(name, 5)]
        hint = f" 비슷한 이름: {', '.join(similar)}" if similar else " hwp_action_help로 이름을 찾아보세요."
        return False, f"'{name}' 액션이 없습니다.{hint}"

    index = load_index()
    s = act.CreateSet()
    with _AutoOk(h):
        if s is None:
            if params:
                return False, f"'{name}'은 매개변수를 받지 않는 액션입니다."
            ok = bool(act.Run())
        else:
            act.GetDefault(s)
            err = _set_items(index, s.SetID, s, params or {})
            if err:
                return False, err
            ok = bool(act.Execute(s))
    if not ok:
        return False, (f"한글이 '{name}' 실행을 거부했습니다(실행할 수 없는 상태일 수 있음: "
                       "선택이 필요하거나, 표·개체 안이어야 하는 등).")
    return True, f"'{name}' 실행 완료."
