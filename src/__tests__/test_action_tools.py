"""hwp_action_tools: 색인 검색·설명·막기·실행(가짜 COM)."""
from unittest.mock import MagicMock

from src.tools import hwp_action_tools as at


def test_index_loads_with_known_entries():
    idx = at.load_index()
    assert "TableSplitCell" in idx["actions"]
    items = [i["id"] for i in idx["psets"]["TableSplitCell"]["items"]]
    assert {"Rows", "Cols", "Merge"} <= set(items)


def test_search_korean_and_name_first():
    names = [r["action"] for r in at.search("셀 나누기")]
    assert names[0] == "TableSplitCell"
    assert at.search("") == []


def test_describe_includes_pset_items_and_block_reason():
    d = at.describe("ParagraphShape")
    assert d["pset"] == "ParaShape" and any(i["id"] == "AlignType" for i in d["pset_detail"]["items"])
    assert "blocked" in at.describe("FileClose")
    assert at.describe("NoSuch") is None


def test_blocked_reason():
    assert at.blocked_reason("FileSaveAs")
    assert at.blocked_reason("ParagraphShapeDialog")
    assert at.blocked_reason("CharShapeBold") is None


def _hwp(set_id=None):
    h = MagicMock()
    act = MagicMock()
    act.Run.return_value = True
    act.Execute.return_value = True
    if set_id:
        s = MagicMock()
        s.SetID = set_id
        act.CreateSet.return_value = s
    else:
        act.CreateSet.return_value = None
    h.CreateAction.return_value = act
    return h, act


def test_run_action_without_pset_uses_run():
    h, act = _hwp()
    ok, msg = at.run_action(h, "CharShapeBold")
    assert ok and act.Run.called and not act.Execute.called


def test_run_action_with_pset_uses_execute_and_sets_items():
    h, act = _hwp("ParaShape")
    ok, msg = at.run_action(h, "ParagraphShape", {"AlignType": 3})
    assert ok
    s = act.CreateSet.return_value
    act.GetDefault.assert_called_with(s)
    s.SetItem.assert_called_with("AlignType", 3)
    act.Execute.assert_called_with(s)
    assert not act.Run.called  # Run은 대화창을 띄우므로 쓰지 않음


def test_run_action_rejects_unknown_item_and_missing_action():
    h, act = _hwp("ParaShape")
    ok, msg = at.run_action(h, "ParagraphShape", {"Align": 3})
    assert not ok and "AlignType" in msg and not act.Execute.called
    h.CreateAction.return_value = None
    ok, msg = at.run_action(h, "TableSplitCel")
    assert not ok and "TableSplitCell" in msg  # 비슷한 이름 제안


def test_run_action_blocked_never_touches_hwp():
    h, _ = _hwp()
    ok, msg = at.run_action(h, "FileClose")
    assert not ok
    h.CreateAction.assert_not_called()
