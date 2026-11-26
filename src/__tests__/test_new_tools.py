# -*- coding: utf-8 -*-
"""hwp_delete_image / hwp_fill_table_grid / hwp_get_text 범위 읽기 단위 테스트(가짜 COM 객체 사용)."""

from unittest.mock import MagicMock

import pytest

from src.tools import hwp_edit_helpers as eh
from src.tools.hwp_controller import HwpController


# ---------------------------------------------------------------- 가짜 표

class FakeAction:
    def __init__(self, owner):
        self.owner = owner

    def Run(self, name):
        return self.owner.run(name)

    def GetDefault(self, *_):
        return True

    def Execute(self, name, _hset):
        if name == "InsertText":
            self.owner.insert(self.owner.pset.Text)
            return True
        if name == "RepeatFind":
            return self.owner.found
        return False


class FakeTableHwp:
    """rows x cols 표. 셀마다 리스트 ID(1부터)를 가짐. 커서는 셀 단위로만 움직임."""

    def __init__(self, rows, cols, found=True, anchor_cell=4, merged_last=False):
        self.rows, self.cols = rows, cols
        self.cells = {i + 1: f"old{i + 1}" for i in range(rows * cols)}
        self.cur = anchor_cell
        self.found = found
        self.selected = False
        self.actions = []
        self.appended_rows = 0
        self.pset = MagicMock()
        self.HAction = FakeAction(self)
        self.HParameterSet = MagicMock()
        self.HParameterSet.HInsertText = self.pset
        self.HParameterSet.HFindReplace = MagicMock()

    @property
    def n(self):
        return self.rows * self.cols

    def GetPos(self):
        return (self.cur, 0, 0)

    def run(self, name):
        self.actions.append(name)
        r, c = divmod(self.cur - 1, self.cols)
        if name == "TableColBegin":
            self.cur = r * self.cols + 1
        elif name == "TableColEnd":
            self.cur = r * self.cols + self.cols
        elif name == "TableColPageUp":
            self.cur = c + 1
        elif name == "TableColPageDown":
            self.cur = (self.rows - 1) * self.cols + c + 1
        elif name == "TableRightCell":
            if self.cur == self.n:  # 실제 한글: 마지막 셀에서 행이 늘어남
                self.appended_rows += 1
            else:
                self.cur += 1
        elif name == "MoveSelListEnd":
            self.selected = True
        elif name == "Delete":
            if self.selected:
                self.cells[self.cur] = ""
                self.selected = False
        elif name == "Cancel":
            self.selected = False
        return True

    def insert(self, text):
        if self.selected:
            self.cells[self.cur] = ""
            self.selected = False
        self.cells[self.cur] += text


def test_fill_table_grid_basic_and_skip():
    h = FakeTableHwp(2, 3)
    ok, msg = eh.fill_table_grid(h, "앵커", [["a", None, "c"], ["", 5, "f"]])
    assert ok, msg
    assert h.cells == {1: "a", 2: "old2", 3: "c", 4: "old4", 5: "5", 6: "f"}
    assert "4칸" in msg
    assert h.appended_rows == 0


def test_fill_table_grid_skip_empty_false_clears():
    h = FakeTableHwp(1, 2)
    ok, msg = eh.fill_table_grid(h, "앵커", [["", "x"]], skip_empty=False)
    assert ok
    assert h.cells == {1: "", 2: "x"}


def test_fill_table_grid_never_passes_last_cell():
    h = FakeTableHwp(2, 2)
    ok, msg = eh.fill_table_grid(h, "앵커", [["1", "2"], ["3", "4"], ["5", "6"]])
    assert ok
    assert h.appended_rows == 0
    assert h.cells == {1: "1", 2: "2", 3: "3", 4: "4"}
    assert "2개" in msg


def test_fill_table_grid_stops_after_last_needed_cell():
    h = FakeTableHwp(2, 2)
    eh.fill_table_grid(h, "앵커", [["1", None], [None, None]])
    assert h.actions.count("TableRightCell") == 0


def test_fill_table_grid_multiline():
    h = FakeTableHwp(1, 1, anchor_cell=1)
    eh.fill_table_grid(h, "앵커", [["가\n나"]])
    assert h.cells[1] == "가나"
    assert "BreakPara" in h.actions


def test_fill_table_grid_errors():
    assert not eh.fill_table_grid(FakeTableHwp(1, 1, found=False), "없음", [["a"]])[0]
    assert not eh.fill_table_grid(FakeTableHwp(1, 1), "a", "not a list")[0]
    body = FakeTableHwp(1, 1)
    body.cur = 0  # 본문(리스트 0) → 표 안이 아님
    assert not eh.fill_table_grid(body, "a", [["x"]])[0]


def test_fill_table_grid_all_skipped_does_nothing():
    h = FakeTableHwp(1, 2)
    ok, msg = eh.fill_table_grid(h, "a", [[None, ""]])
    assert ok and "0칸" in msg
    assert h.actions == []


# ---------------------------------------------------------------- 그림 지우기

def test_delete_picture_with_caption(monkeypatch):
    h = MagicMock()
    pic = MagicMock()
    monkeypatch.setattr(eh, "find_picture_near", lambda *_: (pic, "캡션이 붙은 그림"))
    monkeypatch.setattr(eh, "find_anchor", lambda *_: True)
    monkeypatch.setattr(eh, "read_picture_size",
                        lambda c: {"WidthRelTo": 4, "Width": 28346, "HeightRelTo": 2, "Height": 14173})
    h.DeleteCtrl.return_value = True
    reader = MagicMock(return_value="<그림 1> 누적 순주거비\r\n")
    ok, msg = eh.delete_picture(h, "<그림 1>", text_reader=reader)
    assert ok, msg
    h.DeleteCtrl.assert_called_once_with(pic)
    assert "<그림 1> 누적 순주거비" in msg
    assert "100.0x50.0mm" in msg
    reader.assert_called_once()


def test_delete_picture_not_found(monkeypatch):
    h = MagicMock()
    monkeypatch.setattr(eh, "find_picture_near", lambda *_: (None, "그림을 찾지 못했습니다."))
    ok, msg = eh.delete_picture(h, "x")
    assert not ok
    h.DeleteCtrl.assert_not_called()


def test_delete_picture_bad_which():
    assert not eh.delete_picture(MagicMock(), "x", which="middle")[0]


def test_delete_picture_uses_anchor_list_para():
    """실제 find_picture_near 경로: 캡션에서 CloseEx 후 GetPos의 (list, para)와 그림 anchor가 맞는 그림을 지움."""
    h = MagicMock()
    h.HAction.Execute.return_value = True
    # 찾기 전 위치(save_pos) → 캡션 안 → CloseEx 후 본문
    positions = iter([(0, 0, 0), (5, 0, 3), (0, 7, 1)])
    h.GetPos.side_effect = lambda: next(positions)

    def ctrl(cid, desc, pos, nxt):
        c = MagicMock()
        c.CtrlID, c.UserDesc, c.Next = cid, desc, nxt
        ap = MagicMock()
        ap.Item.side_effect = lambda k: dict(zip(("List", "Para", "Pos"), pos))[k]
        c.GetAnchorPos.return_value = ap
        c.Properties.ItemExist.return_value = False
        return c

    target = ctrl("gso", "그림", (0, 7, 0), None)
    other = ctrl("gso", "그림", (0, 3, 0), target)
    h.HeadCtrl = ctrl("tbl", "표", (0, 1, 0), other)
    h.DeleteCtrl.return_value = True
    ok, msg = eh.delete_picture(h, "<그림 2>")
    assert ok, msg
    h.DeleteCtrl.assert_called_once_with(target)
    h.HAction.Run.assert_any_call("CloseEx")
    # 저장 계열 호출 금지
    for name in ("Save", "SaveAs"):
        getattr(h, name).assert_not_called()


# ---------------------------------------------------------------- 범위 읽기

def _controller(text):
    c = HwpController()
    c.hwp = MagicMock()
    c.is_hwp_running = True
    c.hwp.GetTextFile.return_value = text
    return c


def test_get_text_range_from_start_text():
    c = _controller("머리말\r\n1. 개요\r\n본문\r\n2. 결론\r\n끝")
    ok, text = c.get_text_range("2. 결론")
    assert ok and text == "2. 결론\r\n끝"


def test_get_text_range_max_chars():
    c = _controller("abcdefg")
    assert c.get_text_range("c", 3) == (True, "cde")
    assert c.get_text_range("", 2) == (True, "ab")


def test_get_text_range_not_found():
    ok, msg = _controller("abc").get_text_range("zzz")
    assert not ok and "zzz" in msg


def test_get_text_range_uses_message_box_wrapper():
    c = _controller("abc")
    c.get_text_range("b")
    calls = [k.args[0] for k in c.hwp.SetMessageBoxMode.call_args_list]
    assert calls == [0x00010000, 0x00000000]
    c.hwp.GetTextFile.assert_called_once_with("UNICODE", "")


# ---------------------------------------------------------------- hwp_duplicate 입력 검사

def _dup(**kw):
    h = MagicMock()
    args = dict(unit="table", anchor="남기업", target="ㅇ 자문위원")
    args.update(kw)
    return h, eh.duplicate(h, **args)


def test_duplicate_rejects_bad_unit_and_position():
    _, (ok, msg) = _dup(unit="cell")
    assert not ok and "unit" in msg
    _, (ok, msg) = _dup(position="after_match")
    assert not ok and "position" in msg


def test_duplicate_end_anchor_only_for_paragraphs():
    _, (ok, msg) = _dup(end_anchor="끝")
    assert not ok and "paragraphs" in msg


def test_duplicate_ambiguous_target_does_not_touch_document():
    h, (ok, msg) = _dup(text_reader=lambda: "남기업 ... ㅇ 자문위원 ... ㅇ 자문위원")
    assert not ok and "2곳에서 일치" in msg and "target_occurrence" in msg
    h.HAction.Run.assert_not_called()


def test_duplicate_occurrence_out_of_range():
    h, (ok, msg) = _dup(occurrence=3, text_reader=lambda: "남기업 남기업 ㅇ 자문위원")
    assert not ok and "2곳뿐" in msg
    h.HAction.Run.assert_not_called()


# ---------------------------------------------------------------- 찾기 실패 시 커서 복원

def _search_mock(found: bool):
    h = MagicMock()
    h.GetPos.return_value = (0, 42, 7)  # 사용자가 보던 위치
    h.HAction.Execute.return_value = found
    return h


@pytest.mark.parametrize("find", [
    lambda h: eh.find_anchor(h, "없는글자"),
    lambda h: eh.find_nth(h, "없는글자", 2),
])
def test_failed_find_restores_cursor(find):
    h = _search_mock(found=False)
    assert find(h) is False
    h.SetPos.assert_called_with(0, 42, 7)


def test_successful_find_keeps_found_position():
    h = _search_mock(found=True)
    assert eh.find_anchor(h, "있는글자") is True
    h.SetPos.assert_not_called()


def test_controller_find_text_restores_on_failure():
    c = HwpController()
    c.hwp = _search_mock(found=False)
    c.is_hwp_running = True
    assert c.find_text("없는글자") is False
    c.hwp.SetPos.assert_called_with(0, 42, 7)



# ---------------------------------------------------------------- 표 읽기(HWPML)·표 빠져나오기

from src.tools import hwp_format_helpers as fh


def _cell(r, c, text, rs=1, cs=1, nested=""):
    return (f'<CELL ColAddr="{c}" RowAddr="{r}" ColSpan="{cs}" RowSpan="{rs}"><PARALIST>'
            f'<P><TEXT><CHAR>{text}</CHAR>{nested}</TEXT></P></PARALIST></CELL>')


TABLE_XML = (
    '<HWPML><BODY><SECTION><P><TEXT>'
    '<TABLE RowCount="2" ColCount="3"><ROW>'
    + _cell(0, 0, "구분", rs=2) + _cell(0, 1, "성명") + _cell(0, 2, "전화") +
    '</ROW><ROW>' + _cell(1, 1, "홍길동") +
    _cell(1, 2, "010", nested='<TABLE RowCount="1" ColCount="1"><ROW>' + _cell(0, 0, "안쪽") + '</ROW></TABLE>') +
    '</ROW></TABLE></TEXT></P></SECTION></BODY></HWPML>'
)


def test_read_tables_rows_merged_and_nested():
    tables = fh.read_tables(TABLE_XML)
    assert len(tables) == 2  # 바깥 표 + 셀 안의 표
    outer, inner = tables
    assert outer["rows"] == [["구분", "성명", "전화"], ["홍길동", "010"]]  # 병합으로 가려진 칸은 빠짐
    assert outer["merged"] == [{"row": 1, "col": 1, "rowspan": 2, "colspan": 1, "text": "구분"}]
    assert (outer["row_count"], outer["col_count"]) == (2, 3)
    assert inner["index"] == 2 and inner["rows"] == [["안쪽"]]


def test_find_table_by_anchor_and_occurrence():
    tables = fh.read_tables(TABLE_XML)
    assert fh.find_table(tables, "홍길동")["index"] == 1
    assert fh.find_table(tables, "안쪽")["index"] == 2  # 바깥 셀 글자에는 안쪽 표 글자가 들어가지 않음
    assert fh.find_table(tables, "홍길동", 2) is None
    assert fh.find_table(tables, "없음") is None


def test_leave_table_moves_after_table_control():
    h = MagicMock()
    pos = iter([(8, 0, 3), (0, 1, 0)])  # 셀 안 → CloseEx 후 표 컨트롤 앞
    h.GetPos.side_effect = lambda: next(pos)
    assert eh.leave_table(h)
    h.HAction.Run.assert_any_call("CloseEx")
    h.SetPos.assert_called_with(0, 1, 1)  # 한 칸 뒤로 주면 한글이 표 바로 뒤로 맞춤


def test_leave_table_noop_in_body():
    h = MagicMock()
    h.GetPos.return_value = (0, 3, 2)
    assert eh.leave_table(h)
    h.HAction.Run.assert_not_called()
