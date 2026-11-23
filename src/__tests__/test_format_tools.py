# -*- coding: utf-8 -*-
"""오류 형식·변경 추적·서식 찾기·스타일·메모 단위 테스트(가짜 COM 객체 사용)."""

from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from src.tools import hwp_errors as errors
from src.tools import hwp_format_helpers as fh
from src.tools import hwp_tracking as tr
from src.tools.hwp_controller import HwpController


# ---------------------------------------------------------------- 오류 형식

def test_err_format_has_code_summary_and_next():
    s = errors.err(errors.UNSAVED_DOCUMENT, "새 문서라 저장 경로가 없습니다.")
    assert s.startswith("Error[UNSAVED_DOCUMENT]: 새 문서라")
    assert " / 다음: " in s and "경로를 추측하지" in s


def test_err_custom_next_overrides_default():
    s = errors.err(errors.NOT_FOUND, "x", "다른 안내")
    assert s.endswith("/ 다음: 다른 안내")


def test_every_code_has_next_action():
    codes = [v for k, v in vars(errors).items() if k.isupper() and isinstance(v, str) and k != "NEXT_ACTIONS"]
    for c in codes:
        assert c in errors.NEXT_ACTIONS, c


@pytest.mark.parametrize("msg,code", [
    ("기준 글자를 찾지 못했습니다: 가", errors.NOT_FOUND),
    ("선택된 영역이 없습니다.", errors.NO_SELECTION),
    ("count는 1 이상이어야 합니다.", errors.INVALID_ARG),
    ("Text is required", errors.INVALID_ARG),
    ("3곳에서 일치합니다.", errors.AMBIGUOUS),
    ("행·열 넣기 실패(커서가 표 안에 있는지 확인하세요)", errors.NOT_IN_TABLE),
    ("알 수 없는 실패", errors.COM_ERROR),
])
def test_classify(msg, code):
    assert errors.classify(msg) == code


def test_from_message_strips_old_prefix_and_keeps_standard():
    assert errors.from_message("Error: 찾지 못했습니다").startswith("Error[NOT_FOUND]: 찾지 못했습니다")
    std = errors.err(errors.INVALID_ARG, "x")
    assert errors.from_message(std) == std


# ---------------------------------------------------------------- 변경 추적

class TrackCom:
    def __init__(self, initial=0):
        self.IsTrackChange = initial
        self.history = []

    def __setattr__(self, k, v):
        if k == "IsTrackChange" and "history" in self.__dict__:
            self.history.append(v)
        super().__setattr__(k, v)


@pytest.mark.parametrize("env,expected", [(None, True), ("on", True), ("off", False), ("0", False), ("TRUE", True)])
def test_env_default(monkeypatch, env, expected):
    if env is None:
        monkeypatch.delenv(tr.ENV_NAME, raising=False)
    else:
        monkeypatch.setenv(tr.ENV_NAME, env)
    assert tr.resolve_track(None) is expected


def test_argument_beats_env(monkeypatch):
    monkeypatch.setenv(tr.ENV_NAME, "off")
    assert tr.resolve_track(True) is True
    monkeypatch.setenv(tr.ENV_NAME, "on")
    assert tr.resolve_track(False) is False


def test_track_scope_restores_previous_value_even_on_error():
    h = TrackCom(initial=1)
    with pytest.raises(RuntimeError):
        with tr.track_scope(h, False):
            assert h.IsTrackChange == 0
            raise RuntimeError("boom")
    assert h.IsTrackChange == 1


def test_track_scope_none_does_not_touch():
    h = TrackCom(initial=0)
    with tr.track_scope(h, None):
        pass
    assert h.history == []


def _make_tool(h, behavior):
    @tr.tracked_edit(lambda: h)
    def tool(x: str = "", track: bool = None):
        h.seen = h.IsTrackChange
        return behavior()
    return tool


def test_tracked_edit_default_on_and_note(monkeypatch):
    monkeypatch.delenv(tr.ENV_NAME, raising=False)
    h = TrackCom(initial=0)
    tool = _make_tool(h, lambda: "완료")
    out = tool("a")
    assert h.seen == 1 and h.IsTrackChange == 0
    assert out.startswith("완료") and "hwp_track_changes_resolve" in out


def test_tracked_edit_env_off_no_note(monkeypatch):
    monkeypatch.setenv(tr.ENV_NAME, "off")
    h = TrackCom(initial=0)
    out = _make_tool(h, lambda: "완료")("a")
    assert h.seen == 0 and out == "완료"


def test_tracked_edit_argument_priority(monkeypatch):
    monkeypatch.setenv(tr.ENV_NAME, "off")
    h = TrackCom(initial=0)
    _make_tool(h, lambda: "완료")("a", track=True)
    assert h.seen == 1 and h.IsTrackChange == 0


def test_tracked_edit_restores_on_exception(monkeypatch):
    monkeypatch.delenv(tr.ENV_NAME, raising=False)
    h = TrackCom(initial=0)

    def boom():
        raise RuntimeError("x")

    with pytest.raises(RuntimeError):
        _make_tool(h, boom)("a")
    assert h.IsTrackChange == 0


def test_tracked_edit_no_note_on_error_and_dict(monkeypatch):
    monkeypatch.delenv(tr.ENV_NAME, raising=False)
    h = TrackCom()
    assert _make_tool(h, lambda: "Error[NOT_FOUND]: x")("a") == "Error[NOT_FOUND]: x"
    assert "track_changes" in _make_tool(h, lambda: {"ok": True})("a")
    assert "track_changes" not in _make_tool(h, lambda: {"error": "e"})("a")


def test_tracked_edit_without_connection_just_calls(monkeypatch):
    @tr.tracked_edit(lambda: None)
    def tool(track: bool = None):
        return "Error[HWP_NOT_RUNNING]: x"
    assert tool() == "Error[HWP_NOT_RUNNING]: x"


def test_controller_insert_text_no_longer_forces_tracking():
    c = HwpController.__new__(HwpController)
    c.hwp = TrackCom(initial=0)
    c.is_hwp_running = True
    c.get_selected_text = lambda: ""
    c._insert_text_direct = lambda t: True
    assert c.insert_text("가") is True
    assert c.hwp.history == []  # 건드리지 않음
    assert c.insert_text("가", track=True) is True
    assert c.hwp.history == [1, 0]  # 작업 동안만 켜고 원래 값으로


# ---------------------------------------------------------------- 서식 조건

def test_build_conditions_and_colors():
    cond = fh.build_charshape_conditions(bold=True, italic=False, underline=True, font_size=14, text_color="red")
    assert cond == {"Bold": 1, "Italic": 0, "UnderlineType": 1, "Height": 1400, "TextColor": 0x0000FF}
    assert fh.build_charshape_conditions() == {}
    assert fh.color_to_colorref("#112233") == 0x332211
    assert fh.colorref_to_hex(0x332211) == "#112233"
    with pytest.raises(ValueError):
        fh.color_to_colorref("보라빛")
    with pytest.raises(ValueError):
        fh.build_charshape_conditions(font_size=0)


class FakeSet:
    def __init__(self):
        self.items = {}

    def SetItem(self, k, v):
        self.items[k] = v

    def CreateItemSet(self, k, _type):
        sub = FakeSet()
        self.items[k] = sub
        return sub


class FakeFindAction:
    def __init__(self, owner):
        self.owner = owner

    def CreateSet(self):
        return FakeSet()

    def GetDefault(self, s):
        return True

    def Execute(self, s):
        return self.owner.execute(s)


class FakeShapeFindHwp:
    """미리 정한 일치 위치를 차례로 돌려줌. 끝나면 처음으로 되돌아감(wrap)."""

    def __init__(self, hits):
        self.hits = hits  # [(pos, text, page)]
        self.i = 0
        self.cur = (0, 0, 0)
        self.sel = ""
        self.executed = []
        self.HAction = SimpleNamespace(Run=lambda name: True)

    def CreateAction(self, name):
        return FakeFindAction(self)

    def execute(self, s):
        self.executed.append(s.items)
        if s.items.get("FindString") == "​":
            return False  # 초기화용 찾기
        pos, text, page = self.hits[self.i % len(self.hits)]
        self.i += 1
        self.cur, self.sel, self.page = (0, 0, pos), text, page
        return True

    def GetPos(self):
        return self.cur

    def SetPos(self, *a):
        self.cur = tuple(a)

    def GetTextFile(self, fmt, opt):
        return self.sel

    def KeyIndicator(self):
        return (True, 1, 1, self.page, 1, 1, 1, 0, "")

    def SetMessageBoxMode(self, _):
        return 0


def test_find_by_charshape_stops_on_wrap_and_resets_conditions():
    h = FakeShapeFindHwp([(5, "굵은말", 1), (30, "둘째", 2)])
    res = fh.find_by_charshape(h, "", {"Bold": 1})
    assert [(r["text"], r["page"]) for r in res] == [("굵은말", 1), ("둘째", 2)]
    first = h.executed[0]
    assert first["IgnoreFindString"] == 1 and first["FindCharShape"].items == {"Bold": 1}
    last = h.executed[-1]
    assert last["FindString"] == "​" and last["FindCharShape"].items == {}  # 조건 초기화


# ---------------------------------------------------------------- HWPML(스타일·메모·글꼴)

HWPML = """<?xml version="1.0" encoding="UTF-16" standalone="no" ?><HWPML Version="2.91"><HEAD><MAPPINGTABLE>
<FACENAMELIST><FONTFACE Count="2" Lang="Hangul"><FONT Id="0" Name="HCR Dotum" Type="ttf"/><FONT Id="1" Name="함초롬바탕" Type="ttf"><SUBSTFONT Name="HCR Batang" Type="ttf"/></FONT></FONTFACE>
<FONTFACE Count="2" Lang="Latin"><FONT Id="0" Name="HCR Dotum" Type="ttf"/><FONT Id="1" Name="함초롬바탕" Type="ttf"/></FONTFACE></FACENAMELIST>
<CHARSHAPELIST Count="3">
<CHARSHAPE Height="1000" Id="0" TextColor="0"><FONTID Hangul="1" Latin="1"/></CHARSHAPE>
<CHARSHAPE Height="1400" Id="1" TextColor="255"><FONTID Hangul="1" Latin="1"/><ITALIC/><BOLD/><UNDERLINE Color="0" Shape="Solid" Type="Bottom"/></CHARSHAPE>
<CHARSHAPE Height="1000" Id="2" TextColor="0"><FONTID Hangul="0" Latin="0"/><BOLD/></CHARSHAPE>
</CHARSHAPELIST>
<STYLELIST Count="3"><STYLE EngName="Normal" Id="0" Name="바탕글" Type="Para"/><STYLE EngName="Body" Id="1" Name="본문" Type="Para"/><STYLE EngName="Outline 1" Id="2" Name="개요 1" Type="Para"/></STYLELIST>
</MAPPINGTABLE></HEAD><BODY><SECTION Id="0">
<P Style="0"><TEXT CharShape="0"><CHAR>보통 </CHAR></TEXT><TEXT CharShape="1"><CHAR>모두</CHAR></TEXT><TEXT CharShape="0"><CHAR> 끝</CHAR></TEXT></P>
<P Style="0"><TEXT CharShape="0"><CHAR>둘째 문단 </CHAR><FIELDBEGIN Command="MEMO/65535/1/1/2/홍길동/\\;;" FieldId="7" Type="Memo"><PARALIST><P Style="16"><TEXT CharShape="0"><CHAR>검토 의견</CHAR></TEXT></P></PARALIST></FIELDBEGIN><CHAR>메모대상</CHAR><FIELDEND FieldId="7" Type="Memo"/><CHAR> 글자.</CHAR></TEXT></P>
<P Style="0"><TEXT CharShape="0"><TABLE><ROW><CELL><PARALIST><P Style="0"><TEXT CharShape="2"><CHAR>셀돋움굵음</CHAR></TEXT></P></PARALIST></CELL></ROW></TABLE><CHAR/></TEXT></P>
</SECTION></BODY></HWPML>"""


def test_list_and_resolve_styles():
    styles = fh.list_styles(HWPML)
    assert [s["name"] for s in styles] == ["바탕글", "본문", "개요 1"]
    assert fh.resolve_style(styles, "개요 1")["index"] == 2
    assert fh.resolve_style(styles, "outline1")["index"] == 2
    assert fh.resolve_style(styles, "2")["name"] == "개요 1"
    assert fh.resolve_style(styles, 1)["name"] == "본문"
    assert fh.resolve_style(styles, "없음") is None


def test_list_memos():
    memos = fh.list_memos(HWPML)
    assert memos == [{"index": 1, "author": "홍길동", "content": "검토 의견", "target": "메모대상",
                      "paragraph_text": "둘째 문단 메모대상 글자."}]


def test_find_by_format_hwpml_font_and_shape():
    dotum = fh.find_by_format_hwpml(HWPML, "", {}, font_name="돋움")
    assert [m["text"] for m in dotum] == ["셀돋움굵음"]  # 표 안 문단도 포함
    batang_bold = fh.find_by_format_hwpml(HWPML, "", {"Bold": 1}, font_name="바탕")
    assert [m["text"] for m in batang_bold] == ["모두"]
    red = fh.find_by_format_hwpml(HWPML, "", {"TextColor": 255, "Height": 1400, "UnderlineType": 1, "Italic": 1})
    assert [m["text"] for m in red] == ["모두"]
    # 메모 내용은 본문 검색 대상이 아님
    assert fh.find_by_format_hwpml(HWPML, "검토", {}, font_name="바탕") == []


def test_apply_and_current_style_use_style_action():
    s = FakeSet()
    s.Item = lambda k: 2
    act = MagicMock()
    act.CreateSet.return_value = s
    act.Execute.return_value = True
    h = MagicMock()
    h.CreateAction.return_value = act
    assert fh.apply_style(h, 2) is True
    assert s.items["Apply"] == 2
    assert fh.current_style_index(h) == 2


def test_insert_memo_closes_even_on_error(monkeypatch):
    runs = []
    h = MagicMock()
    h.HAction.Run.side_effect = lambda n: runs.append(n) or True

    def boom(*a, **k):
        raise RuntimeError("x")

    monkeypatch.setattr(fh, "_insert_lines", boom)
    with pytest.raises(RuntimeError):
        fh.insert_memo(h, "의견")
    assert runs == ["InsertFieldMemo", "CloseEx"]
