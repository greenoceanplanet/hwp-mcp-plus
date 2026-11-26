# -*- coding: utf-8 -*-
"""hwp_review_helpers 단위 테스트(가짜 COM 객체 사용)."""

from types import SimpleNamespace
from unittest.mock import MagicMock

from src.tools import hwp_review_helpers as rh
from src.tools.hwp_controller import HwpController


# ---------------------------------------------------------------- 가짜 글자 문서

class FakeAction:
    def __init__(self, owner):
        self.owner = owner

    def Run(self, name):
        return self.owner.run(name)

    def GetDefault(self, *_):
        return True

    def Execute(self, name, _hset):
        if name == "InsertText":
            self.owner.insert(self.owner.ins.Text)
            return True
        if name == "RepeatFind":
            return self.owner.find(self.owner.fr.FindString)
        return False


class FakeTextHwp:
    """한 줄 글자 문서. 커서와 선택(start, end)만 흉내. wrap=True면 찾기가 문서 처음으로 되돌아감."""

    def __init__(self, text, wrap=False):
        self.text = text
        self.cur = 0
        self.sel = None
        self.wrap = wrap
        self.actions = []
        self.IsTrackChange = 0
        self.track_log = []
        self.ins = SimpleNamespace(Text="", HSet=None)
        self.fr = SimpleNamespace(FindString="", HSet=None)
        self.HParameterSet = SimpleNamespace(HInsertText=self.ins, HFindReplace=self.fr)
        self.HAction = FakeAction(self)

    def __setattr__(self, k, v):
        if k == "IsTrackChange" and "track_log" in self.__dict__:
            self.track_log.append(v)
        super().__setattr__(k, v)

    def GetPos(self):
        return (0, 0, self.cur)

    def SetMessageBoxMode(self, _):
        return 0

    def find(self, s):
        i = self.text.find(s, self.cur)
        if i < 0 and self.wrap:
            i = self.text.find(s)
        if i < 0:
            return False
        self.sel = (i, i + len(s))
        self.cur = i + len(s)
        return True

    def _delete_sel(self):
        if self.sel:
            a, b = self.sel
            self.text = self.text[:a] + self.text[b:]
            self.cur = a
            self.sel = None

    def insert(self, s):
        self._delete_sel()
        self.text = self.text[:self.cur] + s + self.text[self.cur:]
        self.cur += len(s)

    def run(self, name):
        self.actions.append(name)
        if name == "MoveDocBegin":
            self.cur, self.sel = 0, None
        elif name == "Cancel":
            self.sel = None
        elif name == "Delete":
            self._delete_sel()
        elif name == "BreakPara":
            self.insert("\n")
        elif name.startswith("TrackChange"):
            return True
        return True


def reader(h):
    return lambda: h.text


# ---------------------------------------------------------------- patch

def test_patch_basic_and_not_found():
    h = FakeTextHwp("가나다 라마바 사아자")
    res = rh.patch_document(h, [{"old": "라마바", "new": "LMB"}, {"old": "없음", "new": "x"}], reader(h))
    assert h.text == "가나다 LMB 사아자"
    assert [r["status"] for r in res] == ["ok", "not_found"]
    assert rh.summarize_patch(res) == "2건 중 1건 성공, 1건 실패"


def test_patch_ambiguous_is_skipped():
    h = FakeTextHwp("사과 배 사과")
    res = rh.patch_document(h, [{"old": "사과", "new": "감"}], reader(h))
    assert res[0]["status"] == "ambiguous"
    assert res[0]["count"] == 2
    assert h.text == "사과 배 사과"


def test_patch_occurrence_and_all():
    h = FakeTextHwp("a1 a1 a1")
    res = rh.patch_document(h, [{"old": "a1", "new": "B", "occurrence": 2}], reader(h))
    assert res[0]["status"] == "ok"
    assert h.text == "a1 B a1"
    res = rh.patch_document(h, [{"old": "a1", "new": "C", "all": True}], reader(h))
    assert res[0]["replaced"] == 2
    assert h.text == "C B C"


def test_patch_all_stops_on_wraparound():
    # 새 글자가 old를 포함해도(문서 처음으로 돌아가는 찾기여도) 무한 반복하지 않음
    h = FakeTextHwp("x y x", wrap=True)
    res = rh.patch_document(h, [{"old": "x", "new": "xx", "all": True}], None)  # 개수 모름 → 되돌아감 감지로 멈춰야 함
    assert res[0]["replaced"] == 2
    assert h.text == "xx y xx"


def test_patch_occurrence_beyond_count_with_wrap():
    h = FakeTextHwp("k k", wrap=True)
    res = rh.patch_document(h, [{"old": "k", "new": "z", "occurrence": 3}], reader(h))
    assert res[0]["status"] == "not_found"
    assert h.text == "k k"


def test_patch_delete_and_multiline_new():
    h = FakeTextHwp("첫째 둘째 셋째")
    res = rh.patch_document(h, [{"old": " 둘째", "new": ""}, {"old": "셋째", "new": "셋\n넷"}], reader(h))
    assert [r["status"] for r in res] == ["ok", "ok"]
    assert h.text == "첫째 셋\n넷"


def test_patch_invalid_items():
    h = FakeTextHwp("abc")
    res = rh.patch_document(h, [{"new": "x"}, "bad", {"old": "a\nb", "new": "x"}], reader(h))
    assert [r["status"] for r in res] == ["invalid", "invalid", "invalid"]
    assert rh.patch_document(h, [], reader(h))[0]["status"] == "invalid"


def test_patch_track_turns_on_and_restores():
    h = FakeTextHwp("abc")
    rh.patch_document(h, [{"old": "b", "new": "B"}], reader(h), track=True)
    assert h.track_log == [1, 0]
    assert h.text == "aBc"


# ---------------------------------------------------------------- 선택 교체·변경 추적

def test_replace_selection_requires_selection():
    h = FakeTextHwp("abc")
    ok, msg = rh.replace_selection(h, "x", "")
    assert not ok and "선택" in msg
    assert h.text == "abc"


def test_replace_selection_replaces_in_place():
    h = FakeTextHwp("하나 둘 셋")
    h.sel, h.cur = (3, 4), 4  # "둘" 선택
    ok, _ = rh.replace_selection(h, "2\n이", "둘")
    assert ok
    assert h.text == "하나 2\n이 셋"


def test_replace_selection_without_linebreaks():
    h = FakeTextHwp("ab")
    h.sel, h.cur = (0, 1), 1
    rh.replace_selection(h, "x\ny", "a", preserve_linebreaks=False)
    assert h.text == "x yb"


def test_resolve_track_changes():
    h = FakeTextHwp("")
    ok, msg = rh.resolve_track_changes(h, "apply_all")
    assert ok and "TrackChangeApplyAll" in h.actions
    h.run = lambda name: False  # 실제 한글은 성공해도 False를 돌려줌
    ok, _ = rh.resolve_track_changes(h, "cancel_all")
    assert ok
    ok, msg = rh.resolve_track_changes(h, "nope")
    assert not ok


# ---------------------------------------------------------------- 누름틀 필드

class FakeFieldHwp:
    def __init__(self, fields):
        self.fields = fields  # {"이름": ["값", ...]}
        self.put_calls = []

    def GetFieldList(self, number, option):
        out = []
        for k, vals in self.fields.items():
            if number == 1 and len(vals) > 1:
                out += [f"{k}{{{{{i}}}}}" for i in range(len(vals))]
            else:
                out.append(k)
        return rh.FIELD_SEP.join(out)

    def GetFieldText(self, names):
        res = []
        for n in names.split(rh.FIELD_SEP):
            base = n.split("{{")[0]
            idx = int(n.split("{{")[1].rstrip("}")) if "{{" in n else 0
            res.append(self.fields[base][idx])
        return rh.FIELD_SEP.join(res)

    def PutFieldText(self, names, texts):
        self.put_calls.append((names, texts))
        for n, t in zip(names.split(rh.FIELD_SEP), texts.split(rh.FIELD_SEP)):
            self.fields[n.split("{{")[0]] = [t] * len(self.fields[n.split("{{")[0]])


def test_fields_list_get_put():
    h = FakeFieldHwp({"성명": ["홍길동"], "날짜": ["", ""]})
    assert rh.list_fields(h, 1) == ["성명", "날짜{{0}}", "날짜{{1}}"]
    assert rh.get_fields(h, ["성명"]) == {"성명": "홍길동"}
    assert rh.get_fields(h, None) == {"성명": "홍길동", "날짜{{0}}": "", "날짜{{1}}": ""}
    ok, msg, missing = rh.put_fields(h, {"성명": "이몽룡", "없는칸": "x", "날짜{{1}}": 2026})
    assert ok and missing == ["없는칸"]
    assert h.put_calls == [("성명\x02날짜{{1}}", "이몽룡\x022026")]


def test_fields_put_empty():
    ok, _, _ = rh.put_fields(FakeFieldHwp({}), {})
    assert not ok


# ---------------------------------------------------------------- 쪽 글자·구조 요약

def test_get_page_text_uses_zero_based_page():
    h = MagicMock()
    h.GetPageText.return_value = "2쪽 글자"
    assert rh.get_page_text(h, 2) == "2쪽 글자"
    h.GetPageText.assert_called_once_with(1, 0xFFFFFFFF)


class FakeCtrl:
    def __init__(self, cid, page, desc="", rows=None, cols=None):
        self.CtrlID, self.UserDesc, self.page = cid, desc, page
        self.Next = None
        self.Properties = MagicMock()
        self.Properties.Item.side_effect = lambda k: {"Rows": rows, "Cols": cols}[k]

    def GetAnchorPos(self, _):
        return self


class FakeScanHwp:
    """컨트롤 목록과 InitScan/GetText 결과를 흉내."""

    def __init__(self, ctrls, scan, para_shapes):
        for a, b in zip(ctrls, ctrls[1:]):
            a.Next = b
        self.HeadCtrl = ctrls[0] if ctrls else None
        self.scan = list(scan)
        self.para_shapes = list(para_shapes)  # 문단 시작마다 하나씩 (heading_type, level, number)
        self.cur_shape = (0, 0, "")
        self.PageCount = 3
        self.page = 1
        self.HAction = MagicMock()
        self.released = False
        self.restored = None
        owner = self

        class Info:
            @property
            def CurrentPage(self):
                return owner.page - 1

        info = Info()
        self.XHwpDocuments = SimpleNamespace(
            Active_XHwpDocument=SimpleNamespace(XHwpDocumentInfo=info))

    def GetPos(self):
        return (0, 5, 7)

    def SetPos(self, *a):
        self.restored = a

    def SetPosBySet(self, ctrl):
        self.page = ctrl.page

    printed_map = {}  # 순번 → 인쇄 번호. 없으면 순번 그대로

    def KeyIndicator(self):
        return (True, 1, 1, self.printed_map.get(self.page, self.page), 1, 1, 1, 0, "")

    def InitScan(self, *_):
        return True

    def ReleaseScan(self):
        self.released = True

    def GetText(self):
        return self.scan.pop(0) if self.scan else (1, "")

    def MovePos(self, *_):
        self.cur_shape = self.para_shapes.pop(0)

    @property
    def ParaShape(self):
        ht, lv, _ = self.cur_shape
        ps = MagicMock()
        ps.Item.side_effect = lambda k: {"HeadingType": ht, "Level": lv}[k]
        return ps

    def GetHeadingString(self):
        return self.cur_shape[2]


def test_scan_document_summary():
    ctrls = [FakeCtrl("secd", 1), FakeCtrl("tbl", 1, rows=3, cols=4),
             FakeCtrl("gso", 2, desc="그림"), FakeCtrl("tbl", 3, rows=2, cols=2)]
    # 실측 형태: 3은 "다음 문단 시작 + 그 문단 글자"
    scan = [
        (2, "1. 개요\r\n"),                 # 개요 1
        (3, "본문\r\n"),                     # 일반 문단
        (4, ""), (2, "표 안 글자"), (3, "표 안 둘째"), (5, ""),  # 개체 안은 건너뜀
        (3, "가. 세부"), (2, " 이어짐\r\n"),   # 개요 2 (한 문단이 두 조각)
        (1, ""),
    ]
    shapes = [(1, 0, "1."), (0, 0, ""), (1, 1, "가.")]
    h = FakeScanHwp(ctrls, scan, shapes)
    r = rh.scan_document(h)
    assert r["pages"] == 3
    assert r["table_count"] == 2
    assert r["tables"] == [{"index": 1, "page": 1, "printed_page": 1},
                           {"index": 2, "page": 3, "printed_page": 3}]
    assert r["pictures"] == [{"index": 1, "page": 2, "printed_page": 2}]
    assert r["page_numbering"] == [{"pages": "1-3", "printed": "1-3"}]
    assert r["other_controls"] == {"secd": 1}
    assert r["headings"] == [{"level": 1, "number": "1.", "text": "1. 개요"},
                             {"level": 2, "number": "가.", "text": "가. 세부 이어짐"}]
    assert h.released
    assert h.restored == (0, 5, 7)


# ---------------------------------------------------------------- 저장

def test_save_format_for():
    assert HwpController.save_format_for("a.pdf") == "PDF"
    assert HwpController.save_format_for("a.HWPX") == "HWPX"
    assert HwpController.save_format_for("a.hwp") == "HWP"
    assert HwpController.save_format_for("a") == "HWP"


def _controller(path):
    c = HwpController.__new__(HwpController)
    c.hwp = MagicMock()
    c.hwp.XHwpDocuments.Active_XHwpDocument.FullName = path
    c.is_hwp_running = True
    c.current_document_path = None
    return c


def test_save_current_document_new_doc_refuses():
    c = _controller("")
    c.hwp.Path = ""
    ok, msg = c.save_current_document()
    assert not ok and "path" in msg
    c.hwp.Save.assert_not_called()


def test_save_current_document_saves_in_place():
    c = _controller(r"C:\doc\a.hwp")
    c.hwp.Save.return_value = True
    ok, path = c.save_current_document()
    assert ok and path == r"C:\doc\a.hwp"
    c.hwp.Save.assert_called_once_with(False)


def test_save_document_pdf_keeps_current_path(tmp_path):
    c = _controller(r"C:\doc\a.hwp")
    c.current_document_path = r"C:\doc\a.hwp"
    c.hwp.SaveAs.return_value = True
    assert c.save_document(str(tmp_path / "out.pdf"))
    args = c.hwp.SaveAs.call_args[0]
    assert args[1] == "PDF"
    assert c.current_document_path == r"C:\doc\a.hwp"


def test_scan_document_page_numbering_restart():
    """표지 2쪽 뒤 3쪽에서 '새 번호로 시작(1)' → 순번 3-20 = 인쇄 1-18 (실제 문서 실측 형태)."""
    ctrls = [FakeCtrl("secd", 1), FakeCtrl("tbl", 1, rows=1, cols=1),
             FakeCtrl("nwno", 3), FakeCtrl("nwno", 3), FakeCtrl("tbl", 10, rows=1, cols=1)]
    h = FakeScanHwp(ctrls, [(1, "")], [])
    h.PageCount = 20
    h.printed_map = {p: p - 2 for p in range(3, 21)}
    r = rh.scan_document(h, include_headings=False)
    assert r["page_numbering"] == [{"pages": "1-2", "printed": "1-2"},
                                   {"pages": "3-20", "printed": "1-18"}]
    assert r["tables"][1] == {"index": 2, "page": 10, "printed_page": 8}
    assert r["other_controls"] == {"secd": 1, "nwno": 2}
    assert h.restored == (0, 5, 7)


def test_page_numbering_single_page_and_unknown():
    assert rh._page_numbering([(1, 1), (2, 5)], 2) == [{"pages": "1", "printed": "1"},
                                                       {"pages": "2", "printed": "5"}]
    assert rh._page_numbering([(None, None)], 3) == []
