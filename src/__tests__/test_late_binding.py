# -*- coding: utf-8 -*-
"""한글 COM 연결이 gen_py 캐시와 관계없이 늦은 바인딩으로 고정되는지 확인."""

from types import SimpleNamespace
from unittest.mock import patch

from src.tools import hwp_controller as hc


def test_late_bound_unwraps_oleobj():
    raw = object()
    wrapped = SimpleNamespace(_oleobj_=raw)  # 이른 바인딩 객체처럼 _oleobj_를 가진 객체
    with patch.object(hc.win32com.client.dynamic, "Dispatch", return_value="LATE") as d:
        assert hc._late_bound(wrapped) == "LATE"
    d.assert_called_once_with(raw)


def test_late_bound_accepts_raw_idispatch():
    raw = object()  # QueryInterface로 얻은 PyIDispatch처럼 _oleobj_가 없는 객체
    with patch.object(hc.win32com.client.dynamic, "Dispatch", return_value="LATE") as d:
        assert hc._late_bound(raw) == "LATE"
    d.assert_called_once_with(raw)


def test_new_and_active_objects_use_dynamic_dispatch():
    with patch.object(hc.win32com.client.dynamic, "Dispatch", return_value="LATE") as d:
        assert hc._new_hwp_object() == "LATE"
        d.assert_called_once_with(hc.HWP_PROGID)

    active = SimpleNamespace(_oleobj_="RAW")
    with patch.object(hc.win32com.client, "GetActiveObject", return_value=active) as g, \
            patch.object(hc.win32com.client.dynamic, "Dispatch", return_value="LATE") as d:
        assert hc._active_hwp_object() == "LATE"
    g.assert_called_once_with(hc.HWP_PROGID)
    d.assert_called_once_with("RAW")
