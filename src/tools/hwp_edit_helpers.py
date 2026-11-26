"""
찾은 글자 기준 삽입·서식 복사용 COM 헬퍼.

모든 함수는 한글 COM 객체(HWPFrame.HwpObject)를 직접 받습니다.
서버 연결과 무관하게 숨김 인스턴스로 시험할 수 있도록 컨트롤러에 의존하지 않습니다.
"""

INSERT_POSITIONS = ("para_before", "para_after", "line_after", "after_match", "before_match")

_LANGS = ("Hangul", "Latin", "Hanja", "Japanese", "Other", "Symbol", "User")

# 복사할 글자 모양 항목(소스에 있는 것만 복사)
CHAR_ITEMS = (
    ["Height", "Bold", "Italic", "TextColor", "ShadeColor",
     "UnderlineType", "UnderlineShape", "UnderlineColor",
     "StrikeOutType", "StrikeOutShape", "StrikeOutColor",
     "SuperScript", "SubScript", "OutlineType", "ShadowType", "Emboss", "Engrave", "SmallCaps"]
    + [f"{k}{lang}" for lang in _LANGS
       for k in ("FaceName", "FontType", "Size", "Ratio", "Spacing", "Offset")]
)

# 복사할 문단 모양 항목
PARA_ITEMS = (
    "AlignType", "LeftMargin", "RightMargin", "Indentation",
    "LineSpacing", "LineSpacingType", "PrevSpacing", "NextSpacing",
    "HeadingType", "Level",
)

# 글머리표 하위 항목(HeadingType=3일 때 글머리 문자 등)
BULLET_ITEMS = ("BulletChar", "HasCharShape", "HasImage", "Checked")


def find_anchor(h, anchor: str, keep_selection: bool = False) -> bool:
    """문서 처음부터 anchor를 찾아 커서를 옮깁니다. 비면 현재 위치를 그대로 씁니다.

    keep_selection=False면 선택을 풀고 커서를 찾은 글자 끝에 둡니다.
    """
    if not anchor:
        return True
    h.HAction.Run("MoveDocBegin")
    pset = h.HParameterSet.HFindReplace
    h.HAction.GetDefault("RepeatFind", pset.HSet)
    pset.FindString = anchor
    pset.FindRegExp = 0
    pset.IgnoreMessage = 1
    pset.Direction = 0
    if not h.HAction.Execute("RepeatFind", pset.HSet):
        return False
    if not keep_selection:
        h.HAction.Run("Cancel")
    return True


def _insert_plain(h, text: str):
    if not text:
        return
    pset = h.HParameterSet.HInsertText
    h.HAction.GetDefault("InsertText", pset.HSet)
    pset.Text = text
    h.HAction.Execute("InsertText", pset.HSet)


def _insert_lines(h, text: str, break_action: str):
    """줄바꿈(\\n)마다 break_action(BreakPara/BreakLine)을 넣으며 글자를 넣습니다."""
    lines = text.replace("\r\n", "\n").split("\n")
    for i, line in enumerate(lines):
        if i:
            h.HAction.Run(break_action)
        _insert_plain(h, line)


def read_shapes(h) -> dict:
    """현재 커서(또는 선택 영역)의 글자 모양과 문단 모양 값을 읽습니다."""
    cs = h.CharShape
    ps = h.ParaShape
    char = {n: cs.Item(n) for n in CHAR_ITEMS if cs.ItemExist(n)}
    para = {n: ps.Item(n) for n in PARA_ITEMS if ps.ItemExist(n)}
    bullet = None
    if ps.ItemExist("Bullet"):
        try:
            b = ps.Item("Bullet")
            bullet = {n: b.Item(n) for n in BULLET_ITEMS if b.ItemExist(n)}
        except Exception:
            bullet = None
    return {"char": char, "para": para, "bullet": bullet}


def read_shapes_at_anchor(h, anchor: str):
    """anchor 글자가 있는 문단의 서식을 읽습니다.

    글자 모양은 찾은 글자(anchor)를 선택한 상태에서 읽고, 문단 모양은 그 문단 것을 씁니다.
    anchor가 비면 현재 커서 위치 기준. 못 찾으면 None.
    """
    if anchor:
        if not find_anchor(h, anchor, keep_selection=True):
            return None
        shapes = read_shapes(h)
        h.HAction.Run("Cancel")
        return shapes
    return read_shapes(h)


def select_current_paragraph(h):
    h.HAction.Run("Cancel")
    h.HAction.Run("MoveParaBegin")
    h.HAction.Run("MoveSelParaEnd")


def apply_shapes(h, shapes: dict, char: bool = True, para: bool = True) -> bool:
    """현재 커서가 있는 문단 전체에 shapes(read_shapes 결과)를 적용합니다."""
    select_current_paragraph(h)
    ok = True
    try:
        if char and shapes.get("char"):
            p = h.HParameterSet.HCharShape
            h.HAction.GetDefault("CharShape", p.HSet)
            for k, v in shapes["char"].items():
                p.HSet.SetItem(k, v)
            ok = bool(h.HAction.Execute("CharShape", p.HSet)) and ok
        if para and shapes.get("para"):
            p = h.HParameterSet.HParaShape
            h.HAction.GetDefault("ParagraphShape", p.HSet)
            for k, v in shapes["para"].items():
                p.HSet.SetItem(k, v)
            bullet = shapes.get("bullet")
            if bullet and shapes["para"].get("HeadingType") == 3:
                # p.Bullet 하위 세트로 넣어야 글머리 문자가 바뀜(CreateItemSet은 실패)
                b = p.Bullet
                for k, v in bullet.items():
                    try:
                        setattr(b, k, v)
                    except Exception:
                        pass
            ok = bool(h.HAction.Execute("ParagraphShape", p.HSet)) and ok
    finally:
        h.HAction.Run("Cancel")
    return ok


def copy_format(h, from_anchor: str, to_anchor: str, char: bool = True, para: bool = True):
    """from_anchor 문단의 서식을 to_anchor 문단 전체에 복사합니다. (성공여부, 메시지)"""
    shapes = read_shapes_at_anchor(h, from_anchor)
    if shapes is None:
        return False, f"원본 글자를 찾지 못했습니다: {from_anchor}"
    if not find_anchor(h, to_anchor):
        return False, f"대상 글자를 찾지 못했습니다: {to_anchor}"
    if not apply_shapes(h, shapes, char=char, para=para):
        return False, "서식 적용 실패"
    return True, "서식을 복사했습니다."


def insert_at(h, anchor: str, text: str, position: str = "para_after", copy_format_from: str = ""):
    """anchor 기준 위치에 text를 넣습니다. (성공여부, 메시지)

    - para_before / para_after: anchor 문단 앞/뒤에 새 문단. text 안의 \\n은 문단 나눔.
    - line_after: anchor 문단 끝에 줄 바꿈(BreakLine, 같은 문단) 후 text. \\n도 줄 바꿈.
    - after_match / before_match: 찾은 글자 바로 뒤/앞에 끼워 넣음. \\n은 문단 나눔.
    - copy_format_from: 넣은 뒤 이 글자가 있는 문단의 서식을 새 문단(넣은 글자가 있는 문단)에 적용.
    """
    if position not in INSERT_POSITIONS:
        return False, f"position은 {list(INSERT_POSITIONS)} 중 하나여야 합니다."
    if not text:
        return False, "text가 비어 있습니다."

    shapes = None
    if copy_format_from:
        shapes = read_shapes_at_anchor(h, copy_format_from)
        if shapes is None:
            return False, f"서식 원본 글자를 찾지 못했습니다: {copy_format_from}"

    if position == "before_match":
        if not find_anchor(h, anchor, keep_selection=True):
            return False, f"기준 글자를 찾지 못했습니다: {anchor}"
        if anchor:
            sel = h.GetSelectedPos()  # (ok, slist, spara, spos, elist, epara, epos)
            h.HAction.Run("Cancel")
            if sel and sel[0]:
                h.SetPos(sel[1], sel[2], sel[3])
            else:
                for _ in range(len(anchor)):
                    h.HAction.Run("MoveLeft")
    elif not find_anchor(h, anchor):
        return False, f"기준 글자를 찾지 못했습니다: {anchor}"

    if position == "para_before":
        h.HAction.Run("MoveParaBegin")
        start_pos = h.GetPos()
        _insert_lines(h, text, "BreakPara")
        end_pos = h.GetPos()
        h.HAction.Run("BreakPara")
    elif position == "para_after":
        h.HAction.Run("MoveParaEnd")
        h.HAction.Run("BreakPara")
        start_pos = h.GetPos()
        _insert_lines(h, text, "BreakPara")
        end_pos = h.GetPos()
    elif position == "line_after":
        h.HAction.Run("MoveParaEnd")
        start_pos = h.GetPos()
        h.HAction.Run("BreakLine")
        _insert_lines(h, text, "BreakLine")
        end_pos = h.GetPos()
    else:  # after_match / before_match
        start_pos = h.GetPos()
        _insert_lines(h, text, "BreakPara")
        end_pos = h.GetPos()

    msg = f"{position} 위치에 글자를 넣었습니다."
    if shapes is not None:
        n_paras = 1 if position == "line_after" else text.replace("\r\n", "\n").count("\n") + 1
        h.SetPos(*start_pos)
        ok = True
        for i in range(n_paras):
            if i:
                h.HAction.Run("MoveNextParaBegin")
            ok = apply_shapes(h, shapes) and ok
        msg += " 서식을 복사했습니다." if ok else " (서식 복사 실패)"
        h.SetPos(*end_pos)
    return True, msg


# ---------------------------------------------------------------- 그림 넣기·바꾸기

PICTURE_POSITIONS = ("para_before", "para_after", "at_cursor")

# 1인치 = 7200 HWPUNIT = 25.4mm → 1mm ≈ 283.46 HWPUNIT
HWPUNIT_PER_MM = 7200 / 25.4


def mm_to_hwpunit(mm: float) -> int:
    return int(round(mm * HWPUNIT_PER_MM))


def hwpunit_to_mm(u: int) -> float:
    return round(u / HWPUNIT_PER_MM, 1)


def _is_picture(ctrl) -> bool:
    """그리기 개체(gso) 중 그림인지. UserDesc는 한글 UI 언어에 따라 '그림'/'Picture'."""
    if ctrl.CtrlID != "gso":
        return False
    desc = str(ctrl.UserDesc or "")
    return desc == "그림" or "picture" in desc.lower()


def _set_picture_props(ctrl, size=None) -> bool:
    """그림을 글자처럼 취급하고, size(dict: Width/Height/WidthRelTo/HeightRelTo)가 있으면 크기를 맞춥니다.

    크기는 ShapeObject의 Width/Height(WidthRelTo=4, HeightRelTo=2일 때 HWPUNIT)로 넣습니다.
    """
    prop = ctrl.Properties
    prop.SetItem("TreatAsChar", 1)
    if size:
        for k in ("WidthRelTo", "HeightRelTo", "Width", "Height"):
            if k in size:
                prop.SetItem(k, size[k])
    ctrl.Properties = prop
    return True


def read_picture_size(ctrl) -> dict:
    """그림 컨트롤의 크기 항목(WidthRelTo, Width, HeightRelTo, Height)을 읽습니다."""
    prop = ctrl.Properties
    return {k: prop.Item(k) for k in ("WidthRelTo", "Width", "HeightRelTo", "Height") if prop.ItemExist(k)}


def _describe_size(size: dict) -> str:
    w, h = size.get("Width"), size.get("Height")
    if size.get("WidthRelTo", 4) == 4 and size.get("HeightRelTo", 2) == 2 and w is not None and h is not None:
        return f"{hwpunit_to_mm(w)}x{hwpunit_to_mm(h)}mm ({w}x{h} HWPUNIT)"
    return f"Width={w}(WidthRelTo={size.get('WidthRelTo')}), Height={h}(HeightRelTo={size.get('HeightRelTo')})"


def _insert_picture_here(h, abs_path: str, size=None, width_mm: float = 0, height_mm: float = 0):
    """현재 커서 위치에 그림을 문서에 포함(embed)해 넣고 글자처럼 취급합니다. 넣은 Ctrl을 돌려줍니다.

    한글 InsertPicture 메서드의 width/height 단위는 mm(정수). 그래서 원래 크기(realSize=0)로 넣은 뒤
    ShapeObject의 Width/Height(HWPUNIT)를 고쳐 정확한 크기로 맞춥니다.
    - size: 그대로 쓸 크기 항목(read_picture_size 결과). 있으면 width_mm/height_mm 무시.
    - width_mm/height_mm: 0이면 원래 크기. 한쪽만 주면 원래 가로세로 비율 유지.
    """
    # 선택 인자도 8개 모두 넘겨야 함(생략하면 'Invalid number of parameters')
    ctrl = h.InsertPicture(abs_path, True, 0, False, False, 0, 0, 0)
    if ctrl is None:
        return None
    if size is None and (width_mm > 0 or height_mm > 0):
        orig = read_picture_size(ctrl)
        ow, oh = orig.get("Width", 0), orig.get("Height", 0)
        w = mm_to_hwpunit(width_mm) if width_mm > 0 else 0
        hh = mm_to_hwpunit(height_mm) if height_mm > 0 else 0
        if not w and oh:
            w = int(round(ow * hh / oh))
        if not hh and ow:
            hh = int(round(oh * w / ow))
        size = {"WidthRelTo": 4, "Width": w, "HeightRelTo": 2, "Height": hh}
    _set_picture_props(ctrl, size)
    return ctrl


def insert_picture(h, image_path: str, width_mm: float = 0, height_mm: float = 0,
                   anchor: str = "", position: str = "para_after"):
    """anchor 기준 위치에 그림을 넣습니다(문서에 포함, 글자처럼 취급). (성공여부, 메시지)

    - para_before / para_after: anchor 문단 앞/뒤에 그림만 있는 새 문단.
    - at_cursor: 현재 커서 위치(anchor가 있으면 찾은 글자 바로 뒤)에 끼워 넣음.
    """
    import os
    if position not in PICTURE_POSITIONS:
        return False, f"position은 {list(PICTURE_POSITIONS)} 중 하나여야 합니다."
    if width_mm < 0 or height_mm < 0:
        return False, "width_mm, height_mm는 0 이상이어야 합니다."
    abs_path = os.path.abspath(image_path)
    if not os.path.exists(abs_path):
        return False, f"그림 파일을 찾을 수 없습니다: {abs_path}"
    if not find_anchor(h, anchor):
        return False, f"기준 글자를 찾지 못했습니다: {anchor}"

    if position == "para_before":
        h.HAction.Run("MoveParaBegin")
        ctrl = _insert_picture_here(h, abs_path, width_mm=width_mm, height_mm=height_mm)
        if ctrl is not None:
            h.HAction.Run("BreakPara")
    elif position == "para_after":
        h.HAction.Run("MoveParaEnd")
        h.HAction.Run("BreakPara")
        ctrl = _insert_picture_here(h, abs_path, width_mm=width_mm, height_mm=height_mm)
    else:
        ctrl = _insert_picture_here(h, abs_path, width_mm=width_mm, height_mm=height_mm)

    if ctrl is None:
        return False, "그림 넣기 실패(InsertPicture가 개체를 돌려주지 않음)"
    return True, f"{position} 위치에 그림을 넣었습니다. 크기: {_describe_size(read_picture_size(ctrl))}"


def _picture_at(h, pos):
    """anchor 위치가 pos(list, para, pos) 또는 그 바로 앞인 그림 컨트롤."""
    ctrl = h.HeadCtrl
    while ctrl is not None:
        if _is_picture(ctrl):
            ap = ctrl.GetAnchorPos(0)
            if ap is not None and (ap.Item("List"), ap.Item("Para")) == tuple(pos[:2])                     and ap.Item("Pos") in (pos[2], pos[2] - 1):
                return ctrl
        ctrl = ctrl.Next
    return None


def find_picture_near(h, near_text: str, which: str = "after"):
    """near_text 기준으로 그림 컨트롤을 찾습니다.

    - near_text가 그림의 캡션 안에 있으면 그 캡션의 그림(which 무시).
    - 아니면 near_text 문단 바로 뒤(after)의 첫 그림 또는 앞(before)의 가장 가까운 그림.
      같은 리스트(본문·같은 셀 등) 안의 그림만 봅니다.
    near_text는 문서 처음부터 첫 번째로 찾은 곳.
    (ctrl, 설명 문자열) 또는 (None, 이유 문자열)을 돌려줍니다.
    """
    if not find_anchor(h, near_text):
        return None, f"기준 글자를 찾지 못했습니다: {near_text}"
    cur = h.GetPos()
    if cur[0] != 0:
        # 캡션(그림 개체 안의 리스트)이면 CloseEx로 나왔을 때 커서가 그림 바로 뒤(또는 앞)에 옴
        h.HAction.Run("CloseEx")
        owner = _picture_at(h, h.GetPos())
        if owner is not None:
            return owner, "캡션이 붙은 그림"
        h.SetPos(*cur)
    a_list, a_para, _ = cur
    best, best_pos = None, None
    ctrl = h.HeadCtrl
    while ctrl is not None:
        if _is_picture(ctrl):
            ap = ctrl.GetAnchorPos(0)
            if ap is not None:
                pos = (ap.Item("List"), ap.Item("Para"), ap.Item("Pos"))
                if pos[0] == a_list:
                    if which == "after" and pos[1] > a_para:
                        if best_pos is None or pos[1:] < best_pos[1:]:
                            best, best_pos = ctrl, pos
                    elif which == "before" and pos[1] < a_para:
                        if best_pos is None or pos[1:] > best_pos[1:]:
                            best, best_pos = ctrl, pos
        ctrl = ctrl.Next
    if best is None:
        return None, f"'{near_text}' 문단 {'뒤' if which == 'after' else '앞'}에서 그림을 찾지 못했습니다."
    gap = abs(best_pos[1] - a_para)
    return best, f"기준 문단에서 {gap}문단 {'뒤' if which == 'after' else '앞'} 그림"


def replace_picture(h, image_path: str, near_text: str, which: str = "after"):
    """그림 개체는 그대로 두고 그림 파일만 바꿉니다(한글의 '그림 바꾸기', PictureChange 액션). (성공여부, 메시지)

    캡션(자동 번호 포함)·위치·글자처럼 취급 등 개체 속성은 그대로 남습니다.
    바꾼 뒤 너비·높이가 달라졌으면 이전 값으로 되돌립니다.
    """
    import os
    if which not in ("after", "before"):
        return False, "which는 'after' 또는 'before'여야 합니다."
    abs_path = os.path.abspath(image_path)
    if not os.path.exists(abs_path):
        return False, f"그림 파일을 찾을 수 없습니다: {abs_path}"
    pic, where = find_picture_near(h, near_text, which)
    if pic is None:
        return False, where
    old_size = read_picture_size(pic)

    # 그림 개체 선택: anchor 위치로 가서 FindCtrl
    h.SetPosBySet(pic.GetAnchorPos(0))
    if h.FindCtrl() != "gso":
        return False, "그림 개체를 선택하지 못했습니다."
    act = h.CreateAction("PictureChange")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("PicturePath", abs_path)
    pset.SetItem("PictureEmbed", 1)
    ok = act.Execute(pset)
    h.HAction.Run("Cancel")
    if not ok:
        return False, "그림 바꾸기(PictureChange) 실패"

    new_size = read_picture_size(pic)
    fixed = ""
    if any(new_size.get(k) != v for k, v in old_size.items()):
        prop = pic.Properties
        for k, v in old_size.items():
            prop.SetItem(k, v)
        pic.Properties = prop
        fixed = f" (바꾼 직후 {_describe_size(new_size)}여서 이전 크기로 되돌림)"
        new_size = read_picture_size(pic)
    return True, (f"그림을 바꿨습니다({where}). 캡션·위치는 그대로입니다. "
                  f"이전 크기: {_describe_size(old_size)} → 새 크기: {_describe_size(new_size)}{fixed}")


def _read_list_text(h, text_reader) -> str:
    """커서가 있는 리스트(캡션·셀 등) 전체를 선택해 text_reader()로 읽고 선택을 풉니다."""
    h.HAction.Run("MoveListBegin")
    h.HAction.Run("MoveSelListEnd")
    try:
        return (text_reader() or "").strip()
    finally:
        h.HAction.Run("Cancel")


def delete_picture(h, near_text: str, which: str = "after", text_reader=None):
    """near_text 기준으로 찾은 그림 개체를 지웁니다(캡션도 함께 지워짐). (성공여부, 메시지)

    그림 찾는 방법은 replace_picture와 같음(find_picture_near).
    text_reader: 선택 영역 글자를 돌려주는 함수(컨트롤러의 _get_text_file_safe("TEXT", "saveblock") 감싼 것).
      주면 near_text가 캡션 안일 때 캡션 글자를 읽어 결과에 넣습니다.
      GetTextFile을 직접 부르지 않음(변경 추적 문서에서 저장 확인 대화창이 뜨기 때문).
    """
    if which not in ("after", "before"):
        return False, "which는 'after' 또는 'before'여야 합니다."
    if not near_text:
        return False, "near_text가 비어 있습니다."
    pic, where = find_picture_near(h, near_text, which)
    if pic is None:
        return False, where
    size = read_picture_size(pic)
    caption = ""
    if where == "캡션이 붙은 그림" and text_reader is not None:
        try:
            if find_anchor(h, near_text):
                caption = _read_list_text(h, text_reader)
        except Exception:
            caption = ""
    if not h.DeleteCtrl(pic):
        return False, "그림 지우기(DeleteCtrl) 실패"
    cap = f" 캡션: \"{caption}\"." if caption else (" 캡션도 함께 지웠습니다." if where == "캡션이 붙은 그림" else "")
    return True, f"그림을 지웠습니다({where}). 크기: {_describe_size(size)}.{cap}"


def _fill_current_cell(h, value: str):
    """커서가 있는 셀의 내용을 value로 바꿉니다. 셀 전체를 선택한 뒤 InsertText로 덮어써서
    셀의 글자 모양을 그대로 씁니다. 줄바꿈(\n)은 문단 나눔(BreakPara)으로 넣습니다."""
    h.HAction.Run("MoveListBegin")
    h.HAction.Run("MoveSelListEnd")
    lines = value.replace("\r\n", "\n").split("\n")
    if lines[0] == "":
        # 첫 줄이 비면 InsertText로 선택을 덮어쓸 수 없으므로 선택 영역을 먼저 지움
        h.HAction.Run("Delete")
    for i, line in enumerate(lines):
        if i:
            h.HAction.Run("BreakPara")
        _insert_plain(h, line)
    h.HAction.Run("Cancel")


def fill_table_grid(h, anchor: str, rows, skip_empty: bool = True):
    """anchor가 든 표의 왼쪽 위 셀부터 행 우선(TableRightCell 순서)으로 rows 값을 채웁니다. (성공여부, 메시지)

    - rows: [[1행 값...], [2행 값...]]. 펼쳐서 셀 순서대로 대응(병합 셀은 한 칸).
    - None(또는 skip_empty일 때 "")은 그 셀을 건드리지 않음.
    - 마지막 셀에서 TableRightCell을 부르면 행이 늘어나므로, 미리 마지막 셀의 리스트 ID를 알아 두고
      거기서 멈춥니다. 남은 값은 버리고 메시지에 알림.
    """
    if not isinstance(rows, (list, tuple)) or any(not isinstance(r, (list, tuple)) for r in rows):
        return False, "rows는 리스트의 리스트여야 합니다. 예: [[\"a\", \"b\"], [null, \"c\"]]"

    def needs(v):
        return v is not None and not (skip_empty and v == "")

    cells = [v for r in rows for v in r]
    targets = [i for i, v in enumerate(cells) if needs(v)]
    if not targets:
        return True, "채울 값이 없습니다(모두 None 또는 빈 문자열). 0칸을 썼습니다."
    last_idx = targets[-1]

    if not find_anchor(h, anchor):
        return False, f"기준 글자를 찾지 못했습니다: {anchor}"
    if h.GetPos()[0] == 0 or not h.HAction.Run("TableColBegin"):
        return False, f"'{anchor}'가 표 안에 있지 않습니다."
    # 마지막 셀(오른쪽 아래)의 리스트 ID
    h.HAction.Run("TableColEnd")
    h.HAction.Run("TableColPageDown")
    last_list = h.GetPos()[0]
    # 왼쪽 위 셀로
    h.HAction.Run("TableColBegin")
    h.HAction.Run("TableColPageUp")

    seen = {h.GetPos()[0]}
    written = 0
    stopped = ""
    for i in range(last_idx + 1):
        if i:
            if h.GetPos()[0] == last_list:
                left = sum(1 for k in targets if k >= i)
                stopped = f" 표의 마지막 셀에 닿아 나머지 값 {left}개는 쓰지 않았습니다."
                break
            h.HAction.Run("TableRightCell")
            cur = h.GetPos()[0]
            if cur in seen:
                stopped = " 셀 이동이 더 진행되지 않아 멈췄습니다."
                break
            seen.add(cur)
        v = cells[i]
        if needs(v):
            _fill_current_cell(h, v if isinstance(v, str) else str(v))
            written += 1
    return True, f"표 셀 {written}칸을 채웠습니다.{stopped}"


# ---------------------------------------------------------------- 표·문단 복제

DUPLICATE_UNITS = ("table", "paragraphs")
DUPLICATE_POSITIONS = ("para_after", "para_before")


def _clipboard_text():
    """클립보드의 글자(CF_UNICODETEXT)를 읽습니다. 없거나 실패하면 None."""
    try:
        import win32clipboard
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            return None
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        return None


def _restore_clipboard_text(text):
    """작업 전에 읽어 둔 글자를 클립보드에 되돌려 놓습니다(글자 외 형식은 복원하지 않음)."""
    try:
        import win32clipboard
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            if text is not None:
                win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        pass


def find_nth(h, text: str, n: int = 1) -> bool:
    """문서 처음부터 text의 n번째(1부터) 일치로 커서를 옮깁니다(선택 해제, 글자 끝). 없으면 False."""
    h.HAction.Run("Cancel")
    h.HAction.Run("MoveDocBegin")
    pset = h.HParameterSet.HFindReplace
    last = None
    for _ in range(max(1, n)):
        h.HAction.GetDefault("RepeatFind", pset.HSet)
        pset.FindString = text
        pset.FindRegExp = 0
        pset.IgnoreMessage = 1
        pset.Direction = 0
        if not h.HAction.Execute("RepeatFind", pset.HSet):
            return False
        h.HAction.Run("Cancel")
        cur = tuple(h.GetPos())
        if last is not None and cur <= last:  # 문서 처음으로 되돌아감
            return False
        last = cur
    return True


def _count_in_text(text_reader, s: str):
    if not s or text_reader is None:
        return None
    try:
        return (text_reader() or "").count(s)
    except Exception:
        return None


def _select_table_at_anchor(h, anchor: str, occurrence: int = 1):
    """anchor가 든 표 전체를 개체로 선택합니다. (성공여부, 메시지)"""
    if not find_nth(h, anchor, occurrence):
        return False, f"복사할 글자를 찾지 못했습니다: {anchor}"
    if h.GetPos()[0] == 0:
        return False, f"'{anchor}'가 표 안에 있지 않습니다."
    try:
        ctrl = h.ParentCtrl
    except Exception:
        ctrl = None
    if ctrl is None or getattr(ctrl, "CtrlID", "") != "tbl":
        return False, f"'{anchor}'가 든 표를 찾지 못했습니다(표 안 글자가 아니거나 글상자 안일 수 있음)."
    h.HAction.Run("Cancel")
    h.SetPosBySet(ctrl.GetAnchorPos(0))
    if not h.FindCtrl():
        return False, "표를 선택하지 못했습니다."
    return True, "표"


def _select_paragraphs(h, anchor: str, end_anchor: str, occurrence: int = 1):
    """anchor 문단부터 end_anchor 문단까지(뒤 문단 나눔 포함)를 선택합니다. (성공여부, 메시지)

    end_anchor는 anchor 문단 이후에서 처음 나오는 곳.
    """
    if not find_nth(h, anchor, occurrence):
        return False, f"복사할 글자를 찾지 못했습니다: {anchor}"
    h.HAction.Run("MoveParaBegin")
    s_list, s_para, _ = h.GetPos()
    if s_list != 0:
        return False, f"'{anchor}'가 본문이 아닌 곳(표·글상자 등)에 있습니다. 문단 복사는 본문에서만 됩니다."
    e_para = s_para
    if end_anchor:
        # anchor 문단 처음부터 앞으로 찾기
        h.SetPos(0, s_para, 0)
        pset = h.HParameterSet.HFindReplace
        h.HAction.GetDefault("RepeatFind", pset.HSet)
        pset.FindString = end_anchor
        pset.FindRegExp = 0
        pset.IgnoreMessage = 1
        pset.Direction = 0
        if not h.HAction.Execute("RepeatFind", pset.HSet):
            return False, f"끝 문단 글자를 찾지 못했습니다: {end_anchor}"
        h.HAction.Run("Cancel")
        e_list, e_para, _ = h.GetPos()
        if e_list != 0:
            return False, f"'{end_anchor}'가 본문이 아닌 곳(표·글상자 등)에 있습니다."
        if e_para < s_para:
            return False, f"끝 문단('{end_anchor}')이 시작 문단('{anchor}')보다 앞에 있습니다."
    # 끝 문단 다음 문단의 처음까지 선택해 문단 나눔까지 포함
    h.SetPos(0, e_para, 0)
    h.HAction.Run("MoveParaEnd")
    end_pos = h.GetPos()
    if h.HAction.Run("MoveNextParaBegin") and h.GetPos()[1] > e_para:
        end_pos = h.GetPos()
    if not h.SelectText(s_para, 0, end_pos[1], end_pos[2]):
        return False, "문단을 선택하지 못했습니다."
    return True, f"문단 {e_para - s_para + 1}개"


def duplicate(h, unit: str, anchor: str, target: str, position: str = "para_after",
              end_anchor: str = "", text_reader=None, occurrence: int = 0, target_occurrence: int = 0):
    """표 하나 또는 문단 범위를 서식 그대로 복사해 target 문단 앞/뒤에 붙여넣습니다. (성공여부, 메시지)

    - unit="table": anchor가 든 표 전체.
    - unit="paragraphs": anchor 문단부터 end_anchor 문단까지(비우면 anchor 문단 하나). 사이의 표·그림 포함.
    - anchor/target이 문서에 여러 번 있으면 occurrence/target_occurrence(1부터)가 없을 때 복사하지 않고
      실패(애매함)로 돌려줍니다. end_anchor는 anchor 문단 이후 처음 나오는 곳.
    - 클립보드를 거치며, 끝나면 작업 전 클립보드의 글자를 되돌려 놓습니다.
    """
    if unit not in DUPLICATE_UNITS:
        return False, f"unit은 {list(DUPLICATE_UNITS)} 중 하나여야 합니다."
    if position not in DUPLICATE_POSITIONS:
        return False, f"position은 {list(DUPLICATE_POSITIONS)} 중 하나여야 합니다."
    if not anchor or not target:
        return False, "anchor와 target이 필요합니다."
    if end_anchor and unit != "paragraphs":
        return False, "end_anchor는 unit=\"paragraphs\"일 때만 씁니다."
    for name, txt, occ, occ_name in (("anchor", anchor, occurrence, "occurrence"),
                                     ("target", target, target_occurrence, "target_occurrence")):
        n = _count_in_text(text_reader, txt)
        if n is not None and n > 1 and not occ:
            return False, (f"{name} '{txt}'가 {n}곳에서 일치합니다. 앞뒤 글자를 넉넉히 넣거나 "
                           f"{occ_name}(1부터)로 몇 번째인지 지정하세요.")
        if n is not None and occ and occ > n:
            return False, f"{name} '{txt}'는 {n}곳뿐입니다({occ_name}={occ})."

    saved_clip = _clipboard_text()
    try:
        if unit == "table":
            ok, what = _select_table_at_anchor(h, anchor, occurrence or 1)
        else:
            ok, what = _select_paragraphs(h, anchor, end_anchor, occurrence or 1)
        if not ok:
            h.HAction.Run("Cancel")
            return False, what
        h.HAction.Run("Copy")
        h.HAction.Run("Cancel")

        if not find_nth(h, target, target_occurrence or 1):
            return False, f"붙여넣을 기준 글자를 찾지 못했습니다: {target}"
        if unit == "table":
            if position == "para_after":
                h.HAction.Run("MoveParaEnd")
                h.HAction.Run("BreakPara")
                h.HAction.Run("Paste")
            else:
                h.HAction.Run("MoveParaBegin")
                h.HAction.Run("Paste")
                h.HAction.Run("BreakPara")
        else:
            if position == "para_after":
                cur_para = h.GetPos()[1]
                if h.HAction.Run("MoveNextParaBegin") and h.GetPos()[1] > cur_para:
                    h.HAction.Run("Paste")
                else:  # 마지막 문단 뒤
                    h.HAction.Run("MoveParaEnd")
                    h.HAction.Run("BreakPara")
                    h.HAction.Run("Paste")
            else:
                h.HAction.Run("MoveParaBegin")
                h.HAction.Run("Paste")
        where = "뒤" if position == "para_after" else "앞"
        return True, f"{what}를 복사해 '{target}' 문단 {where}에 붙여넣었습니다."
    finally:
        _restore_clipboard_text(saved_clip)
