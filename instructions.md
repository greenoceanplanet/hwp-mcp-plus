# HWP MCP Server Instructions

> [!IMPORTANT]
> **CRITICAL RULES FOR LLM AGENTS:**
> 1. **Track Changes Mandatory**: Always activate track changes using `hwp_set_track_changes(active=True)` before performing any HWP edit operations (inserting, replacing, editing text or tables).
> 2. **Ask for User Approval**: Immediately after finishing document edits, you must explicitly request the user to review the track changes and select "Accept All Changes (승인)" inside HWP to finalize the text layer.

This file contains detailed instructions, schemas, syntax guidelines, and examples for complex HWP MCP tools.

---

## 1. hwp_table_read

Reads every cell of one table from HWPML. Does not move the cursor or touch the clipboard.

*   `anchor` (+ `occurrence`): the table whose cell text contains this string.
*   `index`: the n-th table in document order (nested tables count). With neither, returns a list of tables.

### Output Format
```json
{
  "index": 1, "row_count": 3, "col_count": 3,
  "rows": [["이름", "소속", "전화"], ["홍길동", "기획팀", "010-1111"]],
  "merged": [{"row": 1, "col": 1, "rowspan": 2, "colspan": 1, "text": "..."}]
}
```
`rows` lists only real cells per row (cells hidden by a merge are left out), in Tab order —
the same shape `hwp_fill_table_grid` takes, so you can edit it and pass it back.

---

## 2. hwp_fill_cells

Batch fills multiple table cells using path instructions relative to labels.

### Usage Protocol
1.  Check the layout first with `hwp_table_read(anchor="LabelName")` (does not move the cursor).
2.  Fill cells using relative directions.

### Path Syntax & Directional Keywords
*   **Keywords**: `<left>`, `<right>`, `<up>`, `<down>`
*   **Separator**: ` > ` or `/`
*   **Example Keys in `path_value_map`**:
    *   `"Name > <right>"`: Finds cell containing "Name" and moves 1 cell right.
    *   `"Total > <down> > <down>"`: Finds cell containing "Total" and moves 2 cells down.
    *   `"Item1 > <right> / <down>"`: Finds "Item1", moves right, then down.

### Parameters
*   `path_value_map` (dict): Dictionary mapping `{ "PathString": "ValueToInsert" }`
*   `mode` (str):
    *   `"replace"`: Deletes existing content and inserts new value (default).
    *   `"prepend"`: Inserts value before existing content.
    *   `"append"`: Inserts value after existing content.

---

## 3. hwp_table_insert / hwp_table_delete

Insert or delete table rows/columns. Turn on track changes first.

*   `anchor` (str): Text inside the reference cell (first match from the top of the document). Empty = current cursor cell.
*   **hwp_table_insert**
    *   `position`: `"below"` / `"above"` (rows), `"right"` / `"left"` (columns).
    *   `count` (int): How many rows/columns to insert.
    *   `data` (list of lists, rows only): Values for the new rows, top to bottom, each row filled from the first column.
*   **hwp_table_delete**
    *   `target`: `"row"` or `"column"`; deletes `count` rows (downward) / columns (rightward) starting at the anchor cell.

### Example
```json
{"anchor": "162,146", "position": "below", "count": 2,
 "data": [["5", "0", "6,535"], ["…", "…", "…"]]}
```
Choose an anchor that is unique in the document (e.g. a number in the row), since the first match is used.

---

## 4. hwp_insert_at

Insert text relative to found text. Turn on track changes first (`hwp_set_track_changes(active=True)`).

*   `anchor` (str): Reference text (first match from the top of the document). Empty = current cursor position.
*   `text` (str): Text to insert. `\n` = paragraph break (in `line_after`, `\n` = line break within the same paragraph).
*   `position` (str):
    *   `"para_before"`: New paragraph before the anchor's paragraph.
    *   `"para_after"`: New paragraph after the anchor's paragraph (default).
    *   `"line_after"`: At the end of the anchor's paragraph, insert a soft line break (Shift+Enter, `BreakLine`) and then the text. Stays in the same paragraph.
    *   `"after_match"` / `"before_match"`: Inline, right after / before the matched text.
*   `copy_format_from` (str, optional): Reference text of a paragraph whose character + paragraph shape (incl. bullet) is applied to the new paragraph(s) after inserting.

A new paragraph inherits the anchor paragraph's shape, **including auto bullets** (`HeadingType`). If the anchor is a bulleted line but the new paragraph should be plain (or should look like some other paragraph), pass `copy_format_from` pointing at a paragraph with the desired look.

### Example
```json
{"anchor": "개발이익 = 입주", "text": "주: 건축비 제외", "position": "para_after",
 "copy_format_from": "주: 원리금 회수 연차"}
```

---

## 5. hwp_copy_format

Copy formatting from one paragraph to another whole paragraph. Turn on track changes first.

*   `from_anchor` (str): Text inside the source paragraph. Character shape is read at this text; paragraph shape from its paragraph. Empty = current cursor.
*   `to_anchor` (str): Text inside the target paragraph; the whole paragraph is formatted. Empty = current cursor paragraph.
*   `char` (bool, default true): Copy character shape (size, font faces per language, bold/italic, color, ratio/spacing, underline, etc.).
*   `para` (bool, default true): Copy paragraph shape (alignment, left/right margin, indentation, line spacing, spacing before/after, `HeadingType` and bullet character).

Both anchors use the first match from the top of the document, so choose unique text.

---

## 6. Reading text

`hwp_get_text`, `hwp_get_selection` and `hwp_get_context` read text with `GetTextFile("UNICODE", ...)`. (The `"TEXT"` format returns CP949 bytes packed into a string, which shows Korean as mojibake such as `ÁÖ°Å¾ÈÁ¤`.) Bullets/numbering added by paragraph shape appear as characters (e.g. `● `) at the start of the line in the output; soft line breaks (`BreakLine`) do not appear as line breaks.

`hwp_get_text(start_text="", max_chars=0)` can read just a section: `start_text` = start at its first occurrence (from the top), `max_chars` = cut after that many characters (0 = to the end). With both empty it returns the whole document as before. Example: `{"start_text": "3. 용산공원", "max_chars": 3000}`.

Any code calling `GetTextFile` must go through the controller's `_get_text_file_safe` (it sets `SetMessageBoxMode(0x00010000)` before and `0` after). On a document with track changes on, a bare `GetTextFile` pops "변경 추적 정보를 유지하려면 한글 형식으로 저장… 저장할까요?".

---

## 7. hwp_page_image

Render one page of the current document to an image exactly as Hangul draws it (uses `CreatePageImage`). Unsaved edits are included.

*   `page` (int, default 0): Page number, 1-based physical page order (independent of printed page numbering). `0` = page where the cursor is.
*   `resolution` (int, default 96): DPI, 24–300. 96 ≈ 794×1123 px for A4; use 150 for small text.
*   `save_path` (str, optional): Output path (`.png` / `.jpg` / `.bmp` / `.gif`). Empty = `%TEMP%\hwp-mcp\page_NNN.png`.

Returns a text line (page/total and saved path) plus the image (except for `.bmp`).
Note: Hangul's `CreatePageImage` itself only writes `bmp`/`gif` and takes a 0-based page index; the tool writes a temporary bmp and converts it with Pillow.

---

## 8. hwp_insert_image / hwp_replace_image / hwp_delete_image

Insert or replace a picture. The picture is **embedded** in the document (not linked) and **treated as a character** (`TreatAsChar`, 글자처럼 취급), so it flows with the text.

*   **hwp_insert_image**
    *   `image_path` (str): Image file (png, jpg, bmp, gif, ...).
    *   `width_mm` / `height_mm` (float, default 0): Size in mm. Both 0 = original size; only one given = keep the original aspect ratio.
    *   `anchor` (str): Reference text (first match from the top of the document). Empty = current cursor position.
    *   `position` (str): `"para_before"` / `"para_after"` (default; a new paragraph holding only the picture, inheriting the anchor paragraph's shape) or `"at_cursor"` (inline at the cursor, i.e. right after the matched text when `anchor` is given).
*   **hwp_replace_image**
    *   Swaps only the image file of an existing picture (Hangul's "그림 바꾸기", `PictureChange` action with `PicturePath` / `PictureEmbed=1`). The picture object stays, so its caption (incl. the auto-number "그림 N"), position, size and TreatAsChar setting are kept. If the size changes anyway, the old `Width`/`Height` are restored. The new image is stretched to the old frame.
    *   `image_path` (str): New image file (embedded).
    *   `near_text` (str): Text used to find the picture (first match). If it is inside a picture's caption (e.g. `"<그림 1> 거주 연수별 누적 순주거비"`), that picture is replaced and `which` is ignored.
    *   `which` (str): For text in an ordinary paragraph: `"after"` (default) = first picture in a later paragraph; `"before"` = nearest picture in an earlier paragraph. Only pictures in the same list are considered.
    *   Returns old and new sizes.

Units: Hangul's `InsertPicture` method takes width/height in whole **mm** (and all 8 arguments must be passed), so `hwp_insert_image` inserts at original size and then sets the exact size through the control's `ShapeObject` properties in **HWPUNIT** (1 inch = 7200, 1 mm ≈ 283.46).
Pictures are found by `CtrlID == "gso"` and `UserDesc == "그림"` (or containing "Picture" in an English UI); confirmed `UserDesc == "그림"` on a Korean Hangul.

*   **hwp_delete_image**
    *   Deletes a picture object (`DeleteCtrl`); its caption goes with it. Turn on track changes first.
    *   `near_text` (str): Found the same way as `hwp_replace_image` (first match). Inside a caption → that picture (`which` ignored). Internally: find → `Cancel` → `CloseEx` leaves the caption list, `GetPos()` gives the body (list, para), then the `gso` control whose `GetAnchorPos(0)` `List`/`Para` match is taken.
    *   `which` (str): `"after"` (default) / `"before"`, as in `hwp_replace_image`.
    *   Returns the deleted picture's size and, when found through the caption, the caption text.

---

## 9. hwp_fill_table_grid

Fill a whole table at once, row-major from the top-left cell, keeping each cell's character format. Turn on track changes first.

*   `anchor` (str): Text inside any cell of the target table (first match from the top).
*   `rows` (list of lists): Values flattened in cell order (the order `Tab`/`TableRightCell` visits; a merged cell counts once), so give every cell of each row. `null` = leave the cell untouched. Numbers are converted to text; `
` becomes a paragraph break inside the cell.
*   `skip_empty` (bool, default true): `""` also leaves the cell untouched. With false, `""` clears the cell.

How it works: `TableColBegin` + `TableColPageUp` → top-left cell; each cell is filled by `MoveListBegin`, `MoveSelListEnd`, then `InsertText` (replaces the selection, keeps the cell's char shape). The bottom-right cell's list ID is read first (`TableColEnd` + `TableColPageDown`) and the walk never calls `TableRightCell` from it (that would append a row); extra values are dropped and reported. It also stops after the last non-skipped value. Returns the number of cells written.

Example:
```json
{"anchor": "구분", "rows": [[null, "2025", "2030"], ["토지임대료", "1,200", "1,350"]]}
```
