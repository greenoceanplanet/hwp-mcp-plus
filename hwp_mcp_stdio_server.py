#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import json
import traceback
import logging
import ssl
from threading import Thread
import time
import builtins

# Monkey-patch print to prevent stdout corruption in stdio MCP mode
original_print = builtins.print
def safe_print(*args, **kwargs):
    if 'file' not in kwargs or kwargs['file'] is sys.stdout or kwargs['file'] is None:
        kwargs['file'] = sys.stderr
    original_print(*args, **kwargs)
builtins.print = safe_print

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    filename=os.path.join(os.path.dirname(os.path.abspath(__file__)), "hwp_mcp_stdio_server.log"),
    filemode="a",
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)

# 추가 스트림 핸들러 설정 (별도로 추가)
stderr_handler = logging.StreamHandler(sys.stderr)
stderr_handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))
logger = logging.getLogger("hwp-mcp-stdio-server")
logger.addHandler(stderr_handler)

# Optional: Disable SSL certificate validation for development
ssl._create_default_https_context = ssl._create_unverified_context

# Set up paths
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

try:
    # Import FastMCP library
    from mcp.server.fastmcp import FastMCP
    logger.info("FastMCP successfully imported")
except ImportError as e:
    logger.error(f"Failed to import FastMCP: {str(e)}")
    print(f"Error: Failed to import FastMCP. Please install with 'pip install mcp'", file=sys.stderr)
    sys.exit(1)

# Try to import HwpController
try:
    from src.tools.hwp_controller import HwpController
    logger.info("HwpController imported successfully")
except ImportError as e:
    logger.error(f"Failed to import HwpController: {str(e)}")
    # Try alternate paths
    try:
        sys.path.append(os.path.join(current_dir, "src"))
        sys.path.append(os.path.join(current_dir, "src", "tools"))
        from hwp_controller import HwpController
        logger.info("HwpController imported from alternate path")
    except ImportError as e2:
        logger.error(f"Could not find HwpController in any path: {str(e2)}")
        print(f"Error: Could not find HwpController module", file=sys.stderr)
        sys.exit(1)

# 찾은 글자 기준 삽입·서식 복사 헬퍼(COM 객체를 직접 받음)
try:
    from src.tools import hwp_edit_helpers
except ImportError:
    import hwp_edit_helpers

# 변경 추적 수락/거절·일괄 수정·누름틀·문서 구조 헬퍼
try:
    from src.tools import hwp_review_helpers
except ImportError:
    import hwp_review_helpers

# 서식 조건 찾기·스타일·메모 헬퍼, 공통 오류 응답, 변경 추적 처리
try:
    from src.tools import hwp_format_helpers
    from src.tools import hwp_errors as errors
    from src.tools.hwp_tracking import tracked_edit, track_scope, resolve_track, TRACK_NOTE
except ImportError:
    import hwp_format_helpers
    import hwp_errors as errors
    from hwp_tracking import tracked_edit, track_scope, resolve_track, TRACK_NOTE

# Initialize FastMCP server
mcp = FastMCP(
    "hwp-mcp",
    instructions="HWP MCP Server for controlling Hangul Word Processor",
    dependencies=["pywin32>=305"]
)

# Global HWP controller instance
hwp_controller = None

def get_hwp_controller():
    """Get or create HwpController instance. Auto-reconnects if connection is lost."""
    global hwp_controller

    # 연결 상태 확인 및 재연결
    if hwp_controller is not None:
        try:
            # 간단한 연결 테스트
            _ = hwp_controller.hwp.XHwpWindows.Count
        except Exception as e:
            logger.warning(f"HWP connection lost ({e}), attempting to reconnect...")
            hwp_controller = None

    if hwp_controller is None:
        logger.info("Creating HwpController instance...")
        try:
            hwp_controller = HwpController()
            connected = hwp_controller.connect(visible=True)
            if not connected:
                logger.warning("Failed to connect to HWP program on startup, but returning controller instance for manual reconnection/switching.")

            logger.info("HwpController instance created")
        except Exception as e:
            logger.error(f"Error creating HwpController: {str(e)}", exc_info=True)
            return None
    return hwp_controller

def _get_com():
    """현재 한글 COM 객체(연결 실패 시 None). 변경 추적 데코레이터용."""
    hwp = get_hwp_controller()
    return getattr(hwp, "hwp", None) if hwp else None

# 기존 문서를 고치는 도구에 붙이는 변경 추적 데코레이터.
# 도구 인자 track(True/False) > 환경변수 HWP_MCP_TRACK_EDITS(on/off) > 기본 on. 끝나면 작업 전 값으로 되돌림.
track_edits = tracked_edit(_get_com)

@mcp.tool()
def hwp_create() -> str:
    """Create a new HWP document."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        
        if hwp.create_new_document():
            logger.info("Successfully created new document")
            return "New document created successfully"
        else:
            return errors.from_message("Error: Failed to create new document")
    except Exception as e:
        logger.error(f"Error creating document: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_list_tabs() -> str:
    """현재 HWP 창에서 열려있는 탭(문서) 목록을 반환합니다."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()

        success, documents = hwp.get_open_documents()
        if not success:
            return errors.from_message("Error: Failed to get document list")

        if not documents:
            return "열려있는 문서가 없습니다."

        result = f"열려있는 문서 ({len(documents)}개):\n"
        for doc in documents:
            marker = "→ " if doc["is_current"] else "  "
            result += f"{marker}[{doc['index']}] {doc['path']}\n"

        return result
    except Exception as e:
        logger.error(f"Error listing documents: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_switch_tab(index: int) -> str:
    """현재 HWP 창에서 특정 인덱스의 탭으로 전환합니다."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()

        success, message = hwp.switch_document(index)
        if success:
            logger.info(message)
            return message
        else:
            return errors.from_message(message)
    except Exception as e:
        logger.error(f"Error switching document: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_list_windows() -> str:
    """실행 중인 모든 HWP 창(인스턴스) 목록을 반환합니다."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()

        success, instances = hwp.get_all_hwp_instances()
        if not success:
            return errors.from_message("Error: Failed to get HWP instances")

        if not instances:
            return "실행 중인 HWP 인스턴스가 없습니다."

        result = f"HWP 인스턴스 ({len(instances)}개):\n"
        for inst in instances:
            marker = "→ " if inst["is_current"] else "  "
            result += f"{marker}[{inst['index']}] {inst['title']} (hwnd: {inst['hwnd']})\n"

        return result
    except Exception as e:
        logger.error(f"Error listing HWP instances: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_switch_window(hwnd: int) -> str:
    """지정된 윈도우 핸들(hwnd)의 HWP 창으로 전환합니다."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()

        success, message = hwp.connect_to_hwp_instance(hwnd)
        if success:
            logger.info(message)
            return message
        else:
            return errors.from_message(message)
    except Exception as e:
        logger.error(f"Error connecting to HWP instance: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_close_window(hwnd: int) -> str:
    """지정된 HWP 창을 닫습니다(저장 안 함)."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()

        success, message = hwp.close_hwp_window(hwnd)
        if success:
            logger.info(message)
            return message
        else:
            return errors.from_message(message)
    except Exception as e:
        logger.error(f"Error closing HWP window: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_open(path: str) -> str:
    """Open an existing HWP document."""
    try:
        if not path:
            return errors.from_message("Error: File path is required")
        
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        
        if hwp.open_document(path):
            logger.info(f"Successfully opened document: {path}")
            return f"Document opened: {path}"
        else:
            return errors.from_message("Error: Failed to open document")
    except Exception as e:
        logger.error(f"Error opening document: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_save(path: str = None) -> str:
    """
    현재 문서를 저장합니다.
    - path를 비우면 원래 파일에 덮어 저장. 한 번도 저장하지 않은 새 문서면 오류(path 필요).
    - path를 주면 그 경로로 저장. 확장자로 형식 결정: .hwp / .hwpx / .pdf
      (.pdf는 내보내기만 하고, 편집 중인 문서는 원래 파일 그대로)
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()

        if path:
            if hwp.save_document(path):
                abs_path = os.path.abspath(path)
                logger.info(f"Successfully saved document to: {abs_path}")
                return f"Document saved to: {abs_path} ({hwp.save_format_for(abs_path)})"
            else:
                return errors.from_message("Error: Failed to save document")
        else:
            ok, msg = hwp.save_current_document()
            if ok:
                logger.info(f"Successfully saved document: {msg}")
                return f"Document saved to: {msg}"
            return errors.from_message(msg)
    except Exception as e:
        logger.error(f"Error saving document: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
@track_edits
def hwp_insert_text(text: str, preserve_linebreaks: bool = True, track: bool = None) -> str:
    """
    현재 커서 위치에 글자를 넣습니다. 표 셀 안이면 그 셀의 커서 위치에 넣음.
    선택(블록)이 있으면 선택한 글자를 지우고 그 자리에 넣습니다.
    - text: 줄바꿈(\\n)은 문단 나눔. 새 문단만 만들려면 text="\\n".
    - preserve_linebreaks: False면 줄바꿈을 공백으로 바꿔 한 문단으로 넣음.
    찾은 글자를 기준으로 넣으려면 hwp_insert_at.
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        if not text:
            return errors.from_message("Error: Text is required")
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        if not preserve_linebreaks:
            text = " ".join(text.replace("\r\n", "\n").split("\n"))
        if hwp.insert_text(text, preserve_linebreaks):
            return "Text inserted successfully"
        return errors.from_message("Error: Failed to insert text")
    except Exception as e:
        logger.error(f"Error inserting text: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_set_font(
    name: str = None, 
    size: int = None, 
    bold: bool = False, 
    italic: bool = False, 
    underline: bool = False,
    select_previous_text: bool = False
) -> str:
    """Set font properties for selected text."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        
        # 현재 선택된 텍스트에 대해 글자 모양 설정
        if hwp.set_font_style(
            font_name=name,
            font_size=size,
            bold=bold,
            italic=italic,
            underline=underline,
            select_previous_text=select_previous_text
        ):
            logger.info("Successfully set font")
            return "Font set successfully"
        else:
            return errors.from_message("Error: Failed to set font")
    
    except Exception as e:
        logger.error(f"Error setting font: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_insert_table(rows: int, cols: int, data: list = None) -> str:
    """
    현재 커서 위치에 rows x cols 표를 넣고, data가 있으면 채웁니다.
    - data: [[1행 값...], [2행 값...]] (선택). 값이 null이거나 ""이면 빈 칸. 숫자도 됨. 줄바꿈(\\n)은 셀 안 문단 나눔.
      표보다 많은 값은 버림.
    커서가 표 안에 있으면 표 속에 표가 생기므로 거부합니다.
    끝나면 커서는 표 바로 뒤(본문)에 있습니다.
    기존 표를 채우려면 hwp_fill_table_grid, 레이블 옆 칸은 hwp_fill_cells.
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        if isinstance(data, str):
            data = json.loads(data)
        if hwp.is_cursor_in_table():
            return errors.err(errors.INVALID_ARG,
                              "커서가 표 안에 있어 새 표를 만들지 않았습니다(표 속에 표가 생김).",
                              "커서를 표 밖으로 옮긴 뒤 다시 호출하세요. 기존 표를 채우려면 hwp_fill_table_grid.")
        if not hwp.insert_table(rows, cols):
            return errors.from_message(f"Error: 표를 넣지 못했습니다 - {hwp.last_error or '원인 불명'}")
        msg = f"표를 넣었습니다({rows}x{cols})."
        if data:
            ok, fill_msg = hwp_edit_helpers.fill_table_grid(hwp.hwp, "", data, skip_empty=True)
            msg += " " + fill_msg
            if not ok:
                return errors.from_message(f"Error: {msg}")
        hwp_edit_helpers.leave_table(hwp.hwp)
        return msg
    except Exception as e:
        logger.error(f"Error inserting table: {str(e)}", exc_info=True)
        return errors.com_error(e)


@mcp.tool()
def hwp_table_read(anchor: str = "", occurrence: int = 1, index: int = 0) -> dict:
    """
    표 하나의 모든 셀 글자를 읽습니다. 커서와 클립보드를 건드리지 않습니다(문서를 HWPML로 읽음).
    - anchor: 셀 안에 이 글자가 들어 있는 표. occurrence: 그런 표가 여럿이면 몇 번째(1부터).
    - index: 문서 순서상 몇 번째 표(1부터. 표 안의 표도 셈). anchor 대신 씀.
    - 둘 다 비우면 모든 표의 목록(번호, 크기, 첫 행)만 돌려줌.
    Returns: {"index", "row_count", "col_count", "rows", "merged"}
      rows: 행마다 실제 셀 글자 목록(Tab 순서, 병합으로 가려진 칸은 빠짐). hwp_fill_table_grid의 rows와 같은 모양이라
            읽은 rows를 고쳐서 그대로 넘기면 됨.
      merged: 병합 셀 [{"row", "col", "rowspan", "colspan", "text"}] (1부터)
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return {"error": errors.not_running()}
        tables = hwp_format_helpers.read_tables(hwp_format_helpers.read_hwpml(hwp.hwp))
        if not tables:
            return {"error": errors.err(errors.NOT_FOUND, "문서에 표가 없습니다.")}
        if anchor:
            t = hwp_format_helpers.find_table(tables, anchor, max(1, int(occurrence or 1)))
            if t is None:
                return {"error": errors.err(errors.NOT_FOUND, f"'{anchor}'가 든 표를 찾지 못했습니다"
                                            f"{f'({occurrence}번째)' if occurrence and occurrence > 1 else ''}.",
                                            "hwp_table_read()로 표 목록을 먼저 확인하세요.")}
            return t
        if index:
            if not 1 <= index <= len(tables):
                return {"error": errors.err(errors.INVALID_ARG, f"표 번호는 1~{len(tables)}입니다: {index}")}
            return tables[index - 1]
        return {"count": len(tables),
                "tables": [{"index": t["index"], "row_count": t["row_count"], "col_count": t["col_count"],
                            "first_row": t["rows"][0] if t["rows"] else []} for t in tables]}
    except Exception as e:
        logger.error(f"Error reading table: {str(e)}", exc_info=True)
        return {"error": errors.com_error(e)}

@mcp.tool()
def hwp_get_text(start_text: str = "", max_chars: int = 0, page: int = 0) -> str:
    """
    현재 문서의 글자를 읽습니다. 인자를 비우면 문서 전체.
    - start_text: 이 글자가 처음 나오는 곳부터 읽음(문서 처음부터 첫 번째). 비우면 문서 처음부터.
    - max_chars: 최대 글자 수. 0이면 끝까지.
    - page: 쪽 순번(문서 맨 앞부터 1, 2, 3…; 인쇄된 쪽 번호와 다를 수 있음 → hwp_scan_document의 page_numbering).
      주면 그 쪽의 글자만 읽음(표·글상자 포함). start_text와 함께 쓰지 않음.
    긴 문서에서 한 절만 볼 때 예: start_text="3. 용산공원", max_chars=3000
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()

        if page and page > 0:
            total = int(hwp.hwp.PageCount)
            if page > total:
                return errors.from_message(f"Error: 쪽 번호가 범위를 벗어났습니다({page}/{total}쪽).")
            text = hwp_review_helpers.get_page_text(hwp.hwp, page)
            if max_chars and max_chars > 0:
                text = text[:max_chars]
            return text

        if start_text or (max_chars and max_chars > 0):
            ok, text = hwp.get_text_range(start_text, max_chars)
            return text if ok else errors.from_message(text)

        text = hwp.get_text()
        if text is not None:
            logger.info("Successfully retrieved document text")
            return text
        else:
            return errors.from_message("Error: Failed to get document text")
    except Exception as e:
        logger.error(f"Error getting text: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_get_selection() -> str:
    """현재 드래그(블록 지정)하여 선택된 영역의 텍스트를 가져옵니다."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()

        text = hwp.get_selected_text()
        if text:
            logger.info("Successfully retrieved selected text")
            return text
        else:
            return "선택된 텍스트가 없습니다. 한글 창에서 텍스트를 먼저 선택(블록 지정)하세요."
    except Exception as e:
        logger.error(f"Error getting selection: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
@track_edits
def hwp_replace_selection(text: str, preserve_linebreaks: bool = True, track: bool = None) -> str:
    """
    한글 창에서 선택(블록 지정)한 영역을 text로 그 자리에서 바꿉니다.
    선택이 없으면 아무것도 바꾸지 않고 오류를 돌려줍니다.
    - preserve_linebreaks: True면 text 안의 줄바꿈(\\n)을 문단 나눔으로 넣음. False면 공백으로 바꿈.
    흐름 예: hwp_get_selection으로 읽기 → 고쳐 쓰기 → hwp_replace_selection
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        selected = hwp.get_selected_text()
        ok, msg = hwp_review_helpers.replace_selection(hwp.hwp, text, selected, preserve_linebreaks)
        return msg if ok else errors.from_message(msg)
    except Exception as e:
        logger.error(f"Error replacing selection: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_get_context() -> dict:
    """
    현재 커서를 기준으로 같은 문단 내 앞쪽/뒤쪽 문맥을 가져옵니다.

    "이어쓰기"나 "이 위치 기준 교정"을 할 때 앞뒤 맥락 파악용으로 사용합니다.
    커서 위치를 보존합니다.

    Returns:
        dict: {"before": 커서 앞쪽 텍스트, "after": 커서 뒤쪽 텍스트}
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return {"error": errors.not_running()}

        context = hwp.get_cursor_context()
        logger.info("Successfully retrieved cursor context")
        return context
    except Exception as e:
        logger.error(f"Error getting context: {str(e)}", exc_info=True)
        return {"error": errors.com_error(e)}

@mcp.tool()
def hwp_close_document(save: bool = False, suppress_dialog: bool = True) -> str:
    """현재 활성화된 문서를 닫습니다."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()

        if hwp.close_document(save, suppress_dialog):
            logger.info(f"Successfully closed document (save={save}, suppress_dialog={suppress_dialog})")
            return "Document closed successfully"
        else:
            return errors.from_message("Error: Failed to close document")
    except Exception as e:
        logger.error(f"Error closing document: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_close_all_documents(save: bool = False, suppress_dialog: bool = True) -> str:
    """열려 있는 모든 HWP 문서를 닫습니다."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()

        if hwp.close_all_documents(save, suppress_dialog):
            logger.info(f"Successfully closed all documents (save={save}, suppress_dialog={suppress_dialog})")
            return "All documents closed successfully"
        else:
            return errors.from_message("Error: Failed to close all documents")
    except Exception as e:
        logger.error(f"Error closing all documents: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_undo(count: int = 1) -> str:
    """직전 작업을 지정한 횟수만큼 실행 취소(Undo)합니다."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()

        success, message = hwp.undo(count)
        if success:
            logger.info(message)
            return message
        else:
            return errors.from_message(message)
    except Exception as e:
        logger.error(f"Error in undo: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_redo(count: int = 1) -> str:
    """실행 취소한 작업을 지정한 횟수만큼 다시 실행(Redo)합니다."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()

        success, message = hwp.redo(count)
        if success:
            logger.info(message)
            return message
        else:
            return errors.from_message(message)
    except Exception as e:
        logger.error(f"Error in redo: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_find_text(text: str) -> str:
    """문서 내에서 지정한 텍스트를 찾아 커서를 이동합니다."""
    try:
        if not text:
            return errors.from_message("Error: Text is required")

        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()

        if hwp.find_text(text):
            logger.info(f"Found text: {text}")
            return f"Text found: {text}"
        else:
            return errors.err(errors.NOT_FOUND, f"글자를 찾지 못했습니다: {text}")
    except Exception as e:
        logger.error(f"Error finding text: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
@track_edits
def hwp_fill_cells(
    path_value_map: dict,
    mode: str = "replace",
    track: bool = None,
) -> str:
    """
    표에서 지정된 경로(예: "이름 > <right>")를 따라 여러 셀에 값을 입력합니다. 경로 상세 규칙 및 모드는 workspace의 instructions.md 문서를 참조하세요.
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        if not path_value_map:
            return errors.from_message("Error: path_value_map이 필요합니다.")

        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()

        # 배치 처리 (direction은 경로에서 결정되므로 "right"를 기본값으로)
        results = hwp.fill_cells_by_path_batch(path_value_map, "right", mode)

        # 결과 정리
        success_count = 0
        fail_count = 0
        messages = []

        for path_str, (success, message) in results.items():
            if success:
                success_count += 1
                messages.append(f"✓ {path_str}: {message}")
            else:
                fail_count += 1
                messages.append(f"✗ {path_str}: {message}")

        summary = f"\n총 {success_count}개 성공, {fail_count}개 실패"
        result_message = "\n".join(messages) + summary

        logger.info(f"셀 채우기 완료: {success_count}개 성공, {fail_count}개 실패")
        return result_message

    except Exception as e:
        logger.error(f"셀 채우기 중 오류: {str(e)}", exc_info=True)
        return errors.com_error(e)


@mcp.tool()
def hwp_set_track_changes(active: bool) -> str:
    """변경 내용 추적(Track Changes) 모드를 켜거나 끕니다."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        if hwp.set_track_changes(active):
            status = "활성화" if active else "비활성화"
            return f"변경 내용 추적 기능이 {status}되었습니다."
        else:
            return errors.from_message("Error: 변경 내용 추적 설정 실패")
    except Exception as e:
        logger.error(f"Error setting track changes: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_track_changes_resolve(action: str = "apply_all") -> str:
    """
    변경 추적으로 표시된 고친 내용을 반영(수락)하거나 되돌립니다(거절).
    - action:
        "apply_all"   모든 변경 내용 반영
        "cancel_all"  모든 변경 내용 되돌림(원문 복원)
        "apply_next"  커서 다음의 변경 내용 하나 반영
        "cancel_next" 커서 다음의 변경 내용 하나 되돌림
        "apply" / "cancel"  커서 위치의 변경 내용 하나
    흐름 예: hwp_set_track_changes(True) → hwp_patch로 고침 → 사용자 검토 → 이 도구로 반영/되돌림
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        ok, msg = hwp_review_helpers.resolve_track_changes(hwp.hwp, action)
        return msg if ok else errors.from_message(msg)
    except Exception as e:
        logger.error(f"Error resolving track changes: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
@track_edits
def hwp_patch(edits: list, track: bool = None) -> dict:
    """
    원문 일부(old)를 찾아 new로 바꾸는 작업 여러 건을 한 번에 실행하고, 건별 결과를 돌려줍니다.
    - edits: [{"old": "바꿀 원문", "new": "새 글자"}, ...]
        * old: 문서에 있는 글자 그대로(한 문단 안, 줄바꿈 불가). 한 곳에만 있도록 앞뒤 글자를 넉넉히 포함.
        * new: 바꿀 글자. 비우면 지움. 줄바꿈(\\n)은 문단 나눔.
        * occurrence(선택): old가 여러 곳이면 몇 번째(1부터)를 바꿀지.
        * all(선택): True면 모든 곳을 바꿈.
      old가 여러 곳에서 일치하는데 occurrence/all이 없으면 바꾸지 않고 status="ambiguous".
    Returns: {"summary": "...", "results": [{"index", "old", "status"(ok/not_found/ambiguous/invalid/error), "count", "replaced", "message"}]}
    실패한 건만 old를 고쳐 다시 보내면 됩니다.
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return {"error": errors.not_running()}
        if isinstance(edits, str):
            edits = json.loads(edits)
        # 변경 추적은 track_edits 데코레이터가 처리
        results = hwp_review_helpers.patch_document(hwp.hwp, edits, text_reader=hwp.get_text, track=None)
        return {"summary": hwp_review_helpers.summarize_patch(results), "results": results}
    except Exception as e:
        logger.error(f"Error patching document: {str(e)}", exc_info=True)
        return {"error": errors.com_error(e)}

@mcp.tool()
def hwp_fields(action: str = "list", name: str = "", value: str = None, values: dict = None, track: bool = None) -> dict:
    """
    누름틀·셀 필드(이름 붙은 입력 칸)를 읽고 씁니다. 서식 양식 채우기에 가장 정확한 방법입니다.
    - action:
        "list" 필드 이름 목록. 같은 이름이 여러 개면 "이름{{0}}", "이름{{1}}"처럼 번호가 붙음.
        "get"  name 필드 값 읽기. name을 비우면 모든 필드 값.
        "put"  값 쓰기. name+value 한 개, 또는 values={"이름": "값", ...} 여러 개.
               "이름"만 쓰면 같은 이름 필드 전부, "이름{{1}}"이면 그 하나만.
    문서에 필드가 없으면 list가 빈 목록입니다(이때는 hwp_fill_cells 등 다른 도구 사용).
    track: put할 때 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름.
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return {"error": errors.not_running()}
        h = hwp.hwp
        if action == "list":
            names = hwp_review_helpers.list_fields(h, 1)
            return {"count": len(names), "fields": names}
        if action == "get":
            return {"values": hwp_review_helpers.get_fields(h, [name] if name else None)}
        if action == "put":
            data = dict(values or {})
            if name:
                data[name] = "" if value is None else value
            effective = resolve_track(track)
            with track_scope(h, effective):
                ok, msg, missing = hwp_review_helpers.put_fields(h, data)
            out = {"ok": ok, "message": msg, "missing": missing}
            if ok and effective:
                out["track_changes"] = TRACK_NOTE
            if not ok:
                out["error"] = errors.err(errors.NOT_FOUND if missing else errors.INVALID_ARG, msg,
                                          "hwp_fields(action='list')로 실제 필드 이름을 확인하세요." if missing else None)
            return out
        return {"error": errors.err(errors.INVALID_ARG, f"action은 list/get/put 중 하나여야 합니다: {action}")}
    except Exception as e:
        logger.error(f"Error handling fields: {str(e)}", exc_info=True)
        return {"error": errors.com_error(e)}

@mcp.tool()
def hwp_scan_document(include_headings: bool = True) -> dict:
    """
    현재 문서의 구조를 요약합니다. 긴 문서를 고치기 전에 먼저 훑어볼 때 씁니다.
    Returns:
      pages: 쪽 수
      page_numbering: 쪽 순번과 인쇄 쪽 번호의 대응. 예: [{"pages": "1-2", "printed": "1-2"}, {"pages": "3-20", "printed": "1-18"}]
        → 사용자가 말하는 "5쪽"은 보통 인쇄 번호. 이 표로 순번으로 바꿔 hwp_get_text(page=)·hwp_page_image(page=)에 넘길 것.
          같은 인쇄 번호가 여러 구간에 있으면(표지 1쪽, 본문 1쪽 등) 어느 쪽인지 사용자에게 확인.
      table_count / tables: [{"index", "page"(순번), "printed_page"(인쇄 번호)}]  (최대 200개까지 상세. 표 내용은 hwp_table_read)
      picture_count / pictures: [{"index", "page", "printed_page"}]
      other_controls: 그 밖의 개체 종류별 개수(CtrlID 기준)
      headings: 개요 문단 [{"level", "number", "text"}]  (include_headings=False면 생략, 표 안 문단 제외)
    커서 위치는 원래대로 돌려놓습니다. 특정 쪽 글자는 hwp_get_text(page=n).
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return {"error": errors.not_running()}
        return hwp_review_helpers.scan_document(hwp.hwp, include_headings)
    except Exception as e:
        logger.error(f"Error scanning document: {str(e)}", exc_info=True)
        return {"error": errors.com_error(e)}

@mcp.tool()
def hwp_search_format(
    text: str = "",
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    font_size: float = None,
    text_color: str = None,
    font_name: str = "",
    max_results: int = 50,
) -> dict:
    """
    서식 조건으로 글자를 찾습니다. 예: 굵은 글씨, 빨간 글씨, 14pt 글자, 특정 글꼴.
    - text: 비우면 조건에 맞는 모든 글자. 쓰면 그 글자 중 조건에 맞는 것만.
    - bold/italic/underline: True/False. None이면 조건에 넣지 않음.
    - font_size: 글자 크기(pt). 예: 14
    - text_color: "#RRGGBB" 또는 red/blue/green/black/gray 등.
    - font_name: 글꼴 이름 일부(예: "돋움", "HY견고딕"). 이 조건이 있으면 문서를 HWPML로 읽어 판정하므로 쪽 번호(page)는 null.
    여러 조건은 모두 만족(AND). 문서는 바꾸지 않고 커서도 원래 자리로 돌려놓습니다.
    Returns: {"method": "find"|"hwpml", "count", "matches": [{"index", "text", "page"(순번), "printed_page"(인쇄 쪽 번호), ...}]}
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return {"error": errors.not_running()}
        try:
            cond = hwp_format_helpers.build_charshape_conditions(bold, italic, underline, font_size, text_color)
        except ValueError as ve:
            return {"error": errors.err(errors.INVALID_ARG, str(ve))}
        if not cond and not font_name:
            return {"error": errors.err(errors.INVALID_ARG, "서식 조건(bold/italic/underline/font_size/text_color/font_name)을 하나 이상 지정하세요.",
                                        "글자만 찾으려면 hwp_find_text를 쓰세요.")}
        max_results = max(1, min(int(max_results or 50), 500))
        if font_name:
            xml = hwp_format_helpers.read_hwpml(hwp.hwp)
            matches = hwp_format_helpers.find_by_format_hwpml(xml, text, cond, font_name, max_results)
            method = "hwpml"
        else:
            matches = hwp_format_helpers.find_by_charshape(hwp.hwp, text, cond, max_results)
            method = "find"
        out = {"method": method, "count": len(matches), "matches": matches}
        if not matches:
            out["message"] = "조건에 맞는 글자가 없습니다."
        if len(matches) >= max_results:
            out["truncated"] = True
        return out
    except Exception as e:
        logger.error(f"Error searching by format: {str(e)}", exc_info=True)
        return {"error": errors.com_error(e)}

@mcp.tool()
def hwp_list_styles() -> dict:
    """
    문서의 스타일 목록과 커서 위치 문단의 현재 스타일을 돌려줍니다.
    Returns: {"current": {"index", "name"}, "styles": [{"index", "name", "eng_name", "type"(Para/Char)}]}
    문서는 바꾸지 않습니다. 스타일 적용은 hwp_set_style.
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return {"error": errors.not_running()}
        styles = hwp_format_helpers.list_styles(hwp_format_helpers.read_hwpml(hwp.hwp))
        cur = None
        try:
            idx = hwp_format_helpers.current_style_index(hwp.hwp)
            st = hwp_format_helpers.resolve_style(styles, idx)
            cur = {"index": idx, "name": st["name"] if st else ""}
        except Exception as e:
            logger.debug(f"현재 스타일 읽기 실패(무시): {e}")
        return {"current": cur, "count": len(styles), "styles": styles}
    except Exception as e:
        logger.error(f"Error listing styles: {str(e)}", exc_info=True)
        return {"error": errors.com_error(e)}

@mcp.tool()
@track_edits
def hwp_set_style(style: str, anchor: str = "", occurrence: int = 1, track: bool = None) -> str:
    """
    문단에 스타일을 적용합니다(예: "개요 1", "본문", "Outline 2", 또는 번호 2).
    - style: 스타일 이름(한글/영문) 또는 번호. 목록은 hwp_list_styles.
    - anchor: 이 글자가 있는 문단에 적용. 비우면 커서 위치 문단(블록 선택 시 선택한 문단들).
    - occurrence: anchor가 여러 곳이면 몇 번째(1부터).
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        h = hwp.hwp
        styles = hwp_format_helpers.list_styles(hwp_format_helpers.read_hwpml(h))
        st = hwp_format_helpers.resolve_style(styles, style)
        if not st:
            return errors.err(errors.NOT_FOUND, f"없는 스타일입니다: {style}",
                              "hwp_list_styles로 실제 스타일 이름을 확인하세요.")
        if anchor:
            if not hwp_review_helpers._find_nth(h, anchor, max(1, int(occurrence or 1))):
                return errors.err(errors.NOT_FOUND, f"기준 글자를 찾지 못했습니다: {anchor}")
            h.HAction.Run("Cancel")
        if not hwp_format_helpers.apply_style(h, st["index"]):
            return errors.err(errors.COM_ERROR, f"스타일 적용 실패: {st['name']}")
        where = f"'{anchor}' 문단" if anchor else "현재 문단"
        return f"{where}에 스타일 '{st['name']}'(번호 {st['index']})을 적용했습니다."
    except Exception as e:
        logger.error(f"Error setting style: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_memo(action: str = "list", text: str = "", anchor: str = "", occurrence: int = 1) -> dict:
    """
    메모(검토 의견)를 읽거나 넣습니다. 본문은 고치지 않고 의견만 남길 때 씁니다.
    - action="list": 문서의 모든 메모. [{"index", "author", "content", "target"(메모가 붙은 본문 글자), "paragraph_text"}]
    - action="insert": anchor 글자(occurrence번째)에 text 메모를 붙임. anchor를 비우면 현재 선택/커서 위치.
    list는 문서를 바꾸거나 저장하지 않습니다(HWPML로 읽음).
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return {"error": errors.not_running()}
        h = hwp.hwp
        if action == "list":
            memos = hwp_format_helpers.list_memos(hwp_format_helpers.read_hwpml(h))
            return {"count": len(memos), "memos": memos}
        if action == "insert":
            if not text:
                return {"error": errors.err(errors.INVALID_ARG, "메모 내용(text)이 필요합니다.")}
            if anchor:
                if not hwp_review_helpers._find_nth(h, anchor, max(1, int(occurrence or 1))):
                    return {"error": errors.err(errors.NOT_FOUND, f"기준 글자를 찾지 못했습니다: {anchor}")}
            if not hwp_format_helpers.insert_memo(h, text):
                return {"error": errors.err(errors.COM_ERROR, "메모 넣기 실패")}
            return {"ok": True, "message": f"메모를 넣었습니다{f' (대상: {anchor})' if anchor else ''}."}
        return {"error": errors.err(errors.INVALID_ARG, f"action은 list/insert 중 하나여야 합니다: {action}")}
    except Exception as e:
        logger.error(f"Error handling memo: {str(e)}", exc_info=True)
        return {"error": errors.com_error(e)}

_TABLE_INSERT_ACTIONS = {
    "below": "TableInsertLowerRow",
    "above": "TableInsertUpperRow",
    "right": "TableInsertRightColumn",
    "left": "TableInsertLeftColumn",
}

def _move_to_anchor(hwp, anchor: str) -> bool:
    """anchor 글자를 찾아 그 셀로 커서를 옮깁니다. anchor가 비면 현재 커서 위치를 씁니다."""
    if not anchor:
        return True
    if not hwp.find_text(anchor):
        return False
    hwp.hwp.Run("Cancel")  # 찾은 글자 선택 해제(커서는 그 셀 안에 남음)
    return True

@mcp.tool()
@track_edits
def hwp_table_insert(anchor: str = "", position: str = "below", count: int = 1, data: list = None, track: bool = None) -> str:
    """
    표에 행 또는 열을 넣습니다.
    - anchor: 기준 셀 안의 글자(문서 처음부터 첫 번째로 찾은 곳). 비우면 현재 커서가 있는 셀 기준.
    - position: "below"/"above"(행), "right"/"left"(열).
    - count: 넣을 행·열 수.
    - data: (행 추가 때만) 새 행에 채울 값. [[1행 셀값...], [2행 셀값...]] 처럼 위에서 아래 순서.
      각 행은 첫 열부터 채움. 셀 수보다 짧으면 나머지는 빈칸.
    편집 전에 hwp_set_track_changes(active=True)를 먼저 켤 것.
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        if position not in _TABLE_INSERT_ACTIONS:
            return errors.from_message(f"Error: position은 {list(_TABLE_INSERT_ACTIONS)} 중 하나여야 합니다.")
        if count < 1:
            return errors.from_message("Error: count는 1 이상이어야 합니다.")
        if data and position not in ("below", "above"):
            return errors.err(errors.INVALID_ARG, "data는 행 추가(below/above)에서만 쓸 수 있습니다.")
        if data and len(data) > count:
            return errors.err(errors.INVALID_ARG, f"data 행 수({len(data)})가 count({count})보다 많습니다.")
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        if not _move_to_anchor(hwp, anchor):
            return errors.from_message(f"Error: 기준 글자를 찾지 못했습니다: {anchor}")
        h = hwp.hwp
        action = _TABLE_INSERT_ACTIONS[position]
        for _ in range(count):
            if not h.HAction.Run(action):
                return errors.from_message("Error: 행·열 넣기 실패(커서가 표 안에 있는지 확인하세요)")
        if data:
            # 행 추가 후 커서는 기준 셀에 남아 있음 → 새 행 첫 줄로 이동
            if position == "below":
                h.Run("TableLowerCell")
            else:
                for _ in range(count):
                    h.Run("TableUpperCell")
            for i, row in enumerate(data):
                if i:
                    h.Run("TableLowerCell")
                h.Run("TableColBegin")
                for j, val in enumerate(row):
                    if j:
                        h.Run("TableRightCell")
                    if val not in (None, ""):
                        hwp.insert_text(str(val), preserve_linebreaks=False)
                h.Run("TableColBegin")
        kind = "행" if position in ("below", "above") else "열"
        return f"{kind} {count}개를 {position} 쪽에 넣었습니다." + (f" {len(data)}행에 값을 채웠습니다." if data else "")
    except Exception as e:
        logger.error(f"Error inserting table rows/cols: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
@track_edits
def hwp_table_delete(anchor: str = "", target: str = "row", count: int = 1, track: bool = None) -> str:
    """
    표에서 행 또는 열을 지웁니다.
    - anchor: 기준 셀 안의 글자(문서 처음부터 첫 번째로 찾은 곳). 비우면 현재 커서가 있는 셀 기준.
    - target: "row"(기준 셀의 행부터 아래로 count개) 또는 "column"(기준 셀의 열부터 오른쪽으로 count개).
    편집 전에 hwp_set_track_changes(active=True)를 먼저 켤 것.
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        actions = {"row": "TableDeleteRow", "column": "TableDeleteColumn"}
        if target not in actions:
            return errors.from_message("Error: target은 'row' 또는 'column'이어야 합니다.")
        if count < 1:
            return errors.from_message("Error: count는 1 이상이어야 합니다.")
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        if not _move_to_anchor(hwp, anchor):
            return errors.from_message(f"Error: 기준 글자를 찾지 못했습니다: {anchor}")
        for _ in range(count):
            if not hwp.hwp.HAction.Run(actions[target]):
                return errors.from_message("Error: 행·열 지우기 실패(커서가 표 안에 있는지 확인하세요)")
        kind = "행" if target == "row" else "열"
        return f"{kind} {count}개를 지웠습니다."
    except Exception as e:
        logger.error(f"Error deleting table rows/cols: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
@track_edits
def hwp_insert_at(anchor: str = "", text: str = "", position: str = "para_after", copy_format_from: str = "", track: bool = None) -> str:
    """
    찾은 글자(anchor)를 기준으로 글자를 넣습니다.
    - anchor: 기준 글자(문서 처음부터 첫 번째로 찾은 곳). 비우면 현재 커서 위치 기준.
    - text: 넣을 글자. 줄바꿈(\\n)은 문단 나눔(line_after에서는 같은 문단 안 줄 바꿈).
    - position:
        "para_before"  anchor 문단 앞에 새 문단
        "para_after"   anchor 문단 뒤에 새 문단
        "line_after"   anchor 문단 끝에서 줄 바꿈(Shift+Enter, BreakLine) 후 이어 씀(같은 문단)
        "after_match"  찾은 글자 바로 뒤에 끼워 넣음
        "before_match" 찾은 글자 바로 앞에 끼워 넣음
    - copy_format_from: 넣은 뒤, 이 글자가 있는 문단의 글자·문단 모양(글머리표 포함)을
      새로 넣은 글자가 있는 문단 전체에 적용. 비우면 서식은 그대로(새 문단은 anchor 문단의
      문단 모양·글머리표를 물려받음).
    편집 전에 hwp_set_track_changes(active=True)를 먼저 켤 것.
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        ok, msg = hwp_edit_helpers.insert_at(hwp.hwp, anchor, text, position, copy_format_from)
        return msg if ok else errors.from_message(msg)
    except Exception as e:
        logger.error(f"Error inserting at anchor: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
@track_edits
def hwp_duplicate(unit: str, anchor: str, target: str, position: str = "para_after", end_anchor: str = "",
                  occurrence: int = 0, target_occurrence: int = 0, track: bool = None) -> str:
    """
    표 하나 또는 문단 범위를 서식 그대로 복사해 다른 문단 앞/뒤에 붙여넣습니다.
    기존 표와 같은 양식의 표를 하나 더 만들 때 씁니다(복사 후 hwp_table_delete로 줄 수를 맞추고 hwp_fill_table_grid로 내용 교체).
    - unit:
        "table"       anchor 글자가 든 표 전체(선 모양·배경색·열 너비 등 그대로)
        "paragraphs"  anchor 문단부터 end_anchor 문단까지 통째로(사이의 표·그림 포함). 본문 문단만.
    - anchor: 복사할 대상 안의 글자. 예: 표 안의 "남기업"
    - end_anchor: unit="paragraphs"일 때 끝 문단 안의 글자(anchor 문단 이후 처음 나오는 곳). 비우면 anchor 문단 하나.
    - target: 붙여넣을 기준 문단 안의 글자. 예: "ㅇ 자문위원"
    - position: "para_after"(target 문단 뒤, 기본) / "para_before"(target 문단 앞)
    - occurrence / target_occurrence: anchor / target이 여러 곳에 있을 때 문서 처음부터 몇 번째(1부터)인지.
    anchor·target이 여러 곳에 있는데 occurrence를 주지 않으면 복사하지 않고 AMBIGUOUS 오류.
    (본문·별첨에 같은 표가 있으면 표 안 글자가 모두 겹치므로 occurrence로 지정)
    클립보드를 거치며, 끝나면 작업 전 클립보드의 글자를 되돌려 놓음(글자 외 형식은 복원 안 됨).
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        ok, msg = hwp_edit_helpers.duplicate(hwp.hwp, unit, anchor, target, position, end_anchor,
                                             text_reader=hwp.get_text, occurrence=occurrence,
                                             target_occurrence=target_occurrence)
        return msg if ok else errors.from_message(msg)
    except Exception as e:
        logger.error(f"Error duplicating: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
@track_edits
def hwp_copy_format(from_anchor: str, to_anchor: str, char: bool = True, para: bool = True, track: bool = None) -> str:
    """
    한 문단의 서식을 다른 문단 전체에 복사합니다.
    - from_anchor: 서식 원본 문단 안의 글자. 글자 모양은 이 글자 기준, 문단 모양은 그 문단 기준.
      비우면 현재 커서 위치 기준.
    - to_anchor: 서식을 받을 문단 안의 글자(문단 전체에 적용). 비우면 현재 커서 문단.
      (두 anchor 모두 문서 처음부터 첫 번째로 찾은 곳)
    - char: 글자 모양 복사(크기, 글꼴, 굵게·기울임, 색, 장평·자간 등).
    - para: 문단 모양 복사(정렬, 왼쪽·오른쪽 여백, 들여쓰기, 줄 간격, 문단 앞뒤 간격,
      문단 종류 HeadingType과 글머리 문자).
    편집 전에 hwp_set_track_changes(active=True)를 먼저 켤 것.
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        if not char and not para:
            return errors.from_message("Error: char와 para 중 하나는 True여야 합니다.")
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        ok, msg = hwp_edit_helpers.copy_format(hwp.hwp, from_anchor, to_anchor, char=char, para=para)
        return msg if ok else errors.from_message(msg)
    except Exception as e:
        logger.error(f"Error copying format: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
@track_edits
def hwp_insert_image(image_path: str, width_mm: float = 0, height_mm: float = 0,
                     anchor: str = "", position: str = "para_after", track: bool = None) -> str:
    """
    그림 파일을 문서에 포함(링크 아님)해서 넣습니다. 글자처럼 취급되어 글과 함께 움직입니다.
    - image_path: 그림 파일 경로(png, jpg, bmp, gif 등).
    - width_mm / height_mm: 크기(mm). 둘 다 0이면 원래 크기, 한쪽만 주면 원래 가로세로 비율 유지.
    - anchor: 기준 글자(문서 처음부터 첫 번째로 찾은 곳). 비우면 현재 커서 위치 기준.
    - position:
        "para_before"  anchor 문단 앞에 그림만 있는 새 문단
        "para_after"   anchor 문단 뒤에 그림만 있는 새 문단
        "at_cursor"    커서 위치(anchor가 있으면 찾은 글자 바로 뒤)에 끼워 넣음
    새 문단은 anchor 문단의 문단 모양(정렬 등)을 물려받음.
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        ok, msg = hwp_edit_helpers.insert_picture(hwp.hwp, image_path, width_mm, height_mm, anchor, position)
        return msg if ok else errors.from_message(msg)
    except Exception as e:
        logger.error(f"Error inserting image: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
@track_edits
def hwp_replace_image(image_path: str, near_text: str, which: str = "after", track: bool = None) -> str:
    """
    기존 그림의 그림 파일만 바꿉니다(한글의 '그림 바꾸기'). 그림 개체는 그대로라
    캡션(자동 번호 포함)·위치·너비·높이·글자처럼 취급 여부가 그대로 남습니다.
    - image_path: 새 그림 파일 경로. 문서에 포함됨(링크 아님).
    - near_text: 바꿀 그림을 찾을 글자(예: 캡션 "<그림 1> 거주 연수별 누적 순주거비"). 문서 처음부터 첫 번째로 찾은 곳.
      그림의 캡션 안 글자면 그 그림을 바꿈(which 무시).
    - which: near_text가 캡션이 아닌 보통 문단에 있을 때 "after"면 그 문단 뒤의 첫 그림,
      "before"면 그 문단 앞의 가장 가까운 그림(같은 리스트, 즉 본문이면 본문 안의 그림만 찾음).
    이전 크기와 새 크기를 돌려줍니다.
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        ok, msg = hwp_edit_helpers.replace_picture(hwp.hwp, image_path, near_text, which)
        return msg if ok else errors.from_message(msg)
    except Exception as e:
        logger.error(f"Error replacing image: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
@track_edits
def hwp_delete_image(near_text: str, which: str = "after", track: bool = None) -> str:
    """
    그림 개체를 지웁니다. 캡션이 있으면 캡션도 함께 지워집니다.
    - near_text: 지울 그림을 찾을 글자(문서 처음부터 첫 번째로 찾은 곳).
      그림의 캡션 안 글자면 그 그림을 지움(which 무시).
    - which: near_text가 캡션이 아닌 보통 문단에 있을 때 "after"면 그 문단 뒤의 첫 그림,
      "before"면 그 문단 앞의 가장 가까운 그림(같은 리스트 안의 그림만 찾음).
    지운 그림의 크기와 캡션 글자(캡션으로 찾은 경우)를 돌려줍니다.
    편집 전에 hwp_set_track_changes(active=True)를 먼저 켤 것.
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        reader = lambda: hwp._get_text_file_safe("TEXT", "saveblock")
        ok, msg = hwp_edit_helpers.delete_picture(hwp.hwp, near_text, which, text_reader=reader)
        return msg if ok else errors.from_message(msg)
    except Exception as e:
        logger.error(f"Error deleting image: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
@track_edits
def hwp_fill_table_grid(anchor: str, rows: list, skip_empty: bool = True, track: bool = None) -> str:
    """
    표 전체를 왼쪽 위 셀부터 행 순서대로 한꺼번에 채웁니다. 셀의 글자 모양은 그대로 둡니다.
    - anchor: 채울 표의 아무 셀 안의 글자(문서 처음부터 첫 번째로 찾은 곳).
    - rows: [[1행 값...], [2행 값...]]. 펼쳐서 셀 순서(Tab으로 다음 셀 가는 순서)대로 대응.
      병합된 셀은 한 칸으로 셈. 그래서 각 행에 그 행의 셀을 빠짐없이 적을 것.
      값이 null이면 그 셀은 건드리지 않음. 숫자도 됨. 줄바꿈(\\n)은 셀 안 문단 나눔.
    - skip_empty: True면 ""도 건드리지 않음. False면 ""은 셀 내용을 지움.
    표 끝을 넘는 값은 버림(행이 늘어나지 않음). 채운 칸 수를 돌려줍니다.
    편집 전에 hwp_set_track_changes(active=True)를 먼저 켤 것.
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        if isinstance(rows, str):
            rows = json.loads(rows)
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        ok, msg = hwp_edit_helpers.fill_table_grid(hwp.hwp, anchor, rows, skip_empty)
        return msg if ok else errors.from_message(msg)
    except Exception as e:
        logger.error(f"Error filling table grid: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
@track_edits
def hwp_set_paragraph_shape(
    heading_type: str = None,
    alignment: str = None,
    line_spacing: int = None,
    left_margin: int = None,
    right_margin: int = None,
    first_line: int = None,
    space_before: int = None,
    space_after: int = None,
    apply_to_all: bool = False,
    track: bool = None,
) -> str:
    """
    현재 커서 위치 또는 선택된 문단의 모양을 설정합니다.
    지정한 파라미터만 변경되고 나머지는 기존 값이 유지됩니다.

    Args:
        heading_type: 문단 유형.
            "none"=없음, "outline"=개요, "number"=번호매기기, "bullet"=글머리표
        alignment: 정렬.
            "left"=왼쪽, "center"=가운데, "right"=오른쪽, "justify"=양쪽
        line_spacing: 줄간격(%). 예: 160 → 160%
        left_margin: 왼쪽 여백(hwpunit). 1pt=100
        right_margin: 오른쪽 여백(hwpunit).
        first_line: 첫줄 들여쓰기(hwpunit). 양수=들여쓰기, 음수=내어쓰기
        space_before: 문단 앞 간격(hwpunit).
        space_after: 문단 뒤 간격(hwpunit).
        apply_to_all: True이면 전체 문서에 적용 (SelectAll 후 실행).
    track: 변경추적으로 표시할지. None이면 환경변수 HWP_MCP_TRACK_EDITS(on/off, 기본 on)를 따름. 끝나면 원래 설정으로 되돌림.
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return errors.not_running()
        if hwp.set_paragraph_shape(
            heading_type=heading_type,
            alignment=alignment,
            line_spacing=line_spacing,
            left_margin=left_margin,
            right_margin=right_margin,
            first_line=first_line,
            space_before=space_before,
            space_after=space_after,
            apply_to_all=apply_to_all,
        ):
            return "문단 모양이 설정되었습니다."
        else:
            return errors.from_message("Error: 문단 모양 설정 실패")
    except Exception as e:
        logger.error(f"Error setting paragraph shape: {str(e)}", exc_info=True)
        return errors.com_error(e)

@mcp.tool()
def hwp_page_image(page: int = 0, resolution: int = 96, save_path: str = None) -> list:
    """
    현재 문서의 한 쪽을 한글 프로그램이 그린 모습 그대로 PNG 이미지로 만들어 돌려줍니다.
    저장하지 않은 편집 내용도 반영됩니다.
    - page: 쪽 순번(문서 맨 앞부터 1, 2, 3…; 인쇄된 쪽 번호와 다를 수 있음 → hwp_scan_document의 page_numbering).
      0이면 커서가 있는 쪽.
    - resolution: DPI(24~300). 96이면 A4 한 쪽이 약 794x1123픽셀. 작은 글씨를 보려면 150.
    - save_path: 이미지를 저장할 경로(.png/.jpg/.bmp/.gif). 비우면 임시 폴더에 저장.
    """
    try:
        import tempfile
        from mcp.server.fastmcp import Image

        hwp = get_hwp_controller()
        if not hwp:
            return [errors.not_running()]
        if page <= 0:
            page = hwp.get_current_page()
        if not save_path:
            save_path = os.path.join(tempfile.gettempdir(), "hwp-mcp", f"page_{page:03d}.png")

        ok, result = hwp.create_page_image(save_path, page, resolution)
        if not ok:
            return [errors.from_message(result)]
        total = hwp.hwp.PageCount
        info = f"{page}/{total}쪽 이미지 저장: {result}"
        if os.path.splitext(result)[1].lower() in (".png", ".jpg", ".jpeg", ".gif"):
            return [info, Image(path=result)]
        return [info]
    except Exception as e:
        logger.error(f"Error creating page image: {str(e)}", exc_info=True)
        return [errors.com_error(e)]

def main():
    """`hwp-mcp-plus` 명령의 진입점 (pyproject.toml [project.scripts])."""
    logger.info("Starting HWP MCP stdio server")
    try:
        # Run the FastMCP server with stdio transport
        mcp.run(transport="stdio")
    except Exception as e:
        logger.error(f"Error running server: {str(e)}", exc_info=True)
        print(f"Error: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main() 