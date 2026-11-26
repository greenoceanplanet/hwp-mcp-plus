"""
MCP 도구 공통 오류 응답.

형식: "Error[코드]: 요약 / 다음: 권장 행동"
- 코드는 LLM이 다음 행동을 고를 수 있게 하는 기계용 값입니다.
- 권장 행동 문구는 아래 NEXT_ACTIONS 한 곳에 모아 둡니다.
"""

HWP_NOT_RUNNING = "HWP_NOT_RUNNING"
NO_DOCUMENT = "NO_DOCUMENT"
NOT_FOUND = "NOT_FOUND"
AMBIGUOUS = "AMBIGUOUS"
NO_SELECTION = "NO_SELECTION"
INVALID_ARG = "INVALID_ARG"
UNSAVED_DOCUMENT = "UNSAVED_DOCUMENT"
FILE_NOT_FOUND = "FILE_NOT_FOUND"
NOT_IN_TABLE = "NOT_IN_TABLE"
COM_ERROR = "COM_ERROR"

NEXT_ACTIONS = {
    HWP_NOT_RUNNING: "사용자에게 한글 프로그램을 실행해 달라고 요청한 뒤 다시 시도하세요. 한글이 응답하지 않으면 한글을 다시 켜 달라고 요청하세요.",
    NO_DOCUMENT: "hwp_open으로 문서를 열거나 hwp_create로 새 문서를 만든 뒤 다시 시도하세요.",
    NOT_FOUND: "hwp_get_text로 실제 글자를 확인한 뒤, 띄어쓰기·특수문자까지 문서와 똑같이 맞춰 다시 시도하세요.",
    AMBIGUOUS: "기준 글자를 더 길게(앞뒤 글자 포함) 쓰거나, 몇 번째인지(occurrence) 지정하세요.",
    NO_SELECTION: "사용자에게 한글에서 대상 글자를 먼저 선택(블록 지정)해 달라고 요청하거나, 기준 글자(anchor)를 지정하는 도구를 쓰세요.",
    INVALID_ARG: "도구 설명의 인자 형식과 허용값을 확인해 고친 뒤 다시 호출하세요.",
    UNSAVED_DOCUMENT: "경로를 추측하지 말고, 사용자에게 저장할 위치를 물은 뒤 hwp_save(path=...)로 저장하세요.",
    FILE_NOT_FOUND: "경로를 추측하지 말고, 사용자에게 정확한 파일 경로를 확인하세요.",
    NOT_IN_TABLE: "표 안 글자를 anchor로 지정하세요(표 내용은 hwp_table_read로 확인).",
    COM_ERROR: "같은 호출을 그대로 반복하지 마세요. hwp_get_context나 hwp_get_text로 문서 상태를 확인하고, 한글에 대화창이 떠 있는지 사용자에게 확인을 요청하세요.",
}


def err(code: str, summary: str = "", next_action: str = None) -> str:
    """표준 오류 문자열을 만듭니다. next_action이 없으면 NEXT_ACTIONS의 기본 문구를 씁니다."""
    summary = (summary or "").strip() or code
    nxt = next_action or NEXT_ACTIONS.get(code, "")
    return f"Error[{code}]: {summary}" + (f" / 다음: {nxt}" if nxt else "")


def not_running() -> str:
    return err(HWP_NOT_RUNNING, "한글(HWP) 프로그램에 연결하지 못했습니다.")


def com_error(e) -> str:
    return err(COM_ERROR, f"한글 처리 중 오류: {e}")


_CLASSIFY_RULES = [
    (UNSAVED_DOCUMENT, ("한 번도 저장", "저장된 적", "새 문서라", "경로가 없")),
    (NO_SELECTION, ("선택된 영역이 없", "선택 영역이 없", "선택된 글자가 없")),
    (NOT_IN_TABLE, ("표 안", "표 안에", "셀이 아닙", "표가 아닙")),
    (FILE_NOT_FOUND, ("파일이 없", "파일을 찾을 수 없", "존재하지 않는 파일", "No such file")),
    (AMBIGUOUS, ("곳에서 일치", "여러 곳")),
    (NOT_FOUND, ("찾지 못", "찾을 수 없", "not found", "없는 필드", "없는 스타일")),
    (NO_DOCUMENT, ("열린 문서가 없", "문서가 없습니다")),
    (INVALID_ARG, ("이어야 합니다", "필요합니다", "중 하나", "이상이어야", "비어 있", "범위를 벗어", "형식",
                   "is required", "must be", "Unsupported", "Empty data", "Failed to parse")),
]


def classify(message: str) -> str:
    """헬퍼가 돌려준 실패 문구로 오류 코드를 고릅니다. 맞는 규칙이 없으면 COM_ERROR."""
    m = message or ""
    for code, keys in _CLASSIFY_RULES:
        if any(k in m for k in keys):
            return code
    return COM_ERROR


def from_message(message: str, code: str = None) -> str:
    """실패 문구를 표준 오류로 바꿉니다. 이미 표준 형식이면 그대로."""
    m = str(message or "")
    if m.startswith("Error["):
        return m
    if m.startswith("Error:"):
        m = m[len("Error:"):].strip()
    return err(code or classify(m), m)


def is_error(text) -> bool:
    return isinstance(text, str) and text.startswith("Error")
