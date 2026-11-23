# HWP-MCP Plus (한글 Model Context Protocol)

> 이 저장소는 [jkf87/hwp-mcp](https://github.com/jkf87/hwp-mcp)를 포크해서 기능을 추가한 버전입니다.
> 기본 구조와 초기 기능은 원본 저장소에서 왔고, 이 저장소에서는 검토형 편집(변경추적·일괄 수정), 표·그림 편집, 서식·스타일·메모 도구, 한글 창 인식 안정화 등을 더했습니다.

HWP-MCP Plus는 한글 워드 프로세서(HWP)를 Claude와 같은 AI 모델이 제어할 수 있도록 해주는 Model Context Protocol(MCP) 서버입니다. 이 프로젝트는 한글 문서를 자동으로 생성, 편집, 조작하는 기능을 AI에게 제공합니다.

## 주요 기능

- 문서 생성 및 관리: 새 문서 생성, 열기, 저장 기능
- 텍스트 편집: 텍스트 삽입, 글꼴 설정, 단락 추가
- 테이블 작업: 테이블 생성, 데이터 채우기, 셀 내용 설정
- 완성된 문서 생성: 템플릿 기반 보고서 및 편지 자동 생성
- 일괄 작업: 여러 작업을 한 번에 실행하는 배치 기능
- 검토형 편집: 원문 기준 일괄 수정(`hwp_patch`), 선택 영역 교체, 변경 추적 반영·되돌림
- 문서 읽기: 구조 요약(`hwp_scan_document`), 쪽 단위 읽기, 누름틀 필드 읽기·쓰기
- 서식·스타일·메모: 서식 조건 찾기(`hwp_search_format`), 스타일 목록·적용(`hwp_list_styles`, `hwp_set_style`), 메모 넣기·읽기(`hwp_memo`)
- 기존 문서를 고치는 도구는 기본으로 변경추적을 켜고 고침(끝나면 원래 설정으로 되돌림)

## 시스템 요구사항

- Windows 운영체제
- 한글(HWP) 프로그램 설치
- Python 3.11 이상
- [uv](https://docs.astral.sh/uv/) (Python 의존성 관리)
- 필수 Python 패키지 (pyproject.toml 참조)

## 설치 방법

1. 저장소 클론:
```bash
git clone https://github.com/greenoceanplanet/hwp-mcp-plus.git
cd hwp-mcp-plus
```

2. 의존성 설치:
```bash
uv sync          # Python 의존성 (.venv 생성)
```

## 사용 방법

### Claude와 함께 사용하기

Claude 데스크톱 설정 파일에 다음과 같이 HWP-MCP Plus 서버를 등록하세요:

```json
{
  "mcpServers": {
    "hwp": {
      "command": "uv",
      "args": ["run", "--directory", "경로/hwp-mcp-plus", "hwp_mcp_stdio_server.py"],
      "env": { "HWP_MCP_TRACK_EDITS": "on" }
    }
  }
}
```

#### 변경추적 기본값 (`HWP_MCP_TRACK_EDITS`)
기존 문서를 고치는 도구(`hwp_patch`, `hwp_replace_selection`, `hwp_replace_text`, `hwp_insert_at`, `hwp_insert_text`,
`hwp_fill_cells`, `hwp_fill_table_grid`, `hwp_fill_table_with_data`, `hwp_fill_column_numbers`, `hwp_table_insert/delete`,
`hwp_copy_format`, `hwp_set_style`, `hwp_set_paragraph_shape`, 그림 넣기·바꾸기·지우기, `hwp_fields` put)는
작업하는 동안 변경추적을 켜서 고친 내용을 '제안'으로 남기고, 끝나면(오류가 나도) 작업 전 설정으로 되돌립니다.

- 우선순위: 도구 인자 `track`(True/False) > 환경변수 `HWP_MCP_TRACK_EDITS`(on/off) > 기본 on
- 프로젝트마다 다르게 쓰려면 그 프로젝트의 `.mcp.json`에서 `env`의 값을 바꿉니다.
- 새 문서를 만드는 도구(`hwp_create_document_from_text`, `hwp_create_complete_document`, `hwp_create_table_with_data` 등)는 대상이 아닙니다.
- 검토가 끝나면 `hwp_track_changes_resolve("apply_all")`(모두 수락) 또는 `"cancel_all"`(모두 거절).

#### 오류 응답 형식
실패하면 `Error[코드]: 요약 / 다음: 권장 행동` 형태로 돌려줍니다. 코드: `HWP_NOT_RUNNING`, `NO_DOCUMENT`, `NOT_FOUND`,
`AMBIGUOUS`, `NO_SELECTION`, `INVALID_ARG`, `UNSAVED_DOCUMENT`, `FILE_NOT_FOUND`, `NOT_IN_TABLE`, `COM_ERROR`.
코드별 권장 행동 문구는 `src/tools/hwp_errors.py`의 `NEXT_ACTIONS`에 모여 있습니다.

### 주요 기능 예시

#### 새 문서 생성
```python
hwp_create()
```

#### 텍스트 삽입
```python
hwp_insert_text("원하는 텍스트를 입력하세요.")
```

#### 테이블 생성 및 데이터 입력
```python
# 테이블 생성
hwp_insert_table(rows=5, cols=2)

# 테이블에 데이터 채우기
hwp_fill_table_with_data([
    ["월", "판매량"], 
    ["1월", "120"], 
    ["2월", "150"], 
    ["3월", "180"], 
    ["4월", "200"]
], has_header=True)

# 표에 연속된 숫자 채우기
hwp_fill_column_numbers(start=1, end=10, column=1, from_first_cell=True)
```

#### 문서 저장
```python
hwp_save()                      # 원래 파일에 덮어 저장(새 문서면 오류 → path 필요)
hwp_save("경로/문서명.hwpx")    # 확장자로 형식 결정: .hwp / .hwpx / .pdf
hwp_save("경로/문서명.pdf")     # PDF 내보내기(편집 중인 문서 경로는 그대로)
```

#### 검토형 편집 (변경 추적 → 사람 검토 → 반영/되돌림)
```python
hwp_scan_document()                       # 쪽 수, 표·그림 위치, 개요 제목 목록
hwp_get_text(page=2)                      # 2쪽 글자만 읽기
hwp_patch([                               # 원문 일부를 찾아 바꾸기(건별 결과 반환)
    {"old": "추진 계획을 수립함", "new": "추진 계획을 세움"},
    {"old": "2025년", "new": "2026년", "all": True},
])                                        # 기본으로 변경추적 표시(빨간 펜). track=False면 바로 반영
hwp_track_changes_resolve("apply_all")    # 모두 반영 (되돌리려면 "cancel_all")

hwp_get_selection()                       # 한글에서 선택한 글자 읽기
hwp_replace_selection("고쳐 쓴 문장")      # 선택 영역을 그 자리에서 교체
```

#### 서식 조건 찾기·스타일·메모
```python
hwp_search_format(bold=True)                          # 굵은 글자 목록(쪽 번호 포함)
hwp_search_format(text_color="red", font_size=14)     # 빨간 14pt 글자
hwp_search_format(font_name="돋움")                   # 글꼴 조건(HWPML로 판정, 쪽 번호 없음)
hwp_list_styles()                                     # 스타일 목록 + 현재 문단 스타일
hwp_set_style("개요 1", anchor="추진 배경")            # 그 글자가 있는 문단에 스타일 적용
hwp_memo("insert", text="표현 다듬기", anchor="수립함") # 본문은 두고 메모만 남기기
hwp_memo("list")                                      # 메모 내용·붙은 글자 목록
```

#### 누름틀 필드(양식 칸) 채우기
```python
hwp_fields("list")                                    # 필드 이름 목록
hwp_fields("put", values={"성명": "홍길동", "날짜": "2026. 11. 16."})
hwp_fields("get")                                     # 모든 필드 값
```

#### 일괄 작업 예시
```python
hwp_batch_operations([
    {"operation": "hwp_create"},
    {"operation": "hwp_insert_text", "params": {"text": "제목"}},
    {"operation": "hwp_set_font", "params": {"size": 20, "bold": True}},
    {"operation": "hwp_save", "params": {"path": "경로/문서명.hwp"}}
])
```

## 프로젝트 구조

```
hwp-mcp-plus/
├── hwp_mcp_stdio_server.py  # 메인 서버 스크립트
├── pyproject.toml           # Python 의존성 (uv)
├── uv.lock                  # Python 잠금 파일
├── hwp-mcp-구조설명.md       # 프로젝트 구조 설명 문서
├── src/
│   ├── tools/
│   │   ├── hwp_controller.py  # 한글 제어 핵심 컨트롤러
│   │   ├── hwp_table_tools.py # 테이블 관련 기능 전문 모듈
│   │   ├── hwp_edit_helpers.py   # 찾은 글자 기준 삽입·서식 복사·그림 헬퍼
│   │   ├── hwp_review_helpers.py # 일괄 수정·변경 추적·누름틀·문서 구조 헬퍼
│   │   ├── hwp_format_helpers.py # 서식 조건 찾기·스타일·메모 헬퍼(HWPML 읽기 포함)
│   │   ├── hwp_tracking.py       # 고치는 도구의 변경추적 켜기/원복(HWP_MCP_TRACK_EDITS)
│   │   └── hwp_errors.py         # 공통 오류 응답(Error[코드]: 요약 / 다음: 권장 행동)
│   ├── utils/                 # 유틸리티 함수
│   └── __tests__/             # 테스트 모듈
└── security_module/
    └── FilePathCheckerModuleExample.dll  # 보안 모듈
```

## 트러블슈팅

### 보안 모듈 관련 문제
기본적으로 한글 프로그램은 외부에서 파일 접근 시 보안 경고를 표시합니다. 이를 우회하기 위해 `FilePathCheckerModuleExample.dll` 모듈을 사용합니다. 만약 보안 모듈 등록에 실패해도 기능은 작동하지만, 파일 열기/저장 시 보안 대화 상자가 표시될 수 있습니다.

### 한글 연결 실패
한글 프로그램이 실행 중이지 않을 경우 연결에 실패할 수 있습니다. 한글 프로그램이 설치되어 있고 정상 작동하는지 확인하세요.

### 테이블 데이터 입력 문제
테이블에 데이터를 입력할 때 커서 위치가 예상과 다르게 동작하는 경우가 있었으나, 현재 버전에서는 이 문제가 해결되었습니다. 테이블의 모든 셀에 정확하게 데이터가 입력됩니다.

## 변경 로그

### 2026-11-16
- 검토형 편집 도구 추가: `hwp_patch`, `hwp_replace_selection`, `hwp_track_changes_resolve`
- 읽기 도구 추가: `hwp_scan_document`, `hwp_fields`, `hwp_get_text(page=…)`
- `hwp_save`: path 없으면 원래 파일에 저장(이전: 작업 폴더의 temp_document.hwp), .hwpx/.pdf 저장 지원
- 서식·스타일·메모 도구 추가: `hwp_search_format`, `hwp_list_styles`, `hwp_set_style`, `hwp_memo`
- 변경추적: 고치는 도구는 기본 on(`HWP_MCP_TRACK_EDITS`, 인자 `track`), 끝나면 원래 설정으로. `insert_text`가 변경추적을 강제로 켜던 동작 제거
- 오류 응답을 `Error[코드]: 요약 / 다음: 권장 행동`으로 통일

### 2025-03-27
- 표 생성 및 데이터 채우기 기능 개선
  - 표 안에 표가 중첩되는 문제 해결
  - 표 생성과 데이터 채우기 기능 분리
  - 표 생성 전 현재 커서 위치 확인 로직 추가
  - 기존 표에 데이터만 채우는 기능 개선
- 프로젝트 관리 개선
  - .gitignore 파일 추가 (임시 파일, 캐시 파일 등 제외)

### 2025-03-25
- 테이블 데이터 입력 기능 개선
  - 첫 번째 셀부터 정확하게 데이터 입력 가능
  - 셀 선택 및 커서 위치 설정 로직 개선
  - 텍스트 입력 시 커서 위치 유지 기능 추가
- 테이블 전용 도구 모듈(`hwp_table_tools.py`) 추가
- `hwp_fill_column_numbers` 함수에 `from_first_cell` 옵션 추가

## 라이선스

이 프로젝트는 MIT 라이선스에 따라 배포됩니다. 자세한 내용은 [LICENSE](LICENSE) 파일을 참조하세요.

## 기여 방법

1. 이슈 제보 또는 기능 제안: GitHub 이슈를 사용하세요.
2. 코드 기여: 변경사항을 포함한 Pull Request를 제출하세요.

## 관련 프로젝트

- [HWP SDK](https://www.hancom.com/product/sdk): 한글과컴퓨터의 공식 SDK
- [Cursor MCP](https://docs.cursor.com/context/model-context-protocol#configuration-locations)
- [Smithery](https://smithery.ai/server/@jkf87/hwp-mcp)

## 연락처

프로젝트 관련 문의는 GitHub 이슈로 해주세요.
