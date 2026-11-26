mcp 수정 후 재시작이 필요하면 나에게 요청하라.

수정작업후 알아서 적절히 커밋하라

## 한글 시험(실제 한글 조작)

- 사용자가 작업 중인 문서는 건드리지 말고, hwp_create로 만든 테스트용 문서에서 시험하라.
- 편집 전마다 활성 문서가 테스트 문서인지(DocumentID) 확인하고, 아니면 멈춰라. 사용자가 다른 문서를 보고 있으면 전환 전에 물어라.

## 액션 색인(src/tools/data/hwp_actions.json) 다시 만들기

- 한컴 공식 문서가 새로 나오면 참고자료/hancom-api에 넣고, 한글을 켠 상태에서 실행:
  `uv run --with pypdf python scripts/build_action_index.py 참고자료/hancom-api`
  (한글이 켜져 있어야 잘린 이름 합치기·매개변수 이름 보정이 됨. 없으면 --no-verify로 PDF만으로 만듦)

## 배포(PyPI)

- 버전 올리기·태그·PyPI 배포·빌드 설정 변경 작업은 docs/RELEASE.md 절차를 따르라.
