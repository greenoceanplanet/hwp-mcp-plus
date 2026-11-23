"""
기존 문서를 고치는 도구의 변경 추적 처리.

- 고치는 동안 변경 추적(IsTrackChange)을 켜서 수정 내용을 '제안'으로 남기고,
  끝나면(예외가 나도) 작업 전 값으로 되돌립니다.
- 켤지 여부 우선순위: 도구 인자 track(True/False) > 환경변수 HWP_MCP_TRACK_EDITS(on/off) > 기본 on.
"""

import functools
import inspect
import logging
import os
from contextlib import contextmanager

logger = logging.getLogger("hwp-mcp-extended")

ENV_NAME = "HWP_MCP_TRACK_EDITS"
TRACK_NOTE = "(변경추적으로 표시됨. 검토 후 hwp_track_changes_resolve로 수락/거절하세요.)"

_OFF = {"off", "0", "false", "no", "n"}
_ON = {"on", "1", "true", "yes", "y"}


def env_default() -> bool:
    """환경변수 HWP_MCP_TRACK_EDITS 값. 없거나 알 수 없는 값이면 on."""
    v = (os.environ.get(ENV_NAME) or "").strip().lower()
    if v in _OFF:
        return False
    if v in _ON or v == "":
        return True
    logger.warning(f"{ENV_NAME} 값을 알 수 없어 on으로 처리합니다: {v}")
    return True


def resolve_track(track=None) -> bool:
    """도구 인자 > 환경변수 > 기본 on."""
    if track is None:
        return env_default()
    if isinstance(track, str):
        return track.strip().lower() not in _OFF
    return bool(track)


@contextmanager
def track_scope(h, track):
    """
    작업 동안 IsTrackChange를 track 값으로 두고, 끝나면 작업 전 값으로 되돌립니다.
    track=None이면 아무것도 바꾸지 않습니다(문서 현재 설정을 따름).
    """
    if track is None or h is None:
        yield None
        return
    prev = None
    try:
        prev = h.IsTrackChange
    except Exception as e:
        logger.debug(f"IsTrackChange 읽기 실패(무시): {e}")
    try:
        h.IsTrackChange = 1 if track else 0
    except Exception as e:
        logger.warning(f"변경 추적 설정 실패: {e}")
    try:
        yield bool(track)
    finally:
        if prev is not None:
            try:
                h.IsTrackChange = prev
            except Exception as e:
                logger.warning(f"변경 추적 원래 값 복원 실패: {e}")


def _append_note(result, effective: bool):
    if not effective:
        return result
    if isinstance(result, str):
        if result.startswith("Error"):
            return result
        return f"{result}\n{TRACK_NOTE}"
    if isinstance(result, dict) and not result.get("error"):
        result.setdefault("track_changes", TRACK_NOTE)
    return result


def tracked_edit(get_com):
    """
    기존 문서를 고치는 MCP 도구용 데코레이터.
    도구 함수에는 `track: bool = None` 인자가 있어야 합니다.
    get_com: 한글 COM 객체를 돌려주는 함수(연결 실패 시 None).
    """

    def deco(fn):
        sig = inspect.signature(fn)

        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            try:
                bound = sig.bind_partial(*args, **kwargs)
                track_arg = bound.arguments.get("track")
            except TypeError:
                track_arg = kwargs.get("track")
            effective = resolve_track(track_arg)
            try:
                h = get_com()
            except Exception:
                h = None
            if h is None:
                return fn(*args, **kwargs)
            with track_scope(h, effective):
                result = fn(*args, **kwargs)
            return _append_note(result, effective)

        return wrapper

    return deco
