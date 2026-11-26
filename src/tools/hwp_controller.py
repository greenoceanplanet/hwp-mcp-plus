"""
한글(HWP) 문서를 제어하기 위한 컨트롤러 모듈
win32com을 이용하여 한글 프로그램을 자동화합니다.
"""

import os
import logging
import win32com.client
import win32com.client.dynamic
import win32gui
import win32con
import time
import pythoncom
from typing import Optional, List, Dict, Any, Tuple

try:
    from src.tools.hwp_tracking import track_scope
except ImportError:  # src/tools를 sys.path에 직접 넣어 실행하는 경우
    from hwp_tracking import track_scope

logger = logging.getLogger("hwp-controller")

HWP_PROGID = "HWPFrame.HwpObject"


def _late_bound(obj):
    """
    한글 COM 객체를 항상 늦은 바인딩(dynamic)으로 감쌉니다.

    win32com.client.Dispatch/GetActiveObject는 gen_py 캐시(%TEMP%\\gen_py)에 한글 타입 라이브러리가
    생성돼 있으면 조용히 이른 바인딩 객체를 돌려줍니다. 캐시는 같은 파이썬 버전의 다른 프로그램
    (예: pyhwpx)이 만들어도 공유됩니다. 이른 바인딩에서는 타입 라이브러리에 없는 이름
    (예: 일부 HParameterSet 항목)을 찾지 못해 AttributeError가 나므로, 캐시 유무와 관계없이
    늦은 바인딩으로 고정합니다.
    """
    return win32com.client.dynamic.Dispatch(getattr(obj, "_oleobj_", obj))


def _new_hwp_object():
    """새 한글 인스턴스를 늦은 바인딩으로 만듭니다."""
    return win32com.client.dynamic.Dispatch(HWP_PROGID)


def _active_hwp_object():
    """실행 중인 한글 인스턴스(GetActiveObject)를 늦은 바인딩으로 가져옵니다."""
    return _late_bound(win32com.client.GetActiveObject(HWP_PROGID))


class HwpController:
    """한글 문서를 제어하는 클래스"""

    def __init__(self):
        """한글 애플리케이션 인스턴스를 초기화합니다."""
        self.hwp = None
        self.visible = True
        self.is_hwp_running = False
        self.current_document_path = None

    # 노티/헬퍼 창 제목 (실제 문서 창이 아님)
    _HWP_NONDOC_TITLES = {"SystemResourceNotifyWindow", "MediaContextNotificationWindow"}

    @staticmethod
    def _is_hwp_window_class(cls: str) -> bool:
        """
        창 클래스명이 한글(hwp.exe) 창인지 판별합니다.

        한글 2020+ 는 클래스명이 'HwndWrapper[hwp.exe;;<GUID>]' 형태이고,
        구버전/COM 프레임은 'HwpFrame' 입니다. 둘 다 인식합니다.
        """
        if not cls:
            return False
        low = cls.lower()
        return "hwp.exe" in low or cls == "HwpFrame" or low.startswith("hwpframe")

    def _find_hwp_document_windows(self) -> List[Dict[str, Any]]:
        """
        실행 중인 한글 '문서' 창들을 찾습니다.

        보이는(visible) 창이면서 의미 있는 제목이 있는 것만 반환합니다.
        (노티/헬퍼 창과 숨은 창은 제외)

        Returns:
            List[Dict]: [{"hwnd": int, "title": str, "class": str}, ...]
        """
        results = []

        def _enum(hwnd, acc):
            try:
                cls = win32gui.GetClassName(hwnd)
                if not self._is_hwp_window_class(cls):
                    return True
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                title = win32gui.GetWindowText(hwnd)
                if not title or title in self._HWP_NONDOC_TITLES:
                    return True
                acc.append({"hwnd": hwnd, "title": title, "class": cls})
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(_enum, results)
        except Exception as e:
            logger.debug(f"한글 창 열거 실패 (무시): {e}")
        return results

    def _hwp_window_exists(self) -> bool:
        """
        실행 중인 한글 문서 창이 있는지 확인합니다.

        새 인스턴스를 만들지 않고, 이미 떠 있는 한글이 있는지만 판별합니다.

        Returns:
            bool: 한글 문서 창이 하나 이상 존재하면 True
        """
        return len(self._find_hwp_document_windows()) > 0

    def connect(self, visible: bool = True, register_security_module: bool = True) -> bool:
        """
        한글 프로그램에 연결합니다.

        실행 중인 한글 창이 있으면 그 인스턴스에 연결(attach)하며,
        절대 빈 새 창을 만들지 않습니다. 한글이 전혀 실행 중이지 않을 때만
        새 인스턴스를 시작합니다.
        """
        try:
            # 연결 전에 이미 실행 중인 한글 창이 있는지 먼저 확인
            hwp_running = self._hwp_window_exists()
            self.hwp = None

            # 1. ROT(Running Object Table)를 검색하여 한글 인스턴스 찾기
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass

            context = pythoncom.CreateBindCtx(0)
            rot = pythoncom.GetRunningObjectTable()
            
            # 실제 문서(경로가 있는 문서)가 열려있는 HWP 인스턴스를 우선 찾음
            best_hwp = None
            first_hwp = None
            for moniker in rot:
                name = moniker.GetDisplayName(context, None)
                if 'hwpobject' in name.lower():
                    try:
                        obj = rot.GetObject(moniker)
                        hwp_candidate = _late_bound(obj.QueryInterface(pythoncom.IID_IDispatch))
                        if first_hwp is None:
                            first_hwp = hwp_candidate
                        
                        doc_count = hwp_candidate.XHwpDocuments.Count
                        for i in range(doc_count):
                            doc = hwp_candidate.XHwpDocuments.Item(i)
                            if doc.Path:
                                best_hwp = hwp_candidate
                                logger.info(f"ROT에서 문서가 열려있는 한글 인스턴스 발견: {name} (Doc {i}: {doc.Path})")
                                # 문서 탭 전환만 (커서 이동 없이)
                                try:
                                    for _ in range(i):
                                        hwp_candidate.HAction.Run("WindowNext")
                                except Exception as e_act:
                                    logger.debug(f"문서 탭 전환 실패: {e_act}")
                                break
                        if best_hwp:
                            break
                    except Exception as e_rot:
                        logger.debug(f"ROT 객체 바인딩 실패: {e_rot}")

            if best_hwp:
                self.hwp = best_hwp
            elif first_hwp:
                self.hwp = first_hwp
                logger.info("ROT에서 기본 한글 인스턴스에 연결됨")

            # 2. ROT에서 찾지 못한 경우 GetActiveObject 시도
            if self.hwp is None:
                last_err = None
                for attempt in range(5):
                    try:
                        pythoncom.PumpWaitingMessages()
                        self.hwp = _active_hwp_object()
                        logger.info(f"GetActiveObject 성공 - 기존 HWP 인스턴스에 연결됨 (시도 {attempt + 1})")
                        break
                    except Exception as e:
                        last_err = e
                        time.sleep(0.3)

            # 3. 그래도 없으면 Dispatch (새로 생성)
            is_new_instance = False
            if self.hwp is None:
                if hwp_running:
                    # 한글이 떠 있는데 attach 실패 → 새 창을 만들지 않고 실패 반환
                    logger.error(f"실행 중인 한글에 연결 실패 (새 창 생성 방지): {last_err}")
                    print("실행 중인 한글 창을 인식했으나 연결에 실패했습니다. "
                          "한글 창을 한 번 클릭해 활성화한 뒤 다시 시도하세요.")
                    return False
                # 실행 중인 한글이 전혀 없을 때만 새 인스턴스 시작
                self.hwp = _new_hwp_object()
                is_new_instance = True
                logger.info("실행 중인 한글이 없어 새 HWP 인스턴스를 시작함")

            # 보안 모듈 등록 (파일 경로 체크 보안 경고창 방지)
            if register_security_module:
                try:
                    repo_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
                    module_path = os.path.join(repo_root, "security_module", "FilePathCheckerModuleExample.dll")
                    if os.path.exists(module_path):
                        self.hwp.RegisterModule("FilePathCheckerModuleExample", module_path)
                        print("보안 모듈이 등록되었습니다.")
                    else:
                        logger.warning(f"보안 모듈 DLL을 찾을 수 없음 (무시): {module_path}")
                except Exception as e:
                    print(f"보안 모듈 등록 실패 (무시하고 계속 진행): {e}")

            self.visible = visible
            # 새로 만든 인스턴스만 Visible 설정 — 기존 창은 건드리지 않아 전체화면 등 유지
            if is_new_instance:
                try:
                    self.hwp.XHwpWindows.Item(0).Visible = visible
                except Exception as e_vis:
                    logger.debug(f"Visible 설정 실패: {e_vis}")
            self.is_hwp_running = True
            return True
        except Exception as e:
            print(f"한글 프로그램 연결 실패: {e}")
            return False

    def disconnect(self) -> bool:
        """
        한글 프로그램 연결을 종료합니다.

        Returns:
            bool: 종료 성공 여부
        """
        try:
            if self.is_hwp_running:
                # HwpObject를 해제합니다
                self.hwp = None
                self.is_hwp_running = False

            return True
        except Exception as e:
            print(f"한글 프로그램 종료 실패: {e}")
            return False

    def set_message_box_mode(self, mode: int = 0x00020000) -> bool:
        """
        메시지 박스 표시 모드를 설정합니다.

        Args:
            mode (int): 메시지 박스 모드
                - 0x00000000: 기본값 (모든 메시지 박스 표시)
                - 0x00010000: 메시지 박스 표시 안함 (확인 버튼 자동 클릭)
                - 0x00020000: 메시지 박스 표시 안함 (취소 버튼 자동 클릭)
                - 0x00100000: 메시지 박스 표시 안함 (저장 안함 선택)

        Returns:
            bool: 설정 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            self.hwp.SetMessageBoxMode(mode)
            return True
        except Exception as e:
            print(f"메시지 박스 모드 설정 실패: {e}")
            return False

    def close_document(self, save: bool = False, suppress_dialog: bool = True) -> bool:
        """
        현재 문서를 닫습니다.

        Args:
            save (bool): 저장 후 닫을지 여부
            suppress_dialog (bool): 저장 확인 대화상자 표시 안함 (기본값: True)

        Returns:
            bool: 닫기 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            # 대화상자 표시 안함 설정
            if suppress_dialog:
                if save:
                    # 저장하고 닫기: 확인 버튼 자동 클릭
                    self.hwp.SetMessageBoxMode(0x00010000)
                else:
                    # 저장 안하고 닫기: 저장 안함 선택
                    self.hwp.SetMessageBoxMode(0x00100000)

            if save:
                self.hwp.HAction.Run("FileSave")

            result = self.hwp.HAction.Run("FileClose")
            self.current_document_path = None

            # 메시지 박스 모드 복원
            if suppress_dialog:
                self.hwp.SetMessageBoxMode(0x00000000)

            return bool(result)
        except Exception as e:
            print(f"문서 닫기 실패: {e}")
            # 메시지 박스 모드 복원 시도
            try:
                self.hwp.SetMessageBoxMode(0x00000000)
            except Exception as e:
                logger.debug(f"SetMessageBoxMode 복원 실패 (무시): {e}")
            return False

    def close_all_documents(self, save: bool = False, suppress_dialog: bool = True) -> bool:
        """
        모든 문서를 닫습니다.

        Args:
            save (bool): 저장 후 닫을지 여부
            suppress_dialog (bool): 저장 확인 대화상자 표시 안함 (기본값: True)

        Returns:
            bool: 닫기 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            # 대화상자 표시 안함 설정
            if suppress_dialog:
                if save:
                    self.hwp.SetMessageBoxMode(0x00010000)
                else:
                    self.hwp.SetMessageBoxMode(0x00100000)

            if save:
                self.hwp.HAction.Run("FileSaveAll")

            result = self.hwp.HAction.Run("FileCloseAll")
            self.current_document_path = None

            # 메시지 박스 모드 복원
            if suppress_dialog:
                self.hwp.SetMessageBoxMode(0x00000000)

            return bool(result)
        except Exception as e:
            print(f"모든 문서 닫기 실패: {e}")
            try:
                self.hwp.SetMessageBoxMode(0x00000000)
            except Exception as e:
                logger.debug(f"SetMessageBoxMode 복원 실패 (무시): {e}")
            return False

    def create_new_document(self) -> bool:
        """
        새 문서를 생성합니다.
        
        Returns:
            bool: 생성 성공 여부
        """
        try:
            if not self.is_hwp_running:
                self.connect()
            
            self.hwp.Run("FileNew")
            self.current_document_path = None
            return True
        except Exception as e:
            print(f"새 문서 생성 실패: {e}")
            return False

    def get_open_documents(self) -> Tuple[bool, List[Dict[str, Any]]]:
        """
        열려있는 문서 목록을 반환합니다.

        Returns:
            Tuple[bool, List[Dict]]: (성공 여부, 문서 목록)
            각 문서는 {"index": int, "path": str, "is_current": bool} 형태
        """
        try:
            if not self.is_hwp_running:
                return False, []

            documents = []

            # XHwpWindows를 사용하여 열린 윈도우(문서) 목록 가져오기
            try:
                windows = self.hwp.XHwpWindows
                window_count = windows.Count
            except Exception as e:
                logger.debug(f"XHwpWindows 접근 실패, XHwpDocuments로 폴백: {e}")
                window_count = self.hwp.XHwpDocuments.Count
                windows = self.hwp.XHwpDocuments

            # 현재 활성 문서 인덱스
            current_idx = None
            try:
                current_idx = self.hwp.CurDocIndex
            except Exception as e:
                logger.debug(f"CurDocIndex 조회 실패 (무시): {e}")

            for i in range(window_count):
                try:
                    doc = windows.Item(i)
                    doc_path = ""
                    try:
                        doc_path = doc.Path if doc.Path else "(새 문서)"
                    except Exception as e:
                        logger.debug(f"문서 경로 조회 실패: {e}")
                        doc_path = "(새 문서)"

                    is_current = (i == current_idx) if current_idx is not None else (i == 0)
                    documents.append({
                        "index": i,
                        "path": doc_path,
                        "is_current": is_current
                    })
                except Exception as e:
                    documents.append({
                        "index": i,
                        "path": f"(오류: {e})",
                        "is_current": False
                    })

            return True, documents
        except Exception as e:
            print(f"문서 목록 조회 실패: {e}")
            return False, []

    def switch_document(self, index: int) -> Tuple[bool, str]:
        """
        특정 인덱스의 문서로 전환합니다.

        Args:
            index (int): 문서 인덱스

        Returns:
            Tuple[bool, str]: (성공 여부, 메시지)
        """
        try:
            if not self.is_hwp_running:
                return False, "HWP가 실행되지 않았습니다."

            doc_count = self.hwp.XHwpDocuments.Count
            if index < 0 or index >= doc_count:
                return False, f"유효하지 않은 인덱스입니다. (0~{doc_count-1})"

            doc = self.hwp.XHwpDocuments.Item(index)

            # 여러 방법 시도
            try:
                doc.SetActive_OnlyStrongHold()
            except Exception as e1:
                logger.debug(f"SetActive_OnlyStrongHold 실패: {e1}")
                try:
                    doc.SetActive()
                except Exception as e2:
                    logger.debug(f"SetActive 실패, HAction 사용: {e2}")
                    self.hwp.HAction.Run("MoveDocBegin")
                    for _ in range(index):
                        self.hwp.HAction.Run("WindowNext")

            doc_path = doc.Path if doc.Path else "(새 문서)"
            return True, f"문서 전환 완료: {doc_path}"
        except Exception as e:
            return False, f"문서 전환 실패: {e}"

    def get_all_hwp_instances(self) -> Tuple[bool, List[Dict[str, Any]]]:
        """
        Running Object Table에서 모든 HWP 인스턴스를 찾습니다.

        Returns:
            Tuple[bool, List[Dict]]: (성공 여부, 인스턴스 목록)
            각 인스턴스는 {"index": int, "hwnd": int, "title": str, "is_current": bool} 형태
        """
        try:
            instances = []
            current_hwnd = None

            # 현재 연결된 HWP의 윈도우 핸들
            if self.hwp:
                try:
                    current_hwnd = self.hwp.XHwpWindows.Item(0).WindowHandle
                except Exception as e:
                    logger.debug(f"현재 WindowHandle 조회 실패 (무시): {e}")

            # 실제 문서 창만 찾기 (노티/헬퍼 창 제외, 최신/구버전 클래스 모두 인식)
            hwp_windows = self._find_hwp_document_windows()

            for i, win in enumerate(hwp_windows):
                instances.append({
                    "index": i,
                    "hwnd": win["hwnd"],
                    "title": win["title"],
                    "is_current": win["hwnd"] == current_hwnd if current_hwnd else False
                })

            return True, instances
        except Exception as e:
            print(f"HWP 인스턴스 목록 조회 실패: {e}")
            return False, []

    def connect_to_hwp_instance(self, hwnd: int) -> Tuple[bool, str]:
        """
        특정 HWP 윈도우에 연결합니다.

        Args:
            hwnd: 윈도우 핸들

        Returns:
            Tuple[bool, str]: (성공 여부, 메시지)
        """
        try:
            title = win32gui.GetWindowText(hwnd)

            # 기존 연결 해제
            self.hwp = None
            self.is_hwp_running = False

            # 해당 윈도우를 최상위로 가져오기
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                
                # ALT key trick to bypass Windows SetForegroundWindow restrictions
                try:
                    import win32com.client
                    shell = win32com.client.Dispatch("WScript.Shell")
                    shell.SendKeys('%')
                    time.sleep(0.1)
                except Exception as e_sendkeys:
                    logger.debug(f"SendKeys 실패: {e_sendkeys}")
                    
                win32gui.SetForegroundWindow(hwnd)
            except Exception as e:
                logger.warning(f"창 활성화 실패 (계속 진행): {e}")

            time.sleep(0.5)

            # COM 재초기화
            try:
                pythoncom.CoInitialize()
            except Exception as e:
                logger.debug(f"CoInitialize: {e}")  # 이미 초기화된 경우

            # 1. ROT에서 창 핸들(hwnd)이 일치하는 HWP 인스턴스 검색
            found_hwp = None
            try:
                context = pythoncom.CreateBindCtx(0)
                rot = pythoncom.GetRunningObjectTable()
                for moniker in rot:
                    name = moniker.GetDisplayName(context, None)
                    if 'hwpobject' in name.lower():
                        try:
                            obj = rot.GetObject(moniker)
                            hwp_candidate = _late_bound(obj.QueryInterface(pythoncom.IID_IDispatch))
                            hwp_hwnd = hwp_candidate.XHwpWindows.Item(0).WindowHandle
                            if hwp_hwnd == hwnd:
                                found_hwp = hwp_candidate
                                logger.info(f"ROT에서 요청된 hwnd({hwnd})와 일치하는 한글 인스턴스 발견: {name}")
                                break
                        except Exception:
                            pass
            except Exception as e_rot:
                logger.debug(f"ROT 검색 중 오류 (일반 매칭으로 폴백): {e_rot}")

            if found_hwp:
                self.hwp = found_hwp
                self.is_hwp_running = True
                return True, f"HWP 인스턴스(hwnd: {hwnd})에 연결됨: {title}"

            # 방법 1: GetActiveObject 시도
            try:
                self.hwp = _active_hwp_object()
                self.is_hwp_running = True
                logger.info(f"GetActiveObject 성공: {title}")
                return True, f"HWP 인스턴스에 연결됨: {title}"
            except Exception as e:
                logger.warning(f"GetActiveObject 실패: {e}")

            # 방법 2: Dispatch로 연결 (활성화된 HWP에 연결됨)
            try:
                self.hwp = _new_hwp_object()
                self.is_hwp_running = True
                logger.info(f"Dispatch로 연결됨")
                # Dispatch 후 현재 문서 경로로 확인
                try:
                    current_path = self.hwp.Path
                    return True, f"HWP에 연결됨: {title} (문서: {current_path or '새 문서'})"
                except Exception as e:
                    logger.debug(f"Path 가져오기 실패: {e}")
                    return True, f"HWP에 연결됨: {title}"
            except Exception as e:
                logger.error(f"Dispatch 실패: {e}")
                return False, f"연결 실패: {e}"

        except Exception as e:
            return False, f"연결 실패: {e}"

    def close_hwp_window(self, hwnd: int) -> Tuple[bool, str]:
        """
        HWP 윈도우를 닫습니다 (WM_CLOSE 메시지 전송).

        Args:
            hwnd: 윈도우 핸들

        Returns:
            Tuple[bool, str]: (성공 여부, 메시지)
        """
        try:
            title = win32gui.GetWindowText(hwnd)
            WM_CLOSE = 0x0010
            win32gui.PostMessage(hwnd, WM_CLOSE, 0, 0)
            return True, f"창 닫기 요청: {title}"
        except Exception as e:
            return False, f"창 닫기 실패: {e}"

    def open_document(self, file_path: str) -> bool:
        """
        문서를 엽니다.

        Args:
            file_path (str): 열 문서의 경로

        Returns:
            bool: 열기 성공 여부
        """
        try:
            if not self.is_hwp_running:
                self.connect()

            abs_path = os.path.abspath(file_path)
            print(f"[DEBUG] Opening document: {abs_path}")
            print(f"[DEBUG] File exists: {os.path.exists(abs_path)}")

            # Use HAction with FileOpen for reliable file opening
            pset = self.hwp.HParameterSet.HFileOpenSave
            self.hwp.HAction.GetDefault("FileOpen", pset.HSet)
            pset.filename = abs_path
            pset.Format = "HWP"
            result = self.hwp.HAction.Execute("FileOpen", pset.HSet)
            print(f"[DEBUG] FileOpen result: {result}")
            if result:
                self.current_document_path = abs_path
            return result
        except Exception as e:
            print(f"문서 열기 실패: {e}")
            import traceback
            traceback.print_exc()
            return False

    def save_document(self, file_path: Optional[str] = None) -> bool:
        """
        문서를 저장합니다.
        
        Args:
            file_path (str, optional): 저장할 경로. None이면 현재 경로에 저장.
            
        Returns:
            bool: 저장 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            if file_path:
                abs_path = os.path.abspath(file_path)
                fmt = self.save_format_for(abs_path)
                # 파일 형식과 경로 모두 지정하여 저장
                result = self.hwp.SaveAs(abs_path, fmt, "")
                if result is False:
                    return False
                if fmt != "PDF":  # PDF는 내보내기일 뿐 편집 중인 문서 경로는 그대로
                    self.current_document_path = abs_path
            else:
                if self.current_document_path:
                    self.hwp.Save()
                else:
                    # 저장 대화 상자 표시 (파라미터 없이 호출)
                    self.hwp.SaveAs()
                    # 대화 상자에서 사용자가 선택한 경로를 알 수 없으므로 None 유지
            
            return True
        except Exception as e:
            print(f"문서 저장 실패: {e}")
            return False

    @staticmethod
    def save_format_for(path: str) -> str:
        """확장자로 SaveAs 형식 이름을 정합니다(.pdf→PDF, .hwpx→HWPX, 그 밖→HWP)."""
        ext = os.path.splitext(path)[1].lower()
        return {".pdf": "PDF", ".hwpx": "HWPX"}.get(ext, "HWP")

    def get_active_document_path(self) -> str:
        """현재 편집 중인 문서의 파일 경로. 한 번도 저장하지 않은 새 문서면 ""."""
        try:
            path = self.hwp.XHwpDocuments.Active_XHwpDocument.FullName
            if path:
                return str(path)
        except Exception:
            pass
        try:
            return str(self.hwp.Path or "")
        except Exception:
            return ""

    def save_current_document(self) -> Tuple[bool, str]:
        """
        현재 문서를 원래 파일에 저장합니다(덮어쓰기).
        새 문서(경로 없음)면 저장하지 않고 실패를 돌려줍니다(대화창을 띄우지 않음).
        """
        try:
            if not self.is_hwp_running:
                return False, "한글이 실행 중이 아닙니다."
            path = self.get_active_document_path()
            if not path:
                return False, "아직 저장한 적 없는 새 문서입니다. path를 지정하세요."
            result = self.hwp.Save(False)  # False=변경 여부와 관계없이 저장(True면 변경 없을 때 False 반환)
            if result is False:
                return False, f"저장 실패: {path}"
            self.current_document_path = path
            return True, path
        except Exception as e:
            logger.error(f"현재 문서 저장 실패: {e}")
            return False, str(e)

    def insert_text(self, text: str, preserve_linebreaks: bool = True, track: Optional[bool] = None) -> bool:
        """
        현재 커서 위치에 텍스트를 삽입합니다.

        Args:
            text (str): 삽입할 텍스트
            preserve_linebreaks (bool): 줄바꿈 유지 여부
            track (bool|None): None이면 문서의 현재 변경 추적 설정을 그대로 따름.
                True/False면 삽입하는 동안만 그 값으로 두고 끝나면 원래 값으로 되돌림.

        Returns:
            bool: 삽입 성공 여부
        """
        if not self.is_hwp_running:
            return False
        with track_scope(self.hwp, track):
            return self._insert_text_impl(text, preserve_linebreaks)

    def _insert_text_impl(self, text: str, preserve_linebreaks: bool = True) -> bool:
        try:
            # 블록(선택 영역)이 있는 경우 먼저 삭제(변경 추적이 켜져 있으면 삭제 표시) 진행
            try:
                selected = self.get_selected_text()
                if selected:
                    logger.info("선택 영역 감지됨. 삭제(취소선) 처리 후 삽입합니다.")
                    
                    # 만약 선택 영역이 말머리표/번호로 시작하고, 삽입할 텍스트도 말머리표/번호로 시작하면 중복 방지를 위해 삽입 텍스트에서 제거
                    if self._has_leading_bullet(selected) and self._has_leading_bullet(text):
                        logger.info("말머리표 중복 감지: 삽입 텍스트에서 말머리표를 제거합니다.")
                        text = self._strip_leading_bullet(text)
                        
                    self.hwp.Run("Delete")
            except Exception as e_del:
                logger.warning(f"선택 영역 사전 삭제 실패: {e_del}")

            if preserve_linebreaks and '\n' in text:
                # 줄바꿈이 포함된 경우 줄 단위로 처리
                lines = text.split('\n')
                for i, line in enumerate(lines):
                    if i > 0:  # 첫 줄이 아니면 줄바꿈 추가
                        self.insert_paragraph()
                    if line.strip():  # 빈 줄이 아니면 텍스트 삽입
                        self._insert_text_direct(line)
                return True
            else:
                # 줄바꿈이 없거나 유지하지 않는 경우 한 번에 처리
                return self._insert_text_direct(text)
        except Exception as e:
            print(f"텍스트 삽입 실패: {e}")
            return False

    def _set_table_cursor(self) -> bool:
        """
        표 안에서 커서 위치를 제어하는 내부 메서드입니다.
        현재 셀을 선택하고 취소하여 커서를 셀 안에 위치시킵니다.
        
        Returns:
            bool: 성공 여부
        """
        try:
            # 현재 셀 선택
            self.hwp.Run("TableSelCell")
            # 선택 취소 (커서는 셀 안에 위치)
            self.hwp.Run("Cancel")
            # 셀 내부로 커서 이동을 확실히
            self.hwp.Run("CharRight")
            self.hwp.Run("CharLeft")
            return True
        except Exception as e:
            logger.debug(f"셀 내부 커서 이동 실패: {e}")
            return False

    def _insert_text_direct(self, text: str) -> bool:
        """
        텍스트를 직접 삽입하는 내부 메서드입니다.
        
        Args:
            text (str): 삽입할 텍스트
            
        Returns:
            bool: 삽입 성공 여부
        """
        try:
            # 텍스트 삽입을 위한 액션 초기화
            self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            self.hwp.HParameterSet.HInsertText.Text = text
            self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            return True
        except Exception as e:
            print(f"텍스트 직접 삽입 실패: {e}")
            return False

    def set_font(self, font_name: str, font_size: int, bold: bool = False, italic: bool = False, 
                select_previous_text: bool = False) -> bool:
        """
        글꼴 속성을 설정합니다. 현재 위치에서 다음에 입력할 텍스트에 적용됩니다.
        
        Args:
            font_name (str): 글꼴 이름
            font_size (int): 글꼴 크기
            bold (bool): 굵게 여부
            italic (bool): 기울임꼴 여부
            select_previous_text (bool): 이전에 입력한 텍스트를 선택할지 여부
            
        Returns:
            bool: 설정 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 새로운 구현: set_font_style 메서드 사용
            return self.set_font_style(
                font_name=font_name,
                font_size=font_size,
                bold=bold,
                italic=italic,
                underline=False,
                select_previous_text=select_previous_text
            )
        except Exception as e:
            print(f"글꼴 설정 실패: {e}")
            return False

    def set_font_style(self, font_name: str = None, font_size: int = None, 
                     bold: bool = False, italic: bool = False, underline: bool = False,
                     select_previous_text: bool = False) -> bool:
        """
        현재 선택된 텍스트의 글꼴 스타일을 설정합니다.
        선택된 텍스트가 없으면, 다음 입력될 텍스트에 적용됩니다.
        
        Args:
            font_name (str, optional): 글꼴 이름. None이면 현재 글꼴 유지.
            font_size (int, optional): 글꼴 크기. None이면 현재 크기 유지.
            bold (bool): 굵게 여부
            italic (bool): 기울임꼴 여부
            underline (bool): 밑줄 여부
            select_previous_text (bool): 이전에 입력한 텍스트를 선택할지 여부
            
        Returns:
            bool: 설정 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 이전 텍스트 선택 옵션이 활성화된 경우 현재 단락의 이전 텍스트 선택
            if select_previous_text:
                self.select_last_text()
            
            # 글꼴 설정을 위한 액션 초기화
            self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
            
            # 글꼴 이름 설정
            if font_name:
                self.hwp.HParameterSet.HCharShape.FaceNameHangul = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameLatin = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameHanja = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameJapanese = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameOther = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameSymbol = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameUser = font_name
            
            # 글꼴 크기 설정 (hwpunit, 10pt = 1000)
            if font_size:
                self.hwp.HParameterSet.HCharShape.Height = font_size * 100
            
            # 스타일 설정
            self.hwp.HParameterSet.HCharShape.Bold = bold
            self.hwp.HParameterSet.HCharShape.Italic = italic
            self.hwp.HParameterSet.HCharShape.UnderlineType = 1 if underline else 0
            
            # 변경사항 적용
            self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
            
            return True
            
        except Exception as e:
            print(f"글꼴 스타일 설정 실패: {e}")
            return False

    def _get_current_position(self):
        """현재 커서 위치 정보를 가져옵니다."""
        try:
            # GetPos()는 현재 위치 정보를 (위치 유형, List ID, Para ID, CharPos)의 튜플로 반환
            return self.hwp.GetPos()
        except Exception as e:
            logger.debug(f"GetPos 실패: {e}")
            return None

    def _set_position(self, pos):
        """커서 위치를 지정된 위치로 변경합니다."""
        try:
            if pos:
                self.hwp.SetPos(*pos)
            return True
        except Exception as e:
            logger.debug(f"SetPos 실패: {e}")
            return False

    def insert_table(self, rows: int, cols: int) -> bool:
        """
        현재 커서 위치에 표를 삽입합니다.
        
        Args:
            rows (int): 행 수
            cols (int): 열 수
            
        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            self.hwp.HAction.GetDefault("TableCreate", self.hwp.HParameterSet.HTableCreation.HSet)
            self.hwp.HParameterSet.HTableCreation.Rows = rows
            self.hwp.HParameterSet.HTableCreation.Cols = cols
            self.hwp.HParameterSet.HTableCreation.WidthType = 0  # 0: 단에 맞춤, 1: 절대값
            self.hwp.HParameterSet.HTableCreation.HeightType = 1  # 0: 자동, 1: 절대값
            self.hwp.HParameterSet.HTableCreation.WidthValue = 0  # 단에 맞춤이므로 무시됨
            self.hwp.HParameterSet.HTableCreation.HeightValue = 1000  # 셀 높이(hwpunit)
            
            # 각 열의 너비를 설정 (모두 동일하게)
            # PageWidth 대신 고정 값 사용
            col_width = 8000 // cols  # 전체 너비를 열 수로 나눔
            self.hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", cols)
            for i in range(cols):
                self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(i, col_width)
                
            self.hwp.HAction.Execute("TableCreate", self.hwp.HParameterSet.HTableCreation.HSet)
            return True
        except Exception as e:
            print(f"표 삽입 실패: {e}")
            return False

    def undo(self, count: int = 1) -> Tuple[bool, str]:
        """
        실행 취소(Undo)를 수행합니다.

        Args:
            count (int): 취소할 횟수 (기본값: 1)

        Returns:
            Tuple[bool, str]: (성공 여부, 메시지)
        """
        try:
            if not self.is_hwp_running:
                return False, "HWP가 실행되지 않았습니다."

            success_count = 0
            for _ in range(count):
                result = self.hwp.HAction.Run("Undo")
                if result:
                    success_count += 1
                else:
                    break

            if success_count == count:
                return True, f"실행 취소 {success_count}회 완료"
            elif success_count > 0:
                return True, f"실행 취소 {success_count}회 완료 (요청: {count}회)"
            else:
                return False, "실행 취소할 항목이 없습니다."
        except Exception as e:
            return False, f"실행 취소 실패: {e}"

    def redo(self, count: int = 1) -> Tuple[bool, str]:
        """
        다시 실행(Redo)을 수행합니다.

        Args:
            count (int): 다시 실행할 횟수 (기본값: 1)

        Returns:
            Tuple[bool, str]: (성공 여부, 메시지)
        """
        try:
            if not self.is_hwp_running:
                return False, "HWP가 실행되지 않았습니다."

            success_count = 0
            for _ in range(count):
                result = self.hwp.HAction.Run("Redo")
                if result:
                    success_count += 1
                else:
                    break

            if success_count == count:
                return True, f"다시 실행 {success_count}회 완료"
            elif success_count > 0:
                return True, f"다시 실행 {success_count}회 완료 (요청: {count}회)"
            else:
                return False, "다시 실행할 항목이 없습니다."
        except Exception as e:
            return False, f"다시 실행 실패: {e}"

    def find_text(self, text: str) -> bool:
        """
        문서에서 텍스트를 찾습니다.

        Args:
            text (str): 찾을 텍스트

        Returns:
            bool: 찾기 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            # 문서 처음으로 이동
            self.hwp.HAction.Run("MoveDocBegin")

            # HAction으로 찾기
            pset = self.hwp.HParameterSet.HFindReplace
            self.hwp.HAction.GetDefault("RepeatFind", pset.HSet)
            pset.FindString = text
            pset.FindRegExp = 0
            pset.IgnoreMessage = 1
            pset.Direction = 0  # 0: forward
            result = self.hwp.HAction.Execute("RepeatFind", pset.HSet)
            return bool(result)
        except Exception as e:
            print(f"텍스트 찾기 실패: {e}")
            return False

    def get_text(self) -> str:
        """
        현재 문서의 전체 텍스트를 가져옵니다.
        
        Returns:
            str: 문서 텍스트
        """
        try:
            if not self.is_hwp_running:
                return ""
            
            # 선택 영역이 있는 경우 GetTextFile(..., "")가 None을 반환하므로 선택 해제
            # (_get_text_file_safe가 TEXT 요청을 UNICODE로 읽어 한글 깨짐을 막음)
            try:
                self.hwp.Run("Cancel")
            except Exception:
                pass

            return self._get_text_file_safe("TEXT", "")
        except Exception as e:
            logger.error(f"텍스트 가져오기 실패: {e}")
            return ""

    def get_text_range(self, start_text: str = "", max_chars: int = 0) -> Tuple[bool, str]:
        """
        문서 전체 텍스트에서 start_text가 처음 나오는 곳부터 max_chars 글자만 돌려줍니다.

        - start_text가 비면 문서 처음부터.
        - max_chars가 0 이하면 끝까지.
        전체 텍스트는 get_text()(→ _get_text_file_safe)로 읽으므로 변경 추적 문서에서도 대화창이 뜨지 않습니다.

        Returns:
            (성공여부, 텍스트 또는 오류 메시지)
        """
        text = self.get_text()
        if text is None:
            return False, "문서 텍스트를 읽지 못했습니다."
        start = 0
        if start_text:
            start = text.find(start_text)
            if start < 0:
                return False, f"시작 글자를 찾지 못했습니다: {start_text}"
        part = text[start:]
        if max_chars and max_chars > 0:
            part = part[:max_chars]
        return True, part

    def set_page_setup(self, orientation: str = "portrait", margin_left: int = 1000, 
                     margin_right: int = 1000, margin_top: int = 1000, margin_bottom: int = 1000) -> bool:
        """
        페이지 설정을 변경합니다.
        
        Args:
            orientation (str): 용지 방향 ('portrait' 또는 'landscape')
            margin_left (int): 왼쪽 여백(hwpunit)
            margin_right (int): 오른쪽 여백(hwpunit)
            margin_top (int): 위쪽 여백(hwpunit)
            margin_bottom (int): 아래쪽 여백(hwpunit)
            
        Returns:
            bool: 설정 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 매크로 명령 사용
            orient_val = 0 if orientation.lower() == "portrait" else 1
            
            # 페이지 설정 매크로
            result = self.hwp.Run(f"PageSetup3 {orient_val} {margin_left} {margin_right} {margin_top} {margin_bottom}")
            return bool(result)
        except Exception as e:
            print(f"페이지 설정 실패: {e}")
            return False

    def insert_paragraph(self) -> bool:
        """
        새 단락을 삽입합니다.
        
        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            self.hwp.HAction.Run("BreakPara")
            return True
        except Exception as e:
            print(f"단락 삽입 실패: {e}")
            return False

    def select_all(self) -> bool:
        """
        문서 전체를 선택합니다.
        
        Returns:
            bool: 선택 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            self.hwp.Run("SelectAll")
            return True
        except Exception as e:
            print(f"전체 선택 실패: {e}")
            return False

    def fill_cell_field(self, field_name: str, value: str, n: int = 1) -> bool:
        """
        동일한 이름의 셀필드 중 n번째에만 값을 채웁니다.
        위키독스 예제: https://wikidocs.net/261646
        
        Args:
            field_name (str): 필드 이름
            value (str): 채울 값
            n (int): 몇 번째 필드에 값을 채울지 (1부터 시작)
            
        Returns:
            bool: 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
                
            # 1. 필드 목록 가져오기
            # HGO_GetFieldList은 현재 문서에 있는 모든 필드 목록을 가져옵니다.
            self.hwp.HAction.GetDefault("HGo_GetFieldList", self.hwp.HParameterSet.HGo.HSet)
            self.hwp.HAction.Execute("HGo_GetFieldList", self.hwp.HParameterSet.HGo.HSet)
            
            # 2. 필드 이름이 동일한 모든 셀필드 찾기
            field_list = []
            field_count = self.hwp.HParameterSet.HGo.FieldList.Count
            
            for i in range(field_count):
                field_info = self.hwp.HParameterSet.HGo.FieldList.Item(i)
                if field_info.FieldName == field_name:
                    field_list.append((field_info.FieldName, i))
            
            # 3. n번째 필드가 존재하는지 확인 (인덱스는 0부터 시작하므로 n-1)
            if len(field_list) < n:
                print(f"해당 이름의 필드가 충분히 없습니다. 필요: {n}, 존재: {len(field_list)}")
                return False
                
            # 4. n번째 필드의 위치로 이동
            target_field_idx = field_list[n-1][1]
            
            # HGo_SetFieldText를 사용하여 해당 필드 위치로 이동한 후 텍스트 설정
            self.hwp.HAction.GetDefault("HGo_SetFieldText", self.hwp.HParameterSet.HGo.HSet)
            self.hwp.HParameterSet.HGo.HSet.SetItem("FieldIdx", target_field_idx)
            self.hwp.HParameterSet.HGo.HSet.SetItem("Text", value)
            self.hwp.HAction.Execute("HGo_SetFieldText", self.hwp.HParameterSet.HGo.HSet)
            
            return True
        except Exception as e:
            print(f"셀필드 값 채우기 실패: {e}")
            return False
        
    def select_last_text(self) -> bool:
        """
        현재 단락의 마지막으로 입력된 텍스트를 선택합니다.
        
        Returns:
            bool: 선택 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 현재 위치 저장
            current_pos = self.hwp.GetPos()
            if not current_pos:
                return False
                
            # 현재 단락의 시작으로 이동
            self.hwp.Run("MoveLineStart")
            start_pos = self.hwp.GetPos()
            
            # 이전 위치로 돌아가서 선택 영역 생성
            self.hwp.SetPos(*start_pos)
            self.hwp.SelectText(start_pos, current_pos)
            
            return True
        except Exception as e:
            print(f"텍스트 선택 실패: {e}")
            return False

    def _get_text_file_safe(self, fmt: str, option: str) -> str:
        """
        GetTextFile 호출 전후로 메시지박스를 억제합니다.

        변경 추적이 켜진 문서에서 TEXT 형식으로 읽을 때
        '형식 변환 시 변경추적 정보 손실' 경고창이 뜨는 것을 방지합니다.

        "TEXT" 형식은 CP949 바이트를 한 바이트씩 글자로 담아 돌려주므로
        한글이 'ÁÖ°Å¾ÈÁ¤'처럼 깨집니다. 그래서 "TEXT" 요청은 "UNICODE"로 읽고,
        UNICODE가 실패할 때만 TEXT로 읽어 latin-1 → cp949로 되돌립니다.
        """
        try:
            self.hwp.SetMessageBoxMode(0x00010000)  # 대화창 자동 확인(OK)
        except Exception:
            pass
        try:
            if fmt.upper() == "TEXT":
                result = None
                try:
                    result = self.hwp.GetTextFile("UNICODE", option)
                except Exception as e:
                    logger.warning(f"GetTextFile UNICODE 실패, TEXT로 재시도: {e}")
                if result is None:
                    result = self._fix_ansi_mojibake(self.hwp.GetTextFile("TEXT", option))
            else:
                result = self.hwp.GetTextFile(fmt, option)
        finally:
            try:
                self.hwp.SetMessageBoxMode(0x00000000)
            except Exception:
                pass
        return result or ""

    @staticmethod
    def _fix_ansi_mojibake(text):
        """CP949 바이트가 latin-1 글자로 담긴 문자열을 원래 한글로 되돌립니다."""
        if not text:
            return text
        try:
            return text.encode("latin-1").decode("cp949")
        except (UnicodeEncodeError, UnicodeDecodeError):
            return text

    def get_selected_text(self) -> str:
        """
        현재 선택(블록 지정)된 텍스트만 반환합니다.

        사용자가 한글 창에서 마우스로 드래그해 선택한 영역의 텍스트를 가져옵니다.
        선택된 영역이 없으면 빈 문자열을 반환합니다.

        Returns:
            str: 선택된 텍스트 (없으면 "")
        """
        try:
            if not self.is_hwp_running:
                return ""

            return self._get_text_file_safe("TEXT", "saveblock")
        except Exception as e:
            print(f"선택 텍스트 가져오기 실패: {e}")
            return ""

    def get_current_paragraph_text(self) -> str:
        """
        커서가 위치한 현재 문단의 전체 텍스트를 반환합니다.

        커서 위치를 보존합니다(읽기만 하고 원위치로 복귀).

        Returns:
            str: 현재 문단 텍스트 (실패 시 "")
        """
        try:
            if not self.is_hwp_running:
                return ""

            saved_pos = self.hwp.GetPos()

            # 문단 시작으로 이동 후, 문단 끝까지 선택
            self.hwp.Run("MoveParaBegin")
            self.hwp.Run("MoveSelParaEnd")
            text = self._get_text_file_safe("TEXT", "saveblock")

            # 선택 해제 및 커서 원위치 복원
            self.hwp.Run("Cancel")
            if saved_pos:
                self.hwp.SetPos(*saved_pos)

            return text
        except Exception as e:
            print(f"현재 문단 텍스트 가져오기 실패: {e}")
            return ""

    def get_cursor_context(self) -> dict:
        """
        현재 커서를 기준으로 같은 문단 내 앞쪽/뒤쪽 텍스트(문맥)를 반환합니다.

        AI가 "이어쓰기"나 "이 위치 기준 교정"을 할 때 맥락 파악용으로 사용합니다.
        커서 위치를 보존합니다.

        Returns:
            dict: {"before": 앞쪽 텍스트, "after": 뒤쪽 텍스트}
        """
        try:
            if not self.is_hwp_running:
                return {"before": "", "after": ""}

            saved_pos = self.hwp.GetPos()

            # 뒤쪽 문맥: 커서 → 문단 끝까지 선택
            self.hwp.Run("MoveSelParaEnd")
            after = self._get_text_file_safe("TEXT", "saveblock")
            self.hwp.Run("Cancel")
            if saved_pos:
                self.hwp.SetPos(*saved_pos)

            # 앞쪽 문맥: 커서 → 문단 시작까지 선택
            self.hwp.Run("MoveSelParaBegin")
            before = self._get_text_file_safe("TEXT", "saveblock")
            self.hwp.Run("Cancel")
            if saved_pos:
                self.hwp.SetPos(*saved_pos)

            return {"before": before, "after": after}
        except Exception as e:
            print(f"커서 문맥 가져오기 실패: {e}")
            return {"before": "", "after": ""}

    def fill_cell_next_to_label(
        self,
        label: str,
        value: str,
        direction: str = "right",
        occurrence: int = 1,
        mode: str = "replace"
    ) -> Tuple[bool, str]:
        """
        표에서 레이블을 찾아 옆 셀에 값을 입력합니다.

        Args:
            label (str): 찾을 레이블 텍스트 (예: "성명")
            value (str): 입력할 값 (예: "장예준")
            direction (str): 이동 방향 - "right"(오른쪽), "down"(아래), "left"(왼쪽), "up"(위)
                - 레이블이 왼쪽에 있고 값이 오른쪽에 있으면 "right" 사용
                - 레이블이 위에 있고 값이 아래에 있으면 "down" 사용
            occurrence (int): 동일 레이블 중 몇 번째를 사용할지 (1부터 시작, 기본값: 1)
            mode (str): 입력 모드 - "replace"(기존 내용 삭제 후 입력), "prepend"(앞에 추가), "append"(뒤에 추가)

        Returns:
            Tuple[bool, str]: (성공 여부, 결과 메시지)
        """
        try:
            if not self.is_hwp_running:
                return False, "HWP가 연결되어 있지 않습니다."

            # 1. 문서 처음으로 이동
            self.hwp.HAction.Run("MoveDocBegin")

            # 2. 레이블 찾기 (occurrence 횟수만큼 반복)
            found = False
            for i in range(occurrence):
                pset = self.hwp.HParameterSet.HFindReplace
                self.hwp.HAction.GetDefault("RepeatFind", pset.HSet)
                pset.FindString = label
                pset.FindRegExp = 0
                pset.IgnoreMessage = 1
                pset.Direction = 0  # forward
                result = self.hwp.HAction.Execute("RepeatFind", pset.HSet)

                if not result:
                    if i == 0:
                        return False, f"레이블 '{label}'을(를) 찾을 수 없습니다."
                    else:
                        return False, f"레이블 '{label}'의 {occurrence}번째 항목을 찾을 수 없습니다. (총 {i}개 발견)"
                found = True

            if not found:
                return False, f"레이블 '{label}'을(를) 찾을 수 없습니다."

            # 3. 현재 셀(레이블 셀) 전체 선택 후 해제 - 커서 위치 확정
            self.hwp.HAction.Run("TableSelCell")
            self.hwp.HAction.Run("Cancel")

            # 4. 지정된 방향으로 옆 셀로 이동
            direction_lower = direction.lower()
            if direction_lower == "right":
                self.hwp.HAction.Run("TableRightCell")
            elif direction_lower == "left":
                self.hwp.HAction.Run("TableLeftCell")
            elif direction_lower == "down":
                self.hwp.HAction.Run("MoveDown")
            elif direction_lower == "up":
                self.hwp.HAction.Run("TableUpperCell")
            else:
                return False, f"잘못된 방향입니다: {direction}. 'right', 'left', 'down', 'up' 중 하나를 사용하세요."

            # 5. mode에 따라 값 입력
            mode_lower = mode.lower()
            if mode_lower == "replace":
                # 셀 전체 내용 선택 후 잘라내기
                self.hwp.HAction.Run("SelectAll")
                self.hwp.HAction.Run("EditCut")
                self._insert_text_direct(value)
            elif mode_lower == "prepend":
                # 셀 시작으로 이동 후 입력
                self.hwp.HAction.Run("MoveSelCellBegin")
                self.hwp.HAction.Run("Cancel")
                self._insert_text_direct(value)
            elif mode_lower == "append":
                # 셀 끝으로 이동: 전체 선택 후 오른쪽으로 이동하면 끝으로 감
                self.hwp.HAction.Run("SelectAll")
                self.hwp.HAction.Run("Cancel")
                self.hwp.HAction.Run("MoveLineEnd")
                self._insert_text_direct(value)
            else:
                return False, f"잘못된 mode입니다: {mode}. 'replace', 'prepend', 'append' 중 하나를 사용하세요."

            return True, f"'{label}' 옆 셀에 '{value}' 입력 완료"

        except Exception as e:
            print(f"셀 채우기 실패: {e}")
            return False, f"셀 채우기 실패: {str(e)}"

    def fill_cells_from_dict(
        self,
        label_value_map: Dict[str, str],
        direction: str = "right"
    ) -> Dict[str, Tuple[bool, str]]:
        """
        여러 레이블에 대해 옆 셀에 값을 입력합니다.

        Args:
            label_value_map (Dict[str, str]): 레이블과 값의 매핑 (예: {"성명": "장예준", "기업명": "mutual"})
            direction (str): 이동 방향 - "right", "down", "left", "up" (기본값: "right")

        Returns:
            Dict[str, Tuple[bool, str]]: 각 레이블에 대한 (성공 여부, 결과 메시지) 딕셔너리
        """
        results = {}

        for label, value in label_value_map.items():
            success, message = self.fill_cell_next_to_label(label, value, direction)
            results[label] = (success, message)

        return results

    def fill_table_with_data(self, data: List[List[str]], start_row: int = 1, start_col: int = 1, has_header: bool = False) -> bool:
        """
        현재 커서 위치의 표에 데이터를 채웁니다.
        
        Args:
            data (List[List[str]]): 채울 데이터 2차원 리스트 (행 x 열)
            start_row (int): 시작 행 번호 (1부터 시작)
            start_col (int): 시작 열 번호 (1부터 시작)
            has_header (bool): 첫 번째 행을 헤더로 처리할지 여부
            
        Returns:
            bool: 작업 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
                
            # 현재 위치 저장 (나중에 복원을 위해)
            original_pos = self.hwp.GetPos()
            
            # 1. 표 첫 번째 셀로 이동
            self.hwp.Run("TableSelCell")  # 현재 셀 선택
            self.hwp.Run("TableSelTable") # 표 전체 선택
            self.hwp.Run("Cancel")        # 선택 취소 (커서는 표의 시작 부분에 위치)
            self.hwp.Run("TableSelCell")  # 첫 번째 셀 선택
            self.hwp.Run("Cancel")        # 선택 취소
            
            # 시작 위치로 이동
            for _ in range(start_row - 1):
                self.hwp.Run("TableLowerCell")
                
            for _ in range(start_col - 1):
                self.hwp.Run("TableRightCell")
            
            # 데이터 채우기
            for row_idx, row_data in enumerate(data):
                for col_idx, cell_value in enumerate(row_data):
                    # 셀 선택 및 내용 삭제
                    self.hwp.Run("TableSelCell")
                    self.hwp.Run("Delete")
                    
                    # 셀에 값 입력
                    if has_header and row_idx == 0:
                        self.set_font_style(bold=True)
                        self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
                        self.hwp.HParameterSet.HInsertText.Text = cell_value
                        self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
                        self.set_font_style(bold=False)
                    else:
                        self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
                        self.hwp.HParameterSet.HInsertText.Text = cell_value
                        self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
                    
                    # 다음 셀로 이동 (마지막 셀이 아닌 경우)
                    if col_idx < len(row_data) - 1:
                        self.hwp.Run("TableRightCell")
                
                # 다음 행으로 이동 (마지막 행이 아닌 경우)
                if row_idx < len(data) - 1:
                    for _ in range(len(row_data) - 1):
                        self.hwp.Run("TableLeftCell")
                    self.hwp.Run("TableLowerCell")
            
            # 표 밖으로 커서 이동
            self.hwp.Run("TableSelCell")  # 현재 셀 선택
            self.hwp.Run("Cancel")        # 선택 취소
            self.hwp.Run("MoveDown")      # 아래로 이동
            
            return True

        except Exception as e:
            print(f"표 데이터 채우기 실패: {e}")
            return False

    def _move_direction(self, direction: str) -> bool:
        """
        지정된 방향으로 셀 이동.

        Args:
            direction: 이동 방향 ("right", "left", "down", "up")

        Returns:
            bool: 성공 여부
        """
        move_actions = {
            "right": "TableRightCell",
            "left": "TableLeftCell",
            "down": "TableLowerCell",
            "up": "TableUpperCell"
        }
        action = move_actions.get(direction.lower())
        if action:
            self.hwp.HAction.Run(action)
            return True
        return False

    def _get_cell_text_by_clipboard(self) -> str:
        """
        현재 선택된 셀의 텍스트를 클립보드를 통해 가져옵니다.
        (내부 헬퍼 함수 - 셀이 이미 선택된 상태에서 호출)
        """
        import win32clipboard

        # SelectAll로 셀 내용 전체 선택 후 복사
        self.hwp.HAction.Run("SelectAll")
        self.hwp.HAction.Run("Copy")
        self.hwp.HAction.Run("Cancel")

        # 클립보드에서 텍스트 읽기
        win32clipboard.OpenClipboard()
        try:
            text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        except Exception as e:
            logger.debug(f"클립보드 읽기 실패: {e}")
            text = ""
        finally:
            win32clipboard.CloseClipboard()

        return text.strip() if text else "(빈 셀)"

    def navigate_and_get_cell(self, direction: str) -> Tuple[bool, str, str]:
        """
        지정된 방향으로 이동하고 현재 셀의 내용을 반환합니다.

        Args:
            direction: 이동 방향 ("right", "left", "down", "up")

        Returns:
            Tuple[bool, str, str]: (성공 여부, 방향, 셀 내용)
        """
        try:
            if not self.is_hwp_running:
                return False, direction, "HWP가 연결되어 있지 않습니다."

            # 현재 셀 위치 확정 후 이동 (기존 _find_labels_recursive와 동일한 로직)
            self.hwp.HAction.Run("TableSelCell")
            self.hwp.HAction.Run("Cancel")
            self._move_direction(direction)

            # 이동 후 셀 선택하고 내용 가져오기
            self.hwp.HAction.Run("TableSelCell")
            text = self._get_cell_text_by_clipboard()

            return True, direction, text
        except Exception as e:
            return False, direction, f"네비게이션 실패: {str(e)}"

    def get_table_view(self, depth: int = 1) -> Tuple[bool, Dict[str, Any]]:
        """
        현재 위치 기준으로 주변 셀들의 내용을 가져옵니다.

        Args:
            depth: 탐색 깊이 (1이면 상하좌우 1칸씩, 2면 2칸씩...)

        Returns:
            Tuple[bool, Dict]: (성공 여부, 셀 내용 딕셔너리)
            딕셔너리 구조:
            {
                "center": "현재 셀 내용",
                "up_1": "위 1칸",
                "up_2": "위 2칸",
                "down_1": "아래 1칸",
                "left_1": "왼쪽 1칸",
                "right_1": "오른쪽 1칸",
                ...
            }
        """
        try:
            if not self.is_hwp_running:
                return False, {"error": "HWP가 연결되어 있지 않습니다."}

            result = {}

            # 현재 셀 내용 가져오기
            self.hwp.HAction.Run("TableSelCell")
            result["center"] = self._get_cell_text_by_clipboard()
            self.hwp.HAction.Run("Cancel")

            # 각 방향으로 탐색
            directions = [
                ("up", "TableUpperCell"),
                ("down", "TableLowerCell"),
                ("left", "TableLeftCell"),
                ("right", "TableRightCell")
            ]

            for dir_name, action in directions:
                # 현재 위치로 돌아오기 위해 반대 방향 저장
                opposite = {
                    "up": "TableLowerCell",
                    "down": "TableUpperCell",
                    "left": "TableRightCell",
                    "right": "TableLeftCell"
                }

                for d in range(1, depth + 1):
                    # 이동
                    self.hwp.HAction.Run(action)
                    self.hwp.HAction.Run("TableSelCell")
                    cell_text = self._get_cell_text_by_clipboard()
                    result[f"{dir_name}_{d}"] = cell_text
                    self.hwp.HAction.Run("Cancel")

                # 원래 위치로 복귀
                for _ in range(depth):
                    self.hwp.HAction.Run(opposite[dir_name])

            return True, result
        except Exception as e:
            return False, {"error": f"테이블 뷰 가져오기 실패: {str(e)}"}

    def find_and_get_cell(self, text: str) -> Tuple[bool, str]:
        """
        텍스트를 찾고 해당 셀의 내용을 반환합니다.

        Args:
            text: 찾을 텍스트

        Returns:
            Tuple[bool, str]: (성공 여부, 셀 내용 또는 에러 메시지)
        """
        try:
            if not self.is_hwp_running:
                return False, "HWP가 연결되어 있지 않습니다."

            # 문서 처음으로 이동
            self.hwp.HAction.Run("MoveDocBegin")

            # 텍스트 찾기
            pset = self.hwp.HParameterSet.HFindReplace
            self.hwp.HAction.GetDefault("RepeatFind", pset.HSet)
            pset.FindString = text
            pset.FindRegExp = 0
            pset.IgnoreMessage = 1
            pset.Direction = 0
            result = self.hwp.HAction.Execute("RepeatFind", pset.HSet)

            if not result:
                return False, f"'{text}'을(를) 찾을 수 없습니다."

            # 찾은 후 셀 선택하고 내용 가져오기
            self.hwp.HAction.Run("TableSelCell")
            cell_text = self._get_cell_text_by_clipboard()

            return True, cell_text
        except Exception as e:
            return False, f"찾기 실패: {str(e)}"

    def _find_labels_recursive(self, path: List[str], depth: int = 0) -> Tuple[bool, int]:
        """
        경로의 레이블들을 순차적으로 찾는 재귀 함수.
        방향 키워드(<left>, <right>, <up>, <down>)도 지원합니다.

        Args:
            path: 찾을 레이블 경로 (예: ["대표자", "<down>", "<right>"])
            depth: 현재 깊이 (인덱스)

        Returns:
            Tuple[bool, int]: (성공 여부, 찾은 depth)
        """
        # 기저 조건: 모든 항목 처리 완료
        if depth >= len(path):
            return True, depth

        item = path[depth]

        # 방향 키워드 처리: <left>, <right>, <up>, <down>
        if item.startswith("<") and item.endswith(">"):
            direction = item[1:-1].lower()  # "<down>" -> "down"
            if direction in ["left", "right", "up", "down"]:
                # 현재 셀 위치 확정 후 이동
                self.hwp.HAction.Run("TableSelCell")
                self.hwp.HAction.Run("Cancel")
                self._move_direction(direction)
                # 재귀: 다음 항목 처리
                return self._find_labels_recursive(path, depth + 1)
            else:
                return False, depth  # 잘못된 방향 키워드

        # 일반 레이블 찾기
        pset = self.hwp.HParameterSet.HFindReplace
        self.hwp.HAction.GetDefault("RepeatFind", pset.HSet)
        pset.FindString = item
        pset.FindRegExp = 0
        pset.IgnoreMessage = 1
        pset.Direction = 0  # forward
        result = self.hwp.HAction.Execute("RepeatFind", pset.HSet)

        if not result:
            return False, depth

        # 재귀: 다음 항목 처리
        return self._find_labels_recursive(path, depth + 1)

    def fill_cell_by_path(
        self,
        path: List[str],
        value: str,
        direction: str = "right",
        mode: str = "replace"
    ) -> Tuple[bool, str]:
        """
        경로를 따라 레이블과 방향 키워드를 순차적으로 처리하여 셀에 값을 입력합니다.

        **방향 키워드 지원:** <left>, <right>, <up>, <down>
        - 레이블 찾기와 방향 이동을 조합하여 복잡한 표 구조도 정확하게 탐색

        **사용 예시:**
        - path=["대표자", "<down>"] → 대표자 찾고 → 아래로 이동 → direction 방향으로 값 입력
        - path=["대표자", "<down>", "<right>"] → 대표자 → 아래 → 오른쪽 → 값 입력
        - path=["투자"] → 투자 찾고 → direction 방향으로 값 입력 (단위 셀에 prepend)

        Args:
            path: 레이블과 방향 키워드의 경로 (예: ["대표자", "<down>"])
            value: 입력할 값
            direction: 마지막 항목 처리 후 이동 방향 ("right", "down", "left", "up")
            mode: 입력 모드 ("replace", "prepend", "append")

        Returns:
            Tuple[bool, str]: (성공 여부, 결과 메시지)
        """
        try:
            if not self.is_hwp_running:
                return False, "HWP가 연결되어 있지 않습니다."

            if not path or len(path) == 0:
                return False, "경로가 비어있습니다."

            # 1. 문서 처음으로 이동
            self.hwp.HAction.Run("MoveDocBegin")

            # 2. 재귀적으로 경로의 모든 레이블 찾기
            found, found_depth = self._find_labels_recursive(path)
            if not found:
                if found_depth == 0:
                    return False, f"첫 번째 레이블 '{path[0]}'을(를) 찾을 수 없습니다."
                else:
                    found_path = " > ".join(path[:found_depth])
                    missing_label = path[found_depth]
                    return False, f"'{found_path}' 이후에 '{missing_label}'을(를) 찾을 수 없습니다."

            # 3. 현재 셀 선택 후 해제 - 커서 위치 확정
            self.hwp.HAction.Run("TableSelCell")
            self.hwp.HAction.Run("Cancel")

            # 4. 마지막 항목이 방향 키워드가 아닌 경우에만 direction으로 추가 이동
            last_item = path[-1] if path else ""
            is_last_direction = last_item.startswith("<") and last_item.endswith(">")

            if not is_last_direction:
                direction_lower = direction.lower()
                if direction_lower == "right":
                    self.hwp.HAction.Run("TableRightCell")
                elif direction_lower == "left":
                    self.hwp.HAction.Run("TableLeftCell")
                elif direction_lower == "down":
                    self.hwp.HAction.Run("TableLowerCell")
                elif direction_lower == "up":
                    self.hwp.HAction.Run("TableUpperCell")

            # 5. mode에 따라 값 입력
            mode_lower = mode.lower()
            if mode_lower == "replace":
                self.hwp.HAction.Run("SelectAll")
                self.hwp.HAction.Run("EditCut")
                self._insert_text_direct(value)
            elif mode_lower == "prepend":
                self.hwp.HAction.Run("MoveSelCellBegin")
                self.hwp.HAction.Run("Cancel")
                self._insert_text_direct(value)
            elif mode_lower == "append":
                # 셀 끝으로 이동: 전체 선택 후 오른쪽으로 이동하면 끝으로 감
                self.hwp.HAction.Run("SelectAll")
                self.hwp.HAction.Run("Cancel")
                self.hwp.HAction.Run("MoveLineEnd")
                self._insert_text_direct(value)
            else:
                return False, f"잘못된 mode입니다: {mode}. 'replace', 'prepend', 'append' 중 하나를 사용하세요."

            path_str = " > ".join(path)
            return True, f"'{path_str}' 경로의 셀에 '{value}' 입력 완료"

        except Exception as e:
            return False, f"셀 채우기 실패: {str(e)}"

    def fill_cells_by_path_batch(
        self,
        path_value_map: Dict[str, str],
        direction: str = "right",
        mode: str = "replace"
    ) -> Dict[str, Tuple[bool, str]]:
        """
        여러 경로에 대해 값을 일괄 입력합니다.

        Args:
            path_value_map: 경로(문자열)와 값의 매핑
                - 경로는 " > " 또는 "/"로 구분 (예: "대표자 > 총 인원" 또는 "대표자/총 인원")
            direction: 이동 방향 ("right", "down", "left", "up")
            mode: 입력 모드 ("replace", "prepend", "append")

        Returns:
            Dict[str, Tuple[bool, str]]: 각 경로에 대한 (성공 여부, 결과 메시지)
        """
        results = {}

        for path_str, value in path_value_map.items():
            # 경로 문자열을 리스트로 변환
            if " > " in path_str:
                path = [p.strip() for p in path_str.split(" > ")]
            elif "/" in path_str:
                path = [p.strip() for p in path_str.split("/")]
            else:
                path = [path_str]

            success, message = self.fill_cell_by_path(path, value, direction, mode)
            results[path_str] = (success, message)

        return results

    def get_current_page(self) -> int:
        """커서가 있는 쪽 번호(1부터, 실제 쪽 순서)를 반환합니다. 알 수 없으면 1."""
        try:
            info = self.hwp.XHwpDocuments.Active_XHwpDocument.XHwpDocumentInfo
            return int(info.CurrentPage) + 1
        except Exception as e:
            logger.debug(f"현재 쪽 번호 확인 실패: {e}")
            return 1

    def create_page_image(self, out_path: str, page: int, resolution: int = 96) -> Tuple[bool, str]:
        """
        한글이 그린 모습 그대로 한 쪽을 이미지 파일로 저장합니다.

        Args:
            out_path: 저장할 경로(.png 권장, .bmp/.gif도 가능)
            page: 쪽 번호(1부터, 문서 안의 실제 쪽 순서. 쪽 번호 매기기와 무관)
            resolution: DPI (24~300)
        """
        try:
            if not self.is_hwp_running:
                return False, "한글에 연결되어 있지 않습니다."
            page_count = int(self.hwp.PageCount)
            if page < 1 or page > page_count:
                return False, f"쪽 번호는 1~{page_count} 사이여야 합니다(요청: {page})."
            resolution = max(24, min(300, int(resolution)))

            out_path = os.path.abspath(out_path)
            os.makedirs(os.path.dirname(out_path), exist_ok=True)
            ext = os.path.splitext(out_path)[1].lower().lstrip(".")

            # 한글의 CreatePageImage는 bmp/gif만 지원(쪽 번호는 0부터)
            if ext in ("bmp", "gif"):
                ok = self.hwp.CreatePageImage(out_path, page - 1, resolution, 24, ext)
                if not ok or not os.path.exists(out_path):
                    return False, f"한글이 쪽 이미지를 만들지 못했습니다(저장 시도 경로: {out_path})."
                return True, out_path

            from PIL import Image as PILImage
            tmp_bmp = os.path.splitext(out_path)[0] + ".__tmp__.bmp"
            try:
                ok = self.hwp.CreatePageImage(tmp_bmp, page - 1, resolution, 24, "bmp")
                if not ok or not os.path.exists(tmp_bmp):
                    return False, f"한글이 쪽 이미지를 만들지 못했습니다(저장 시도 경로: {tmp_bmp})."
                with PILImage.open(tmp_bmp) as img:
                    img.convert("RGB").save(out_path, format="JPEG" if ext in ("jpg", "jpeg") else "PNG")
            finally:
                if os.path.exists(tmp_bmp):
                    os.remove(tmp_bmp)
            return True, out_path
        except Exception as e:
            logger.error(f"쪽 이미지 생성 실패: {e}", exc_info=True)
            return False, str(e)

    def set_track_changes(self, active: bool) -> bool:
        """
        변경 내용 추적(Track Changes) 기능을 설정합니다.
        """
        try:
            if not self.is_hwp_running:
                return False
            self.hwp.IsTrackChange = 1 if active else 0
            return True
        except Exception as e:
            logger.error(f"변경 내용 추적 설정 실패: {e}")
            return False

    def set_paragraph_shape(
        self,
        heading_type: str = None,
        alignment: str = None,
        line_spacing: int = None,
        left_margin: int = None,
        right_margin: int = None,
        first_line: int = None,
        space_before: int = None,
        space_after: int = None,
        apply_to_all: bool = False,
    ) -> bool:
        """현재 커서 또는 선택 문단의 모양을 설정합니다. 지정된 값만 변경됩니다."""
        _HEADING = {"none": 0, "outline": 1, "number": 2, "bullet": 3}
        _ALIGN   = {"left": 0, "right": 1, "center": 2, "justify": 3}
        try:
            if not self.is_hwp_running:
                return False
            if apply_to_all:
                self.hwp.HAction.Run("SelectAll")
            pset = self.hwp.HParameterSet.HParaShape
            self.hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
            if heading_type is not None:
                code = _HEADING.get(heading_type.lower())
                if code is None:
                    logger.error(f"잘못된 heading_type: {heading_type!r}")
                    return False
                pset.HeadingType = code
            if alignment is not None:
                code = _ALIGN.get(alignment.lower())
                if code is None:
                    logger.error(f"잘못된 alignment: {alignment!r}")
                    return False
                pset.Alignment = code
            if line_spacing is not None:
                pset.LineSpacing = line_spacing
            if left_margin is not None:
                pset.LeftMargin = left_margin
            if right_margin is not None:
                pset.RightMargin = right_margin
            if first_line is not None:
                pset.FirstLineIndent = first_line
            if space_before is not None:
                pset.ParaSpaceBefore = space_before
            if space_after is not None:
                pset.ParaSpaceAfter = space_after
            return bool(self.hwp.HAction.Execute("ParagraphShape", pset.HSet))
        except Exception as e:
            logger.error(f"문단 모양 설정 실패: {e}")
            return False

    def _has_leading_bullet(self, text: str) -> bool:
        """텍스트가 말머리표나 번호 패턴으로 시작하는지 감지합니다."""
        if not text:
            return False
        text_clean = text.lstrip()
        if not text_clean:
            return False
        bullet_chars = "□■○●◎◇◆▷▶▲▼★＊※·•-*"
        if text_clean[0] in bullet_chars:
            return True
        import re
        numbering_pattern = re.compile(r'^([0-9a-zA-Z가-힣]+[\.\)]|①|②|③|④|⑤|⑥|⑦|⑧|⑨|⑩|⑪|⑫|⑬|⑭|⑮|Ⅰ|Ⅱ|Ⅲ|Ⅳ|Ⅴ)\s*')
        if numbering_pattern.match(text_clean):
            return True
        return False

    def _strip_leading_bullet(self, text: str) -> str:
        """텍스트의 선두 말머리표나 번호 패턴을 제거합니다."""
        if not text:
            return ""
        text_clean = text.lstrip()
        bullet_chars = "□■○●◎◇◆▷▶▲▼★＊※·•-*"
        if text_clean and text_clean[0] in bullet_chars:
            return text_clean[1:].lstrip()
        import re
        numbering_pattern = re.compile(r'^([0-9a-zA-Z가-힣]+[\.\)]|①|②|③|④|⑤|⑥|⑦|⑧|⑨|⑩|⑪|⑫|⑬|⑭|⑮|Ⅰ|Ⅱ|Ⅲ|Ⅳ|Ⅴ)\s*')
        match = numbering_pattern.match(text_clean)
        if match:
            return text_clean[match.end():].lstrip()
        return text