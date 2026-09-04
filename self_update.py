r"""
qdrant_register_gui.exe 자기 자신의 자동 업데이트 확인/설치.

mcp_config_helper.py의 MediaWiki MCP 서버 업데이트와 같은 방식(이 git 저장소 자체를
"최신 버전" 기준으로 삼음)을 이 프로그램 자신에도 적용한다. 다른 점은, 실행 중인 프로그램이
자기 자신의 exe 파일을 직접 덮어쓸 수 없다는 것 - Windows는 실행 중인 exe의 "이름 변경"은
허용하지만 "내용 덮어쓰기"는 막는다. 그래서:
    1) 새 exe를 별도 파일(qdrant_register_gui_new.exe)로 받아 해시 검증
    2) 현재 실행 중인 exe와 새 exe를 맞바꾸고 재실행하는 작은 배치 스크립트를 임시 폴더에 생성
    3) 그 배치 스크립트를 별도 프로세스로 띄우고, 이 프로그램은 스스로 종료
    4) 배치 스크립트가 (프로세스가 완전히 끝나 파일 잠금이 풀리길 잠깐 기다린 뒤) 파일을
       교체하고 새 버전을 다시 실행

현재는 Windows onefile exe만 지원한다 - macOS는 .app 번들 전체를 바꿔야 해서 방식이 다르고,
아직 이 자기-업데이트 기능의 대상이 아니다(소스에서 직접 실행 중일 때도 지원 안 함 - 그럴
땐 git pull로 업데이트하면 되므로).
"""
import hashlib
import subprocess
import sys
import tempfile
from pathlib import Path

import requests

APP_MANIFEST_URL = "https://raw.githubusercontent.com/yuseungil-a11y/qdrant_rag/main/releases/manifest.json"
RAW_BASE_URL = "https://raw.githubusercontent.com/yuseungil-a11y/qdrant_rag/main/"


def _version_tuple(v: str) -> tuple[int, ...]:
    try:
        return tuple(int(p) for p in v.strip().split("."))
    except ValueError:
        return (0,)


def check_for_app_update(current_version: str) -> dict:
    """반환: {"ok": bool, "update_available": bool, "message": str, "manifest_entry": dict|None,
    "latest_version": str|None}
    - Windows에서 이 프로그램이 exe로 실행 중일 때만 의미가 있다(소스 실행/다른 OS는 대상 아님)."""
    if sys.platform != "win32" or not getattr(sys, "frozen", False):
        return {
            "ok": False, "update_available": False,
            "message": "자기 업데이트는 Windows exe 실행 시에만 지원됩니다.",
            "manifest_entry": None, "latest_version": None,
        }
    try:
        resp = requests.get(APP_MANIFEST_URL, timeout=10)
        resp.raise_for_status()
        manifest = resp.json()
    except Exception as e:
        return {
            "ok": False, "update_available": False, "message": f"업데이트 정보 확인 실패: {e}",
            "manifest_entry": None, "latest_version": None,
        }

    entry = manifest.get("windows")
    latest_version = manifest.get("version")
    if not entry or not latest_version:
        return {
            "ok": False, "update_available": False, "message": "manifest.json에 Windows용 항목이 없습니다.",
            "manifest_entry": None, "latest_version": None,
        }

    update_available = _version_tuple(latest_version) > _version_tuple(current_version)
    message = f"새 버전 있음: v{latest_version}" if update_available else "최신 버전입니다."
    return {
        "ok": True, "update_available": update_available, "message": message,
        "manifest_entry": entry, "latest_version": latest_version,
    }


def download_and_apply_app_update(manifest_entry: dict) -> None:
    """새 exe를 받아 해시 검증한 뒤, 현재 실행 파일과 맞바꿔서 재실행하는 배치 스크립트를
    띄우고 이 프로세스를 종료 준비 상태로 만든다(실제 종료는 호출부가 root.destroy() 등으로
    수행). 배치 스크립트가 파일 교체+재실행을 담당하므로, 이 함수 자체는 파일을 직접
    덮어쓰지 않는다(실행 중인 자기 자신을 덮어쓸 수 없기 때문)."""
    if not (sys.platform == "win32" and getattr(sys, "frozen", False)):
        raise RuntimeError("자기 업데이트는 Windows exe 실행 시에만 지원됩니다.")

    url = RAW_BASE_URL + manifest_entry["path"]
    resp = requests.get(url, timeout=120)
    resp.raise_for_status()
    data = resp.content

    expected_hash = manifest_entry.get("sha256")
    actual_hash = hashlib.sha256(data).hexdigest()
    if expected_hash and actual_hash != expected_hash:
        raise RuntimeError("다운로드한 파일의 해시가 일치하지 않습니다 (받은 파일이 손상되었을 수 있음)")

    current_exe = Path(sys.executable).resolve()
    new_exe = current_exe.with_name(current_exe.stem + "_new.exe")
    new_exe.write_bytes(data)

    # 배치 스크립트: 실행 중인 exe는 "내용 덮어쓰기"는 막혀도 "이름 변경"은 이 프로세스가
    # 완전히 종료된 뒤에야 가능하다 - move가 성공할 때까지(최대 15초) 1초 간격으로 재시도해서
    # "프로세스 종료 대기"를 구현한다(별도 PID 추적 없이 파일 잠금 해제 자체를 신호로 씀).
    bat_path = Path(tempfile.gettempdir()) / f"{current_exe.stem}_update.bat"
    old_exe = current_exe.with_name(current_exe.stem + "_old.exe")
    bat_content = (
        "@echo off\r\n"
        "chcp 65001 > nul\r\n"
        "setlocal enabledelayedexpansion\r\n"
        f'if exist "{old_exe}" del /f /q "{old_exe}" >nul 2>&1\r\n'
        "set RETRIES=0\r\n"
        ":retry\r\n"
        f'move /y "{current_exe}" "{old_exe}" >nul 2>&1\r\n'
        "if errorlevel 1 (\r\n"
        "    set /a RETRIES+=1\r\n"
        "    if !RETRIES! GEQ 15 goto :giveup\r\n"
        "    timeout /t 1 /nobreak > nul\r\n"
        "    goto :retry\r\n"
        ")\r\n"
        f'move /y "{new_exe}" "{current_exe}"\r\n'
        f'start "" "{current_exe}"\r\n'
        f'del /f /q "{old_exe}" >nul 2>&1\r\n'
        "goto :cleanup\r\n"
        ":giveup\r\n"
        f'echo 업데이트 실패: {current_exe.name}이 계속 사용 중입니다.\r\n'
        f'if exist "{old_exe}" move /y "{old_exe}" "{current_exe}" >nul 2>&1\r\n'
        "pause\r\n"
        ":cleanup\r\n"
        'del "%~f0"\r\n'
    )
    bat_path.write_text(bat_content, encoding="utf-8")

    subprocess.Popen(
        ["cmd", "/c", str(bat_path)],
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS,
        close_fds=True,
    )
