r"""
qdrant_register_gui.exe 자기 자신의 자동 업데이트 확인/설치.

mcp_config_helper.py의 MediaWiki MCP 서버 업데이트와 같은 방식(이 git 저장소 자체를
"최신 버전" 기준으로 삼음)을 이 프로그램 자신에도 적용한다.

2026-09-06부터 Windows 빌드가 onefile(단일 exe)에서 onedir(설치 폴더)로 바뀌었다
(qdrant_register_gui.spec 참고 - onefile의 실행 시점 압축해제가 일부 환경에서 불안정해서
매번 다른 파일이 빠진 것처럼 보이는 에러가 여러 PC/VM에서 재현됨). 그래서 자기업데이트도
"exe 파일 하나 교체"가 아니라 "설치 폴더 전체를 새 폴더로 통째 교체"하는 방식으로 바뀌었다.
실행 중인 프로그램이 자기 자신을 담은 폴더를 스스로 지울 수는 없다는 제약은 여전하므로:
    1) 새 배포판(zip)을 받아 해시 검증 후, 설치 폴더의 형제 폴더(<설치폴더명>_new)에 압축 해제
    2) 현재 설치 폴더와 새 폴더를 맞바꾸고 재실행하는 배치 스크립트를 임시 폴더에 생성
    3) 그 배치 스크립트를 별도 프로세스로 띄우고, 이 프로그램은 스스로 종료
    4) 배치 스크립트가 (실행 중인 exe가 폴더를 붙잡고 있어 폴더 이름 변경이 실패하는 동안
       재시도하며 "프로세스 종료 대기"를 구현) 폴더를 통째로 교체하고 새 버전을 다시 실행

주의: 이전(2.18.3 이전) onefile 버전을 쓰던 사용자는 이 방식 변경 이후 버전으로 자기업데이트를
통해 넘어올 수 없다(자기업데이트로 받은 zip 바이트를 그대로 exe로 착각해 실행하려 하면 깨짐).
그 사용자들은 최초 1회 GitHub에서 새 zip을 수동으로 받아 재설치해야 한다.

현재는 Windows onedir exe만 지원한다 - macOS는 이미 onedir(.app 번들)이라 방식이 또 다르고,
아직 이 자기-업데이트 기능의 대상이 아니다(소스에서 직접 실행 중일 때도 지원 안 함 - 그럴
땐 git pull로 업데이트하면 되므로).
"""
import hashlib
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import requests

APP_MANIFEST_URL = "https://raw.githubusercontent.com/yuseungil-a11y/qdrant_rag/main/releases/manifest.json"
RAW_BASE_URL = "https://raw.githubusercontent.com/yuseungil-a11y/qdrant_rag/main/"


def _version_tuple(v: str) -> tuple[int, ...]:
    try:
        return tuple(int(p) for p in v.strip().split("."))
    except ValueError:
        return (0,)


def _app_dir() -> Path:
    """현재 실행 중인 exe가 들어있는 설치 폴더(onedir 배포의 최상위 폴더)."""
    return Path(sys.executable).resolve().parent


def cleanup_stale_update_files() -> None:
    """이전 업데이트에서 남은 <설치폴더>_old/<exe명>_update_failed.log가 있으면 프로그램
    시작 시 조용히 지운다. 이 시점엔 이미 새 폴더로 실행 중이라 old 폴더를 잠글 프로세스가
    없으므로 안전하게 지울 수 있다(배치 스크립트의 삭제 재시도가 실패했거나 남은 부산물 등 대비)."""
    if not (sys.platform == "win32" and getattr(sys, "frozen", False)):
        return
    try:
        app_dir = _app_dir()
        old_dir = app_dir.with_name(app_dir.name + "_old")
        if old_dir.exists():
            shutil.rmtree(old_dir, ignore_errors=True)
        fail_log = app_dir.with_name(app_dir.name + "_update_failed.log")
        if fail_log.exists():
            fail_log.unlink(missing_ok=True)
    except Exception:
        pass  # 정리 실패해도 앱 실행 자체를 막을 이유는 없음


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
        # raw.githubusercontent.com은 파일이 바뀐 뒤에도 몇 분간 CDN에 캐시된 이전 내용을
        # 돌려줄 수 있다(실측 확인) - 매번 다른 쿼리 문자열을 붙여 캐시를 우회한다.
        resp = requests.get(APP_MANIFEST_URL, params={"_": str(time.time())}, timeout=10)
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
    """새 배포판(zip)을 받아 해시 검증한 뒤, 설치 폴더 전체를 새 폴더로 맞바꿔서 재실행하는
    배치 스크립트를 띄우고 이 프로세스를 종료 준비 상태로 만든다(실제 종료는 호출부가
    root.destroy() 등으로 수행). 배치 스크립트가 폴더 교체+재실행을 담당하므로, 이 함수
    자체는 실행 중인 설치 폴더를 직접 건드리지 않는다(자기 자신이 들어있는 폴더를 통째로
    지우거나 옮길 수 없기 때문)."""
    if not (sys.platform == "win32" and getattr(sys, "frozen", False)):
        raise RuntimeError("자기 업데이트는 Windows exe 실행 시에만 지원됩니다.")

    url = RAW_BASE_URL + manifest_entry["path"]
    resp = requests.get(url, timeout=180)
    resp.raise_for_status()
    data = resp.content

    expected_hash = manifest_entry.get("sha256")
    actual_hash = hashlib.sha256(data).hexdigest()
    if expected_hash and actual_hash != expected_hash:
        raise RuntimeError("다운로드한 파일의 해시가 일치하지 않습니다 (받은 파일이 손상되었을 수 있음)")

    app_dir = _app_dir()
    exe_name = Path(sys.executable).name
    new_dir = app_dir.with_name(app_dir.name + "_new")
    old_dir = app_dir.with_name(app_dir.name + "_old")
    if new_dir.exists():
        shutil.rmtree(new_dir, ignore_errors=True)

    zip_path = Path(tempfile.gettempdir()) / f"{app_dir.name}_update.zip"
    zip_path.write_bytes(data)
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(new_dir)
    zip_path.unlink(missing_ok=True)

    if not (new_dir / exe_name).exists():
        shutil.rmtree(new_dir, ignore_errors=True)
        raise RuntimeError(f"받은 배포판 안에 {exe_name}이 없습니다 (배포판 구조가 예상과 다름)")

    # 배치 스크립트: 실행 중인 exe가 든 폴더는 이 프로세스가 완전히 종료된 뒤에야 이름
    # 변경이 가능하다 - move가 성공할 때까지(최대 30초) 1초 간격으로 재시도해서 "프로세스
    # 종료 대기"를 구현한다(별도 PID 추적 없이 폴더 잠금 해제 자체를 신호로 씀). 폴더
    # 안에서 하나라도 파일이 열려 있으면 폴더 자체의 이름 변경도 막히는 경우가 있어,
    # 기존 단일 exe 교체 때와 동일한 재시도 전략을 그대로 폴더 단위로 적용한다.
    bat_path = Path(tempfile.gettempdir()) / f"{app_dir.name}_update.bat"
    fail_log = app_dir.with_name(app_dir.name + "_update_failed.log")
    # 주의(2026-09-06, 실제 검증 중 발견): `timeout /t N`은 표준입력이 콘솔이어야 동작하는데,
    # 이 배치는 콘솔 창 없이(CREATE_NO_WINDOW) 백그라운드 프로세스로 뜨는 부모(윈도우 서브
    # 시스템 GUI exe, console=False)에서 실행되어 콘솔 입력이 없다 - 그 상태에서 `timeout`은
    # "입력 리디렉션은 지원 안 함"으로 즉시 실패하고 대기 없이 곧장 다음 줄로 넘어간다. 그
    # 결과 "30초 재시도"가 실제로는 거의 대기 없이 30번을 순식간에 소진하고 포기해버려서,
    # 구버전 프로세스가 채 종료되기도 전에 이미 포기하는 사고가 났을 가능성이 높다(로컬에서
    # 직접 재현 확인: 프로세스 종료 직후에도 채 1초가 안 돼 giveup 로그가 생김). 콘솔 없이도
    # 안정적으로 대기하는 `ping -n <초+1> 127.0.0.1`로 교체 - 이 트릭은 콘솔 유무와 무관하게
    # 동작해 배치 스크립트에서 흔히 쓰인다.
    bat_content = (
        "@echo off\r\n"
        "chcp 65001 > nul\r\n"
        "setlocal enabledelayedexpansion\r\n"
        f'if exist "{old_dir}" rmdir /s /q "{old_dir}" >nul 2>&1\r\n'
        f'if exist "{fail_log}" del /f /q "{fail_log}" >nul 2>&1\r\n'
        "set RETRIES=0\r\n"
        ":retry\r\n"
        f'move /y "{app_dir}" "{old_dir}" >nul 2>&1\r\n'
        "if errorlevel 1 (\r\n"
        "    set /a RETRIES+=1\r\n"
        "    if !RETRIES! GEQ 30 goto :giveup\r\n"
        "    ping -n 2 127.0.0.1 > nul\r\n"
        "    goto :retry\r\n"
        ")\r\n"
        f'move /y "{new_dir}" "{app_dir}"\r\n'
        # 방금 막 새로 생긴(인터넷에서 받은) 폴더/파일들을 백신(Windows Defender 등)이
        # 실시간으로 스캔 중일 수 있어, 교체 직후 곧바로 실행하면 그 스캔과 겹쳐 오류가
        # 날 수 있다(2026-09-06 실사용 보고) - 스캔이 끝날 시간을 벌기 위해 실행 전 대기.
        "ping -n 3 127.0.0.1 > nul\r\n"
        f'start "" "{app_dir}\\{exe_name}"\r\n'
        # {old_dir} 삭제도 방금 막 이름 바뀐 폴더라 백신 검사 등으로 아주 잠깐 잠길 수 있어
        # 재시도 없이 한 번만 시도하던 것을 최대 10초 재시도로 바꿈 - 그래도 실패하면
        # 기능상 문제는 없고(다음 실행 폴더는 이미 정상) 남은 폴더는 그냥 무시.
        "set DELRETRIES=0\r\n"
        ":delretry\r\n"
        f'rmdir /s /q "{old_dir}" >nul 2>&1\r\n'
        f'if exist "{old_dir}" (\r\n'
        "    set /a DELRETRIES+=1\r\n"
        "    if !DELRETRIES! LSS 10 (\r\n"
        "        ping -n 2 127.0.0.1 > nul\r\n"
        "        goto :delretry\r\n"
        "    )\r\n"
        ")\r\n"
        "goto :cleanup\r\n"
        ":giveup\r\n"
        f'echo update failed - {app_dir.name} was still in use after 30s > "{fail_log}"\r\n'
        f'if exist "{old_dir}" move /y "{old_dir}" "{app_dir}" >nul 2>&1\r\n'
        ":cleanup\r\n"
        'del "%~f0"\r\n'
    )
    bat_path.write_text(bat_content, encoding="utf-8")

    subprocess.Popen(
        ["cmd", "/c", str(bat_path)],
        creationflags=subprocess.CREATE_NO_WINDOW,
        close_fds=True,
    )
