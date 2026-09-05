r"""
Claude 데스크톱 앱의 claude_desktop_config.json에 사내 MCP 서버(PMS_redmine/groupware/mediawiki)를
연동하기 위한 설정 도우미. 비개발자가 raw JSON을 직접 열어 command/args/env 구조를 편집하지
않아도, GUI에서 서버별 URL/계정/키 값만 입력하면 이 모듈이 올바른 구조로 claude_desktop_config.json에
병합해 저장한다.

지원 범위: 사내 자체 서버 3종(PMS_redmine, groupware, mediawiki)만 다룬다. github 등 다른 서버나
coworkUserFilesPath/preferences 같은 다른 최상위 설정은 절대 건드리지 않고 그대로 보존한다.

실제 MCP 서버 실행 파일(mediawiki-mcp-server-windows.exe, UTGroupwareMCP.exe) 자체는 이 모듈이
만들거나 포함하지 않는다 - 이미 별도로 빌드/배포된 실행 파일이 PC 어딘가에 있다고 가정하고,
흔한 기본 경로에 있으면 자동 인식하거나 사용자가 직접 경로를 지정하게 한다
(register.py의 find_tesseract/find_soffice와 동일한 패턴).
"""
import hashlib
import json
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

import requests

import register

# mediawiki-mcp-server 실행 파일은 이 프로그램(Python) 코드가 아니라 별도 빌드된 Node.js
# 기반 바이너리라 자체 버전 정보가 없다 - 대신 qdrant_rag git 저장소에 실행 파일 자체와
# 해시값을 커밋해두고, 그 저장소를 "최신 버전"의 기준으로 삼는다(명시적 사용자 확인:
# "예, qdrant_rag 저장소에 그대로 커밋"). manifest.json만 먼저 가볍게 받아서 로컬 파일의
# sha256과 비교하고, 다를 때만 실제 실행 파일(수십MB)을 내려받는다.
MANIFEST_URL = "https://raw.githubusercontent.com/yuseungil-a11y/qdrant_rag/main/mcp_servers/manifest.json"
RAW_BASE_URL = "https://raw.githubusercontent.com/yuseungil-a11y/qdrant_rag/main/"


def find_claude_config_path() -> Path:
    """Claude 데스크톱 앱의 claude_desktop_config.json 경로 (OS별 표준 위치).
    Windows Store 패키지형 설치는 AppData\\Local\\Packages\\Claude_*\\LocalCache\\Roaming\\Claude로도
    보이지만 실제로는 표준 %APPDATA%\\Claude 경로와 동일한 파일(리다이렉션)이므로 표준 경로 하나만 쓴다."""
    if sys.platform == "win32":
        appdata = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
        return Path(appdata) / "Claude" / "claude_desktop_config.json"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "Claude" / "claude_desktop_config.json"
    return Path.home() / ".config" / "Claude" / "claude_desktop_config.json"


def find_npx() -> str | None:
    """PMS_redmine처럼 npx로 실행하는 서버에 Node.js/npx가 설치돼 있는지 확인."""
    return shutil.which("npx")


def find_server_exe(candidates: list[str]) -> str | None:
    """exe로 직접 실행하는 서버(mediawiki/groupware)의 실행 파일을 흔한 기본 경로에서 탐색."""
    for c in candidates:
        if Path(c).exists():
            return c
    return None


# (env 키, 표시 라벨, 비밀값 여부) - 사용자가 GUI에서 입력하는 값만 여기 나열한다.
# command/args/fixed_env는 서버마다 고정이라 사용자가 몰라도 되게 여기서 관리한다.
SERVER_TEMPLATES: dict[str, dict] = {
    "PMS_redmine": {
        "label": "PMS (Redmine)",
        "kind": "npx",
        "command": "npx",
        # 주의(2026-09-05, 사용자와 논의 후 보류): @latest라 서드파티 개발자가 이 npm 패키지를
        # 내리거나 문제 있는 업데이트를 올리면 그대로 영향받는다. 필요해지면 버전 고정 또는
        # mediawiki-mcp-server처럼 mcp_servers/에 미러링(vendoring)하는 방식으로 전환 고려.
        "args": ["-y", "@chspower1/mcp-for-redmine@latest"],
        "fixed_env": {},
        # (env 키, 표시 라벨, 비밀값 여부, 기본값) - 기본값은 사내에서 공통으로 쓰는 고정 주소라
        # 사용자가 매번 직접 입력하지 않아도 되게 미리 채워둔다(API 키처럼 사람마다 다른 값만
        # 기본값을 비워둠).
        "fields": [
            ("REDMINE_BASE_URL", "Redmine 서버 주소", False, "http://pms.utinfo.co.kr:1177/redmine"),
            ("REDMINE_API_KEY", "API 키", True, ""),
        ],
    },
    "groupware": {
        "label": "그룹웨어",
        "kind": "exe",
        "exe_candidates": [r"C:\UTGroupwareMCP.exe"],
        "fixed_env": {},
        "fields": [
            ("GROUPWARE_API_KEY", "API 키", True, ""),
        ],
    },
    "mediawiki": {
        "label": "MediaWiki",
        "kind": "exe",
        # manifest.json에 이 키로 등록돼 있으면 "업데이트 확인" 버튼이 뜬다. groupware는
        # "최신 버전"의 출처가 없어(내부에서만 빌드) manifest에 없고, 자동 업데이트 대상도 아니다.
        "manifest_key": "mediawiki",
        # OS별로 실행 파일 형태가 다르다(Windows: .exe / macOS: 확장자 없는 유닉스 실행 파일,
        # dist/mac_mediawiki-mcp-server/ 하위에 배포됨) - 같은 dist 폴더 안에 양쪽이 섞여
        # 있어도 실행 중인 OS에 맞는 후보만 찾도록 분리한다. 이 프로그램과 같은 폴더에
        # 함께 배포되는 경우를 먼저 찾고, 없으면 예전부터 흔히 쓰던 고정 경로도 확인한다(하위 호환).
        "exe_candidates": (
            [
                str(register.SCRIPT_DIR / "mac_mediawiki-mcp-server" / "mediawiki-mcp-server"),
                str(register.SCRIPT_DIR / "mediawiki-mcp-server"),
                "/usr/local/bin/mediawiki-mcp-server",
            ]
            if sys.platform == "darwin"
            else [
                str(register.SCRIPT_DIR / "mediawiki-mcp-server-windows.exe"),
                r"C:\mediawiki-mcp-server-windows.exe",
            ]
        ),
        "fixed_env": {
            "NODE_OPTIONS": "--input-type=module",
            "LC_ALL": "ko_KR.UTF-8",
            "LANG": "ko_KR.UTF-8",
            "PYTHONIOENCODING": "utf-8",
        },
        "fields": [
            ("MEDIAWIKI_URL", "위키 API 주소", False, "https://pms.utinfo.co.kr/mediawiki/api.php"),
            ("MEDIAWIKI_USERNAME", "계정 (예: User@BotName)", False, ""),
            ("MEDIAWIKI_PASSWORD", "비밀번호", True, ""),
        ],
    },
}


def load_claude_config() -> dict:
    path = find_claude_config_path()
    if not path.exists():
        return {"mcpServers": {}}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        raise RuntimeError(f"claude_desktop_config.json 읽기 실패: {e}")


def resolve_exe_path(server_key: str) -> str | None:
    """이미 claude_desktop_config.json에 저장된 실행 파일 경로가 있으면 그걸 쓰고, 없으면
    흔한 기본 경로에서 자동탐지한다. 둘 다 없으면 None(= 이 서버를 아직 설정해본 적 없음)."""
    template = SERVER_TEMPLATES[server_key]
    if template["kind"] != "exe":
        return None
    existing = get_existing_server_config(server_key)
    existing_exe = existing.get("command") or ""
    return existing_exe or find_server_exe(template.get("exe_candidates", []))


def get_existing_server_config(server_key: str) -> dict:
    """이미 설정돼 있으면 그 값을 돌려준다 (GUI 입력칸을 기존 값으로 미리 채울 때 사용).
    없으면 빈 dict."""
    try:
        data = load_claude_config()
    except RuntimeError:
        return {}
    return data.get("mcpServers", {}).get(server_key, {})


def save_server_config(server_key: str, field_values: dict[str, str], exe_path: str | None = None) -> Path:
    """field_values: 사용자가 입력한 (ENV_KEY -> 값) 중 이 서버의 fields에 해당하는 것만.
    exe_path: kind="exe" 서버일 때 실행 파일 경로 (자동탐지됐거나 사용자가 지정한 값) - 필수.
    기존 claude_desktop_config.json에서 이 서버 블록 하나만 만들거나 덮어쓰고, 다른 서버/다른
    최상위 키는 절대 건드리지 않는다. 저장된 경로를 반환한다."""
    template = SERVER_TEMPLATES[server_key]
    path = find_claude_config_path()
    data = load_claude_config()
    data.setdefault("mcpServers", {})

    env = dict(template.get("fixed_env", {}))
    env.update(field_values)

    if template["kind"] == "npx":
        entry = {"command": template["command"], "args": list(template["args"]), "env": env}
    else:
        if not exe_path:
            raise ValueError("실행 파일 경로가 필요합니다.")
        entry = {"command": exe_path, "env": env}

    data["mcpServers"][server_key] = entry

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _file_sha256(path: str | Path) -> str | None:
    p = Path(path)
    if not p.exists():
        return None
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def check_for_server_update(server_key: str, current_exe_path: str) -> dict:
    """server_key의 실행 파일을 qdrant_rag 저장소의 manifest.json과 비교한다.
    반환: {"ok": bool, "update_available": bool, "message": str, "manifest_entry": dict|None}
    - ok=False면 확인 자체가 실패한 것(네트워크 오류 등)이라 update_available은 의미 없음.
    - manifest.json에 이 서버가 없으면(예: groupware) ok=False, message로 그 사실을 알린다."""
    template = SERVER_TEMPLATES[server_key]
    manifest_key = template.get("manifest_key")
    if not manifest_key:
        return {"ok": False, "update_available": False, "message": "이 서버는 자동 업데이트를 지원하지 않습니다.", "manifest_entry": None}

    os_key = "mac" if sys.platform == "darwin" else "windows"
    try:
        # raw.githubusercontent.com은 파일이 바뀐 뒤에도 몇 분간 CDN에 캐시된 이전 내용을
        # 돌려줄 수 있다(실측 확인) - 매번 다른 쿼리 문자열을 붙여 캐시를 우회한다.
        resp = requests.get(MANIFEST_URL, params={"_": str(time.time())}, timeout=10)
        resp.raise_for_status()
        manifest = resp.json()
    except Exception as e:
        return {"ok": False, "update_available": False, "message": f"업데이트 정보 확인 실패: {e}", "manifest_entry": None}

    entry = manifest.get(manifest_key, {}).get(os_key)
    if not entry:
        return {"ok": False, "update_available": False, "message": "manifest.json에 이 OS용 항목이 없습니다.", "manifest_entry": None}

    local_hash = _file_sha256(current_exe_path) if current_exe_path else None
    remote_hash = entry.get("sha256")
    if local_hash is None:
        return {"ok": True, "update_available": True, "message": "로컬에 실행 파일이 없습니다 - 새로 받을 수 있습니다.", "manifest_entry": entry}
    if local_hash == remote_hash:
        return {"ok": True, "update_available": False, "message": "최신 버전입니다.", "manifest_entry": entry}
    return {"ok": True, "update_available": True, "message": "새 버전이 있습니다.", "manifest_entry": entry}


def download_server_update(manifest_entry: dict, dest_path: str) -> None:
    """manifest_entry(check_for_server_update가 돌려준 것)가 가리키는 파일을 내려받아
    dest_path에 원자적으로 교체한다. dest_path가 Claude 데스크톱에 의해 실행 중이라
    잠겨 있으면 교체가 실패할 수 있는데, 이 경우 예외 메시지로 "Claude 데스크톱 앱을
    먼저 종료하라"는 안내가 나가도록 한다."""
    url = RAW_BASE_URL + manifest_entry["path"]
    resp = requests.get(url, timeout=60)
    resp.raise_for_status()
    data = resp.content

    actual_hash = hashlib.sha256(data).hexdigest()
    expected_hash = manifest_entry.get("sha256")
    if expected_hash and actual_hash != expected_hash:
        raise RuntimeError(f"다운로드한 파일의 해시가 일치하지 않습니다 (받은 파일이 손상되었을 수 있음)")

    dest = Path(dest_path)
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp_fd, tmp_name = tempfile.mkstemp(dir=str(dest.parent), prefix=".update_", suffix=dest.suffix)
    try:
        with os.fdopen(tmp_fd, "wb") as f:
            f.write(data)
        if sys.platform != "win32":
            os.chmod(tmp_name, 0o755)
        try:
            os.replace(tmp_name, dest)
        except PermissionError as e:
            raise PermissionError(
                "실행 파일을 교체할 수 없습니다 - Claude 데스크톱 앱이 이 파일을 사용 중일 수 있습니다. "
                "Claude 데스크톱 앱을 완전히 종료한 뒤 다시 시도하세요."
            ) from e
    finally:
        Path(tmp_name).unlink(missing_ok=True)
