r"""
문서를 텍스트로 추출해서 사내 MediaWiki에 자동 업로드하는 모듈. Windows/macOS/Linux 공용.

예전 Proposal_to_wikiUpload.py는 win32com으로 Excel/한글/PowerPoint를 직접 열어 PDF로
변환한 뒤 텍스트를 뽑았음 - Windows + 해당 오피스 프로그램이 실제 설치되어 있어야만
동작하고 자동화가 쉽게 깨지는 방식이었다. 이 모듈은 대신 register.py가 이미 갖고 있는
안정적인 텍스트 추출기(PyMuPDF/openpyxl/python-pptx/pyhwp/lxml 등, 전부 순수 Python)를
그대로 재사용해서 OS나 오피스 설치 여부와 무관하게 동작한다.

설정:
    config.json에 아래 키가 있어야 함 (없으면 register.py의 CONFIG_PATH 로직과 동일하게
    자동으로 플레이스홀더 값을 채워 넣으니, 채워진 뒤 실제 값으로 수정):
    wiki_site_url, wiki_path, wiki_username, wiki_password, wiki_category
    예전 key.txt는 더 이상 쓰지 않음 (config.json으로 통합).
"""
import json
from pathlib import Path

import mwclient
import urllib3

import register

# 사내 위키가 사설 인증서를 쓰는 경우가 많아 검증을 끔 (Proposal_to_wikiUpload.py와 동일)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

WIKI_DEFAULT_CONFIG = {
    "wiki_site_url": "example.com",
    "wiki_path": "/mediawiki/",
    "wiki_username": "REPLACE_ME",
    "wiki_password": "REPLACE_ME",
    "wiki_category": "제안서",
}


def get_wiki_config() -> dict:
    """register.py가 읽어둔 config.json에서 wiki_* 값을 꺼낸다. 키가 없으면 플레이스홀더로
    채워서 config.json에 저장해둔다 (register.py의 최초 config.json 자동 생성과 같은 방식)."""
    changed = False
    for key, default in WIKI_DEFAULT_CONFIG.items():
        if key not in register.CONFIG:
            register.CONFIG[key] = default
            changed = True
    if changed:
        register.CONFIG_PATH.write_text(
            json.dumps(register.CONFIG, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"config.json에 wiki_* 설정이 없어 기본값으로 채웠습니다: {register.CONFIG_PATH}")
    return {key: register.CONFIG[key] for key in WIKI_DEFAULT_CONFIG}


def get_wiki_site(cfg: dict | None = None) -> mwclient.Site:
    """위키에 로그인한 mwclient.Site 객체를 반환."""
    cfg = cfg or get_wiki_config()
    site = mwclient.Site(cfg["wiki_site_url"], path=cfg["wiki_path"], connection_options={"verify": False})
    site.login(cfg["wiki_username"], cfg["wiki_password"])
    return site


def list_wiki_categories(site: mwclient.Site, limit: int = 200) -> list[str]:
    """위키에 이미 존재하는 분류 이름 목록을 가져온다 ("분류선택" 버튼용)."""
    names = []
    for cat in site.allcategories():
        # mwclient 버전에 따라 문자열 또는 카테고리 객체를 줄 수 있어 둘 다 처리
        name = cat if isinstance(cat, str) else getattr(cat, "name", str(cat))
        name = name.split(":", 1)[-1] if ":" in name else name
        names.append(name)
        if len(names) >= limit:
            break
    return names


EXTRACT_EXTS = {
    ".pdf", ".pptx", ".ppt", ".xlsx", ".hwpx",
    ".docx", ".doc", ".hwp", ".html", ".htm", ".txt", ".md",
}


def extract_text(path: Path) -> str:
    """register.py의 리더를 재사용해 파일에서 텍스트만 뽑는다 (위키 업로드는 이미지 불필요)."""
    ext = path.suffix.lower()
    if ext in register.TEXT_AND_IMAGE_PROCESSORS:
        old_flag = register.PROCESS_IMAGES
        register.PROCESS_IMAGES = False  # 텍스트만 필요하므로 이미지 OCR/캡션은 생략해 속도 확보
        try:
            text, _images_meta = register.TEXT_AND_IMAGE_PROCESSORS[ext](path)
        finally:
            register.PROCESS_IMAGES = old_flag
        return text
    return register.read_text_only(path)


def upload_text_to_wiki(site: mwclient.Site, title: str, text: str, category: str) -> None:
    if not text.strip():
        raise ValueError("추출된 텍스트가 없습니다.")
    body = f"{text}\n\n[[분류:{category}]]" if category else text
    page = site.pages[title]
    page.save(body, summary="Qdrant 문서 등록 프로그램에서 자동 업로드")


def upload_paths_to_wiki(paths: list[Path], category: str, progress_callback=None) -> tuple[int, int]:
    """파일/폴더 목록을 위키에 업로드. (성공 개수, 전체 개수) 반환.
    progress_callback은 gui.py의 update_progress_display가 기대하는 것과 같은
    {"file_index","file_total","file_name","unit_index","unit_total","unit_label"} dict를 받는다."""
    files = []
    for p in paths:
        if not p.exists():
            print(f"경로가 존재하지 않습니다: {p}")
            continue
        if p.is_dir():
            files.extend(f for f in p.rglob("*") if f.is_file() and f.suffix.lower() in EXTRACT_EXTS)
        else:
            files.append(p)

    if not files:
        print("업로드할 파일이 없습니다.")
        return 0, 0

    cfg = get_wiki_config()
    print(f"위키 로그인 중: {cfg['wiki_username']} @ {cfg['wiki_site_url']}")
    site = get_wiki_site(cfg)
    print(f"로그인 성공. 분류: {category or '(없음)'}")

    total = len(files)
    ok = 0
    for i, f in enumerate(files, start=1):
        title = f.stem
        try:
            text = extract_text(f)
            if not text.strip():
                print(f"건너뜀(내용 없음): {f.name}")
            else:
                upload_text_to_wiki(site, title, text, category)
                print(f"업로드 완료: {title}")
                ok += 1
        except Exception as e:
            print(f"[오류] {f.name}: {e}")

        if progress_callback:
            try:
                progress_callback({
                    "file_index": i, "file_total": total, "file_name": f.name,
                    "unit_index": 1, "unit_total": 1, "unit_label": "업로드",
                })
            except Exception:
                pass

    print(f"모든 작업 완료. {ok}/{total}개 업로드 성공.")
    return ok, total
