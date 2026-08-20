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
import re
from datetime import datetime
from pathlib import Path

import mwclient
import requests
import urllib3
from bs4 import BeautifulSoup

import register

# 사내 위키가 사설 인증서를 쓰는 경우가 많아 검증을 끔 (Proposal_to_wikiUpload.py와 동일)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 참고: 예전 site_to_mediawikiupload.py는 truststore.inject_into_ssl()로 OS 인증서 저장소를
# 쓰게 했었는데, 그건 전역으로 ssl.SSLContext를 몽키패치하는 방식이라 register.py에서 이미
# 겪었던 "Python 3.14 + truststore 조합에서 SSL 요청이 무한 재귀에 빠지는" 버그를 이 모듈에서도
# 다시 일으킬 수 있다. 뉴스 스크래핑엔 필요 없는 기능이라 일부러 쓰지 않는다.

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


# ---------------------------------------------------------------
# URL 목록(제목/URL 번갈아 붙여넣기)을 스크랩해서 각각 위키 페이지로 업로드.
# 예전 site_to_mediawikiupload.py + sites.txt 파일을 대체 - 파일 대신 GUI 텍스트
# 붙여넣기 박스에서 바로 읽어온다.
# ---------------------------------------------------------------
SITE_UPLOAD_DEFAULT_CATEGORY = "ITS동향"

_TITLE_STRIP_RE = re.compile(r"[\[\]{}|#<>%+?]")


def clean_wiki_title(title: str) -> str:
    """미디어위키 문서 제목에 쓸 수 없는 특수문자를 제거하고, 맨 앞 '-'도 지운다."""
    cleaned = _TITLE_STRIP_RE.sub("", title).strip()
    if cleaned.startswith("-"):
        cleaned = cleaned.lstrip("-").strip()
    return cleaned


def escape_mediawiki_brackets(text: str) -> str:
    """본문 속 [ ]가 위키 링크로 오작동하지 않도록 이스케이프."""
    if not text:
        return ""
    return text.replace("[", "&#91;").replace("]", "&#93;")


def scrape_url_content(url: str) -> str:
    """URL을 받아 본문으로 보이는 텍스트만 뽑는다 (article/main/#content 우선, 없으면 body 전체)."""
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) ScraperBot/1.1"}
    response = requests.get(url, headers=headers, timeout=15)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    content_tag = soup.find("article") or soup.find("main") or soup.find("div", id="content")
    if content_tag is None:
        content_tag = soup.body
    return content_tag.get_text(separator="\n", strip=True) if content_tag else ""


def parse_site_pairs(text: str) -> list[tuple[str, str]]:
    """"제목" 한 줄, "URL" 한 줄이 번갈아 나오는 텍스트를 파싱 (빈 줄/#주석 무시).
    예전 sites.txt와 완전히 같은 형식 - 파일 대신 텍스트박스에서 붙여넣은 내용을 그대로 받는다."""
    lines = [line.strip() for line in text.splitlines() if line.strip() and not line.strip().startswith("#")]
    pairs = []
    for i in range(0, len(lines), 2):
        if i + 1 >= len(lines):
            print(f"[안내] 짝이 맞지 않아 건너뜀(URL 누락): {lines[i]}")
            break
        pairs.append((lines[i], lines[i + 1]))
    return pairs


def upload_sites_to_wiki(
    text: str, category: str = SITE_UPLOAD_DEFAULT_CATEGORY, progress_callback=None
) -> tuple[int, int]:
    """붙여넣은 "제목/URL" 목록을 하나씩 스크랩해서 위키 페이지로 업로드.
    (성공 개수, 전체 개수) 반환."""
    pairs = parse_site_pairs(text)
    if not pairs:
        print("업로드할 URL이 없습니다. '제목' 한 줄, 'URL' 한 줄 형식으로 번갈아 입력하세요.")
        return 0, 0

    print(f"위키 로그인 중...")
    site = get_wiki_site()
    print(f"로그인 성공. 분류: {category or '(없음)'}, 대상 {len(pairs)}건")

    date_prefix = datetime.now().strftime("%Y-%m-%d")
    total = len(pairs)
    ok = 0

    for i, (raw_title, url) in enumerate(pairs, start=1):
        doc_name = f"{date_prefix} {clean_wiki_title(raw_title)}"
        try:
            body_text = scrape_url_content(url)
            safe_body = escape_mediawiki_brackets(body_text)
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            category_tag = f"\n\n[[Category:{category}]]" if category else ""
            wiki_content = (
                f"== {doc_name} 자동 업데이트 ==\n\n"
                f"* '''출처:''' {url}\n"
                f"* '''마지막 수집:''' {now_str}\n\n"
                f"----\n\n"
                f"{safe_body}"
                f"{category_tag}"
            )
            page = site.pages[doc_name]
            page.save(wiki_content, summary=f"URL 붙여넣기 업로드 ({now_str})")
            print(f"업로드 완료: {doc_name} ({url})")
            ok += 1
        except Exception as e:
            print(f"[오류] {raw_title} ({url}): {e}")

        if progress_callback:
            try:
                progress_callback({
                    "file_index": i, "file_total": total, "file_name": doc_name,
                    "unit_index": 1, "unit_total": 1, "unit_label": "업로드",
                })
            except Exception:
                pass

    print(f"모든 작업 완료. {ok}/{total}개 업로드 성공.")
    return ok, total
