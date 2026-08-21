# 변경 이력

버전 규칙: `major.minor.patch`
- **major**: 주요 기능 변화
- **minor**: 마이너 기능 변화
- **patch**: 버그 수정

## 2.2.0

- **마이너 기능**: "위키 문서 등록"(파일/폴더 선택·드래그앤드롭)에서 텍스트 등록에 이어 **원본 파일 자체도
  첨부파일로 함께 업로드**하도록 함 (`wiki_upload.upload_file_attachment()`, `mwclient.Site.upload()` 사용).
  업로드 성공 시 문서 본문 끝에 `== 첨부파일 ==` 절로 `[[Media:...]]` 링크를 자동으로 추가.
  위키 서버 설정(허용 확장자/업로드 용량 제한 등)에 따라 첨부만 거부될 수 있는데, 그 경우 예외 없이
  안내만 찍고 본문 텍스트 등록은 그대로 진행 (기존 항목 재등록 시 덮어쓰도록 `ignore=True` 사용).
  "URL 목록으로 위키 등록"(뉴스 스크랩)은 로컬 원본 파일이 없으므로 대상에서 제외.

## 2.1.0

- **마이너 기능**: "위키 문서 등록" 섹션 아래에 **"URL 목록으로 위키 등록"** 추가. 제목/URL을 한 줄씩
  번갈아 텍스트박스에 붙여넣으면 각 URL을 스크랩해서 별도의 위키 페이지로 업로드. 예전
  `site_to_mediawikiupload.py` + `sites.txt` 파일 조합을 대체 - 파일 편집 없이 GUI에서 바로 붙여넣기
  (`wiki_upload.upload_sites_to_wiki()`, 형식/제목 정제/카테고리 태그는 원본 스크립트와 동일하게 유지)
  - 원본 스크립트가 쓰던 `truststore.inject_into_ssl()`은 일부러 빼고 순수 `requests`만 사용 - 그
    함수는 전역으로 `ssl.SSLContext`를 패치해서, 1.1.0에서 고쳤던 "Python 3.14 + truststore 조합
    SSL 무한 재귀" 버그를 다시 일으킬 위험이 있었음

## 2.0.0

- **주요 기능**: 하나의 GUI(`gui.py`)에서 Qdrant 등록과 **위키(MediaWiki) 문서 등록을 함께 사용**할 수 있도록 통합.
  - 새 모듈 `wiki_upload.py`: 문서를 MediaWiki 페이지로 자동 업로드. 예전 `Proposal_to_wikiUpload.py`가
    win32com으로 Excel/한글/PowerPoint를 직접 열어 PDF로 변환한 뒤 텍스트를 뽑던 것과 달리, 이미 있는
    `register.py`의 텍스트 추출기(PDF/PPTX/PPT/XLSX/HWPX/HWP/DOCX/DOC/HTML/TXT/MD, 전부 순수 Python)를
    재사용 - Office 설치나 Windows 종속성 없이 macOS/Linux에서도 동일하게 동작
  - GUI에 "위키 문서 등록" 섹션 추가: 전용 드래그 앤 드롭 영역, 파일/폴더 선택, "분류선택"(위키에 있는
    분류 중 골라서 적용)/"분류텍스트 입력"(새 분류명 직접 입력) 버튼. 기존 진행률 표시줄·로그창을 그대로 공유
  - 접속 정보(`key.txt`)를 `config.json`의 `wiki_site_url`/`wiki_path`/`wiki_username`/`wiki_password`/
    `wiki_category`로 이전 - 다른 설정과 마찬가지로 한 곳에서 관리하고 git에는 올라가지 않음
- **마이너 기능**: Qdrant 등록 지원 형식에 HTML(`.html`/`.htm`) 추가. `lxml.html`로 `<script>`/`<style>`을
  제거하고 블록 요소(문단/제목/리스트 등) 경계에 줄바꿈을 넣어 읽기 좋은 텍스트로 추출 (`register.read_html`).

## 1.1.3

- **버그 수정**: macOS `.app` 번들로 빌드하면 `config.json`/`incoming`/`extracted_images`/`extracted_text`가
  `qdrant_register_gui.app/Contents/MacOS/` 안쪽 깊숙이 생성되어 Finder에서 접근하기 불편했던 문제 수정.
  이제 macOS `.app` 번들에서 실행 중이면 번들이 놓인(사용자 눈에 보이는) 폴더를 기준으로 잡아, Windows exe와
  동일하게 "실행 파일과 같은 경로"에 생성/조회되도록 함 (`register.py`의 `SCRIPT_DIR` 계산 수정).

## 1.1.2

- **버그 수정**: 1.1.1의 Command-v 키 바인딩 수정만으로는 일부 macOS 환경에서 여전히 붙여넣기가 안 되는
  것으로 확인됨(Tk 기본 메뉴가 Cmd+V를 먼저 가로채는 것으로 추정, 정확한 원인 미확정). 키보드 단축키에
  의존하지 않도록 "텍스트 붙여넣기로 등록" 섹션의 제목/본문 입력창 옆에 **붙여넣기 버튼**을 추가해서
  클립보드 내용을 직접 삽입하도록 함.

## 1.1.1

- **버그 수정**: macOS에서 Entry/Text 위젯(텍스트 붙여넣기 제목/본문, 검색창 등)에 Command-V(붙여넣기)/Command-C(복사)/Command-X(잘라내기)/Command-A(전체선택)가
  동작하지 않던 문제 수정. 일부 Tcl/Tk 빌드는 macOS 기본 단축키 바인딩이 빠져있어 명시적으로 바인딩 추가.

## 1.1.0

- **마이너 기능**: Windows/macOS/Linux 크로스플랫폼 지원.
  - Tesseract, LibreOffice 실행 파일을 PATH 우선 탐색 후 OS별(Windows/macOS Homebrew/Linux)
    흔한 설치 경로로 자동 탐색 (`find_tesseract`, `find_soffice`)
  - GUI 폰트(한글 표시용, 로그/검색결과 고정폭용)를 OS별로 자동 선택
  - 코드 자체는 이제 플랫폼 무관하게 동작하지만, macOS/Linux용 실행파일은
    해당 OS에서 직접 PyInstaller로 빌드해야 함 (크로스 컴파일 불가 — `python gui.py`로는
    바로 실행 가능)

## 1.0.1

- **버그 수정**: 청크 분할(`chunk_text`)이 문자 수로만 잘라서 엑셀 행이나 문장이 청크 경계
  중간에서 잘리던 문제 수정. 이제 줄(line) 단위를 지켜서 나누고, 한 줄 자체가 청크 크기보다
  긴 경우에만 예외적으로 그 줄만 문자 단위로 쪼갠다. (네이버시스템_제안서_리스트.xlsx 등록 후
  검색 결과가 잘려 보인다는 리포트로 발견)

## 1.0.0

최초 릴리스 기준선. 이 시점까지 만들어진 기능:

- PDF/PPTX/PPT/XLSX/HWPX/HWP/DOCX/DOC/TXT/MD/이미지 파일 등록 (텍스트+이미지 함께 추출)
- 이미지 OCR(Tesseract) + 캡션(로컬 Ollama moondream) 자동 생성
- 페이지/슬라이드/시트당 가장 큰 이미지 3장으로 제한 (성능)
- OCR/캡션 저장용 이미지 다운스케일
- Qdrant에 등록된 내용을 로컬 텍스트로 덤프 (`extracted_text/`)
- 드래그 앤 드롭 + 파일/폴더 선택 + 텍스트 붙여넣기 등록
- 개인 저장소 / 팀 공유 저장소 선택
- 파일 단위 삭제, 키워드 검색 후 선택 삭제 (개인/팀 각각)
- 실시간 진행률(파일 단위 + 페이지/슬라이드/이미지 단위)
- 저장 속도 개선: 서버 배치 저장 API 우선 사용, 없으면 동시 저장으로 자동 폴백
- 화면 상단 버전 표시
