# Qdrant 문서 등록 프로그램

로컬 문서(PDF/오피스 파일/이미지)를 분석해서 Qdrant 벡터 DB에 등록하는 프로그램. 텍스트뿐 아니라 문서 안에 포함된 이미지도 의미를 해석(캡션)해서 함께 검색되게 만드는 것이 목표.

## 구성 파일

| 파일 | 역할 |
|---|---|
| `register.py` | 핵심 로직 — 파일을 읽어서 텍스트/이미지를 추출하고 Qdrant에 등록. CLI로도 직접 실행 가능 |
| `gui.py` | Windows GUI — 드래그 앤 드롭으로 등록, 진행상태/오류를 리스트박스에 표시 (현재 주력 사용 방식) |
| `config.json` | 실행 설정 (없으면 최초 실행 시 기본값으로 자동 생성) |
| `dist/qdrant_register_gui.exe` | GUI를 PyInstaller onefile로 빌드한 배포용 실행파일 |

CLI(`register.exe`)는 더 이상 빌드하지 않음 — GUI만 사용.

## 지원 파일 형식

| 형식 | 텍스트 추출 | 내부 이미지 추출(OCR+캡션) |
|---|---|---|
| PDF (`.pdf`) | PyMuPDF | O — 페이지당 가장 큰 이미지 3장 |
| 파워포인트 (`.pptx`) | python-pptx | O — 슬라이드당 가장 큰 이미지 3장 (그룹 안 중첩 이미지 포함) |
| 파워포인트 구버전 (`.ppt`) | LibreOffice로 `.pptx` 변환 후 동일 처리 | O (LibreOffice 설치 시) |
| 엑셀 (`.xlsx`) | openpyxl | O — 시트당 가장 큰 이미지 3장 |
| 한글 (`.hwpx`) | zip 내부 XML 직접 파싱 | O — 문서 전체에서 최대 10장 (페이지 개념이 없어 `page=None`) |
| 한글 구버전 (`.hwp`) | pyhwp 라이브러리 직접 호출 | X |
| 워드 (`.docx`) | python-docx | X |
| 워드 구버전 (`.doc`) | antiword, 없으면 LibreOffice로 `.docx` 변환 | X |
| 텍스트 (`.txt`, `.md`) | 그대로 읽기 | X |
| 이미지 (`.jpg/.jpeg/.png`) | — | O — 파일 자체를 OCR+캡션 |

`.doc`/`.ppt`는 antiword 또는 LibreOffice가 없으면 안내 메시지만 찍고 건너뜀 (프로그램은 죽지 않음).

## 이미지 처리 파이프라인

문서에서 추출된 이미지마다:
1. **OCR** (Tesseract, 한국어+영어) — 이미지 안에 있는 글자를 텍스트로
2. **캡션** (로컬 Ollama `moondream` 모델) — 이미지가 "무엇을 의미하는지" 영어로 설명 생성 (moondream이 한국어 출력 품질이 낮아 캡션은 영어, 원문 한글은 OCR이 담당)
3. 위 둘을 합쳐서 하나의 검색 가능한 텍스트로 Qdrant에 등록

Qdrant에 저장되는 이미지 항목의 metadata:
```json
{
  "source": "원본 파일 경로",
  "title": "파일명",
  "type": "image",
  "page": "페이지/슬라이드/시트 번호 (없으면 null)",
  "image_index": "같은 page 안에서의 순서 (1부터)",
  "image_path": "로컬에 저장된 추출 이미지 파일 경로",
  "caption": "생성된 캡션",
  "image_base64": "이미지 파일 자체 (5MB 이하일 때만, base64)"
}
```
`image_base64`가 있으면 로컬 파일이 사라져도 Qdrant만으로 이미지를 복원할 수 있음. `source`+`page`(+`image_index`)로 정확히 특정 이미지를 filter 조회할 수 있게 설계됨 — 별도로 구현 중인 `qdrant_view_image(source, page)` 같은 "이미지를 실제로 보여주는" MCP 도구와 연동 가능.

Tesseract나 Ollama가 설치/실행되어 있지 않으면 해당 기능만 조용히 빠지고 나머지는 정상 동작.

## 설정 (`config.json`)

```json
{
  "mcp_url": "MCP 서버 주소 (cloudflared 재시작 시 갱신 필요)",
  "ollama_url": "http://localhost:11434/api/generate",
  "vision_model": "moondream",
  "tesseract_cmd": "Tesseract 실행 파일 경로",
  "process_images": "y",
  "store_image_base64": "y"
}
```
- `process_images`: `n`으로 바꾸면 이미지 추출/OCR/캡션을 전부 건너뛰고 텍스트만 등록 (GUI에서는 체크박스로 매 실행마다 덮어쓸 수 있음)
- `store_image_base64`: `n`으로 바꾸면 이미지 자체는 저장 안 하고 캡션/OCR 텍스트만 등록

## GUI (`gui.py` / `qdrant_register_gui.exe`)

- 탐색기에서 파일/폴더를 **여러 개 드래그 앤 드롭**하면 바로 등록 시작 (`tkinterdnd2`)
- "파일 선택.../폴더 선택..." 버튼으로도 등록 가능
- **"이미지 처리" 체크박스**로 실행마다 이미지 처리 on/off (config.json 편집 불필요, 기본값은 항상 꺼짐)
- **실시간 진행률**: 프로그레스바 %와 함께 "파일 X/Y", "{페이지/슬라이드/시트/이미지} A/B"를 숫자로 따로 표시 (파일 하나짜리 작업도 그 안의 이미지 단위로 갱신됨)
- 등록 진행 상태·오류가 전부 하단 리스트박스에 실시간 표시 (콘솔 창 없음)
- 등록은 백그라운드 스레드에서 실행되어 처리 중에도 창이 멈추지 않음

## 성능 관련 설계 결정

- **페이지당 이미지 3장 제한**: 실제 테스트한 50페이지 PDF에서 페이지당 이미지가 최대 80장까지 나오는 경우가 있어(반복되는 배경 이미지 등), 면적 기준 가장 큰 이미지만 골라 처리 시간을 크게 단축
- **파일별 스레드 병렬화는 채택 안 함**: 진짜 병목인 이미지 캡션(Ollama, CPU 추론)이 결국 하나의 CPU에서 순차 처리되기 때문에 스레드를 늘려도 체감 효과가 작고, MCP `ClientSession`이 asyncio 전용이라 스레드로 건드리면 안전성 문제만 생김. asyncio 동시성으로 개선할 여지는 남아있음(보류 상태)

## 알려진 제약

- `qdrant-store` MCP 툴은 텍스트만 임베딩함 — 이미지 자체로 유사도 검색(예: "이 사진과 비슷한 이미지 찾기")은 안 되고, 캡션/OCR 텍스트를 통한 검색만 가능
- `.docx`/`.doc`/`.hwp` 내부에 포함된 이미지는 아직 추출하지 않음 (필요시 추가 가능)
- LibreOffice 동시 실행 시 프로필 잠금 충돌 가능성 있음 (여러 `.doc`/`.ppt`를 동시에 변환하는 구조는 아직 아님)

## 빌드된 실행 파일 이력

| exe | 상태 |
|---|---|
| `register.exe` (CLI) | 빌드 중단 — 더 이상 사용 안 함 |
| `qdrant_register_gui.exe` (GUI) | 현재 사용 중, pptx/xlsx/hwpx 이미지 추출까지 반영된 최신 버전 |

재빌드 명령:
```bash
cd C:\Claude_DEV\qdrant_rag
py -m PyInstaller register.spec        # CLI (참고용, 미사용)
py -m PyInstaller qdrant_register_gui.spec   # GUI
```

## 이번 개발 과정에서 고친 주요 버그

- 설치된 `mcp` 2.0.0에서 `streamablehttp_client` 함수명/반환값 개수가 바뀐 것 (구버전 API 기준 코드가 깨져 있었음)
- `hwp5txt`가 PATH에 없어 `.hwp` 읽기가 항상 실패하던 것 → 라이브러리 직접 호출로 전환
- Python 3.14 + `truststore` 조합에서 HTTPS 요청 시 무한 재귀가 나는 환경 버그 → `SSL_CERT_FILE`로 `certifi` 인증서를 쓰도록 우회
- PyInstaller onefile exe에서 `__file__` 기준 경로가 임시 압축 해제 폴더를 가리키던 것 → exe 실제 위치 기준으로 수정
- PyInstaller `--windowed` 빌드에서 `sys.stdout`이 `None`이라 `print()`가 즉시 죽던 것 → 안전한 기본값으로 대체
- 등록이 끝난 뒤 MCP 연결을 정리하는 순간 SSE 스트림 타이밍 이슈로 `ClosedResourceError`가 나며 "Error parsing SSE message"가 로그에 뜨던 것 → mcp 라이브러리 자체 주석에도 명시된 알려진 무해한 동작이라 해당 로그만 필터링 (실제 등록에는 영향 없음)
