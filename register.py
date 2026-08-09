r"""
윈도우 로컬에서 실행하는 문서 등록 프로그램 (이미지 추출 + OCR 지원)
- 인자 없이 실행 -> 이 스크립트가 있는 폴더의 "incoming" 하위 폴더를 통째로 등록
- 특정 파일/폴더를 지정하면 그것만 등록
- PDF 안의 이미지는 별도 파일로 추출 저장 + OCR 텍스트를 함께 Qdrant에 등록

지원 파일 형식:
    PDF(.pdf), Word(.docx/.doc), 한글(.hwp/.hwpx), 엑셀(.xlsx),
    파워포인트(.pptx/.ppt), 텍스트(.txt/.md), 이미지(.jpg/.jpeg/.png)

사용법:
    python register.py
    python register.py "문서.pdf"
    python register.py "C:\Docs\등록할문서들"

설정:
    이 스크립트와 같은 폴더의 config.json 에서 mcp_url 등을 읽어옵니다.
    파일이 없으면 최초 실행 시 기본값으로 자동 생성됩니다.

사전 설치:
    pip install pymupdf python-docx pyhwp mcp pytesseract pillow requests openpyxl python-pptx

OCR을 쓰려면 Tesseract 엔진 자체도 설치해야 합니다(파이썬 패키지와는 별도):
    1) https://github.com/UB-Mannheim/tesseract/wiki 에서 윈도우 설치파일 다운로드/설치
    2) 설치 중 "Additional language data" 에서 Korean 체크 (또는 설치 후 kor.traineddata를
       tessdata 폴더에 추가)
    3) 설치 경로(보통 C:\Program Files\Tesseract-OCR\tesseract.exe)를 config.json의
       tesseract_cmd 값으로 수정
    Tesseract가 없어도 스크립트는 동작합니다 - 이 경우 이미지는 저장되지만
    OCR 텍스트 없이 "페이지/파일명"만으로 등록됩니다.

이미지 "의미" 캡션(비전 모델)을 쓰려면:
    1) Ollama 설치: https://ollama.com/download (winget install Ollama.Ollama)
    2) 모델 다운로드: ollama pull moondream
       (CPU 전용 PC 기준 가벼운 모델. GPU 있으면 llava, qwen2.5vl 등으로 교체 가능)
    Ollama가 꺼져 있거나 모델이 없어도 스크립트는 동작합니다 - 이 경우 캡션 없이
    OCR 텍스트(있으면)만으로 등록됩니다.

구버전 .doc/.ppt 지원(선택):
    antiword(.doc) 또는 LibreOffice(둘 다 변환용)가 설치되어 있으면 자동 사용됩니다.
    없으면 해당 파일은 건너뛰고 안내 메시지를 출력합니다.
"""
import asyncio
import base64
import io
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import certifi

# mcp(httpx2)가 기본으로 쓰는 truststore가 Python 3.14 + Windows 조합에서
# SSLContext.verify_mode 설정 시 무한 재귀에 빠지는 버그가 있음. SSL_CERT_FILE을
# 지정해 truststore 대신 certifi 인증서로 검증하도록 우회. (mcp import 전에 설정해야 함)
os.environ.setdefault("SSL_CERT_FILE", certifi.where())

import docx
import fitz  # PyMuPDF
import openpyxl
import requests
from mcp import ClientSession
from mcp.client.streamable_http import streamable_http_client as streamablehttp_client
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


class _SuppressSSETeardownNoise(logging.Filter):
    """mcp 라이브러리 자체 주석에 명시된 알려진 동작: 등록이 끝나고 연결을
    정리(취소)하는 순간 SSE 스트림에 마지막 이벤트가 도착하면 ClosedResourceError가
    나면서 'Error parsing SSE message'로 로그된다. 실제 데이터 등록에는 영향 없는
    타이밍성 노이즈라 이 메시지만 걸러낸다."""

    def filter(self, record: logging.LogRecord) -> bool:
        return record.getMessage() != "Error parsing SSE message"


logging.getLogger("mcp.client.streamable_http").addFilter(_SuppressSSETeardownNoise())

try:
    import pytesseract
    OCR_LIB_AVAILABLE = True
except ImportError:
    OCR_LIB_AVAILABLE = False

# ---------------------------------------------------------------
# 환경설정 (config.json에서 로드, 없으면 기본값으로 자동 생성)
# ---------------------------------------------------------------
if getattr(sys, "frozen", False):
    # PyInstaller onefile로 실행 중이면 __file__은 임시 압축 해제 폴더를 가리키므로,
    # config.json/incoming/extracted_images는 실제 exe가 있는 폴더 기준으로 잡는다.
    SCRIPT_DIR = Path(sys.executable).resolve().parent
else:
    SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = SCRIPT_DIR / "config.json"

DEFAULT_CONFIG = {
    "mcp_url": "https://blank-closing-techrepublic-lot.trycloudflare.com/mcp",  # cloudflared 재실행 시 갱신 필요
    "ollama_url": "http://localhost:11434/api/generate",
    "vision_model": "moondream",  # GPU 있으면 llava, qwen2.5vl 등으로 교체 가능
    "tesseract_cmd": r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    "process_images": "y",  # y: 이미지 추출/OCR/캡션 처리, n: 이미지는 건너뛰고 텍스트만 등록
    "store_image_base64": "y",  # y: 이미지 파일 자체를 base64로 인코딩해 Qdrant metadata에 함께 저장
}

MAX_BASE64_IMAGE_BYTES = 5 * 1024 * 1024  # 이보다 큰 이미지는 용량 문제로 base64 저장 생략


def parse_yn(value, default: bool = True) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in ("y", "yes", "true", "1")


def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception as e:
            print(f"config.json 읽기 실패, 기본값 사용: {e}")
            data = {}
    else:
        data = {}
        CONFIG_PATH.write_text(json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"config.json이 없어 기본값으로 생성했습니다: {CONFIG_PATH}")
    return {**DEFAULT_CONFIG, **data}


CONFIG = load_config()
MCP_URL = CONFIG["mcp_url"]
OLLAMA_URL = CONFIG["ollama_url"]
VISION_MODEL = CONFIG["vision_model"]
TESSERACT_CMD = CONFIG["tesseract_cmd"]
PROCESS_IMAGES = parse_yn(CONFIG["process_images"])
STORE_IMAGE_BASE64 = parse_yn(CONFIG["store_image_base64"])

if OCR_LIB_AVAILABLE:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
    OCR_AVAILABLE = Path(TESSERACT_CMD).exists()
else:
    OCR_AVAILABLE = False

CHUNK_SIZE = 800
CHUNK_OVERLAP = 120

CAPTION_PROMPT = "Describe what this image shows and what information or meaning it conveys, in detail."
# 참고: moondream은 한국어 출력 품질이 낮아(예: 한 단어만 반환) 캡션은 영어로 생성합니다.
# 이미지 속 한국어 원문은 OCR(pytesseract, kor+eng)이 별도로 처리합니다.

DEFAULT_INCOMING_DIR = SCRIPT_DIR / "incoming"
IMAGES_DIR = SCRIPT_DIR / "extracted_images"
EXTRACTED_TEXT_DIR = SCRIPT_DIR / "extracted_text"  # Qdrant에 실제로 보낸 내용을 파일별로 남기는 폴더
IMAGE_EXTS = {".jpg", ".jpeg", ".png"}


def find_soffice() -> str | None:
    """구버전 .doc/.ppt 변환에 쓸 LibreOffice(soffice) 실행 파일을 찾는다."""
    exe = shutil.which("soffice")
    if exe:
        return exe
    for candidate in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ):
        if Path(candidate).exists():
            return candidate
    return None


SOFFICE_CMD = find_soffice()


def chunk_text(text: str, size: int = CHUNK_SIZE, overlap: int = CHUNK_OVERLAP) -> list[str]:
    chunks, i = [], 0
    while i < len(text):
        chunks.append(text[i:i + size])
        i += size - overlap
    return [c for c in chunks if c.strip()]


BASE64_MAX_SIDE = 1600  # Qdrant에 저장할 base64 이미지는 원본 대신 이 크기로 축소해서 용량 절약


def encode_image_base64(img_path: Path) -> str:
    """이미지를 (필요하면) 축소한 뒤 base64 문자열로 인코딩. 로컬에 저장된 원본 파일 자체는 그대로 둔다."""
    try:
        img = Image.open(img_path)
        fmt = img.format or "PNG"
        if max(img.size) > BASE64_MAX_SIDE:
            scale = BASE64_MAX_SIDE / max(img.size)
            new_size = (round(img.width * scale), round(img.height * scale))
            img = img.resize(new_size, Image.LANCZOS)

        buf = io.BytesIO()
        try:
            img.save(buf, format=fmt)
        except Exception:
            buf = io.BytesIO()
            fmt = "PNG"
            img.save(buf, format=fmt)
        data = buf.getvalue()

        if len(data) > MAX_BASE64_IMAGE_BYTES:
            print(f"    [base64 저장 생략] {img_path.name}: 축소 후에도 {len(data) / 1024 / 1024:.1f}MB (제한 {MAX_BASE64_IMAGE_BYTES / 1024 / 1024:.0f}MB 초과)")
            return ""
        return base64.b64encode(data).decode("utf-8")
    except Exception as e:
        print(f"    [base64 인코딩 실패] {img_path.name}: {e}")
        return ""


OCR_MAX_SIDE = 2000  # OCR은 300dpi 안팎이면 충분해서, 이보다 큰 이미지는 축소 후 OCR (정확도 손실은 미미, 속도는 크게 향상)


def ocr_image(img_path: Path) -> str:
    if not OCR_AVAILABLE:
        return ""
    try:
        img = Image.open(img_path)
        if max(img.size) > OCR_MAX_SIDE:
            scale = OCR_MAX_SIDE / max(img.size)
            new_size = (round(img.width * scale), round(img.height * scale))
            img = img.resize(new_size, Image.LANCZOS)
        return pytesseract.image_to_string(img, lang="kor+eng").strip()
    except Exception as e:
        print(f"    [OCR 실패] {img_path.name}: {e}")
        return ""


def caption_image(img_path: Path) -> str:
    """로컬 Ollama 비전 모델로 이미지의 의미를 설명하는 캡션을 생성."""
    try:
        img_b64 = base64.b64encode(img_path.read_bytes()).decode("utf-8")
        resp = requests.post(
            OLLAMA_URL,
            json={
                "model": VISION_MODEL,
                "prompt": CAPTION_PROMPT,
                "images": [img_b64],
                "stream": False,
            },
            timeout=120,
        )
        resp.raise_for_status()
        return resp.json().get("response", "").strip()
    except requests.exceptions.ConnectionError:
        # Ollama가 꺼져 있음 - 조용히 건너뜀 (파일마다 반복 출력 방지)
        return ""
    except Exception as e:
        print(f"    [캡션 생성 실패] {img_path.name}: {e}")
        return ""


MAX_IMAGES_PER_PAGE = 3  # 페이지/슬라이드/시트당 처리(추출+OCR+캡션)할 최대 이미지 수, 큰 순서대로
MAX_IMAGES_NO_PAGE = 10  # hwpx처럼 고정 페이지 개념이 없는 문서에서 처리할 전체 이미지 수 상한
MIN_IMAGE_SIDE = 80  # 이보다 작은 이미지(아이콘/장식용 라인 등)는 노이즈로 보고 건너뜀


def _report(progress_callback, unit_index: int, unit_total: int, unit_label: str):
    """파일 내부 진행률 콜백을 안전하게 호출. unit_label 예: "페이지"/"슬라이드"/"시트"/"이미지".
    콜백이 없거나 실패해도 등록 작업엔 영향 없음."""
    if progress_callback is None:
        return
    try:
        progress_callback(unit_index, unit_total, unit_label)
    except Exception:
        pass


def save_and_analyze_image(image_bytes: bytes, ext: str, img_out_dir: Path, base_name: str,
                            page: int | None, index: int) -> dict:
    """이미지 바이트를 파일로 저장하고 OCR+캡션까지 처리해서 register_images용 meta dict 생성."""
    img_out_dir.mkdir(parents=True, exist_ok=True)
    img_filename = f"{base_name}_p{page}_{index}.{ext}" if page is not None else f"{base_name}_{index}.{ext}"
    img_path = img_out_dir / img_filename
    img_path.write_bytes(image_bytes)

    ocr_text = ocr_image(img_path)
    caption = caption_image(img_path)
    status = [f"OCR {len(ocr_text)}자" if ocr_text else "OCR 없음", "캡션 O" if caption else "캡션 없음"]
    print(f"    이미지 추출: {img_filename} (" + ", ".join(status) + ")")

    return {
        "page": page, "image_index": index, "image_path": str(img_path),
        "ocr_text": ocr_text, "caption": caption,
    }


def process_pdf(path: Path, progress_callback=None) -> tuple[str, list[dict]]:
    """PDF에서 (본문 텍스트, 이미지 메타데이터 목록)을 함께 추출."""
    doc = fitz.open(str(path))
    total_pages = doc.page_count
    full_text = []
    images_meta = []
    img_out_dir = IMAGES_DIR / path.stem

    for page_num, page in enumerate(doc, start=1):
        full_text.append(page.get_text())

        if not PROCESS_IMAGES:
            _report(progress_callback, page_num, total_pages, "페이지")
            continue

        # get_images(full=True) 튜플: (xref, smask, width, height, ...)
        # 너무 작은 이미지는 제외하고, 면적(width*height) 기준 큰 순서로
        # 페이지당 MAX_IMAGES_PER_PAGE장만 골라 처리 (캡션 생성이 느려서 제한)
        candidates = [img for img in page.get_images(full=True) if img[2] >= MIN_IMAGE_SIDE and img[3] >= MIN_IMAGE_SIDE]
        candidates.sort(key=lambda img: img[2] * img[3], reverse=True)
        top_images = candidates[:MAX_IMAGES_PER_PAGE]

        for img_idx, img in enumerate(top_images, start=1):
            xref = img[0]
            try:
                base_image = doc.extract_image(xref)
            except Exception as e:
                print(f"    [이미지 추출 실패] p{page_num}-{img_idx}: {e}")
                continue
            meta = save_and_analyze_image(
                base_image["image"], base_image.get("ext", "png"), img_out_dir, path.stem, page_num, img_idx,
            )
            images_meta.append(meta)
            # 페이지 하나에 느린 이미지(캡션)가 여러 장 있어도 이미지 단위로 실시간 갱신
            _report(progress_callback, img_idx, len(top_images), "이미지")

        _report(progress_callback, page_num, total_pages, "페이지")

    doc.close()
    return "\n".join(full_text), images_meta


def iter_picture_shapes(shapes):
    """pptx 도형 목록을 순회하며 그림 도형만 뽑아낸다 (그룹 안에 중첩된 그림도 재귀적으로 포함)."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            yield shape
        elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_picture_shapes(shape.shapes)


def process_pptx(path: Path, progress_callback=None) -> tuple[str, list[dict]]:
    """pptx에서 (본문 텍스트, 이미지 메타데이터 목록)을 함께 추출."""
    prs = Presentation(str(path))
    slides = list(prs.slides)
    total_slides = len(slides) or 1
    slides_text = []
    images_meta = []
    img_out_dir = IMAGES_DIR / path.stem

    for slide_num, slide in enumerate(slides, start=1):
        parts = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                parts.append(shape.text_frame.text)
            elif shape.has_table:
                for row in shape.table.rows:
                    parts.append("\t".join(cell.text for cell in row.cells))
        if parts:
            slides_text.append(f"[슬라이드 {slide_num}]\n" + "\n".join(parts))

        if not PROCESS_IMAGES:
            _report(progress_callback, slide_num, total_slides, "슬라이드")
            continue

        candidates = []
        for shape in iter_picture_shapes(slide.shapes):
            try:
                image = shape.image
                w, h = image.size
            except Exception as e:
                print(f"    [이미지 읽기 실패] 슬라이드{slide_num}: {e}")
                continue
            if w < MIN_IMAGE_SIDE or h < MIN_IMAGE_SIDE:
                continue
            candidates.append((w * h, image.blob, image.ext))
        candidates.sort(key=lambda c: c[0], reverse=True)
        top_images = candidates[:MAX_IMAGES_PER_PAGE]

        for idx, (_, blob, ext) in enumerate(top_images, start=1):
            meta = save_and_analyze_image(blob, ext, img_out_dir, path.stem, slide_num, idx)
            images_meta.append(meta)
            _report(progress_callback, idx, len(top_images), "이미지")

        _report(progress_callback, slide_num, total_slides, "슬라이드")

    return "\n".join(slides_text), images_meta


def process_xlsx(path: Path, progress_callback=None) -> tuple[str, list[dict]]:
    """xlsx에서 (본문 텍스트, 이미지 메타데이터 목록)을 함께 추출.
    이미지는 read_only 워크북에서 접근할 수 없어 이미지 처리가 켜져 있을 때만
    일반 모드로 다시 연다."""
    img_out_dir = IMAGES_DIR / path.stem
    images_meta = []

    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    try:
        lines = []
        for ws in wb.worksheets:
            lines.append(f"[시트: {ws.title}]")
            for row in ws.iter_rows(values_only=True):
                cells = [str(c) for c in row if c is not None]
                if cells:
                    lines.append("\t".join(cells))
        text = "\n".join(lines)
    finally:
        wb.close()

    if not PROCESS_IMAGES:
        _report(progress_callback, 1, 1, "시트")
        return text, images_meta

    wb_full = openpyxl.load_workbook(str(path), data_only=True)
    try:
        total_sheets = len(wb_full.worksheets) or 1
        for sheet_idx, ws in enumerate(wb_full.worksheets, start=1):
            candidates = []
            for img in getattr(ws, "_images", []):
                try:
                    data = img._data()
                    w, h = Image.open(io.BytesIO(data)).size
                except Exception as e:
                    print(f"    [이미지 읽기 실패] {ws.title}: {e}")
                    continue
                if w < MIN_IMAGE_SIDE or h < MIN_IMAGE_SIDE:
                    continue
                candidates.append((w * h, data))
            candidates.sort(key=lambda c: c[0], reverse=True)
            top_images = candidates[:MAX_IMAGES_PER_PAGE]

            for idx, (_, data) in enumerate(top_images, start=1):
                meta = save_and_analyze_image(data, "png", img_out_dir, path.stem, sheet_idx, idx)
                images_meta.append(meta)
                _report(progress_callback, idx, len(top_images), "이미지")

            _report(progress_callback, sheet_idx, total_sheets, "시트")
    finally:
        wb_full.close()

    return text, images_meta


def convert_with_libreoffice(path: Path, target_ext: str) -> Path | None:
    """LibreOffice로 구버전 .doc/.ppt 등을 target_ext로 변환. 실패 시 None."""
    if not SOFFICE_CMD:
        return None
    out_dir = Path(tempfile.mkdtemp(prefix="lo_convert_"))
    try:
        subprocess.run(
            [SOFFICE_CMD, "--headless", "--convert-to", target_ext, "--outdir", str(out_dir), str(path)],
            capture_output=True, timeout=120,
        )
        converted = out_dir / f"{path.stem}.{target_ext}"
        if converted.exists():
            return converted
    except Exception as e:
        print(f"    [LibreOffice 변환 실패] {path.name}: {e}")
    shutil.rmtree(out_dir, ignore_errors=True)
    return None


def read_hwpx(path: Path) -> str:
    """hwpx(zip+xml)에서 본문 텍스트만 추출."""
    texts = []
    with zipfile.ZipFile(path) as z:
        section_files = sorted(n for n in z.namelist() if re.match(r"Contents/section\d+\.xml$", n))
        for name in section_files:
            root = ET.fromstring(z.read(name))
            for elem in root.iter():
                if (elem.tag.endswith("}t") or elem.tag == "t") and elem.text:
                    texts.append(elem.text)
    return "\n".join(texts)


def process_hwpx(path: Path, progress_callback=None) -> tuple[str, list[dict]]:
    """hwpx에서 (본문 텍스트, 이미지 메타데이터 목록)을 함께 추출.
    hwpx는 워드처럼 흐르는 문서라 고정된 "페이지"가 없어, 이미지는 page=None으로
    저장하고 문서 내 순서(image_index)만 부여한다."""
    text = read_hwpx(path)
    images_meta = []
    if not PROCESS_IMAGES:
        _report(progress_callback, 1, 1, "이미지")
        return text, images_meta

    img_out_dir = IMAGES_DIR / path.stem
    candidates = []
    with zipfile.ZipFile(path) as z:
        bindata_names = sorted(n for n in z.namelist() if n.startswith("BinData/") and not n.endswith("/"))
        for name in bindata_names:
            data = z.read(name)
            try:
                w, h = Image.open(io.BytesIO(data)).size
            except Exception:
                continue  # 폰트 등 이미지가 아닌 바이너리는 건너뜀
            if w < MIN_IMAGE_SIDE or h < MIN_IMAGE_SIDE:
                continue
            ext = Path(name).suffix.lstrip(".").lower() or "png"
            candidates.append((w * h, data, ext))

    candidates.sort(key=lambda c: c[0], reverse=True)
    selected = candidates[:MAX_IMAGES_NO_PAGE]
    for idx, (_, data, ext) in enumerate(selected, start=1):
        meta = save_and_analyze_image(data, ext, img_out_dir, path.stem, None, idx)
        images_meta.append(meta)
        _report(progress_callback, idx, len(selected) or 1, "이미지")

    return text, images_meta


def read_doc(path: Path) -> str:
    try:
        out = subprocess.run(
            ["antiword", str(path)], capture_output=True, text=True, encoding="utf-8", errors="ignore"
        )
        if out.returncode == 0 and out.stdout.strip():
            return out.stdout
    except FileNotFoundError:
        pass

    converted = convert_with_libreoffice(path, "docx")
    if converted:
        try:
            return "\n".join(p.text for p in docx.Document(str(converted)).paragraphs)
        finally:
            shutil.rmtree(converted.parent, ignore_errors=True)

    print(f"  [건너뜀] {path.name}: .doc 읽기 도구 없음 (antiword 또는 LibreOffice 설치 필요)")
    return ""


def read_hwp(path: Path) -> str:
    """pyhwp를 라이브러리로 직접 호출 (exe로 빌드 시 sys.executable이 파이썬이 아니게 되므로
    'python -m hwp5.hwp5txt' 서브프로세스 대신 이 방식을 사용)."""
    from contextlib import closing

    from hwp5.hwp5txt import TextTransform
    from hwp5.xmlmodel import Hwp5File

    transform = TextTransform().transform_hwp5_to_text
    tmp_path = Path(tempfile.mktemp(suffix=".txt"))
    try:
        with closing(Hwp5File(str(path))) as hwp5file, open(tmp_path, "wb") as dest:
            transform(hwp5file, dest)
        return tmp_path.read_text(encoding="utf-8", errors="ignore")
    finally:
        tmp_path.unlink(missing_ok=True)


def process_ppt(path: Path, progress_callback=None) -> tuple[str, list[dict]]:
    """구버전 .ppt는 LibreOffice로 .pptx 변환 후 process_pptx로 텍스트+이미지 함께 추출."""
    converted = convert_with_libreoffice(path, "pptx")
    if not converted:
        print(f"  [건너뜀] {path.name}: .ppt 변환 불가 (LibreOffice 설치 필요)")
        _report(progress_callback, 1, 1, "슬라이드")
        return "", []
    try:
        return process_pptx(converted, progress_callback=progress_callback)
    finally:
        shutil.rmtree(converted.parent, ignore_errors=True)


def read_text_only(path: Path) -> str:
    """이미지 처리가 필요 없는 포맷(또는 텍스트만 필요한 호출)을 위한 텍스트 전용 추출.
    .pdf/.pptx/.ppt/.xlsx/.hwpx는 이미지까지 함께 뽑는 process_*() 함수를 register_file()에서 직접 사용."""
    ext = path.suffix.lower()
    try:
        if ext == ".docx":
            return "\n".join(p.text for p in docx.Document(str(path)).paragraphs)
        if ext == ".doc":
            return read_doc(path)
        if ext == ".hwp":
            return read_hwp(path)
        if ext in (".txt", ".md"):
            return path.read_text(encoding="utf-8", errors="ignore")
    except Exception as e:
        print(f"  [읽기 실패] {path.name}: {e}")
    return ""


async def register_text_chunks(session: ClientSession, path: Path, text: str) -> list[tuple[str, dict]]:
    """Qdrant에 실제로 보낸 (information, metadata) 목록을 그대로 반환 (덤프 파일 작성에 재사용)."""
    if not text.strip():
        return []
    chunks = chunk_text(text)
    records = []
    for idx, chunk in enumerate(chunks):
        metadata = {"source": str(path), "title": path.name, "type": "text", "chunk_index": idx}
        await session.call_tool("qdrant-store", {"information": chunk, "metadata": metadata})
        records.append((chunk, metadata))
    return records


async def register_images(session: ClientSession, path: Path, images_meta: list[dict]) -> list[tuple[str, dict]]:
    """Qdrant에 실제로 보낸 (information, metadata) 목록을 그대로 반환 (덤프 파일 작성에 재사용)."""
    records = []
    for meta in images_meta:
        if meta["page"] is not None:
            parts = [f"[{path.name} {meta['page']}페이지에 포함된 이미지]"]
        else:
            parts = [f"[{path.name} 이미지]"]
        if meta["caption"]:
            parts.append(f"설명: {meta['caption']}")
        if meta["ocr_text"]:
            parts.append(f"이미지 내 텍스트(OCR): {meta['ocr_text']}")
        if not meta["caption"] and not meta["ocr_text"]:
            parts.append("(캡션/OCR 텍스트 없음, 이미지 파일로만 저장됨)")
        info = "\n".join(parts)

        metadata = {
            "source": str(path), "title": path.name,
            "type": "image", "page": meta["page"],
            "image_index": meta["image_index"],
            "image_path": meta["image_path"],
            "caption": meta["caption"],
        }
        if STORE_IMAGE_BASE64:
            img_b64 = encode_image_base64(Path(meta["image_path"]))
            if img_b64:
                metadata["image_base64"] = img_b64

        await session.call_tool("qdrant-store", {"information": info, "metadata": metadata})
        records.append((info, metadata))
    return records


def write_extraction_dump(path: Path, text_records: list[tuple[str, dict]], image_records: list[tuple[str, dict]]):
    """이 파일에서 실제로 Qdrant에 등록된 내용을 사람이 읽기 좋은 텍스트 파일로 남긴다."""
    if not text_records and not image_records:
        return
    EXTRACTED_TEXT_DIR.mkdir(parents=True, exist_ok=True)
    lines = [f"=== 원본 파일: {path} ===", ""]

    lines.append(f"--- 텍스트 청크 (총 {len(text_records)}개) ---")
    for chunk, metadata in text_records:
        lines.append(f"[청크 {metadata['chunk_index']}]")
        lines.append(chunk)
        lines.append("")

    lines.append(f"--- 이미지 (총 {len(image_records)}개) ---")
    for info, metadata in image_records:
        page = metadata["page"]
        loc = f"page={page}, index={metadata['image_index']}" if page is not None else f"index={metadata['image_index']}"
        lines.append(f"[이미지: {loc}] {metadata['image_path']}")
        lines.append(info)
        lines.append("")

    dump_path = EXTRACTED_TEXT_DIR / f"{path.stem}.txt"
    dump_path.write_text("\n".join(lines), encoding="utf-8")


# 텍스트와 이미지를 함께 뽑아내는 포맷들 (이미지가 있으면 OCR+캡션까지 처리)
TEXT_AND_IMAGE_PROCESSORS = {
    ".pdf": process_pdf,
    ".pptx": process_pptx,
    ".ppt": process_ppt,
    ".xlsx": process_xlsx,
    ".hwpx": process_hwpx,
}


async def register_file(session: ClientSession, path: Path, progress_callback=None):
    ext = path.suffix.lower()

    if ext in TEXT_AND_IMAGE_PROCESSORS:
        text, images_meta = TEXT_AND_IMAGE_PROCESSORS[ext](path, progress_callback=progress_callback)
        text_records = await register_text_chunks(session, path, text)
        image_records = await register_images(session, path, images_meta) if images_meta else []
        write_extraction_dump(path, text_records, image_records)
        print(f"등록 완료: {path.name} (텍스트 {len(text_records)}청크, 이미지 {len(image_records)}개)")
        return

    if ext in IMAGE_EXTS:
        if not PROCESS_IMAGES:
            print(f"건너뜀(이미지 처리 꺼짐): {path.name}")
            return
        ocr_text = ocr_image(path)
        caption = caption_image(path)
        meta = [{"page": None, "image_index": 1, "image_path": str(path), "ocr_text": ocr_text, "caption": caption}]
        image_records = await register_images(session, path, meta)
        write_extraction_dump(path, [], image_records)
        print(f"등록 완료: {path.name} (이미지 {len(image_records)}개)")
        return

    text = read_text_only(path)
    if not text.strip():
        print(f"건너뜀(내용 없음): {path.name}")
        return
    text_records = await register_text_chunks(session, path, text)
    write_extraction_dump(path, text_records, [])
    print(f"등록 완료: {path.name} (텍스트 {len(text_records)}청크)")


async def register_targets(session: ClientSession, targets: list[Path], progress_callback=None):
    """파일/폴더가 섞인 경로 목록을 받아 실제 파일 목록으로 펼친 뒤 등록."""
    files = []
    for target in targets:
        if not target.exists():
            print(f"경로가 존재하지 않습니다: {target}")
            continue
        if target.is_dir():
            sub_files = [f for f in target.rglob("*") if f.is_file()]
            print(f"{target} 안 {len(sub_files)}개 파일 발견")
            files.extend(sub_files)
        else:
            files.append(target)

    if not files:
        print("등록할 파일이 없습니다.")
        return

    total = len(files)
    print(f"총 {total}개 파일 등록 시작...")
    for i, f in enumerate(files):
        def _unit_progress(unit_index, unit_total, unit_label, i=i, f=f):
            if progress_callback is None:
                return
            try:
                progress_callback({
                    "file_index": i + 1, "file_total": total, "file_name": f.name,
                    "unit_index": unit_index, "unit_total": unit_total, "unit_label": unit_label,
                })
            except Exception:
                pass

        await register_file(session, f, progress_callback=_unit_progress)
        _unit_progress(1, 1, "완료")  # 세부 진행률을 못 받는 포맷(.doc/.txt 등)도 파일 완료 시 확실히 갱신


async def main(targets: Path | list[Path], progress_callback=None):
    if isinstance(targets, Path):
        targets = [targets]

    if not PROCESS_IMAGES:
        print("이미지 처리: config.json의 process_images=n 설정으로 꺼져 있음 (텍스트만 등록)")
    else:
        if not OCR_LIB_AVAILABLE:
            print("참고: pytesseract/Pillow가 없어 OCR 없이 이미지만 저장합니다. (pip install pytesseract pillow)")
        elif not OCR_AVAILABLE:
            print(f"참고: Tesseract 엔진을 찾을 수 없어 OCR 없이 진행합니다. (config.json의 tesseract_cmd 확인: {TESSERACT_CMD})")

        ollama_base = OLLAMA_URL.rsplit("/api/", 1)[0]
        try:
            requests.get(f"{ollama_base}/api/tags", timeout=2)
            print(f"이미지 캡션: Ollama 연결됨 (모델: {VISION_MODEL})")
        except requests.exceptions.ConnectionError:
            print("참고: Ollama가 실행 중이 아니라 이미지 캡션 없이 진행합니다. (ollama serve 로 실행하세요)")

    if not SOFFICE_CMD:
        print("참고: LibreOffice가 없어 구버전 .ppt(및 antiword 없을 때 .doc)는 건너뜁니다.")

    async with streamablehttp_client(MCP_URL) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            await register_targets(session, targets, progress_callback=progress_callback)

    print("모든 작업 완료.")
    print(f"추출된 이미지는 여기 저장됨: {IMAGES_DIR}")
    print(f"Qdrant에 등록된 내용(텍스트+이미지 설명)은 여기서 확인 가능: {EXTRACTED_TEXT_DIR}")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        target_path = Path(sys.argv[1]).resolve()
    else:
        target_path = DEFAULT_INCOMING_DIR
        target_path.mkdir(exist_ok=True)
        print(f"인자 없이 실행됨 -> 기본 폴더 사용: {target_path}")

    asyncio.run(main(target_path))