r"""
Qdrant 문서 등록 GUI (+ wiki 문서 등록 통합)
- 탐색기에서 파일/폴더를 드래그 앤 드롭하면 register.py의 등록 로직을 그대로 실행
- 하단 "위키 문서 등록" 섹션에서는 같은 파일들을 wiki_upload.py로 MediaWiki에 자동 업로드
- register.py/wiki_upload.py가 print()로 찍는 진행 상태/오류 메시지를 하단 리스트박스에 그대로 표시

사전 설치:
    pip install tkinterdnd2

실행:
    python gui.py
"""
import asyncio
import os
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

# PyInstaller --windowed(콘솔 없음)로 빌드하면 sys.stdout/stderr가 None이 되어,
# register.py가 import 시점(config.json 로드 등)에 찍는 print()가 바로 죽는다.
# register import 전에 먼저 안전한 값으로 채워둔다.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

import mcp_config_helper
import register
import self_update
import wiki_upload

APP_VERSION = "2.18.0"

# OS별 한글 표시가 자연스러운 기본 폰트 (없는 폰트를 지정해도 tkinter가 조용히
# 시스템 기본 폰트로 대체하긴 하지만, 지정 가능한 경우 더 자연스럽게 보이도록)
if sys.platform == "win32":
    KOREAN_FONT = "맑은 고딕"
elif sys.platform == "darwin":
    KOREAN_FONT = "Apple SD Gothic Neo"
else:
    KOREAN_FONT = "NanumGothic"

# 로그/검색 결과용 고정폭 폰트도 OS별로 실제 설치돼 있는 이름으로
if sys.platform == "win32":
    MONO_FONT = "Consolas"
elif sys.platform == "darwin":
    MONO_FONT = "Menlo"
else:
    MONO_FONT = "DejaVu Sans Mono"


class QueueWriter:
    """print()가 쓰는 내용을 줄 단위로 큐에 담아 GUI 스레드가 읽어가게 한다."""

    def __init__(self, q: queue.Queue):
        self.q = q
        self._buf = ""

    def write(self, s: str):
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            if line.strip():
                self.q.put(line)

    def flush(self):
        pass


def parse_drop_paths(data: str) -> list[Path]:
    """tkinterdnd2가 넘기는 드롭 문자열 파싱. 공백 포함 경로는 {..}로 감싸져 온다."""
    paths, buf, in_brace = [], "", False
    for ch in data:
        if ch == "{":
            in_brace = True
            buf = ""
        elif ch == "}":
            in_brace = False
            paths.append(buf)
            buf = ""
        elif ch == " " and not in_brace:
            if buf:
                paths.append(buf)
                buf = ""
        else:
            buf += ch
    if buf:
        paths.append(buf)
    return [Path(p) for p in paths if p]


class App:
    def __init__(self, root):
        self.root = root
        root.title(f"Qdrant 문서 등록 v{APP_VERSION}")
        root.geometry("820x1000")
        root.minsize(650, 500)

        # macOS의 일부 Tcl/Tk 빌드에서는 Entry/Text에 Command-c/v/x/a가
        # 기본 바인딩되어 있지 않아 붙여넣기가 먹지 않는 경우가 있어 명시적으로 등록
        for widget_class in ("Entry", "Text"):
            root.bind_class(widget_class, "<Command-c>", lambda e: e.widget.event_generate("<<Copy>>"))
            root.bind_class(widget_class, "<Command-v>", lambda e: e.widget.event_generate("<<Paste>>"))
            root.bind_class(widget_class, "<Command-x>", lambda e: e.widget.event_generate("<<Cut>>"))
            root.bind_class(widget_class, "<Command-a>", lambda e: e.widget.event_generate("<<SelectAll>>"))

        # 창을 스크롤해도 항상 보이도록 최상단에 고정하는 버전 표시줄
        version_bar = tk.Frame(root)
        version_bar.pack(side="top", fill="x", padx=10, pady=(4, 0))
        tk.Label(version_bar, text=f"v{APP_VERSION}", fg="#888888", font=(KOREAN_FONT, 8)).pack(side="right")
        tk.Button(version_bar, text="설정...", command=self.open_settings_dialog).pack(side="right", padx=(0, 8))

        # 명시적 사용자 요청("인증되지 않은 키로 접속할 때 ... 프로그램/GUI 시작 시 한 번",
        # "개인키, 공용키 구분해서 메시지 알려줘") - 개인/공용 저장소 키를 각각 별도로 표시.
        # 초기 텍스트는 "확인 중..."이고, _check_key_status_on_startup()이 백그라운드
        # 스레드에서 register.check_key_status()를 호출해 실제 결과로 갱신한다.
        self.personal_key_status_label = tk.Label(
            version_bar, text="개인키: 확인 중...", fg="#888888", font=(KOREAN_FONT, 8),
        )
        self.personal_key_status_label.pack(side="left")
        self.shared_key_status_label = tk.Label(
            version_bar, text="공용키: 확인 중...", fg="#888888", font=(KOREAN_FONT, 8),
        )
        self.shared_key_status_label.pack(side="left", padx=(12, 0))
        self.proposal_key_status_label = tk.Label(
            version_bar, text="전략기획실키: 확인 중...", fg="#888888", font=(KOREAN_FONT, 8),
        )
        self.proposal_key_status_label.pack(side="left", padx=(12, 0))

        # 명시적 사용자 요청: "업데이트 확인 버튼을 메인UI에 표기하고 실행할때 자동으로
        # 버전체크해서 버전이 다르면 애니메이션으로 표기해줘" - qdrant_register_gui.exe
        # 자기 자신의 버전 대상(명시적으로 재확인함: "qdrant_register_gui.exe를 원한거야").
        # Windows exe로 실행 중일 때만 의미 있음(self_update.py 참고) - 소스 실행/다른 OS는
        # 조용히 아무것도 표시하지 않는다. 클릭하면 업데이트 확인 결과에 따라 설치를 시작한다.
        self.app_update_label = tk.Label(
            version_bar, text="", font=(KOREAN_FONT, 8), cursor="hand2",
        )
        self.app_update_label.pack(side="left", padx=(12, 0))
        self.app_update_label.bind("<Button-1>", lambda e: self._on_click_app_update_label())
        self._app_update_blinking = False
        self._app_update_blink_idx = 0
        self._app_update_manifest_entry = None

        # 등록/삭제 관련 섹션이 계속 늘어나도 진행 상태 로그가 항상 보이도록,
        # 아래쪽(진행률+로그)은 창에 고정하고 위쪽 콘텐츠만 스크롤되게 분리한다.
        # side="bottom"으로 먼저 pack하는 것부터 실제 화면 맨 아래에 붙으므로,
        # "보이는 순서"의 역순(로그 -> 라벨 -> 세부진행률 -> 진행률바)으로 pack한다.
        # 로그 영역은 expand=True를 주지 않고 높이를 고정해서, 위쪽 기능 섹션들이
        # 화면 대부분을 차지하도록 한다 (로그가 길어지면 로그 자체 스크롤바로 확인).
        list_frame = tk.Frame(root)
        list_frame.pack(side="bottom", fill="x", padx=10, pady=(0, 10))
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        self.listbox = tk.Listbox(list_frame, height=8, yscrollcommand=scrollbar.set, font=(MONO_FONT, 9))
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)

        tk.Label(root, text="진행 상태 / 오류").pack(side="bottom", anchor="w", padx=10, pady=(10, 0))

        detail_frame = tk.Frame(root)
        detail_frame.pack(side="bottom", fill="x", padx=10, pady=(2, 0))
        self.file_progress_label = tk.Label(detail_frame, text="파일 -/-", fg="#333333", anchor="w")
        self.file_progress_label.pack(side="left")
        self.unit_progress_label = tk.Label(detail_frame, text="", fg="#333333", anchor="w")
        self.unit_progress_label.pack(side="left", padx=(16, 0))

        progress_frame = tk.Frame(root)
        progress_frame.pack(side="bottom", fill="x", padx=10, pady=(8, 0))
        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.progress_bar.pack(side="left", fill="x", expand=True)
        self.progress_label = tk.Label(progress_frame, text="0%", width=5, anchor="e")
        self.progress_label.pack(side="left", padx=(8, 0))

        # --- 여기부터는 스크롤 가능한 위쪽 영역 (등록/삭제 섹션들) ---
        scroll_outer = tk.Frame(root)
        scroll_outer.pack(side="top", fill="both", expand=True)
        canvas = tk.Canvas(scroll_outer, highlightthickness=0)
        outer_scrollbar = tk.Scrollbar(scroll_outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=outer_scrollbar.set)
        outer_scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        top = tk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=top, anchor="nw")
        top.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width))
        # 마우스가 스크롤 영역 위에 있을 때만 휠 스크롤을 가로챈다 (하단 로그 리스트박스 등
        # 다른 위젯의 자체 휠 스크롤과 충돌하지 않도록 진입/이탈 시에만 bind_all).
        canvas.bind("<Enter>", lambda e: canvas.bind_all(
            "<MouseWheel>", lambda ev: canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")
        ))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # --- Qdrant 문서 등록 (register.py, 벡터DB 등록 - AI가 내용을 검색/답변에 활용할 수 있게 함) ---
        qdrant_frame = tk.LabelFrame(top, text="벡터저장소 파일등록")
        qdrant_frame.pack(fill="x", padx=10, pady=(10, 0))

        # 개인/공용 체크박스는 파일/폴더/드래그앤드롭/텍스트 붙여넣기 등록 전부에 적용되는
        # 가장 중요한 설정이라 이 섹션 맨 위에 배치 (등록 전에 먼저 확인하게 함)
        target_frame = tk.Frame(qdrant_frame)
        target_frame.pack(fill="x", padx=5, pady=(5, 4))
        # 개인/공용은 서로 독립적인 체크박스 - 둘 다 체크하면 같은 내용을 두 저장소 모두에 등록.
        self.personal_store_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            target_frame, text="개인 저장소에 등록",
            variable=self.personal_store_var, font=(KOREAN_FONT, 10, "bold"),
        ).pack(side="left")
        # 공용 저장소는 팀 전체가 보게 되므로 실수로 올리는 일이 없도록 기본값은 항상 꺼짐
        self.shared_store_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            target_frame, text="공용 저장소에 등록",
            variable=self.shared_store_var, font=(KOREAN_FONT, 10, "bold"),
        ).pack(side="left", padx=(15, 0))
        # 명시적 사용자 요청("벡터 공용 저장소를 한개 더 만들고 싶어, 저장소 이름은
        # proposal_data") - 팀 공용 저장소(utinfo_docs)와는 별도인 두 번째 공용 저장소.
        # 마찬가지로 팀 전체가 보게 되므로 기본값은 항상 꺼짐.
        self.proposal_store_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            target_frame, text="전략기획실 자료저장소 등록",
            variable=self.proposal_store_var, font=(KOREAN_FONT, 10, "bold"),
        ).pack(side="left", padx=(15, 0))

        tk.Label(
            qdrant_frame,
            text="문서를 Qdrant 벡터DB에 등록해서 AI가 내용을 검색하고 답변에 활용할 수 있게 합니다",
            fg="#666666", anchor="w",
        ).pack(fill="x", padx=5, pady=(0, 0))

        self.drop_label = tk.Label(
            qdrant_frame,
            text="여기로 파일/폴더를 드래그 앤 드롭하면 Qdrant에 등록\n(여러 개 동시 선택 가능)",
            relief="ridge", bd=2, height=6, bg="#f5f5f5", fg="#333333",
            font=(KOREAN_FONT, 12), justify="center",
        )
        self.drop_label.pack(fill="x", padx=5, pady=5)

        btn_frame = tk.Frame(qdrant_frame)
        btn_frame.pack(fill="x", padx=5, pady=(0, 5))
        tk.Button(btn_frame, text="파일 선택...", command=self.browse_files).pack(side="left")
        tk.Button(btn_frame, text="폴더 선택...", command=self.browse_folder).pack(side="left", padx=5)

        # GUI 기본값은 항상 꺼짐(체크 안 함) - 이미지 처리는 느리므로 필요할 때만 켜서 사용
        self.process_images_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            btn_frame, text="이미지 처리(추출/OCR/캡션)", variable=self.process_images_var,
        ).pack(side="left", padx=10)

        self.status_label = tk.Label(btn_frame, text="● 대기 중", fg="#555555")
        self.status_label.pack(side="right")

        paste_frame = tk.LabelFrame(top, text="벡터저장소 텍스트 등록")
        paste_frame.pack(fill="x", padx=10, pady=(10, 0))

        paste_title_frame = tk.Frame(paste_frame)
        paste_title_frame.pack(fill="x", padx=5, pady=(5, 2))
        tk.Label(paste_title_frame, text="제목:").pack(side="left")
        self.paste_title_entry = tk.Entry(paste_title_frame)
        self.paste_title_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        tk.Button(
            paste_title_frame, text="붙여넣기",
            command=lambda: self._paste_from_clipboard(self.paste_title_entry),
        ).pack(side="left", padx=(5, 0))

        paste_text_toolbar = tk.Frame(paste_frame)
        paste_text_toolbar.pack(fill="x", padx=5, pady=(2, 0))
        tk.Button(
            paste_text_toolbar, text="붙여넣기",
            command=lambda: self._paste_from_clipboard(self.paste_text),
        ).pack(side="right")

        paste_text_frame = tk.Frame(paste_frame)
        paste_text_frame.pack(fill="x", padx=5)
        paste_scrollbar = tk.Scrollbar(paste_text_frame)
        paste_scrollbar.pack(side="right", fill="y")
        self.paste_text = tk.Text(
            paste_text_frame, height=6, wrap="word", yscrollcommand=paste_scrollbar.set,
            relief="solid", bd=1,
        )
        self.paste_text.pack(side="left", fill="both", expand=True)
        paste_scrollbar.config(command=self.paste_text.yview)

        tk.Button(paste_frame, text="텍스트 등록", command=self.start_pasted_registration).pack(
            anchor="e", padx=5, pady=5
        )

        delete_frame = tk.LabelFrame(top, text="벡터저장소 파일 단위삭제(되돌릴 수 없음)", fg="#a33")
        delete_frame.pack(fill="x", padx=10, pady=(10, 0))

        delete_row = tk.Frame(delete_frame)
        delete_row.pack(fill="x", padx=5, pady=5)
        tk.Label(delete_row, text="삭제할 파일:").pack(side="left")
        self.delete_source_entry = tk.Entry(delete_row)
        self.delete_source_entry.pack(side="left", fill="x", expand=True, padx=(5, 5))
        tk.Button(delete_row, text="파일 선택...", command=self.browse_delete_source).pack(side="left")
        tk.Button(delete_row, text="삭제", fg="#a33", command=self.start_delete).pack(side="left", padx=(5, 0))

        # 명시적 사용자 요청: "키워드로 검색해서 삭제 기능은 벡터 저장소 체크 박스가 있는 것
        # 중 검색해서 삭제하는 기능은 어때, 단 체크박스는 한개만 되어 있어야 함" - 예전엔
        # "공용 전용"/"내 개인 전용" 두 섹션이 따로 있었는데(전략기획실 저장소는 검색할
        # 방법 자체가 없었음), 상단 개인/공용/전략기획실 체크박스 중 정확히 1개가 켜져 있을
        # 때 그 저장소 하나를 대상으로 검색/삭제하는 섹션 하나로 통합.
        search_frame = tk.LabelFrame(top, text="키워드로 검색해서 삭제 (되돌릴 수 없음)", fg="#a33")
        search_frame.pack(fill="x", padx=10, pady=(10, 0))
        tk.Label(
            search_frame,
            text="상단 개인/공용/전략기획실 체크박스 중 정확히 1개를 체크한 상태에서 검색/삭제됩니다",
            fg="#666666", anchor="w",
        ).pack(fill="x", padx=5, pady=(5, 0))

        search_row = tk.Frame(search_frame)
        search_row.pack(fill="x", padx=5, pady=5)
        tk.Label(search_row, text="검색어:").pack(side="left")
        self.search_entry = tk.Entry(search_row)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(5, 5))
        self.search_entry.bind("<Return>", lambda e: self.start_search())
        tk.Button(search_row, text="검색", command=self.start_search).pack(side="left")

        search_result_frame = tk.Frame(search_frame)
        search_result_frame.pack(fill="x", padx=5)
        search_scrollbar = tk.Scrollbar(search_result_frame)
        search_scrollbar.pack(side="right", fill="y")
        self.search_listbox = tk.Listbox(
            search_result_frame, height=5, selectmode="extended",
            yscrollcommand=search_scrollbar.set, font=(MONO_FONT, 9),
        )
        self.search_listbox.pack(side="left", fill="both", expand=True)
        search_scrollbar.config(command=self.search_listbox.yview)
        self.search_results: list[dict] = []
        self.search_target: str | None = None  # 검색 당시 대상이었던 저장소("personal"/"shared"/"proposal")

        tk.Button(
            search_frame, text="선택 항목 삭제", fg="#a33", command=self.start_delete_selected,
        ).pack(anchor="e", padx=5, pady=5)

        # --- 위키 문서 등록 (wiki_upload.py, MediaWiki 자동 업로드) ---
        wiki_frame = tk.LabelFrame(top, text="위키 문서 등록")
        wiki_frame.pack(fill="x", padx=10, pady=(10, 0))

        self.wiki_drop_label = tk.Label(
            wiki_frame,
            text="여기로 파일/폴더를 드래그 앤 드롭하면 위키에 업로드\n(여러 개 동시 선택 가능)",
            relief="ridge", bd=2, height=4, bg="#f5f5f5", fg="#333333",
            font=(KOREAN_FONT, 11), justify="center",
        )
        self.wiki_drop_label.pack(fill="x", padx=5, pady=5)

        wiki_btn_frame = tk.Frame(wiki_frame)
        wiki_btn_frame.pack(fill="x", padx=5)
        tk.Button(wiki_btn_frame, text="파일 선택...", command=self.browse_wiki_files).pack(side="left")
        tk.Button(wiki_btn_frame, text="폴더 선택...", command=self.browse_wiki_folder).pack(side="left", padx=5)
        # 기본값은 항상 꺼짐(체크 안 함) - 첨부 업로드는 위키 서버 용량/속도 부담이 있어 필요할 때만 켜서 사용
        self.wiki_attach_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            wiki_btn_frame, text="첨부파일 업로드", variable=self.wiki_attach_var,
        ).pack(side="left", padx=(10, 0))

        wiki_category_row = tk.Frame(wiki_frame)
        wiki_category_row.pack(fill="x", padx=5, pady=5)
        tk.Label(wiki_category_row, text="현재 분류:").pack(side="left")
        self.wiki_category_var = tk.StringVar(value=wiki_upload.get_wiki_config()["wiki_category"])
        tk.Label(wiki_category_row, textvariable=self.wiki_category_var, fg="#0a5", anchor="w").pack(
            side="left", fill="x", expand=True, padx=(5, 5)
        )
        tk.Button(wiki_category_row, text="분류선택", command=self.choose_wiki_category).pack(side="left")
        tk.Button(wiki_category_row, text="분류텍스트 입력", command=self.type_wiki_category).pack(
            side="left", padx=(5, 0)
        )

        # --- URL 목록으로 위키 등록 (sites.txt 형식을 파일 대신 텍스트박스에 붙여넣기) ---
        site_frame = tk.LabelFrame(top, text="URL 목록으로 위키 등록")
        site_frame.pack(fill="x", padx=10, pady=(10, 0))

        site_label_row = tk.Frame(site_frame)
        site_label_row.pack(fill="x", padx=5, pady=(5, 0))
        tk.Label(
            site_label_row, text="제목 한 줄, URL 한 줄을 번갈아 붙여넣으세요 (빈 줄/# 주석은 무시됨)",
            fg="#666666", anchor="w",
        ).pack(side="left", fill="x", expand=True)
        tk.Button(
            site_label_row, text="붙여넣기",
            command=lambda: self._paste_from_clipboard(self.site_text),
        ).pack(side="right")

        site_text_frame = tk.Frame(site_frame)
        site_text_frame.pack(fill="x", padx=5, pady=(2, 5))
        site_scrollbar = tk.Scrollbar(site_text_frame)
        site_scrollbar.pack(side="right", fill="y")
        self.site_text = tk.Text(
            site_text_frame, height=8, wrap="word", yscrollcommand=site_scrollbar.set,
            relief="solid", bd=1,
        )
        self.site_text.pack(side="left", fill="both", expand=True)
        site_scrollbar.config(command=self.site_text.yview)

        site_category_row = tk.Frame(site_frame)
        site_category_row.pack(fill="x", padx=5, pady=(0, 5))
        tk.Label(site_category_row, text="분류:").pack(side="left")
        self.site_category_var = tk.StringVar(value=wiki_upload.SITE_UPLOAD_DEFAULT_CATEGORY)
        self.site_category_entry = tk.Entry(site_category_row, textvariable=self.site_category_var)
        self.site_category_entry.pack(side="left", fill="x", expand=True, padx=(5, 5))
        tk.Button(site_category_row, text="URL 위키에 업로드", command=self.start_site_upload).pack(side="left")

        # MCP 연동(PMS/그룹웨어/MediaWiki) 설정은 별도 섹션이 아니라 상단 "설정..." 버튼
        # 화면(탭)에 통합되어 있음 - open_settings_dialog() 참고.
        self.mcp_cfg_widgets: dict[str, dict] = {}

        if DND_AVAILABLE:
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind("<<Drop>>", self.on_drop)
            self.wiki_drop_label.drop_target_register(DND_FILES)
            self.wiki_drop_label.dnd_bind("<<Drop>>", self.on_wiki_drop)
        else:
            self.drop_label.config(
                text="드래그 앤 드롭을 쓰려면 'pip install tkinterdnd2' 필요\n"
                     "(지금은 아래 버튼으로 파일/폴더를 선택하세요)"
            )
            self.wiki_drop_label.config(text="드래그 앤 드롭을 쓰려면 'pip install tkinterdnd2' 필요")

        self.msg_queue: queue.Queue[str] = queue.Queue()
        self.progress_queue: queue.Queue[dict] = queue.Queue()
        self.busy = False
        self._status_base = ""
        self._spinner_idx = 0
        self.root.after(100, self.poll_queue)
        threading.Thread(target=self._check_key_status_on_startup, daemon=True).start()
        self_update.cleanup_stale_update_files()
        threading.Thread(target=self._check_app_update_on_startup, daemon=True).start()

    # 상태 표시줄을 정적 텍스트 대신 회전 스피너 + 색상으로 표시해, 처리 중인지 대기 중인지
    # 한눈에 구분되도록 한다. self.busy가 True인 동안 150ms마다 스스로 다시 예약해서 도는
    # 애니메이션이고, self.busy가 False가 되는 순간 다음 틱에서 스스로 멈추고 "대기 중"으로 복귀한다.
    _SPINNER_FRAMES = ["◐", "◓", "◑", "◒"]

    def set_status(self, text: str):
        """작업 시작 시 호출 - 상태 표시줄에 회전 스피너 애니메이션을 켠다."""
        self._status_base = text
        self._animate_status()

    def _animate_status(self):
        if not self.busy:
            self.status_label.config(text="● 대기 중", fg="#555555")
            return
        frame = self._SPINNER_FRAMES[self._spinner_idx % len(self._SPINNER_FRAMES)]
        self._spinner_idx += 1
        self.status_label.config(text=f"{frame} {self._status_base}", fg="#1a73e8")
        self.root.after(150, self._animate_status)

    def _check_key_status_on_startup(self):
        """프로그램 시작 시 한 번, 실제 등록/검색을 시도하기 전에 개인/공용 키가 유효한지
        미리 확인해서 상단에 표시한다(명시적 사용자 요청: "인증되지 않은 키로 접속할 때
        응답 API", "프로그램/GUI 시작 시 한 번", "개인키, 공용키 구분해서 메시지 알려줘").
        네트워크 요청(register.check_key_status)이 블로킹이라 백그라운드 스레드에서 돈다 -
        UI 갱신은 self.root.after(0, ...)로 메인 스레드에 넘기고, 로그는 self.log()가
        이미 스레드-세이프한 큐라 바로 써도 된다."""
        self._check_one_key_status(
            "개인", register.MCP_URL, self.personal_key_status_label,
        )
        # 공용/제안서 자료 저장소는 필수가 아니라 비워둘 수 있음(register.py DEFAULT_CONFIG
        # 참고) - 비어 있으면 확인 자체를 건너뛰고 "미설정"으로만 표시, 에러로 취급하지 않는다.
        if register.MCP_URL_SHARED:
            self._check_one_key_status(
                "공용", register.MCP_URL_SHARED, self.shared_key_status_label,
            )
        else:
            self.root.after(
                0,
                lambda: self.shared_key_status_label.config(text="공용키: 미설정", fg="#888888"),
            )
        if register.MCP_URL_PROPOSAL:
            self._check_one_key_status(
                "전략기획실", register.MCP_URL_PROPOSAL, self.proposal_key_status_label,
            )
        else:
            self.root.after(
                0,
                lambda: self.proposal_key_status_label.config(text="전략기획실키: 미설정", fg="#888888"),
            )

    def _check_one_key_status(self, label: str, mcp_url: str, status_widget: tk.Label):
        result = register.check_key_status(mcp_url)
        if result["ok"]:
            self.log(f"[안내] {label} 저장소 키 인증 확인됨")
            self.root.after(
                0, lambda: status_widget.config(text=f"{label}키: ✓ 정상", fg="#2e7d32"),
            )
        else:
            reason = result["reason"]
            self.log(f"[경고] {label} 저장소 키 인증 실패: {reason}")
            self.root.after(
                0, lambda: status_widget.config(text=f"{label}키: ✗ {reason}", fg="#c0392b"),
            )

    def _check_app_update_on_startup(self):
        """프로그램 시작 시 한 번, qdrant_register_gui.exe 자신의 새 버전이 있는지 확인한다
        (명시적 사용자 요청: "업데이트 확인 버튼을 메인UI에 표기하고 실행할때 자동으로
        버전체크해서 버전이 다르면 애니메이션으로 표기해줘" - qdrant_register_gui.exe 대상
        임을 명시적으로 재확인함). Windows exe로 실행 중이 아니면(소스 실행/다른 OS)
        self_update.check_for_app_update()가 ok=False를 주므로 조용히 아무것도 표시 안 함."""
        result = self_update.check_for_app_update(APP_VERSION)
        self.root.after(0, lambda: self._show_app_update_indicator(result))

    def _show_app_update_indicator(self, result: dict):
        self._app_update_manifest_entry = result.get("manifest_entry")
        if not result["ok"]:
            return  # 대상이 아니거나(다른 OS/소스 실행) 확인 실패 - 메인 화면은 조용히 넘어감
        if result["update_available"]:
            self.app_update_label.config(text=f"⬆ 새 버전 있음 (v{result['latest_version']}, 클릭해서 업데이트)")
            self._start_app_update_blink(True)
            self.log(f"[안내] 새 버전이 있습니다: v{result['latest_version']}")
        else:
            self._start_app_update_blink(False)
            self.app_update_label.config(text="", fg="#888888")

    _APP_UPDATE_BLINK_COLORS = ["#e67e22", "#ffb74d"]

    def _start_app_update_blink(self, enable: bool):
        was_blinking = self._app_update_blinking
        self._app_update_blinking = enable
        if enable and not was_blinking:
            self._animate_app_update_blink()

    def _animate_app_update_blink(self):
        if not self._app_update_blinking:
            return
        self._app_update_blink_idx = (self._app_update_blink_idx + 1) % len(self._APP_UPDATE_BLINK_COLORS)
        self.app_update_label.config(fg=self._APP_UPDATE_BLINK_COLORS[self._app_update_blink_idx])
        self.root.after(500, self._animate_app_update_blink)

    def _on_click_app_update_label(self):
        if not self._app_update_manifest_entry:
            return
        if self.busy:
            self.log("[알림] 이미 처리 중입니다. 완료 후 다시 시도하세요.")
            return
        if not messagebox.askyesno(
            "업데이트 설치",
            "새 버전을 내려받아 설치합니다. 설치가 끝나면 프로그램이 자동으로 재시작됩니다.\n"
            "계속하시겠습니까?",
        ):
            return
        self._start_app_update_blink(False)
        self.app_update_label.config(text="다운로드 중...", fg="#888888")
        threading.Thread(target=self._run_app_update_install, daemon=True).start()

    def _run_app_update_install(self):
        try:
            self_update.download_and_apply_app_update(self._app_update_manifest_entry)
        except Exception as e:
            self.root.after(0, lambda: self._show_app_update_install_result(False, str(e)))
            return
        # 배치 스크립트가 파일 교체+재실행을 맡고, 이 프로세스는 스스로 종료해야 파일 잠금이
        # 풀려서 배치 스크립트가 진행될 수 있다.
        self.root.after(0, self.root.destroy)

    def _show_app_update_install_result(self, success: bool, error: str):
        if success:
            return
        self.app_update_label.config(text="⬆ 업데이트 설치 실패 (클릭해서 재시도)", fg="#c0392b")
        self.log(f"[오류] 자동 업데이트 설치 실패: {error}")

    def open_settings_dialog(self, initial_tab: int | None = None):
        """개인/공용/제안서 자료 저장소 MCP 서버 URL, 위키 로그인 계정/비밀번호를
        config.json 파일을 직접 열지 않고 GUI에서 등록/편집(명시적 사용자 요청:
        "윈도우 등록 프로그램에 저장소 편집 기능도 추가해")."""
        win = tk.Toplevel(self.root)
        win.title("설정")
        win.geometry("680x540")
        win.minsize(560, 400)
        win.transient(self.root)

        notebook = ttk.Notebook(win)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # --- 탭 1: Qdrant 저장소 URL ---
        qdrant_tab = tk.Frame(notebook)
        notebook.add(qdrant_tab, text="Qdrant 저장소")

        tk.Label(
            qdrant_tab, text="개인 저장소 URL (mcp_url):", anchor="w", font=(KOREAN_FONT, 10),
        ).pack(fill="x", padx=10, pady=(12, 0))
        personal_entry = tk.Entry(qdrant_tab)
        personal_entry.insert(0, register.CONFIG.get("mcp_url", ""))
        personal_entry.pack(fill="x", padx=10, pady=(2, 10))

        tk.Label(
            qdrant_tab, text="공용 저장소 URL (mcp_url_shared):", anchor="w", font=(KOREAN_FONT, 10),
        ).pack(fill="x", padx=10, pady=(0, 0))
        shared_entry = tk.Entry(qdrant_tab)
        shared_entry.insert(0, register.CONFIG.get("mcp_url_shared", ""))
        shared_entry.pack(fill="x", padx=10, pady=(2, 10))

        tk.Label(
            qdrant_tab, text="전략기획실 자료저장소 URL (mcp_url_proposal):", anchor="w", font=(KOREAN_FONT, 10),
        ).pack(fill="x", padx=10, pady=(0, 0))
        proposal_entry = tk.Entry(qdrant_tab)
        proposal_entry.insert(0, register.CONFIG.get("mcp_url_proposal", ""))
        proposal_entry.pack(fill="x", padx=10, pady=(2, 10))

        def save_qdrant_urls():
            personal = personal_entry.get().strip()
            shared = shared_entry.get().strip()
            proposal = proposal_entry.get().strip()
            if not personal:
                messagebox.showwarning("입력 필요", "개인 저장소 URL은 비워둘 수 없습니다.", parent=win)
                return
            register.save_mcp_urls(personal, shared, proposal)
            self.log("[안내] Qdrant 저장소 설정을 저장했습니다 (config.json)")
            # register.save_mcp_urls()가 register.MCP_URL 등 전역값은 이미 즉시 갱신하지만,
            # 상단 "개인키/공용키/제안서키" 상태 표시줄은 시작할 때 한 번만 확인하고는 자동으로
            # 다시 확인하지 않아서, 저장 직후에는 여전히 예전 상태(또는 "미설정")로 남아있어
            # 마치 재시작해야만 반영되는 것처럼 보였다 - 저장 성공 시 바로 재확인해서 갱신한다.
            self.personal_key_status_label.config(text="개인키: 확인 중...", fg="#888888")
            self.shared_key_status_label.config(text="공용키: 확인 중...", fg="#888888")
            self.proposal_key_status_label.config(text="전략기획실키: 확인 중...", fg="#888888")
            threading.Thread(target=self._check_key_status_on_startup, daemon=True).start()

        tk.Button(qdrant_tab, text="저장", command=save_qdrant_urls).pack(anchor="e", padx=10, pady=(4, 10))

        # --- 탭 2: 위키 문서 등록 로그인 (wiki_upload.py가 mwclient로 직접 로그인할 때 씀) ---
        wiki_tab = tk.Frame(notebook)
        notebook.add(wiki_tab, text="위키(문서등록)")

        wiki_cfg = wiki_upload.get_wiki_config()

        tk.Label(
            wiki_tab, text="위키 로그인 계정 (wiki_username):", anchor="w", font=(KOREAN_FONT, 10),
        ).pack(fill="x", padx=10, pady=(12, 0))
        wiki_user_entry = tk.Entry(wiki_tab)
        wiki_user_entry.insert(0, wiki_cfg.get("wiki_username", ""))
        wiki_user_entry.pack(fill="x", padx=10, pady=(2, 10))

        tk.Label(
            wiki_tab, text="위키 로그인 비밀번호 (wiki_password):", anchor="w", font=(KOREAN_FONT, 10),
        ).pack(fill="x", padx=10, pady=(0, 0))
        wiki_pw_row = tk.Frame(wiki_tab)
        wiki_pw_row.pack(fill="x", padx=10, pady=(2, 10))
        wiki_pw_entry = tk.Entry(wiki_pw_row, show="*")
        wiki_pw_entry.insert(0, wiki_cfg.get("wiki_password", ""))
        wiki_pw_entry.pack(side="left", fill="x", expand=True)
        wiki_pw_show_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            wiki_pw_row, text="표시", variable=wiki_pw_show_var,
            command=lambda: wiki_pw_entry.config(show="" if wiki_pw_show_var.get() else "*"),
        ).pack(side="left", padx=(6, 0))

        def save_wiki_login():
            wiki_upload.save_wiki_credentials(wiki_user_entry.get().strip(), wiki_pw_entry.get())
            self.log("[안내] 위키 로그인 설정을 저장했습니다 (config.json)")

        tk.Button(wiki_tab, text="저장", command=save_wiki_login).pack(anchor="e", padx=10, pady=(4, 10))

        # --- 탭 3: MCP 연동 (Claude 데스크톱 앱의 claude_desktop_config.json) ---
        # 명시적 사용자 요청: "config MCP를 연동하려면 비 개발자가 설정하기 너무 어려워
        # 각각의 모듈로 만들어서 바로 설정값만 입력해서 사용하고 싶어" - raw JSON을 직접
        # 편집하지 않고 서버별 값(URL/계정/키)만 입력하면 mcp_config_helper.py가 올바른
        # command/args/env 구조로 claude_desktop_config.json에 병합해준다.
        mcp_tab_outer = tk.Frame(notebook)
        notebook.add(mcp_tab_outer, text="MCP 연동")

        tk.Label(
            mcp_tab_outer,
            text="사내 MCP 서버 접속 정보를 입력하면 claude_desktop_config.json에 자동으로 반영합니다"
                 " (저장 후 Claude 데스크톱 앱을 재시작해야 적용됨)",
            fg="#c0392b", font=(KOREAN_FONT, 9, "bold"), anchor="w", wraplength=560, justify="left",
        ).pack(fill="x", padx=10, pady=(10, 5))

        # 탭 안 내용이 창보다 길어질 수 있어 자체 스크롤 캔버스로 감싼다.
        mcp_canvas = tk.Canvas(mcp_tab_outer, highlightthickness=0)
        mcp_scrollbar = tk.Scrollbar(mcp_tab_outer, orient="vertical", command=mcp_canvas.yview)
        mcp_canvas.configure(yscrollcommand=mcp_scrollbar.set)
        mcp_scrollbar.pack(side="right", fill="y")
        mcp_canvas.pack(side="left", fill="both", expand=True, padx=(10, 0))
        mcp_tab = tk.Frame(mcp_canvas)
        mcp_canvas_window = mcp_canvas.create_window((0, 0), window=mcp_tab, anchor="nw")
        mcp_tab.bind("<Configure>", lambda e: mcp_canvas.configure(scrollregion=mcp_canvas.bbox("all")))
        mcp_canvas.bind("<Configure>", lambda e: mcp_canvas.itemconfig(mcp_canvas_window, width=e.width))

        self.mcp_cfg_widgets = {}
        for server_key in mcp_config_helper.SERVER_TEMPLATES:
            self._build_mcp_server_section(mcp_tab, server_key)

        tk.Button(win, text="닫기", command=win.destroy).pack(anchor="e", padx=10, pady=(0, 10))

        if initial_tab is not None:
            notebook.select(initial_tab)

    @staticmethod
    def _mcp_status_text(is_saved: bool, has_value: bool, kind: str) -> tuple[str, str]:
        """입력칸에 보이는 값이 실제로 claude_desktop_config.json에 저장된 값인지, 아니면
        자동탐지/기본값이라 아직 저장 전인지 구분해서 보여줄 (문구, 색상)을 만든다.
        명시적 사용자 요청: "(예: '자동탐지됨 - 저장 필요' / '저장됨')를 추가해"."""
        if is_saved:
            return "✓ 저장됨", "#2e7d32"
        if has_value:
            label = "자동탐지됨 - 저장 필요" if kind == "exe" else "기본값 - 저장 필요"
            return label, "#e67e22"
        return "", "#888888"

    def _build_mcp_server_section(self, parent, server_key: str):
        """server_key(mcp_config_helper.SERVER_TEMPLATES의 키) 하나에 대한 입력 폼을 만든다.
        exe로 실행하는 서버(mediawiki/groupware)는 실행 파일 경로 입력칸도 함께 넣고, 흔한
        기본 경로에 있으면 자동으로 채워준다(register.py의 Tesseract/LibreOffice 자동탐지와
        동일한 패턴) - 실행 파일 자체는 이 프로그램이 만들지 않고 이미 설치돼 있다고 가정한다.
        각 입력칸 아래에는 그 값이 이미 저장된 값인지, 자동탐지/기본값이라 아직 저장 전인지
        보여주는 작은 상태 라벨을 붙인다."""
        template = mcp_config_helper.SERVER_TEMPLATES[server_key]
        existing = mcp_config_helper.get_existing_server_config(server_key)
        existing_env = existing.get("env", {})

        box = tk.LabelFrame(parent, text=template["label"])
        box.pack(fill="x", padx=5, pady=(0, 8))

        widgets = {"fields": {}, "field_status": {}}

        if template["kind"] == "exe":
            exe_row = tk.Frame(box)
            exe_row.pack(fill="x", padx=5, pady=(5, 0))
            tk.Label(exe_row, text="실행 파일:", width=14, anchor="w").pack(side="left")
            exe_entry = tk.Entry(exe_row)
            existing_exe = existing.get("command") or ""
            default_exe = existing_exe or mcp_config_helper.find_server_exe(template["exe_candidates"])
            if default_exe:
                exe_entry.insert(0, default_exe)
            exe_entry.pack(side="left", fill="x", expand=True, padx=(5, 5))
            tk.Button(
                exe_row, text="찾아보기...",
                command=lambda e=exe_entry: self._browse_mcp_exe(e),
            ).pack(side="left")
            # 라벨+입력칸+버튼만으로도 한 줄이 창 너비를 넘기기 쉬워, 상태 문구는 같은 줄이
            # 아니라 그 아래 별도 줄에 둔다(같은 줄에 넣었더니 창 오른쪽으로 잘려 보이던 문제).
            exe_status = tk.Label(box, anchor="w", font=(KOREAN_FONT, 8))
            exe_status.pack(fill="x", padx=(19 + 5, 5), pady=(0, 4))
            text, color = self._mcp_status_text(bool(existing_exe), bool(default_exe), "exe")
            exe_status.config(text=text, fg=color)
            widgets["exe_entry"] = exe_entry
            widgets["exe_status"] = exe_status

            # 실행 파일 자체(예: mediawiki-mcp-server)의 최신 버전을 qdrant_rag 저장소
            # manifest.json과 비교해서 업데이트할 수 있는 서버만 이 UI를 보여준다 - groupware처럼
            # "최신 버전"의 출처가 없는 서버는 manifest_key가 없어 자동으로 숨겨짐.
            if template.get("manifest_key"):
                update_row = tk.Frame(box)
                update_row.pack(fill="x", padx=5, pady=(0, 4))
                update_check_btn = tk.Button(
                    update_row, text="업데이트 확인",
                    command=lambda k=server_key: self._check_mcp_server_update(k),
                )
                update_check_btn.pack(side="left")
                update_install_btn = tk.Button(
                    update_row, text="업데이트 설치",
                    command=lambda k=server_key: self._install_mcp_server_update(k),
                    state="disabled",
                )
                update_install_btn.pack(side="left", padx=(6, 0))
                update_status = tk.Label(update_row, anchor="w", font=(KOREAN_FONT, 8))
                update_status.pack(side="left", padx=(6, 0))
                widgets["update_check_btn"] = update_check_btn
                widgets["update_install_btn"] = update_install_btn
                widgets["update_status"] = update_status
                widgets["update_manifest_entry"] = None
        else:
            npx_path = mcp_config_helper.find_npx()
            note = "Node.js/npx 확인됨" if npx_path else "Node.js/npx를 찾을 수 없음 - 설치 필요 (https://nodejs.org)"
            tk.Label(box, text=note, fg="#2e7d32" if npx_path else "#c0392b", anchor="w").pack(
                fill="x", padx=5, pady=(5, 2)
            )

        for env_key, disp_label, is_secret, default_value in template["fields"]:
            row = tk.Frame(box)
            row.pack(fill="x", padx=5, pady=(2, 0))
            tk.Label(row, text=f"{disp_label}:", width=14, anchor="w").pack(side="left")
            entry = tk.Entry(row, show="*" if is_secret else "")
            existing_value = existing_env.get(env_key) or ""
            # 이미 설정된 값이 있으면 그걸 우선하고, 없으면 사내 공통 기본 주소로 미리 채운다
            # (API 키처럼 사람마다 다른 값은 default_value가 빈 문자열이라 그냥 비워둠).
            entry.insert(0, existing_value or default_value)
            entry.pack(side="left", fill="x", expand=True, padx=(5, 5))
            if is_secret:
                show_var = tk.BooleanVar(value=False)
                tk.Checkbutton(
                    row, text="표시", variable=show_var,
                    command=lambda e=entry, v=show_var: e.config(show="" if v.get() else "*"),
                ).pack(side="left")
            widgets["fields"][env_key] = entry

            # 상태 문구도 exe와 동일하게 입력칸 줄이 아니라 그 아래 별도 줄에 둬서 창 너비를
            # 넘기지 않게 한다.
            status_label = tk.Label(box, anchor="w", font=(KOREAN_FONT, 8))
            status_label.pack(fill="x", padx=(19 + 5, 5), pady=(0, 2))
            text, color = self._mcp_status_text(bool(existing_value), bool(existing_value or default_value), "field")
            status_label.config(text=text, fg=color)
            widgets["field_status"][env_key] = status_label

        tk.Button(
            box, text="저장", command=lambda k=server_key: self._save_mcp_server(k),
        ).pack(anchor="e", padx=5, pady=(2, 5))

        self.mcp_cfg_widgets[server_key] = widgets

    def _browse_mcp_exe(self, entry: tk.Entry):
        # macOS 실행 파일은 보통 확장자가 없어 "*.exe" 필터가 의미 없으므로, Windows에서만
        # exe 필터를 앞세우고 macOS/Linux에서는 모든 파일을 기본으로 보여준다.
        if sys.platform == "win32":
            filetypes = [("실행 파일", "*.exe"), ("모든 파일", "*.*")]
        else:
            filetypes = [("모든 파일", "*.*")]
        file = filedialog.askopenfilename(title="MCP 서버 실행 파일 선택", filetypes=filetypes)
        if file:
            entry.delete(0, "end")
            entry.insert(0, file)

    def _save_mcp_server(self, server_key: str):
        template = mcp_config_helper.SERVER_TEMPLATES[server_key]
        widgets = self.mcp_cfg_widgets[server_key]

        field_values = {}
        for env_key, disp_label, _, _ in template["fields"]:
            value = widgets["fields"][env_key].get().strip()
            if not value:
                messagebox.showwarning("입력 필요", f"{disp_label} 값을 입력하세요.", parent=self.root)
                return
            field_values[env_key] = value

        exe_path = None
        if template["kind"] == "exe":
            exe_path = widgets["exe_entry"].get().strip()
            if not exe_path:
                messagebox.showwarning("입력 필요", "실행 파일 경로를 입력하거나 선택하세요.", parent=self.root)
                return
            if not Path(exe_path).exists():
                if not messagebox.askyesno(
                    "파일을 찾을 수 없음",
                    f"지정한 경로에 실행 파일이 없습니다:\n{exe_path}\n\n그래도 이 경로로 저장하시겠습니까?",
                    parent=self.root,
                ):
                    return

        try:
            saved_paths = mcp_config_helper.save_server_config(server_key, field_values, exe_path=exe_path)
        except Exception as e:
            self.log(f"[오류] {template['label']} MCP 설정 저장 실패: {e}")
            return

        # 방금 저장한 값들은 이제 "자동탐지/기본값"이 아니라 "저장된 값"이므로 상태 라벨을 갱신
        for env_key in widgets["field_status"]:
            widgets["field_status"][env_key].config(text="✓ 저장됨", fg="#2e7d32")
        if "exe_status" in widgets:
            widgets["exe_status"].config(text="✓ 저장됨", fg="#2e7d32")

        paths_text = ", ".join(str(p) for p in saved_paths)
        self.log(
            f"[안내] {template['label']} MCP 설정을 저장했습니다 ({len(saved_paths)}개 경로: {paths_text}). "
            "Claude 데스크톱 앱을 재시작해야 적용됩니다."
        )

    def _check_mcp_server_update(self, server_key: str):
        widgets = self.mcp_cfg_widgets[server_key]
        exe_path = widgets["exe_entry"].get().strip()
        widgets["update_check_btn"].config(state="disabled")
        widgets["update_status"].config(text="확인 중...", fg="#888888")
        threading.Thread(target=self._run_check_update, args=(server_key, exe_path), daemon=True).start()

    def _run_check_update(self, server_key: str, exe_path: str):
        result = mcp_config_helper.check_for_server_update(server_key, exe_path)
        self.root.after(0, lambda: self._show_update_check_result(server_key, result))

    def _show_update_check_result(self, server_key: str, result: dict):
        widgets = self.mcp_cfg_widgets[server_key]
        widgets["update_check_btn"].config(state="normal")
        widgets["update_manifest_entry"] = result.get("manifest_entry")
        if not result["ok"]:
            widgets["update_status"].config(text=result["message"], fg="#c0392b")
            widgets["update_install_btn"].config(state="disabled")
            return
        if result["update_available"]:
            widgets["update_status"].config(text=result["message"], fg="#e67e22")
            widgets["update_install_btn"].config(state="normal")
        else:
            widgets["update_status"].config(text=result["message"], fg="#2e7d32")
            widgets["update_install_btn"].config(state="disabled")

    def _install_mcp_server_update(self, server_key: str):
        widgets = self.mcp_cfg_widgets[server_key]
        manifest_entry = widgets.get("update_manifest_entry")
        if not manifest_entry:
            return
        exe_path = widgets["exe_entry"].get().strip()
        if not exe_path:
            messagebox.showwarning("입력 필요", "실행 파일 경로를 먼저 입력하거나 선택하세요.", parent=self.root)
            return
        widgets["update_install_btn"].config(state="disabled")
        widgets["update_check_btn"].config(state="disabled")
        widgets["update_status"].config(text="다운로드 중...", fg="#888888")
        threading.Thread(
            target=self._run_install_update, args=(server_key, manifest_entry, exe_path), daemon=True
        ).start()

    def _run_install_update(self, server_key: str, manifest_entry: dict, exe_path: str):
        try:
            mcp_config_helper.download_server_update(manifest_entry, exe_path)
        except Exception as e:
            self.root.after(0, lambda: self._show_update_install_result(server_key, False, str(e)))
            return
        self.root.after(0, lambda: self._show_update_install_result(server_key, True, ""))

    def _show_update_install_result(self, server_key: str, success: bool, error: str):
        widgets = self.mcp_cfg_widgets[server_key]
        widgets["update_check_btn"].config(state="normal")
        if success:
            widgets["update_install_btn"].config(state="disabled")
            widgets["update_status"].config(text="✓ 설치 완료 - Claude 재시작 필요", fg="#2e7d32")
            self.log(f"[안내] {mcp_config_helper.SERVER_TEMPLATES[server_key]['label']} 실행 파일을 최신 버전으로 교체했습니다.")
        else:
            widgets["update_install_btn"].config(state="normal")
            widgets["update_status"].config(text="설치 실패", fg="#c0392b")
            self.log(f"[오류] {mcp_config_helper.SERVER_TEMPLATES[server_key]['label']} 업데이트 설치 실패: {error}")

    def log(self, msg: str):
        self.msg_queue.put(msg)

    def poll_queue(self):
        # 이 메서드는 100ms마다 스스로 재예약하는 방식으로 도는데, 안에서 예외가 나서 맨 아래
        # self.root.after(100, self.poll_queue) 줄에 도달 못 하면 재예약이 끊겨서 로그/진행률
        # 표시가 그 순간부터 앱을 재시작할 때까지 영구적으로 멈춘다 (실제로 이런 사례가 있었음 -
        # progress_queue에 dict가 아닌 값이 잘못 들어와 update_progress_display가 예외를 던진
        # 경우). 그래서 본문 전체를 try/finally로 감싸서, 무슨 일이 있어도 재예약만은 반드시
        # 일어나게 한다.
        try:
            try:
                while True:
                    line = self.msg_queue.get_nowait()
                    self.listbox.insert("end", line)
                    self.listbox.see("end")
            except queue.Empty:
                pass

            # 진행률은 짧은 시간에 여러 번 들어올 수 있어 마지막 값만 반영
            latest = None
            try:
                while True:
                    latest = self.progress_queue.get_nowait()
            except queue.Empty:
                pass
            if latest is not None:
                try:
                    self.update_progress_display(latest)
                except Exception:
                    pass
        finally:
            self.root.after(100, self.poll_queue)

    def update_progress_display(self, info: dict):
        file_index, file_total = info["file_index"], info["file_total"]
        unit_index, unit_total, unit_label = info["unit_index"], info["unit_total"], info["unit_label"]

        self.file_progress_label.config(text=f"파일 {file_index}/{file_total}")
        self.unit_progress_label.config(text=f"{unit_label} {unit_index}/{unit_total}" if unit_label else "")

        unit_frac = (unit_index / unit_total) if unit_total else 1.0
        overall_frac = ((file_index - 1) + unit_frac) / file_total if file_total else 0.0
        pct = round(min(max(overall_frac, 0.0), 1.0) * 100)
        self.progress_bar["value"] = pct
        self.progress_label.config(text=f"{pct}%")

    def on_drop(self, event):
        self.start_registration(parse_drop_paths(event.data))

    def browse_files(self):
        files = filedialog.askopenfilenames(title="등록할 파일 선택")
        if files:
            self.start_registration([Path(f) for f in files])

    def browse_folder(self):
        folder = filedialog.askdirectory(title="등록할 폴더 선택")
        if folder:
            self.start_registration([Path(folder)])

    def current_store_flags(self) -> tuple[bool, bool, bool]:
        return self.personal_store_var.get(), self.shared_store_var.get(), self.proposal_store_var.get()

    _STORE_TARGET_LABELS = {"personal": "개인", "shared": "공용", "proposal": "전략기획실"}

    def single_checked_store(self) -> str | None:
        """개인/공용/전략기획실 체크박스 중 정확히 1개만 켜져 있으면 그 이름
        ("personal"/"shared"/"proposal")을 돌려주고, 0개거나 2개 이상이면 None을 돌려준다
        (명시적 사용자 요청: "체크박스는 한개만 되어 있어야 함" - 파일 단위 삭제/키워드
        검색삭제가 어느 저장소를 대상으로 할지 상단 체크박스로 정하기 때문에, 이 기능들을
        쓸 때는 등록용 다중 체크와 달리 정확히 하나만 골라야 한다)."""
        personal, shared, proposal = self.current_store_flags()
        checked = [name for name, v in (("personal", personal), ("shared", shared), ("proposal", proposal)) if v]
        return checked[0] if len(checked) == 1 else None

    def start_registration(self, paths: list[Path]):
        if self.busy:
            self.log("[알림] 이미 처리 중입니다. 완료 후 다시 시도하세요.")
            return
        if not paths:
            return
        personal, shared, proposal = self.current_store_flags()
        if not personal and not shared and not proposal:
            self.log("[알림] 저장소를 하나 이상 체크하세요.")
            return
        self.busy = True
        self.set_status("처리 중")
        self.progress_bar["value"] = 0
        self.progress_label.config(text="0%")
        self.file_progress_label.config(text="파일 -/-")
        self.unit_progress_label.config(text="")
        # tkinter 변수는 메인 스레드에서 읽고, 백그라운드 스레드에는 순수 값만 넘긴다
        process_images = self.process_images_var.get()
        threading.Thread(
            target=self.run_registration, args=(paths, process_images, personal, shared, proposal), daemon=True
        ).start()

    def run_registration(
        self, paths: list[Path], process_images: bool, personal: bool, shared: bool, proposal: bool,
    ):
        register.PROCESS_IMAGES = process_images
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            self.log(f"이미지 처리: {'켜짐' if process_images else '꺼짐 (텍스트만 등록)'}")
            asyncio.run(
                register.main(
                    paths, progress_callback=self.progress_queue.put,
                    personal=personal, shared=shared, proposal=proposal,
                )
            )
        except Exception as e:
            self.log(f"[오류] {e}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.busy = False
            self.root.after(0, lambda: self.status_label.config(text="● 대기 중", fg="#555555"))

    def _paste_from_clipboard(self, widget):
        try:
            content = self.root.clipboard_get()
        except tk.TclError:
            return
        widget.insert(tk.INSERT, content)

    def start_pasted_registration(self):
        if self.busy:
            self.log("[알림] 이미 처리 중입니다. 완료 후 다시 시도하세요.")
            return
        text = self.paste_text.get("1.0", "end").strip()
        if not text:
            self.log("[알림] 등록할 텍스트가 비어 있습니다.")
            return
        personal, shared, proposal = self.current_store_flags()
        if not personal and not shared and not proposal:
            self.log("[알림] 저장소를 하나 이상 체크하세요.")
            return
        title = self.paste_title_entry.get().strip()
        if not title:
            # 제목 없이 등록하면 "(직접 입력) 붙여넣은 시각" 형식으로 자동 생성되어, 나중에
            # 검색/삭제할 때 이게 무슨 내용인지 알아보기 어렵다 - 실수로 빈 채로 등록하는 것을
            # 막기 위해 확인창을 띄운다.
            proceed = messagebox.askyesno(
                "제목 없음",
                "제목을 입력하지 않으면 붙여넣은 시각으로 제목이 자동 생성됩니다\n"
                "(나중에 검색·삭제할 때 내용을 알아보기 어려울 수 있습니다).\n\n"
                "제목 없이 계속 등록하시겠습니까?",
            )
            if not proceed:
                return
        self.busy = True
        self.set_status("처리 중")
        self.progress_bar["value"] = 0
        self.progress_label.config(text="0%")
        self.file_progress_label.config(text="파일 -/-")
        self.unit_progress_label.config(text="")
        threading.Thread(
            target=self.run_pasted_registration, args=(title, text, personal, shared, proposal), daemon=True
        ).start()

    def run_pasted_registration(
        self, title: str, text: str, personal: bool, shared: bool, proposal: bool,
    ):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            asyncio.run(
                register.register_pasted_text(
                    title, text, progress_callback=self.progress_queue.put,
                    personal=personal, shared=shared, proposal=proposal,
                )
            )
            self.root.after(0, lambda: self.paste_text.delete("1.0", "end"))
            self.root.after(0, lambda: self.paste_title_entry.delete(0, "end"))
        except Exception as e:
            self.log(f"[오류] {e}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.busy = False
            self.root.after(0, lambda: self.status_label.config(text="● 대기 중", fg="#555555"))

    def browse_delete_source(self):
        # 등록 때 파일 선택/드래그로 넘어가는 경로 형식(str(Path(...)))과 맞춰서
        # 여기서 고른 파일도 register.py가 저장한 source 값과 그대로 일치하게 한다.
        file = filedialog.askopenfilename(title="삭제할 파일 선택 (등록할 때와 같은 파일)")
        if file:
            self.delete_source_entry.delete(0, "end")
            self.delete_source_entry.insert(0, str(Path(file)))

    def start_delete(self):
        if self.busy:
            self.log("[알림] 이미 처리 중입니다. 완료 후 다시 시도하세요.")
            return
        source = self.delete_source_entry.get().strip()
        if not source:
            self.log("[알림] 삭제할 파일을 선택하거나 경로를 입력하세요.")
            return
        target = self.single_checked_store()
        if target is None:
            self.log("[알림] 상단 개인/공용/전략기획실 체크박스 중 정확히 하나만 체크하세요.")
            return
        label = self._STORE_TARGET_LABELS[target]
        if not messagebox.askyesno(
            "삭제 확인",
            f"다음 파일에서 등록된 모든 내용(텍스트+이미지)을 {label} 저장소에서 삭제합니다.\n"
            "이 작업은 되돌릴 수 없습니다.\n\n" + source,
        ):
            return
        self.busy = True
        self.set_status("삭제 중")
        threading.Thread(target=self.run_delete, args=(source, target), daemon=True).start()

    def run_delete(self, source: str, target: str):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            if target == "personal":
                asyncio.run(register.delete_mine_by_source(source))
            elif target == "shared":
                asyncio.run(register.delete_by_source(source))
            else:
                asyncio.run(register.delete_proposal_by_source(source))
        except Exception as e:
            self.log(f"[오류] {e}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.busy = False
            self.root.after(0, lambda: self.status_label.config(text="● 대기 중", fg="#555555"))

    def start_search(self):
        if self.busy:
            self.log("[알림] 이미 처리 중입니다. 완료 후 다시 시도하세요.")
            return
        query = self.search_entry.get().strip()
        if not query:
            self.log("[알림] 검색어를 입력하세요.")
            return
        target = self.single_checked_store()
        if target is None:
            self.log("[알림] 상단 개인/공용/전략기획실 체크박스 중 정확히 하나만 체크하세요.")
            return
        self.busy = True
        self.set_status("검색 중")
        threading.Thread(target=self.run_search, args=(query, target), daemon=True).start()

    def run_search(self, query: str, target: str):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            if target == "personal":
                results = asyncio.run(register.search_my_qdrant(query))
            elif target == "shared":
                results = asyncio.run(register.search_qdrant(query))
            else:
                results = asyncio.run(register.search_proposal_qdrant(query))
            self.search_target = target
            self.root.after(0, lambda: self.populate_search_results(results))
        except Exception as e:
            self.log(f"[오류] {e}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.busy = False
            self.root.after(0, lambda: self.status_label.config(text="● 대기 중", fg="#555555"))

    def populate_search_results(self, results: list[dict]):
        self.search_results = results
        self.search_listbox.delete(0, "end")
        label = self._STORE_TARGET_LABELS.get(self.search_target, "?")
        if not results:
            self.log(f"[알림] {label} 저장소 검색 결과가 없습니다.")
            return
        for r in results:
            meta = r.get("metadata", {})
            title = meta.get("title") or meta.get("source") or "(제목 없음)"
            if meta.get("type") == "image":
                loc = f" [이미지 p{meta.get('page')}-{meta.get('image_index')}]"
            elif meta.get("chunk_index") is not None:
                loc = f" [청크 {meta.get('chunk_index')}]"
            else:
                loc = ""
            snippet = (r.get("content") or "").replace("\n", " ").strip()[:60]
            self.search_listbox.insert("end", f"{title}{loc} — {snippet}")
        self.log(f"{label} 저장소 검색 결과 {len(results)}건 (목록에서 선택 후 '선택 항목 삭제')")

    def start_delete_selected(self):
        if self.busy:
            self.log("[알림] 이미 처리 중입니다. 완료 후 다시 시도하세요.")
            return
        indices = self.search_listbox.curselection()
        if not indices:
            self.log("[알림] 삭제할 항목을 목록에서 선택하세요.")
            return
        selected = [self.search_results[i] for i in indices]
        label = self._STORE_TARGET_LABELS.get(self.search_target, "?")
        preview = "\n".join(
            f"- {r['metadata'].get('title') or r['metadata'].get('source') or '(제목 없음)'}" for r in selected[:10]
        )
        if len(selected) > 10:
            preview += f"\n... 외 {len(selected) - 10}건"
        if not messagebox.askyesno(
            "삭제 확인",
            f"{label} 저장소에서 선택한 {len(selected)}개 항목을 삭제합니다.\n"
            "이 작업은 되돌릴 수 없습니다.\n\n" + preview,
        ):
            return
        self.busy = True
        self.set_status("삭제 중")
        threading.Thread(target=self.run_delete_selected, args=(selected, self.search_target), daemon=True).start()

    def run_delete_selected(self, selected: list[dict], target: str):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            asyncio.run(self._delete_selected(selected, target))
            self.root.after(0, lambda: self.search_listbox.delete(0, "end"))
            self.search_results = []
        except Exception as e:
            self.log(f"[오류] {e}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.busy = False
            self.root.after(0, lambda: self.status_label.config(text="● 대기 중", fg="#555555"))

    @staticmethod
    async def _delete_selected(selected: list[dict], target: str) -> int:
        total = 0
        for item in selected:
            meta = item.get("metadata", {})
            if target == "personal":
                total += await register.delete_mine_by_metadata(meta)
            elif target == "shared":
                total += await register.delete_by_metadata(meta)
            else:
                total += await register.delete_proposal_by_metadata(meta)
        print(f"선택 삭제 완료: 총 {total}개 항목 삭제")
        return total

    # --- 위키 문서 등록 ---

    def on_wiki_drop(self, event):
        self.start_wiki_upload(parse_drop_paths(event.data))

    def browse_wiki_files(self):
        files = filedialog.askopenfilenames(title="위키에 올릴 파일 선택")
        if files:
            self.start_wiki_upload([Path(f) for f in files])

    def browse_wiki_folder(self):
        folder = filedialog.askdirectory(title="위키에 올릴 폴더 선택")
        if folder:
            self.start_wiki_upload([Path(folder)])

    def choose_wiki_category(self):
        if self.busy:
            self.log("[알림] 이미 처리 중입니다. 완료 후 다시 시도하세요.")
            return
        self.busy = True
        self.set_status("분류 조회 중")
        threading.Thread(target=self.run_fetch_wiki_categories, daemon=True).start()

    def run_fetch_wiki_categories(self):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        cats = None
        try:
            site = wiki_upload.get_wiki_site()
            cats = wiki_upload.list_wiki_categories(site)
        except Exception as e:
            self.log(f"[오류] 위키 분류 목록 조회 실패: {e}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.busy = False
            self.root.after(0, lambda: self.status_label.config(text="● 대기 중", fg="#555555"))
        if cats is not None:
            self.root.after(0, lambda: self.show_wiki_category_picker(cats))

    def show_wiki_category_picker(self, categories: list[str]):
        win = tk.Toplevel(self.root)
        win.title("위키 분류 선택")
        win.geometry("300x400")
        tk.Label(win, text="목록에서 분류를 선택하세요 (더블클릭으로도 선택)").pack(anchor="w", padx=10, pady=(10, 5))

        frame = tk.Frame(win)
        frame.pack(fill="both", expand=True, padx=10)
        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side="right", fill="y")
        listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set)
        for c in categories:
            listbox.insert("end", c)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)

        def on_select():
            sel = listbox.curselection()
            if sel:
                self.wiki_category_var.set(listbox.get(sel[0]))
                win.destroy()

        tk.Button(win, text="선택", command=on_select).pack(pady=10)
        listbox.bind("<Double-Button-1>", lambda e: on_select())

    def type_wiki_category(self):
        value = simpledialog.askstring(
            "분류 입력", "적용할 분류명을 입력하세요:", initialvalue=self.wiki_category_var.get(), parent=self.root,
        )
        if value is not None and value.strip():
            self.wiki_category_var.set(value.strip())

    def start_wiki_upload(self, paths: list[Path]):
        if self.busy:
            self.log("[알림] 이미 처리 중입니다. 완료 후 다시 시도하세요.")
            return
        if not paths:
            return
        self.busy = True
        self.set_status("위키 업로드 중")
        self.progress_bar["value"] = 0
        self.progress_label.config(text="0%")
        self.file_progress_label.config(text="파일 -/-")
        self.unit_progress_label.config(text="")
        category = self.wiki_category_var.get().strip()
        upload_attachment = self.wiki_attach_var.get()
        threading.Thread(
            target=self.run_wiki_upload, args=(paths, category, upload_attachment), daemon=True
        ).start()

    def run_wiki_upload(self, paths: list[Path], category: str, upload_attachment: bool):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            wiki_upload.upload_paths_to_wiki(
                paths, category, progress_callback=self.progress_queue.put, upload_attachment=upload_attachment
            )
        except Exception as e:
            self.log(f"[오류] {e}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.busy = False
            self.root.after(0, lambda: self.status_label.config(text="● 대기 중", fg="#555555"))

    # --- URL 목록으로 위키 등록 ---

    def start_site_upload(self):
        if self.busy:
            self.log("[알림] 이미 처리 중입니다. 완료 후 다시 시도하세요.")
            return
        text = self.site_text.get("1.0", "end").strip()
        if not text:
            self.log("[알림] 제목/URL을 붙여넣으세요.")
            return
        self.busy = True
        self.set_status("URL 위키 업로드 중")
        self.progress_bar["value"] = 0
        self.progress_label.config(text="0%")
        self.file_progress_label.config(text="파일 -/-")
        self.unit_progress_label.config(text="")
        category = self.site_category_var.get().strip()
        threading.Thread(target=self.run_site_upload, args=(text, category), daemon=True).start()

    def run_site_upload(self, text: str, category: str):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            wiki_upload.upload_sites_to_wiki(text, category, progress_callback=self.progress_queue.put)
        except Exception as e:
            self.log(f"[오류] {e}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.busy = False
            self.root.after(0, lambda: self.status_label.config(text="● 대기 중", fg="#555555"))


def main():
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
