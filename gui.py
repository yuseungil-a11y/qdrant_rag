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

import register
import wiki_upload

APP_VERSION = "2.3.1"

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

        self.drop_label = tk.Label(
            top,
            text="여기로 파일/폴더를 드래그 앤 드롭하세요\n(여러 개 동시 선택 가능)",
            relief="ridge", bd=2, height=6, bg="#f5f5f5", fg="#333333",
            font=(KOREAN_FONT, 12), justify="center",
        )
        self.drop_label.pack(fill="x", padx=10, pady=10)

        target_frame = tk.Frame(top)
        target_frame.pack(fill="x", padx=10, pady=(0, 4))
        # 기본값 개인 저장소 - 파일/폴더/드래그앤드롭/텍스트 붙여넣기 등록 전부 이 설정을 따른다.
        # 팀 전체가 봐야 할 자료를 등록할 때만 체크 해제.
        self.personal_store_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            target_frame, text="개인 저장소에 등록 (체크 해제 시 팀 공유 저장소)",
            variable=self.personal_store_var, font=(KOREAN_FONT, 10, "bold"),
        ).pack(side="left")

        btn_frame = tk.Frame(top)
        btn_frame.pack(fill="x", padx=10)
        tk.Button(btn_frame, text="파일 선택...", command=self.browse_files).pack(side="left")
        tk.Button(btn_frame, text="폴더 선택...", command=self.browse_folder).pack(side="left", padx=5)

        # GUI 기본값은 항상 꺼짐(체크 안 함) - 이미지 처리는 느리므로 필요할 때만 켜서 사용
        self.process_images_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            btn_frame, text="이미지 처리(추출/OCR/캡션)", variable=self.process_images_var,
        ).pack(side="left", padx=10)

        self.status_label = tk.Label(btn_frame, text="대기 중", fg="#555555")
        self.status_label.pack(side="right")

        paste_frame = tk.LabelFrame(top, text="텍스트 붙여넣기로 등록")
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

        delete_frame = tk.LabelFrame(top, text="파일 단위 삭제 (되돌릴 수 없음)", fg="#a33")
        delete_frame.pack(fill="x", padx=10, pady=(10, 0))

        delete_row = tk.Frame(delete_frame)
        delete_row.pack(fill="x", padx=5, pady=5)
        tk.Label(delete_row, text="삭제할 파일:").pack(side="left")
        self.delete_source_entry = tk.Entry(delete_row)
        self.delete_source_entry.pack(side="left", fill="x", expand=True, padx=(5, 5))
        tk.Button(delete_row, text="파일 선택...", command=self.browse_delete_source).pack(side="left")
        tk.Button(delete_row, text="삭제", fg="#a33", command=self.start_delete).pack(side="left", padx=(5, 0))

        search_frame = tk.LabelFrame(top, text="키워드로 검색해서 삭제 (되돌릴 수 없음)", fg="#a33")
        search_frame.pack(fill="x", padx=10, pady=(10, 0))

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

        tk.Button(
            search_frame, text="선택 항목 삭제", fg="#a33", command=self.start_delete_selected,
        ).pack(anchor="e", padx=5, pady=5)

        my_search_frame = tk.LabelFrame(top, text="내 개인 저장소에서 검색해서 삭제 (되돌릴 수 없음)", fg="#a33")
        my_search_frame.pack(fill="x", padx=10, pady=(10, 0))

        my_search_row = tk.Frame(my_search_frame)
        my_search_row.pack(fill="x", padx=5, pady=5)
        tk.Label(my_search_row, text="검색어:").pack(side="left")
        self.my_search_entry = tk.Entry(my_search_row)
        self.my_search_entry.pack(side="left", fill="x", expand=True, padx=(5, 5))
        self.my_search_entry.bind("<Return>", lambda e: self.start_my_search())
        tk.Button(my_search_row, text="검색", command=self.start_my_search).pack(side="left")

        my_search_result_frame = tk.Frame(my_search_frame)
        my_search_result_frame.pack(fill="x", padx=5)
        my_search_scrollbar = tk.Scrollbar(my_search_result_frame)
        my_search_scrollbar.pack(side="right", fill="y")
        self.my_search_listbox = tk.Listbox(
            my_search_result_frame, height=4, selectmode="extended",
            yscrollcommand=my_search_scrollbar.set, font=(MONO_FONT, 9),
        )
        self.my_search_listbox.pack(side="left", fill="both", expand=True)
        my_search_scrollbar.config(command=self.my_search_listbox.yview)
        self.my_search_results: list[dict] = []

        tk.Button(
            my_search_frame, text="선택 항목 삭제", fg="#a33", command=self.start_delete_my_selected,
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

        tk.Label(
            site_frame, text="제목 한 줄, URL 한 줄을 번갈아 붙여넣으세요 (빈 줄/# 주석은 무시됨)",
            fg="#666666", anchor="w",
        ).pack(fill="x", padx=5, pady=(5, 0))

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
        self.root.after(100, self.poll_queue)

    def log(self, msg: str):
        self.msg_queue.put(msg)

    def poll_queue(self):
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
            self.update_progress_display(latest)

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

    def current_store_tool(self) -> str:
        return "qdrant_store_mine" if self.personal_store_var.get() else "qdrant-store"

    def start_registration(self, paths: list[Path]):
        if self.busy:
            self.log("[알림] 이미 처리 중입니다. 완료 후 다시 시도하세요.")
            return
        if not paths:
            return
        self.busy = True
        self.status_label.config(text="처리 중...")
        self.progress_bar["value"] = 0
        self.progress_label.config(text="0%")
        self.file_progress_label.config(text="파일 -/-")
        self.unit_progress_label.config(text="")
        # tkinter 변수는 메인 스레드에서 읽고, 백그라운드 스레드에는 순수 값만 넘긴다
        process_images = self.process_images_var.get()
        store_tool = self.current_store_tool()
        threading.Thread(target=self.run_registration, args=(paths, process_images, store_tool), daemon=True).start()

    def run_registration(self, paths: list[Path], process_images: bool, store_tool: str):
        register.PROCESS_IMAGES = process_images
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            label = "개인 저장소" if store_tool == "qdrant_store_mine" else "팀 공유 저장소"
            self.log(f"등록 대상: {label} / 이미지 처리: {'켜짐' if process_images else '꺼짐 (텍스트만 등록)'}")
            asyncio.run(register.main(paths, progress_callback=self.progress_queue.put, store_tool=store_tool))
        except Exception as e:
            self.log(f"[오류] {e}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.busy = False
            self.root.after(0, lambda: self.status_label.config(text="대기 중"))

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
        title = self.paste_title_entry.get().strip()
        store_tool = self.current_store_tool()
        self.busy = True
        self.status_label.config(text="처리 중...")
        self.progress_bar["value"] = 0
        self.progress_label.config(text="0%")
        self.file_progress_label.config(text="파일 -/-")
        self.unit_progress_label.config(text="")
        threading.Thread(target=self.run_pasted_registration, args=(title, text, store_tool), daemon=True).start()

    def run_pasted_registration(self, title: str, text: str, store_tool: str):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            asyncio.run(
                register.register_pasted_text(
                    title, text, progress_callback=self.progress_queue.put, store_tool=store_tool
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
            self.root.after(0, lambda: self.status_label.config(text="대기 중"))

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
        if not messagebox.askyesno(
            "삭제 확인",
            "다음 파일에서 등록된 모든 내용(텍스트+이미지)을 Qdrant에서 삭제합니다.\n"
            "이 작업은 되돌릴 수 없습니다.\n\n" + source,
        ):
            return
        self.busy = True
        self.status_label.config(text="삭제 중...")
        threading.Thread(target=self.run_delete, args=(source,), daemon=True).start()

    def run_delete(self, source: str):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            asyncio.run(register.delete_by_source(source))
        except Exception as e:
            self.log(f"[오류] {e}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.busy = False
            self.root.after(0, lambda: self.status_label.config(text="대기 중"))

    def start_search(self):
        if self.busy:
            self.log("[알림] 이미 처리 중입니다. 완료 후 다시 시도하세요.")
            return
        query = self.search_entry.get().strip()
        if not query:
            self.log("[알림] 검색어를 입력하세요.")
            return
        self.busy = True
        self.status_label.config(text="검색 중...")
        threading.Thread(target=self.run_search, args=(query,), daemon=True).start()

    def run_search(self, query: str):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            results = asyncio.run(register.search_qdrant(query))
            self.root.after(0, lambda: self.populate_search_results(results))
        except Exception as e:
            self.log(f"[오류] {e}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.busy = False
            self.root.after(0, lambda: self.status_label.config(text="대기 중"))

    def populate_search_results(self, results: list[dict]):
        self.search_results = results
        self.search_listbox.delete(0, "end")
        if not results:
            self.log("[알림] 검색 결과가 없습니다.")
            return
        for r in results:
            meta = r.get("metadata", {})
            title = meta.get("title") or meta.get("source", "?")
            if meta.get("type") == "image":
                loc = f" [이미지 p{meta.get('page')}-{meta.get('image_index')}]"
            elif meta.get("chunk_index") is not None:
                loc = f" [청크 {meta.get('chunk_index')}]"
            else:
                loc = ""
            snippet = (r.get("content") or "").replace("\n", " ").strip()[:60]
            self.search_listbox.insert("end", f"{title}{loc} — {snippet}")
        self.log(f"검색 결과 {len(results)}건 (목록에서 선택 후 '선택 항목 삭제')")

    def start_delete_selected(self):
        if self.busy:
            self.log("[알림] 이미 처리 중입니다. 완료 후 다시 시도하세요.")
            return
        indices = self.search_listbox.curselection()
        if not indices:
            self.log("[알림] 삭제할 항목을 목록에서 선택하세요.")
            return
        selected = [self.search_results[i] for i in indices]
        preview = "\n".join(
            f"- {r['metadata'].get('title') or r['metadata'].get('source')}" for r in selected[:10]
        )
        if len(selected) > 10:
            preview += f"\n... 외 {len(selected) - 10}건"
        if not messagebox.askyesno(
            "삭제 확인",
            f"선택한 {len(selected)}개 항목을 Qdrant에서 삭제합니다.\n"
            "이 작업은 되돌릴 수 없습니다.\n\n" + preview,
        ):
            return
        self.busy = True
        self.status_label.config(text="삭제 중...")
        threading.Thread(target=self.run_delete_selected, args=(selected,), daemon=True).start()

    def run_delete_selected(self, selected: list[dict]):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            asyncio.run(self._delete_selected(selected))
            self.root.after(0, lambda: self.search_listbox.delete(0, "end"))
            self.search_results = []
        except Exception as e:
            self.log(f"[오류] {e}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.busy = False
            self.root.after(0, lambda: self.status_label.config(text="대기 중"))

    @staticmethod
    async def _delete_selected(selected: list[dict]) -> int:
        total = 0
        for item in selected:
            total += await register.delete_by_metadata(item.get("metadata", {}))
        print(f"선택 삭제 완료: 총 {total}개 항목 삭제")
        return total

    def start_my_search(self):
        if self.busy:
            self.log("[알림] 이미 처리 중입니다. 완료 후 다시 시도하세요.")
            return
        query = self.my_search_entry.get().strip()
        if not query:
            self.log("[알림] 검색어를 입력하세요.")
            return
        self.busy = True
        self.status_label.config(text="검색 중...")
        threading.Thread(target=self.run_my_search, args=(query,), daemon=True).start()

    def run_my_search(self, query: str):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            results = asyncio.run(register.search_my_qdrant(query))
            self.root.after(0, lambda: self.populate_my_search_results(results))
        except Exception as e:
            self.log(f"[오류] {e}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.busy = False
            self.root.after(0, lambda: self.status_label.config(text="대기 중"))

    def populate_my_search_results(self, results: list[dict]):
        self.my_search_results = results
        self.my_search_listbox.delete(0, "end")
        if not results:
            self.log("[알림] 개인 저장소 검색 결과가 없습니다.")
            return
        for r in results:
            meta = r.get("metadata", {})
            title = meta.get("title") or meta.get("source") or "(제목 없음)"
            snippet = (r.get("content") or "").replace("\n", " ").strip()[:60]
            self.my_search_listbox.insert("end", f"{title} — {snippet}")
        self.log(f"개인 저장소 검색 결과 {len(results)}건 (목록에서 선택 후 '선택 항목 삭제')")

    def start_delete_my_selected(self):
        if self.busy:
            self.log("[알림] 이미 처리 중입니다. 완료 후 다시 시도하세요.")
            return
        indices = self.my_search_listbox.curselection()
        if not indices:
            self.log("[알림] 삭제할 항목을 목록에서 선택하세요.")
            return
        selected = [self.my_search_results[i] for i in indices]
        preview = "\n".join(
            f"- {r['metadata'].get('title') or r['metadata'].get('source') or '(제목 없음)'}" for r in selected[:10]
        )
        if len(selected) > 10:
            preview += f"\n... 외 {len(selected) - 10}건"
        if not messagebox.askyesno(
            "삭제 확인",
            f"내 개인 저장소에서 선택한 {len(selected)}개 항목을 삭제합니다.\n"
            "이 작업은 되돌릴 수 없습니다.\n\n" + preview,
        ):
            return
        self.busy = True
        self.status_label.config(text="삭제 중...")
        threading.Thread(target=self.run_delete_my_selected, args=(selected,), daemon=True).start()

    def run_delete_my_selected(self, selected: list[dict]):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            asyncio.run(self._delete_my_selected(selected))
            self.root.after(0, lambda: self.my_search_listbox.delete(0, "end"))
            self.my_search_results = []
        except Exception as e:
            self.log(f"[오류] {e}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.busy = False
            self.root.after(0, lambda: self.status_label.config(text="대기 중"))

    @staticmethod
    async def _delete_my_selected(selected: list[dict]) -> int:
        total = 0
        for item in selected:
            total += await register.delete_mine_by_metadata(item.get("metadata", {}))
        print(f"개인 저장소 선택 삭제 완료: 총 {total}개 항목 삭제")
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
        self.status_label.config(text="분류 조회 중...")
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
            self.root.after(0, lambda: self.status_label.config(text="대기 중"))
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
        self.status_label.config(text="위키 업로드 중...")
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
            self.root.after(0, lambda: self.status_label.config(text="대기 중"))

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
        self.status_label.config(text="URL 위키 업로드 중...")
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
            self.root.after(0, lambda: self.status_label.config(text="대기 중"))


def main():
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
