r"""
Qdrant 문서 등록 GUI
- 탐색기에서 파일/폴더를 드래그 앤 드롭하면 register.py의 등록 로직을 그대로 실행
- register.py가 print()로 찍는 진행 상태/오류 메시지를 하단 리스트박스에 그대로 표시

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
from tkinter import filedialog, ttk

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
        root.title("Qdrant 문서 등록")
        root.geometry("640x480")
        root.minsize(480, 360)

        self.drop_label = tk.Label(
            root,
            text="여기로 파일/폴더를 드래그 앤 드롭하세요\n(여러 개 동시 선택 가능)",
            relief="ridge", bd=2, height=6, bg="#f5f5f5", fg="#333333",
            font=("맑은 고딕", 12), justify="center",
        )
        self.drop_label.pack(fill="x", padx=10, pady=10)

        btn_frame = tk.Frame(root)
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

        progress_frame = tk.Frame(root)
        progress_frame.pack(fill="x", padx=10, pady=(8, 0))
        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.progress_bar.pack(side="left", fill="x", expand=True)
        self.progress_label = tk.Label(progress_frame, text="0%", width=5, anchor="e")
        self.progress_label.pack(side="left", padx=(8, 0))

        detail_frame = tk.Frame(root)
        detail_frame.pack(fill="x", padx=10, pady=(2, 0))
        self.file_progress_label = tk.Label(detail_frame, text="파일 -/-", fg="#333333", anchor="w")
        self.file_progress_label.pack(side="left")
        self.unit_progress_label = tk.Label(detail_frame, text="", fg="#333333", anchor="w")
        self.unit_progress_label.pack(side="left", padx=(16, 0))

        tk.Label(root, text="진행 상태 / 오류").pack(anchor="w", padx=10, pady=(10, 0))
        list_frame = tk.Frame(root)
        list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=("Consolas", 9))
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)

        if DND_AVAILABLE:
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind("<<Drop>>", self.on_drop)
        else:
            self.drop_label.config(
                text="드래그 앤 드롭을 쓰려면 'pip install tkinterdnd2' 필요\n"
                     "(지금은 아래 버튼으로 파일/폴더를 선택하세요)"
            )

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
        # tkinter 변수는 메인 스레드에서 읽고, 백그라운드 스레드에는 순수 값만 넘긴다
        process_images = self.process_images_var.get()
        threading.Thread(target=self.run_registration, args=(paths, process_images), daemon=True).start()

    def run_registration(self, paths: list[Path], process_images: bool):
        register.PROCESS_IMAGES = process_images
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.msg_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            self.log(f"이미지 처리: {'켜짐' if process_images else '꺼짐 (텍스트만 등록)'}")
            asyncio.run(register.main(paths, progress_callback=self.progress_queue.put))
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
