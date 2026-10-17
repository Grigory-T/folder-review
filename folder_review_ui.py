"""Desktop folder selection and progress for the shared Folder Review scanner."""

from __future__ import annotations

import os
import queue
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import folder_review as core

BACKGROUND = "#F3F4F6"
CARD = "#FFFFFF"
TEXT = "#17212F"
MUTED = "#586574"
ACTION_WIDTH = 20


class Application:
    def __init__(self, window, directory):
        self.window = window
        self.messages = queue.Queue()
        self.busy = False
        self.last_report = None
        self.poll_id = None
        self.folder = tk.StringVar()
        self.output = tk.StringVar(value=str(Path(directory) / core.OUTPUT_FILE))
        self.workers = tk.StringVar(value="auto")
        self.full_metadata = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="Choose a folder to begin.")
        self.report_name = tk.StringVar(value="No report created in this session.")
        window.title("Folder Review")
        window.configure(background=BACKGROUND)
        window.protocol("WM_DELETE_WINDOW", self.close)
        style = ttk.Style(window)
        style.configure(".", font=("Segoe UI", 10))
        style.configure("App.TFrame", background=BACKGROUND)
        style.configure("App.TLabel", background=BACKGROUND, foreground=TEXT)
        style.configure("Card.TFrame", background=CARD)
        style.configure("Card.TLabel", background=CARD, foreground=TEXT)
        style.configure(
            "Muted.Card.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9)
        )
        style.configure(
            "Title.Card.TLabel",
            background=CARD,
            foreground=TEXT,
            font=("Segoe UI", 14, "bold"),
        )
        style.configure(
            "Step.Card.TLabel",
            background="#E5EFFB",
            foreground="#175C9E",
            font=("Segoe UI", 11, "bold"),
            padding=(9, 4),
        )
        style.configure("Action.TButton", padding=(10, 6))
        style.configure("Card.TCheckbutton", background=CARD)
        outer = ttk.Frame(window, style="App.TFrame", padding=20)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1)
        ttk.Label(
            outer,
            text="Folder Review",
            style="App.TLabel",
            font=("Segoe UI", 18, "bold"),
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            outer,
            text="Choose a folder to create an Excel inventory.",
            style="App.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 16))
        self.canvas = tk.Canvas(
            outer, background=BACKGROUND, highlightthickness=0, width=1
        )
        self.canvas.grid(row=2, column=0, sticky="nsew")
        self.scrollbar = ttk.Scrollbar(
            outer, orient="vertical", command=self.canvas.yview
        )
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        border = tk.Frame(
            self.canvas,
            background=CARD,
            highlightbackground="#D7DDE4",
            highlightthickness=1,
        )
        self.card_border = border
        self.canvas_item = self.canvas.create_window(0, 0, window=border, anchor="nw")
        self.canvas.bind("<Configure>", self.resize_card)
        border.bind("<Configure>", self.update_scroll)
        window.bind("<MouseWheel>", self.scroll_card)
        window.bind("<Button-4>", self.scroll_card)
        window.bind("<Button-5>", self.scroll_card)
        card = ttk.Frame(border, style="Card.TFrame", padding=14)
        card.pack(fill="both", expand=True)
        card.columnconfigure(1, weight=1)
        card.rowconfigure(9, weight=1)
        heading = ttk.Frame(card, style="Card.TFrame")
        heading.grid(row=0, column=0, columnspan=3, sticky="ew")
        ttk.Label(heading, text="1", style="Step.Card.TLabel").pack(
            side="left", padx=(0, 10)
        )
        ttk.Label(heading, text="Review a folder", style="Title.Card.TLabel").pack(
            side="left"
        )
        explanation = ttk.Label(
            card,
            text="Browse for a folder or paste its path. Local folders and SMB shares are supported.",
            style="Muted.Card.TLabel",
        )
        explanation.grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 12))
        card.bind(
            "<Configure>",
            lambda event: explanation.configure(wraplength=max(100, event.width - 28)),
        )
        ttk.Label(card, text="Folder", style="Card.TLabel").grid(
            row=2, column=0, sticky="w", padx=(0, 14)
        )
        self.folder_entry = ttk.Entry(card, textvariable=self.folder)
        self.folder_entry.grid(row=2, column=1, sticky="ew")
        self.browse_button = ttk.Button(
            card,
            text="Browse…",
            width=ACTION_WIDTH,
            style="Action.TButton",
            command=self.browse,
        )
        self.browse_button.grid(row=2, column=2, padx=(12, 0))
        ttk.Label(card, text="Excel report", style="Card.TLabel").grid(
            row=3, column=0, sticky="w", padx=(0, 14), pady=(8, 0)
        )
        self.output_entry = ttk.Entry(card, textvariable=self.output)
        self.output_entry.grid(row=3, column=1, sticky="ew", pady=(8, 0))
        self.save_button = ttk.Button(
            card,
            text="Save as…",
            width=ACTION_WIDTH,
            style="Action.TButton",
            command=self.save_as,
        )
        self.save_button.grid(row=3, column=2, padx=(12, 0), pady=(8, 0))
        ttk.Label(card, text="Read workers", style="Card.TLabel").grid(
            row=4, column=0, sticky="w", pady=(10, 0)
        )
        options = ttk.Frame(card, style="Card.TFrame")
        options.grid(row=4, column=1, columnspan=2, sticky="w", pady=(10, 0))
        self.worker_box = ttk.Combobox(
            options,
            textvariable=self.workers,
            values=("auto", "4", "8", "16", "32", "64"),
            state="readonly",
            width=6,
        )
        self.worker_box.pack(side="left")
        ttk.Label(
            options,
            text="Automatic by default; manual override available.",
            style="Muted.Card.TLabel",
        ).pack(side="left", padx=(12, 0))
        self.metadata_box = ttk.Checkbutton(
            card,
            text="Include advanced filesystem metadata",
            variable=self.full_metadata,
            style="Card.TCheckbutton",
        )
        self.metadata_box.grid(row=5, column=1, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Label(
            card,
            text="An existing report at this path will be replaced after a successful save.",
            style="Muted.Card.TLabel",
        ).grid(row=6, column=0, columnspan=3, sticky="w", pady=(10, 0))
        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        actions.columnconfigure(0, weight=1)
        self.status_label = ttk.Label(
            actions, textvariable=self.status, style="Card.TLabel"
        )
        self.status_label.grid(row=0, column=0, sticky="w", padx=(0, 12))
        actions.bind(
            "<Configure>",
            lambda event: self.status_label.configure(
                wraplength=max(
                    100, event.width - self.start_button.winfo_reqwidth() - 24
                )
            ),
        )
        self.start_button = ttk.Button(
            actions,
            text="Create Excel report",
            width=ACTION_WIDTH,
            style="Action.TButton",
            command=self.start,
        )
        self.start_button.grid(row=0, column=1, sticky="e")
        self.progress = ttk.Progressbar(card, mode="indeterminate")
        self.progress.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self.progress.grid_remove()
        self.log = ScrolledText(
            card,
            height=7,
            width=1,
            wrap="word",
            state="disabled",
            font=("Consolas", 9),
            background=CARD,
            foreground=MUTED,
            relief="flat",
        )
        self.log.grid(row=9, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        footer = ttk.Frame(outer, style="App.TFrame")
        footer.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(16, 0))
        footer.columnconfigure(0, weight=1)
        ttk.Label(footer, text="Last report", style="App.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.report_entry = ttk.Entry(
            footer, textvariable=self.report_name, state="readonly"
        )
        self.report_entry.grid(row=1, column=0, sticky="ew", padx=(0, 12))
        self.open_button = ttk.Button(
            footer,
            text="Open report folder",
            width=ACTION_WIDTH,
            style="Action.TButton",
            command=self.open_folder,
        )
        self.open_button.grid(row=0, column=1, rowspan=2, sticky="e")
        ttk.Label(
            footer,
            text="Local execution • SMB folders supported",
            style="App.TLabel",
            foreground=MUTED,
            font=("Segoe UI", 9),
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))
        self.folder.trace_add("write", self.changed)
        self.output.trace_add("write", self.changed)
        window.bind("<Alt-s>", lambda event: self.start())
        window.bind("<Control-Return>", lambda event: self.start())
        available = max(440, window.winfo_screenheight() - 100)
        window.geometry(f"920x{min(620, available)}")
        window.minsize(760, min(540, available))
        self.changed()
        self.folder_entry.focus_set()
        self.poll_id = window.after(100, self.poll)

    def resize_card(self, event):
        self.canvas.itemconfigure(self.canvas_item, width=event.width)
        self.update_scroll()

    def update_scroll(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        if self.card_border.winfo_reqheight() > self.canvas.winfo_height() + 1:
            self.scrollbar.grid(row=2, column=1, sticky="ns", padx=(8, 0))
        else:
            self.scrollbar.grid_remove()
            self.canvas.yview_moveto(0)

    def scroll_card(self, event):
        if (
            isinstance(event.widget, (ttk.Combobox, tk.Entry, ttk.Entry, tk.Text))
            or not self.scrollbar.winfo_ismapped()
        ):
            return
        units = (
            (-1 if event.num == 4 else 1)
            if event.num in (4, 5)
            else -int(event.delta / 120)
        )
        self.canvas.yview_scroll(units, "units")

    def changed(self, *unused):
        if not self.busy:
            self.status.set(
                "Ready to create a report."
                if self.folder.get().strip()
                else "Choose a folder to begin."
            )
        self.enable()

    def enable(self):
        for entry in (self.folder_entry, self.output_entry):
            entry.configure(state="disabled" if self.busy else "normal")
        self.worker_box.configure(state="disabled" if self.busy else "readonly")
        for button in (self.browse_button, self.save_button, self.metadata_box):
            button.state(["disabled" if self.busy else "!disabled"])
        ready = (
            not self.busy and self.folder.get().strip() and self.output.get().strip()
        )
        self.start_button.state(["!disabled" if ready else "disabled"])
        self.open_button.state(
            ["!disabled" if self.last_report and not self.busy else "disabled"]
        )

    def browse(self):
        if not self.busy:
            chosen = filedialog.askdirectory(
                parent=self.window, title="Choose a folder to inspect", mustexist=True
            )
            if chosen:
                self.folder.set(chosen)

    def save_as(self):
        if not self.busy:
            chosen = filedialog.asksaveasfilename(
                parent=self.window,
                title="Save Excel report",
                defaultextension=".xlsx",
                filetypes=[("Excel workbook", "*.xlsx")],
                initialfile=core.OUTPUT_FILE,
            )
            if chosen:
                self.output.set(chosen)

    def append_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start(self):
        if self.busy or not self.folder.get().strip() or not self.output.get().strip():
            return
        folder, output = self.folder.get(), self.output.get()
        workers, metadata = self.workers.get(), self.full_metadata.get()
        self.busy = True
        self.status.set("Reading folder and preparing the report…")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.progress.grid()
        self.progress.start(12)
        self.enable()

        def work():
            try:
                result = core.review_folder(
                    folder,
                    output,
                    workers,
                    metadata,
                    lambda text: self.messages.put(("progress", text)),
                )
                self.messages.put(
                    ("done", (result, core.absolute_path(core.pasted_path(output))))
                )
            except Exception as exc:  # noqa: BLE001 -- surface worker/start failures in the UI
                self.messages.put(("error", f"{type(exc).__name__}: {exc}"))

        try:
            threading.Thread(target=work, name="folder-review", daemon=False).start()
        except Exception as exc:  # noqa: BLE001 -- surface worker/start failures in the UI
            self.messages.put(("error", f"{type(exc).__name__}: {exc}"))

    def poll(self):
        try:
            while True:
                kind, value = self.messages.get_nowait()
                if kind == "progress":
                    self.status.set(value)
                    self.append_log(value)
                    continue
                self.busy = False
                self.progress.stop()
                self.progress.grid_remove()
                if kind == "done":
                    code, output = value
                    self.last_report = Path(output)
                    self.report_name.set(output)
                    self.status.set(
                        "COMPLETE — report created."
                        if code == 0
                        else "PARTIAL — report created; review scan_status for issues."
                    )
                else:
                    self.status.set("Report was not saved. See the error details.")
                    self.append_log(value)
                    messagebox.showerror("Folder Review", value, parent=self.window)
                self.enable()
        except queue.Empty:
            pass
        self.poll_id = self.window.after(100, self.poll)

    def open_folder(self):
        if self.last_report and not self.busy:
            try:
                if os.name == "nt":
                    os.startfile(str(self.last_report.parent))
                else:
                    subprocess.Popen(
                        ["xdg-open", str(self.last_report.parent)],
                        stdin=subprocess.DEVNULL,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                    )
            except OSError as exc:
                messagebox.showerror("Folder Review", str(exc), parent=self.window)

    def close(self):
        if self.busy:
            self.status.set("Please let this review finish before closing.")
            return
        if self.poll_id is not None:
            self.window.after_cancel(self.poll_id)
        self.window.destroy()


def launch():
    window = tk.Tk()
    Application(window, Path(core.__file__).resolve().parent)
    window.mainloop()
    return 0
