import os
import time
import tkinter as tk
from contextlib import suppress
from tkinter import filedialog, ttk


def create_path_entry(
    parent, label_text, variable, browse_command=None, row=0, column_span=1, readonly=False, button_text="Sfoglia"
):
    """
    Creates a standardized path entry widget with an optional 'Sfoglia' button.
    """
    frame = ttk.Frame(parent)
    frame.grid(row=row, column=0, columnspan=column_span, sticky="ew", pady=5)
    frame.columnconfigure(1, weight=1)

    state = "readonly" if readonly else "normal"
    ttk.Label(frame, text=label_text, width=25).grid(row=0, column=0, sticky="w", padx=(0, 5))
    ttk.Entry(frame, textvariable=variable, state=state).grid(row=0, column=1, sticky="ew", padx=5)

    if browse_command is not None:
        ttk.Button(frame, text=button_text, command=browse_command, width=10).grid(
            row=0, column=2, sticky="e", padx=(5, 0)
        )


class ProgressWithETA(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        # Non configuriamo pesi qui, lo farà il chiamante
        
        self.progress_label = ttk.Label(self, text="Progresso:", width=20, anchor="e")
        self.progress_label.pack(side=tk.LEFT, padx=5)

        self.progressbar = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=150)
        self.progressbar.pack(side=tk.LEFT, padx=5)

        self.percent_label = ttk.Label(self, text="0%", width=5)
        self.percent_label.pack(side=tk.LEFT, padx=2)

        self.eta_label = ttk.Label(self, text="ETA: --:--", width=12)
        self.eta_label.pack(side=tk.LEFT, padx=5)

        self.start_time = 0.0
        self.max_value = 0

    def setup(self, max_value, label_text="Progresso:"):
        self.progress_label["text"] = label_text
        self.max_value = max_value
        self.progressbar["maximum"] = max_value
        self.progressbar["value"] = 0
        self.percent_label["text"] = "0%"
        self.eta_label["text"] = "ETA: --:--"
        self.start_time = time.time()

    def update_progress(self, value):
        self.progressbar["value"] = value
        with suppress(ValueError, TypeError):
            max_val = float(self.progressbar["maximum"])
            if max_val > 0:
                percent = (value / max_val) * 100
                self.percent_label["text"] = f"{percent:.0f}%"

                # Calcolo ETA
                elapsed = time.time() - self.start_time
                if value > 0 and elapsed > 1:  # Calcola dopo 1 sec per stabilità
                    rate = elapsed / value
                    remaining = (max_val - value) * rate
                    mins, secs = divmod(int(remaining), 60)
                    hrs, mins = divmod(mins, 60)
                    if hrs > 0:
                        self.eta_label["text"] = f"ETA: {hrs:02d}:{mins:02d}:{secs:02d}"
                    else:
                        self.eta_label["text"] = f"ETA: {mins:02d}:{secs:02d}"

    def setup_indeterminate(self, label_text="Elaborazione in corso..."):
        self.progress_label["text"] = label_text
        self.progressbar.config(mode="indeterminate")
        self.percent_label["text"] = ""
        self.eta_label["text"] = ""
        self.progressbar.start(10)

    def stop_indeterminate(self):
        self.progressbar.stop()
        self.progressbar.config(mode="determinate")


def select_file_dialog(variable, file_types=(("Tutti i file", "*.*"),)):
    """
    Opens a file selection dialog and updates the provided variable.
    """
    path = filedialog.askopenfilename(filetypes=file_types)
    if path:
        variable.set(os.path.normpath(path))


def select_folder_dialog(variable):
    """
    Opens a folder selection dialog and updates the provided variable.
    """
    path = filedialog.askdirectory()
    if path:
        variable.set(os.path.normpath(path))


def create_log_widget(parent):
    """
    Creates a ScrolledText-like widget using standard Tkinter Text and Scrollbar
    to display logs with colors.
    """
    frame = ttk.Frame(parent)
    frame.pack(fill=tk.BOTH, expand=True)

    scrollbar = ttk.Scrollbar(frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    text_widget = tk.Text(
        frame,
        height=10,
        wrap=tk.WORD,
        yscrollcommand=scrollbar.set,
        font=("Consolas", 10),
        bg="#f8f9fa",
        state=tk.DISABLED,
    )
    text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=text_widget.yview)

    # Configure tags for log levels
    text_widget.tag_configure("INFO", foreground="black")
    text_widget.tag_configure("SUCCESS", foreground="green", font=("Consolas", 10, "bold"))
    text_widget.tag_configure("WARNING", foreground="orange", font=("Consolas", 10, "bold"))
    text_widget.tag_configure("ERROR", foreground="red", font=("Consolas", 10, "bold"))
    text_widget.tag_configure("HEADER", foreground="blue", font=("Consolas", 10, "bold"))

    return text_widget


def log_message(text_widget, message, level="INFO"):
    """
    Appends a message to the log widget with the appropriate level color.
    """
    if not text_widget or not text_widget.winfo_exists():
        return

    text_widget.config(state=tk.NORMAL)
    text_widget.insert(tk.END, f"[{level}] {message}\n", level)
    text_widget.see(tk.END)
    text_widget.config(state=tk.DISABLED)


def open_folder_in_explorer(path_to_open):
    """
    Opens the specified path in Windows File Explorer.
    """
    if not path_to_open:
        return

    if not os.path.isdir(path_to_open):
        return
    with suppress(Exception):
        # os.startfile is Windows-specific, which is appropriate for this app
        os.startfile(path_to_open)
