import os
import tkinter as tk
from contextlib import suppress
from tkinter import filedialog, ttk


def create_path_entry(parent, label_text, variable, browse_command, row, column_span=1, readonly=False):
    """
    Creates a standardized path entry widget with a 'Sfoglia' button.
    """
    frame = ttk.Frame(parent)
    frame.grid(row=row, column=0, columnspan=column_span, sticky="ew", pady=5)
    frame.columnconfigure(1, weight=1)

    state = "readonly" if readonly else "normal"
    ttk.Label(frame, text=label_text, width=20).grid(row=0, column=0, sticky="w")
    ttk.Entry(frame, textvariable=variable, state=state).grid(row=0, column=1, sticky="ew", padx=5)
    ttk.Button(frame, text="Sfoglia", command=browse_command).grid(row=0, column=2, sticky="e")


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
