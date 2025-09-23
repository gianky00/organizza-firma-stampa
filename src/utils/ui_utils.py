import tkinter as tk
from tkinter import scrolledtext, filedialog

def create_log_widget(parent):
    """
    Creates and configures a ScrolledText widget for logging.
    """
    log_widget = scrolledtext.ScrolledText(parent, wrap=tk.WORD, height=10, font=("Courier New", 9))
    log_widget.pack(fill=tk.BOTH, expand=True)
    log_widget.config(state='disabled')

    # Configure tags for different log levels
    tags = {
        'INFO': 'black',
        'SUCCESS': '#008744',
        'WARNING': '#ffa700',
        'ERROR': '#d62d20',
        'HEADER': '#0057e7'
    }
    for tag, color in tags.items():
        fw = "bold" if tag in ['SUCCESS', 'ERROR', 'HEADER'] else "normal"
        log_widget.tag_config(tag.upper(), foreground=color, font=("Courier New", 9, fw))

    return log_widget

def log_message(widget, message, level='INFO'):
    """
    Inserts a message into the log widget with the appropriate tag.
    Must be called from the main GUI thread.
    """
    if not widget:
        return
    widget.config(state='normal')
    widget.insert(tk.END, f"> {message}\n", level.upper())
    widget.config(state='disabled')
    widget.see(tk.END)

def clear_log(widget):
    """Clears all text from the log widget."""
    if not widget:
        return
    widget.config(state='normal')
    widget.delete(1.0, tk.END)
    widget.config(state='disabled')

def create_path_entry(parent, label_text, string_var, row, readonly=False, browse_command=None):
    """
    Creates a standardized row with a label, entry, and optional browse button.
    """
    from tkinter import ttk

    ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)

    entry_width = 60 if browse_command or readonly else 70
    entry = ttk.Entry(parent, textvariable=string_var, width=entry_width)
    if readonly:
        entry.config(state='readonly')
    entry.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

    if browse_command:
        button = ttk.Button(parent, text="Sfoglia...", command=browse_command, width=10)
        button.grid(row=row, column=2, sticky=tk.E, padx=5, pady=3)
        parent.columnconfigure(1, weight=1)
    else:
        # Ensure the entry column still expands
        parent.grid_columnconfigure(1, weight=1)

def select_folder_dialog(string_var, title):
    """Opens a dialog to select a folder and sets the string_var."""
    folder_selected = filedialog.askdirectory(title=title)
    if folder_selected:
        string_var.set(folder_selected)

def select_file_dialog(string_var, title, filetypes, initialdir=None):
    """Opens a dialog to select a file and sets the string_var."""
    file_selected = filedialog.askopenfilename(title=title, filetypes=filetypes, initialdir=initialdir)
    if file_selected:
        string_var.set(file_selected)
