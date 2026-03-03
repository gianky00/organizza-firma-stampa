import os
import time
from contextlib import suppress

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QTextCharFormat, QTextCursor
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QWidget,
)


def create_path_entry(parent, label_text, variable, browse_command=None, readonly=False, button_text="Sfoglia"):
    """
    Creates a standardized path entry widget with an optional 'Sfoglia' button.
    Returns a layout containing the widgets, which can be added to a parent layout.
    """
    layout = QHBoxLayout()
    layout.setContentsMargins(0, 0, 0, 0)

    label = QLabel(label_text, parent)
    label.setMinimumWidth(150)
    layout.addWidget(label)

    entry = QLineEdit(parent)
    entry.setReadOnly(readonly)
    # Sync with variable
    entry.setText(variable.get())
    variable.value_changed.connect(entry.setText)

    # We only want to update the variable if it's not readonly
    if not readonly:
        entry.textChanged.connect(variable.set)

    layout.addWidget(entry, 1)  # stretch=1

    if browse_command is not None:
        btn = QPushButton(button_text, parent)
        btn.clicked.connect(browse_command)
        btn.setFixedWidth(80)
        layout.addWidget(btn)

    return layout


class ProgressWithETA(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_layout = QHBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)

        self.progress_label = QLabel("Progresso:", self)
        self.progress_label.setMinimumWidth(100)
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.main_layout.addWidget(self.progress_label)

        self.progressbar = QProgressBar(self)
        self.progressbar.setMinimumWidth(200)
        self.progressbar.setTextVisible(False)
        self.main_layout.addWidget(self.progressbar)

        self.percent_label = QLabel("0%", self)
        self.percent_label.setMinimumWidth(40)
        self.main_layout.addWidget(self.percent_label)

        self.eta_label = QLabel("ETA: --:--", self)
        self.eta_label.setMinimumWidth(80)
        self.main_layout.addWidget(self.eta_label)

        self.start_time = 0.0
        self.max_value = 0

    def setup(self, max_value, label_text="Progresso:"):
        self.progress_label.setText(label_text)
        self.max_value = max_value
        self.progressbar.setRange(0, int(max_value))
        self.progressbar.setValue(0)
        self.percent_label.setText("0%")
        self.eta_label.setText("ETA: --:--")
        self.start_time = time.time()

    def update_progress(self, value):
        self.progressbar.setValue(int(value))
        if self.max_value > 0:
            percent = (value / self.max_value) * 100
            self.percent_label.setText(f"{percent:.0f}%")

            elapsed = time.time() - self.start_time
            if value > 0 and elapsed > 1:
                rate = elapsed / value
                remaining = (self.max_value - value) * rate
                mins, secs = divmod(int(remaining), 60)
                hrs, mins = divmod(mins, 60)
                if hrs > 0:
                    self.eta_label.setText(f"ETA: {hrs:02d}:{mins:02d}:{secs:02d}")
                else:
                    self.eta_label.setText(f"ETA: {mins:02d}:{secs:02d}")

    def setup_indeterminate(self, label_text="Elaborazione in corso..."):
        self.progress_label.setText(label_text)
        self.progressbar.setRange(0, 0)
        self.percent_label.setText("")
        self.eta_label.setText("")

    def stop_indeterminate(self):
        self.progressbar.setRange(0, 100)


def select_file_dialog(variable, file_types="Tutti i file (*.*)", parent=None):
    """
    Opens a file selection dialog and updates the provided variable.
    """
    path, _ = QFileDialog.getOpenFileName(parent, "Seleziona file", "", file_types)
    if path:
        variable.set(os.path.normpath(path))


def select_folder_dialog(variable, parent=None):
    """
    Opens a folder selection dialog and updates the provided variable.
    """
    path = QFileDialog.getExistingDirectory(parent, "Seleziona cartella")
    if path:
        variable.set(os.path.normpath(path))


def create_log_widget(parent):
    """
    Creates a QTextEdit widget to display logs with colors.
    """
    text_widget = QTextEdit(parent)
    text_widget.setReadOnly(True)
    font = QFont("Consolas", 10)
    text_widget.setFont(font)
    text_widget.setStyleSheet("background-color: #f8f9fa;")
    return text_widget


def log_message(text_widget, message, level="INFO"):
    """
    Appends a message to the log widget with the appropriate level color.
    """
    if not text_widget:
        return

    cursor = text_widget.textCursor()
    cursor.movePosition(QTextCursor.MoveOperation.End)

    format = QTextCharFormat()
    font = QFont("Consolas", 10)

    if level == "SUCCESS":
        format.setForeground(QColor("green"))
        font.setBold(True)
    elif level == "WARNING":
        format.setForeground(QColor("orange"))
        font.setBold(True)
    elif level == "ERROR":
        format.setForeground(QColor("red"))
        font.setBold(True)
    elif level == "HEADER":
        format.setForeground(QColor("blue"))
        font.setBold(True)
    else:  # INFO
        format.setForeground(QColor("black"))

    format.setFont(font)
    cursor.insertText(f"[{level}] {message}\n", format)
    text_widget.setTextCursor(cursor)
    text_widget.ensureCursorVisible()


def open_folder_in_explorer(path_to_open):
    """
    Opens the specified path in Windows File Explorer.
    """
    if not path_to_open:
        return

    if not os.path.isdir(path_to_open):
        return
    with suppress(Exception):
        os.startfile(path_to_open)
