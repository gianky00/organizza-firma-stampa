from functools import partial

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from src.utils.qt_vars import BooleanVar, StringVar
from src.utils.ui_utils import create_path_entry, select_folder_dialog


class SettingsTab(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.app_config = parent  # MainApplication is passed as parent
        self._create_widgets()

    def _create_widgets(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Create a scroll area for the whole settings tab
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)

        self.scrollable_widget = QWidget()
        self.scrollable_layout = QVBoxLayout(self.scrollable_widget)
        self.scrollable_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.scroll_area.setWidget(self.scrollable_widget)

        main_layout.addWidget(self.scroll_area)

        container = self.scrollable_layout

        # --- Description ---
        desc_label = QLabel(
            "Personalizza i percorsi di rete e i TCL per i canoni mensili. Le modifiche ai TCL appariranno nella scheda Canoni dopo il riavvio o l'aggiornamento."
        )
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #333333;")
        container.addWidget(desc_label)

        # --- Paths Settings Frame ---
        paths_frame = QGroupBox("1. Percorsi di Rete e Cartelle")
        paths_layout = QVBoxLayout(paths_frame)
        container.addWidget(paths_frame)

        p1 = create_path_entry(
            paths_frame,
            "Cartella Giornaliere (Rete):",
            self.app_config.canoni_giornaliera_base_dir,
            lambda: select_folder_dialog(self.app_config.canoni_giornaliera_base_dir, self),
            readonly=False,
        )
        paths_layout.addLayout(p1)

        p2 = create_path_entry(
            paths_frame,
            "Cartella Consuntivi (Rete):",
            self.app_config.canoni_consuntivi_base_dir,
            lambda: select_folder_dialog(self.app_config.canoni_consuntivi_base_dir, self),
            readonly=False,
        )
        paths_layout.addLayout(p2)

        p3 = create_path_entry(
            paths_frame,
            "Archivio ODC (Rete):",
            self.app_config.organizza_base_dir,
            lambda: select_folder_dialog(self.app_config.organizza_base_dir, self),
            readonly=False,
        )
        paths_layout.addLayout(p3)

        # --- TCL Management Frame ---
        ref_frame = QGroupBox("2. Gestione Lista TCL (Canoni)")
        ref_layout = QVBoxLayout(ref_frame)
        container.addWidget(ref_frame)

        # Header TCL
        header_layout = QHBoxLayout()
        lbl1 = QLabel("Nome Visualizzato")
        lbl1.setStyleSheet("font-weight: bold;")
        lbl1.setMinimumWidth(200)
        header_layout.addWidget(lbl1)

        lbl2 = QLabel("Chiave TCL (Ricerca)")
        lbl2.setStyleSheet("font-weight: bold;")
        lbl2.setMinimumWidth(150)
        header_layout.addWidget(lbl2)

        header_layout.addStretch()
        ref_layout.addLayout(header_layout)

        self.ref_list_container = QWidget()
        self.ref_list_layout = QVBoxLayout(self.ref_list_container)
        self.ref_list_layout.setContentsMargins(0, 0, 0, 0)
        ref_layout.addWidget(self.ref_list_container)

        self._refresh_tcl_list_settings()

        btn_layout = QHBoxLayout()
        add_tcl_btn = QPushButton("+ Aggiungi TCL")
        add_tcl_btn.clicked.connect(self._add_tcl_settings)
        btn_layout.addWidget(add_tcl_btn)

        update_btn = QPushButton("🔄 Aggiorna Vista")
        update_btn.clicked.connect(self._apply_to_fees_tab)
        btn_layout.addWidget(update_btn)

        btn_layout.addStretch()
        ref_layout.addLayout(btn_layout)

        # --- Models Management Frame ---
        models_frame = QGroupBox("3. Configurazione Modelli Schede (Ridenominazione, TCL e Stampa)")
        models_layout = QVBoxLayout(models_frame)
        container.addWidget(models_frame)

        self.models_container = QWidget()
        self.models_layout = QGridLayout(self.models_container)
        self.models_layout.setContentsMargins(0, 0, 0, 0)
        models_layout.addWidget(self.models_container)

        self._refresh_models_list_settings()

        m_btn_layout = QHBoxLayout()
        add_model_btn = QPushButton("+ Aggiungi Nuovo Modello")
        add_model_btn.clicked.connect(self._add_model_settings)
        m_btn_layout.addWidget(add_model_btn)
        m_btn_layout.addStretch()
        models_layout.addLayout(m_btn_layout)

    def _refresh_models_list_settings(self):
        for i in reversed(range(self.models_layout.count())):
            item = self.models_layout.itemAt(i)
            if item is None:
                continue
            widget_to_remove = item.widget()
            if widget_to_remove is not None:
                widget_to_remove.setParent(None)  # type: ignore
                widget_to_remove.deleteLater()  # type: ignore

        # Configurazione pesi colonne per la griglia
        self.models_layout.setColumnStretch(0, 3)  # Nome
        self.models_layout.setColumnStretch(1, 3)  # Match
        self.models_layout.setColumnStretch(2, 1)  # Celle ID
        self.models_layout.setColumnStretch(3, 1)  # Area Stampa
        self.models_layout.setColumnStretch(4, 1)  # Cella TCL
        self.models_layout.setColumnStretch(5, 2)  # Celle Data
        self.models_layout.setColumnStretch(6, 0)  # X

        headers = [
            ("Nome Modello", 0),
            ("Match ID", 1),
            ("Celle ID", 2),
            ("Area Stampa", 3),
            ("Cella TCL", 4),
            ("Celle Data", 5),
        ]

        for text, col in headers:
            lbl = QLabel(text)
            lbl.setStyleSheet("font-weight: bold;")
            self.models_layout.addWidget(lbl, 0, col)

        for i, mod in enumerate(self.app_config.rename_models_vars):
            row_idx = i + 1

            w0 = QLineEdit()
            w0.setText(mod["name"].get())
            w0.textChanged.connect(mod["name"].set)
            self.models_layout.addWidget(w0, row_idx, 0)

            w1 = QLineEdit()
            w1.setText(mod["match_value"].get())
            w1.textChanged.connect(mod["match_value"].set)
            self.models_layout.addWidget(w1, row_idx, 1)

            w2 = QLineEdit()
            w2.setText(mod["id_cells"].get())
            w2.textChanged.connect(mod["id_cells"].set)
            self.models_layout.addWidget(w2, row_idx, 2)

            w3 = QLineEdit()
            w3.setText(mod["print_area"].get())
            w3.textChanged.connect(mod["print_area"].set)
            self.models_layout.addWidget(w3, row_idx, 3)

            w4 = QLineEdit()
            w4.setText(mod["tcl_cell"].get())
            w4.textChanged.connect(mod["tcl_cell"].set)
            self.models_layout.addWidget(w4, row_idx, 4)

            w5 = QLineEdit()
            w5.setText(mod["date_cells"].get())
            w5.textChanged.connect(mod["date_cells"].set)
            self.models_layout.addWidget(w5, row_idx, 5)

            btn = QPushButton("❌")
            btn.setFixedWidth(30)
            btn.clicked.connect(partial(self._remove_model_settings, i))
            self.models_layout.addWidget(btn, row_idx, 6)

    def _add_model_settings(self):
        new_mod = {
            "name": StringVar(value="Nuovo Modello", parent=self.app_config),
            "match_value": StringVar(value="testo", parent=self.app_config),
            "id_cells": StringVar(value="E2, T2", parent=self.app_config),
            "print_area": StringVar(value="A1:N50", parent=self.app_config),
            "tcl_cell": StringVar(value="L45", parent=self.app_config),
            "date_cells": StringVar(value="B50", parent=self.app_config),
        }
        self.app_config.rename_models_vars.append(new_mod)
        self._refresh_models_list_settings()

    def _remove_model_settings(self, index):
        if len(self.app_config.rename_models_vars) <= 1:
            return
        self.app_config.rename_models_vars.pop(index)
        self._refresh_models_list_settings()

    def _refresh_tcl_list_settings(self):
        for i in reversed(range(self.ref_list_layout.count())):
            item = self.ref_list_layout.itemAt(i)
            if item is None:
                continue
            if item.widget() is not None:
                item.widget().setParent(None)  # type: ignore
                item.widget().deleteLater()  # type: ignore
            elif item.layout() is not None:
                layout = item.layout()
                while layout.count():
                    child = layout.takeAt(0)
                    if child is not None and child.widget() is not None:
                        child.widget().deleteLater()  # type: ignore
                layout.deleteLater()

        for i, ref in enumerate(self.app_config.canoni_tcl_vars):
            row_layout = QHBoxLayout()
            row_layout.setContentsMargins(0, 0, 0, 0)

            w1 = QLineEdit()
            w1.setMinimumWidth(200)
            w1.setText(ref["name"].get())
            w1.textChanged.connect(ref["name"].set)
            row_layout.addWidget(w1)

            w2 = QLineEdit()
            w2.setMinimumWidth(150)
            w2.setText(ref["tcl"].get())
            w2.textChanged.connect(ref["tcl"].set)
            row_layout.addWidget(w2)

            btn = QPushButton("❌")
            btn.setFixedWidth(30)
            btn.clicked.connect(partial(self._remove_tcl_settings, i))
            row_layout.addWidget(btn)

            row_layout.addStretch()
            self.ref_list_layout.addLayout(row_layout)

    def _add_tcl_settings(self):
        new_ref = {
            "name": StringVar(value="Nuovo", parent=self.app_config),
            "tcl": StringVar(value="TCL", parent=self.app_config),
            "num": StringVar(value="", parent=self.app_config),
            "print": BooleanVar(value=False, parent=self.app_config),
            "path": StringVar(value="", parent=self.app_config),
        }
        self.app_config.canoni_tcl_vars.append(new_ref)
        self._refresh_tcl_list_settings()

    def _remove_tcl_settings(self, index):
        if len(self.app_config.canoni_tcl_vars) <= 1:
            QMessageBox.warning(self, "Attenzione", "Deve esserci almeno un TCL in lista.")
            return
        self.app_config.canoni_tcl_vars.pop(index)
        self._refresh_tcl_list_settings()

    def _apply_to_fees_tab(self):
        if hasattr(self.app_config, "fees_tab"):
            self.app_config.fees_tab._refresh_dynamic_tcl_ui()
            self.app_config.fees_tab._setup_tcl_traces()
            QMessageBox.information(self, "Successo", "Interfaccia Canoni aggiornata.")
