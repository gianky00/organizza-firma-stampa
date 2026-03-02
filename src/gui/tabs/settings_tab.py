import tkinter as tk
from functools import partial
from tkinter import messagebox, ttk

from src.utils.ui_utils import create_path_entry, select_folder_dialog


class SettingsTab(ttk.Frame):
    def __init__(self, parent, app_config):
        super().__init__(parent)
        self.app_config = app_config
        self._create_widgets()

    def _create_widgets(self):
        self.columnconfigure(0, weight=1)

        # --- Description ---
        desc_label = ttk.Label(
            self,
            text="Personalizza i percorsi di rete e i referenti per i canoni mensili. Le modifiche ai referenti appariranno nella scheda Canoni dopo il riavvio o l'aggiornamento.",
            wraplength=800,
            justify=tk.LEFT,
            style="info.TLabel",
        )
        desc_label.pack(fill=tk.X, pady=(0, 15), anchor="w")

        # --- Paths Settings Frame ---
        paths_frame = ttk.LabelFrame(self, text="1. Percorsi di Rete e Cartelle", padding=15)
        paths_frame.pack(fill=tk.X, pady=5)
        paths_frame.columnconfigure(0, weight=1)

        create_path_entry(
            paths_frame,
            "Cartella Giornaliere (Rete):",
            self.app_config.canoni_giornaliera_base_dir,
            lambda: select_folder_dialog(self.app_config.canoni_giornaliera_base_dir),
            0,
            readonly=False,
        )

        create_path_entry(
            paths_frame,
            "Cartella Consuntivi (Rete):",
            self.app_config.canoni_consuntivi_base_dir,
            lambda: select_folder_dialog(self.app_config.canoni_consuntivi_base_dir),
            1,
            readonly=False,
        )

        create_path_entry(
            paths_frame,
            "Archivio ODC (Rete):",
            self.app_config.organizza_base_dir,
            lambda: select_folder_dialog(self.app_config.organizza_base_dir),
            2,
            readonly=False,
        )

        # --- Referenti Management Frame ---
        ref_frame = ttk.LabelFrame(self, text="2. Gestione Lista TCL (Canoni)", padding=15)
        ref_frame.pack(fill=tk.X, pady=10)
        ref_frame.columnconfigure(0, weight=1)

        # Header TCL
        header_f = ttk.Frame(ref_frame)
        header_f.pack(fill="x")
        ttk.Label(header_f, text="Nome Visualizzato", width=30, font=self.app_config.font_bold).pack(side=tk.LEFT, padx=5)
        ttk.Label(header_f, text="Chiave TCL (Ricerca)", width=20, font=self.app_config.font_bold).pack(side=tk.LEFT, padx=5)

        self.ref_list_container = ttk.Frame(ref_frame)
        self.ref_list_container.pack(fill="x")

        self._refresh_tcl_list_settings()

        btn_f = ttk.Frame(ref_frame)
        btn_f.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_f, text="➕ Aggiungi TCL", command=self._add_tcl_settings).pack(side=tk.LEFT)
        ttk.Button(btn_f, text="🔄 Aggiorna Vista", command=self._apply_to_fees_tab).pack(side=tk.LEFT, padx=10)

        # --- Models Management Frame ---
        models_frame = ttk.LabelFrame(self, text="3. Configurazione Modelli Schede (Ridenominazione e TCL)", padding=15)
        models_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Area scorrevole per i modelli
        m_canvas = tk.Canvas(models_frame, borderwidth=0, highlightthickness=0, height=300)
        m_scrollbar = ttk.Scrollbar(models_frame, orient="vertical", command=m_canvas.yview)
        self.models_container = ttk.Frame(m_canvas)
        
        self.models_window = m_canvas.create_window((0, 0), window=self.models_container, anchor="nw")
        m_canvas.configure(yscrollcommand=m_scrollbar.set)
        
        self.models_container.bind("<Configure>", lambda e: m_canvas.configure(scrollregion=m_canvas.bbox("all")))
        m_canvas.bind("<Configure>", lambda e: m_canvas.itemconfig(self.models_window, width=e.width))

        m_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        m_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._refresh_models_list_settings()

        m_btn_f = ttk.Frame(models_frame)
        m_btn_f.pack(fill=tk.X, side=tk.BOTTOM, pady=(5, 0))
        ttk.Button(m_btn_f, text="➕ Aggiungi Nuovo Modello", command=self._add_model_settings).pack(side=tk.LEFT)

    def _refresh_models_list_settings(self):
        for widget in self.models_container.winfo_children():
            widget.destroy()

        # Headers
        h_f = ttk.Frame(self.models_container)
        h_f.pack(fill="x", pady=(0, 5))
        headers = [("Nome Modello", 25), ("Cella ID", 8), ("Match ID", 20), ("Cella TCL", 10), ("Celle Data (sep. virgola)", 25)]
        for text, w in headers:
            ttk.Label(h_f, text=text, width=w, font=self.app_config.font_bold).pack(side=tk.LEFT, padx=5)

        for i, mod in enumerate(self.app_config.rename_models_vars):
            row_f = ttk.Frame(self.models_container)
            row_f.pack(fill="x", pady=2)

            ttk.Entry(row_f, textvariable=mod["name"], width=25).pack(side=tk.LEFT, padx=5)
            ttk.Entry(row_f, textvariable=mod["id_cell"], width=8).pack(side=tk.LEFT, padx=5)
            ttk.Entry(row_f, textvariable=mod["match_value"], width=20).pack(side=tk.LEFT, padx=5)
            ttk.Entry(row_f, textvariable=mod["tcl_cell"], width=10).pack(side=tk.LEFT, padx=5)
            ttk.Entry(row_f, textvariable=mod["date_cells"], width=25).pack(side=tk.LEFT, padx=5)

            ttk.Button(row_f, text="❌", width=3, command=partial(self._remove_model_settings, i)).pack(side=tk.LEFT, padx=5)

    def _add_model_settings(self):
        new_mod = {
            "name": tk.StringVar(value="Nuovo Modello"),
            "id_cell": tk.StringVar(value="E2"),
            "match_value": tk.StringVar(value="testo"),
            "tcl_cell": tk.StringVar(value="L45"),
            "date_cells": tk.StringVar(value="B50"),
        }
        self.app_config.rename_models_vars.append(new_mod)
        self._refresh_models_list_settings()

    def _remove_model_settings(self, index):
        if len(self.app_config.rename_models_vars) <= 1: return
        self.app_config.rename_models_vars.pop(index)
        self._refresh_models_list_settings()

    def _refresh_tcl_list_settings(self):
        for widget in self.ref_list_container.winfo_children():
            widget.destroy()

        for i, ref in enumerate(self.app_config.canoni_tcl_vars):
            row_f = ttk.Frame(self.ref_list_container)
            row_f.pack(fill="x", pady=2)

            ttk.Entry(row_f, textvariable=ref["name"], width=30).pack(side=tk.LEFT, padx=5)
            ttk.Entry(row_f, textvariable=ref["tcl"], width=20).pack(side=tk.LEFT, padx=5)

            ttk.Button(row_f, text="❌", width=3, command=partial(self._remove_tcl_settings, i)).pack(
                side=tk.LEFT, padx=5
            )

    def _add_tcl_settings(self):
        import tkinter as tk

        new_ref = {
            "name": tk.StringVar(value="Nuovo"),
            "tcl": tk.StringVar(value="TCL"),
            "num": tk.StringVar(value=""),
            "print": tk.BooleanVar(value=False),
            "path": tk.StringVar(value=""),
        }
        self.app_config.canoni_tcl_vars.append(new_ref)
        self._refresh_tcl_list_settings()

    def _remove_tcl_settings(self, index):
        if len(self.app_config.canoni_tcl_vars) <= 1:
            messagebox.showwarning("Attenzione", "Deve esserci almeno un TCL in lista.")
            return
        self.app_config.canoni_tcl_vars.pop(index)
        self._refresh_tcl_list_settings()

    def _apply_to_fees_tab(self):
        if hasattr(self.app_config, "fees_tab"):
            self.app_config.fees_tab._refresh_dynamic_tcl_ui()
            self.app_config.fees_tab._setup_tcl_traces()
            messagebox.showinfo("Successo", "Interfaccia Canoni aggiornata.")
