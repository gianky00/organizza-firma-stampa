import tkinter as tk
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
        ref_frame = ttk.LabelFrame(self, text="2. Gestione Lista TCL (Aggiungi/Rimuovi/Rinomina)", padding=15)
        ref_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        ref_frame.columnconfigure(0, weight=1)

        # Descrizione
        ttk.Label(ref_frame, text="Modifica qui i nomi e le chiavi di ricerca. Le modifiche appariranno nella scheda Canoni.", 
                  style="info.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        # Header
        header_f = ttk.Frame(ref_frame)
        header_f.grid(row=1, column=0, sticky="ew")
        ttk.Label(header_f, text="Nome Visualizzato", width=30, font=self.app_config.font_bold).pack(side=tk.LEFT, padx=5)
        ttk.Label(header_f, text="Chiave TCL (Ricerca)", width=20, font=self.app_config.font_bold).pack(side=tk.LEFT, padx=5)

        self.ref_list_container = ttk.Frame(ref_frame)
        self.ref_list_container.grid(row=2, column=0, sticky="nsew")
        
        self._refresh_tcl_list_settings()

        # Add button
        btn_f = ttk.Frame(ref_frame)
        btn_f.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(btn_f, text="➕ Aggiungi Nuovo TCL", command=self._add_tcl_settings).pack(side=tk.LEFT)
        ttk.Button(btn_f, text="🔄 Aggiorna Vista Canoni", command=self._apply_to_fees_tab).pack(side=tk.LEFT, padx=10)

    def _refresh_tcl_list_settings(self):
        for widget in self.ref_list_container.winfo_children():
            widget.destroy()

        for i, ref in enumerate(self.app_config.canoni_tcl_vars):
            row_f = ttk.Frame(self.ref_list_container)
            row_f.pack(fill="x", pady=2)
            
            ttk.Entry(row_f, textvariable=ref["name"], width=30).pack(side=tk.LEFT, padx=5)
            ttk.Entry(row_f, textvariable=ref["tcl"], width=20).pack(side=tk.LEFT, padx=5)
            
            ttk.Button(row_f, text="❌", width=3, command=lambda idx=i: self._remove_tcl_settings(idx)).pack(side=tk.LEFT, padx=5)

    def _add_tcl_settings(self):
        import tkinter as tk
        new_ref = {
            "name": tk.StringVar(value="Nuovo"),
            "tcl": tk.StringVar(value="TCL"),
            "num": tk.StringVar(value=""),
            "print": tk.BooleanVar(value=False),
            "path": tk.StringVar(value="")
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
