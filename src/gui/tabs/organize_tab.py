import tkinter as tk
from tkinter import ttk
import threading
import os
from src.logic.organization import OrganizationProcessor
from src.utils.ui_utils import create_path_entry, select_folder_dialog, open_folder_in_explorer

class OrganizeTab(ttk.Frame):
    """
    GUI for the Organize and Print tab.
    """
    def __init__(self, parent, app_config, logger):
        super().__init__(parent)
        self.app_config = app_config
        self.log_widget = logger
        self.stampa_checkbox_vars = {}

        self._create_widgets()
        self.populate_stampa_list()

        self.processor = OrganizationProcessor(
            self,
            app_config,
            self.setup_progress,
            self.update_progress,
            self.hide_progress
        )

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        desc_text = "Questa sezione analizza i file Excel dalla 'Cartella di origine', legge un codice ODC e li copia in sottocartelle. Poi permette di selezionare le cartelle per la stampa di gruppo."
        desc_label = ttk.Label(main_frame, text=desc_text, wraplength=850, justify=tk.LEFT, style='info.TLabel')
        desc_label.pack(fill=tk.X, pady=(0, 15), anchor='w')

        # --- Organization Frame ---
        org_frame = ttk.LabelFrame(main_frame, text="1. Elabora e Organizza per ODC", padding="15")
        org_frame.pack(fill=tk.X, pady=(0, 10))
        create_path_entry(org_frame, "Cartella di Origine:", self.app_config.organizza_source_dir, 0, readonly=False,
                          browse_command=lambda: select_folder_dialog(self.app_config.organizza_source_dir, "Seleziona cartella schede da organizzare"))
        self.organize_button = ttk.Button(org_frame, text="🚀 Avvia Organizzazione", style='primary.TButton', command=self.start_organization_process)
        self.organize_button.grid(row=1, column=0, columnspan=3, sticky="we", pady=(10, 5))

        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=15, padx=5)

        # --- Printing Frame ---
        print_frame = ttk.LabelFrame(main_frame, text="2. Stampa Schede Organizzate", padding="15")
        print_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        print_controls_frame = ttk.Frame(print_frame)
        print_controls_frame.pack(fill=tk.X, pady=(0, 10))
        self.print_button = ttk.Button(print_controls_frame, text="🖨️ Stampa Selezionate", command=self.start_printing_process)
        self.print_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.refresh_button = ttk.Button(print_controls_frame, text="🔄 Aggiorna Lista", command=self.populate_stampa_list)
        self.refresh_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))
        self.open_folder_button = ttk.Button(print_controls_frame, text="📂 Apri Cartella", command=lambda: open_folder_in_explorer(self.app_config.organizza_dest_dir.get()))
        self.open_folder_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

        # --- Checkbox list for printing ---
        list_container = ttk.Frame(print_frame)
        list_container.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(list_container, borderwidth=0, highlightthickness=0)
        self.stampa_checkbox_frame = ttk.Frame(canvas)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        self.canvas_window = canvas.create_window((0, 0), window=self.stampa_checkbox_frame, anchor="nw")
        self.stampa_checkbox_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(self.canvas_window, width=e.width))

        # --- Progress Bar ---
        self.progress_frame = ttk.Frame(main_frame)
        # We will .pack() this frame dynamically when needed
        self.progress_label = ttk.Label(self.progress_frame, text="Progresso:")
        self.progress_label.pack(side=tk.LEFT, padx=(5, 5))
        self.progressbar = ttk.Progressbar(self.progress_frame, orient='horizontal', mode='determinate')
        self.progressbar.pack(fill=tk.X, expand=True)

    def populate_stampa_list(self):
        for widget in self.stampa_checkbox_frame.winfo_children():
            widget.destroy()
        self.stampa_checkbox_vars.clear()
        dest_path = self.app_config.organizza_dest_dir.get()
        if not os.path.isdir(dest_path):
            return
        try:
            folders = sorted([d for d in os.listdir(dest_path) if os.path.isdir(os.path.join(dest_path, d))])
            for folder_name in folders:
                var = tk.IntVar()
                cb = ttk.Checkbutton(self.stampa_checkbox_frame, text=folder_name, variable=var)
                cb.pack(anchor="w", padx=5, fill='x')
                self.stampa_checkbox_vars[folder_name] = {"var": var, "path": os.path.join(dest_path, folder_name)}
        except Exception as e:
            self.log_organizza(f"Errore durante la lettura delle cartelle organizzate: {e}", "ERROR")

    def start_organization_process(self):
        self.toggle_organizza_buttons('disabled')
        threading.Thread(target=self.processor.run_organization_process, daemon=True).start()

    def start_printing_process(self):
        selected_folders = [d["path"] for d in self.stampa_checkbox_vars.values() if d["var"].get() == 1]
        self.toggle_organizza_buttons('disabled')
        threading.Thread(target=self.processor.run_printing_process, args=(selected_folders,), daemon=True).start()

    def toggle_organizza_buttons(self, state):
        self.organize_button.config(state=state)
        self.print_button.config(state=state)
        self.refresh_button.config(state=state)

    def log_organizza(self, message, level='INFO'):
        self.master.after(0, self.log_widget, message, level)

    def setup_progress(self, max_value, label_text="Progresso:"):
        self.progress_label['text'] = label_text
        self.progress_frame.pack(fill=tk.X, pady=(10, 5), after=self.print_frame)
        self.progressbar['maximum'] = max_value
        self.progressbar['value'] = 0

    def update_progress(self, value):
        self.progressbar['value'] = value

    def hide_progress(self):
        self.progress_frame.pack_forget()
