import tkinter as tk
from tkinter import ttk
import threading
import os
from src.logic.organization import OrganizationProcessor
from src.utils.ui_utils import create_path_entry, select_folder_dialog, open_folder_in_explorer

class OrganizeTab(ttk.Frame):
    def __init__(self, parent, app_config, logger):
        super().__init__(parent)
        self.app_config = app_config
        self.log_widget = logger
        self.stampa_checkbox_vars = {}
        self.cancel_event = threading.Event()
        self.active_process_type = None

        self._create_widgets()
        self.processor = OrganizationProcessor(self, app_config, self.setup_progress, self.update_progress, self.hide_progress)
        self.after(100, self.populate_stampa_list)

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        desc_label = ttk.Label(main_frame, text="Questa sezione analizza i file Excel dalla 'Cartella di origine', legge un codice ODC e li copia in sottocartelle. Poi permette di selezionare le cartelle per la stampa di gruppo.", wraplength=850, justify=tk.LEFT, style='info.TLabel')
        desc_label.pack(fill=tk.X, pady=(0, 15), anchor='w')

        self.org_frame = ttk.LabelFrame(main_frame, text="1. Elabora e Organizza per ODC", padding="15")
        self.org_frame.pack(fill=tk.X, pady=(0, 10))
        create_path_entry(self.org_frame, "Cartella di Origine:", self.app_config.organizza_source_dir, 0, readonly=False, browse_command=lambda: select_folder_dialog(self.app_config.organizza_source_dir, "Seleziona cartella schede da organizzare"))

        self.organize_button = ttk.Button(self.org_frame, text="🚀 Avvia Organizzazione", style='primary.TButton', command=self.start_organization_process)
        self.organize_button.grid(row=1, column=0, columnspan=3, sticky="we", pady=(10, 5))

        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=15, padx=5)

        self.print_frame = ttk.LabelFrame(main_frame, text="2. Stampa Schede Organizzate", padding="15")
        self.print_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.print_controls_frame = ttk.Frame(self.print_frame)
        self.print_controls_frame.pack(fill=tk.X, pady=(0, 10))
        self.print_button = ttk.Button(self.print_controls_frame, text="🖨️ Stampa Selezionate", command=self.start_printing_process)
        self.print_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.refresh_button = ttk.Button(self.print_controls_frame, text="🔄 Aggiorna Lista", command=self.populate_stampa_list)
        self.refresh_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))
        self.open_folder_button = ttk.Button(self.print_controls_frame, text="📂 Apri Cartella", command=lambda: open_folder_in_explorer(self.app_config.organizza_dest_dir.get()))
        self.open_folder_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

        self.cancel_org_button = ttk.Button(self.org_frame, text="Annulla Organizzazione", command=self.cancel_process)
        self.cancel_print_button = ttk.Button(self.print_controls_frame, text="Annulla Stampa", command=self.cancel_process)

        list_container = ttk.Frame(self.print_frame)
        list_container.pack(fill=tk.BOTH, expand=False)
        canvas = tk.Canvas(list_container, borderwidth=0, highlightthickness=0, height=300)
        self.stampa_checkbox_frame = ttk.Frame(canvas)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        self.canvas_window = canvas.create_window((0, 0), window=self.stampa_checkbox_frame, anchor="nw")
        self.stampa_checkbox_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(self.canvas_window, width=e.width))

        self.progress_frame = ttk.Frame(main_frame)
        self.progress_label = ttk.Label(self.progress_frame, text="Progresso:")
        self.progress_label.pack(side=tk.LEFT, padx=(5, 5))
        self.progressbar = ttk.Progressbar(self.progress_frame, orient='horizontal', mode='determinate')
        self.progressbar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.percent_label = ttk.Label(self.progress_frame, text="0%", width=5)
        self.percent_label.pack(side=tk.LEFT)

        self.on_process_finished()

    def start_process(self, process_type, target_func, *args):
        self.cancel_event.clear()
        self.active_process_type = process_type
        self.toggle_buttons(is_running=True)

        thread_args = (self.cancel_event,) + args
        threading.Thread(target=target_func, args=thread_args, daemon=True).start()

    def start_organization_process(self):
        self.start_process('organize', self.processor.run_organization_process)

    def start_printing_process(self):
        selected_folders = [d["path"] for d in self.stampa_checkbox_vars.values() if d["var"].get() == 1]
        self.start_process('print', self.processor.run_printing_process, selected_folders)

    def cancel_process(self):
        self.log_organizza("Annullamento richiesto...", "WARNING")
        self.cancel_event.set()
        self.cancel_org_button.config(state='disabled')
        self.cancel_print_button.config(state='disabled')

    def on_process_finished(self):
        self.toggle_buttons(is_running=False)
        self.active_process_type = None

    def toggle_buttons(self, is_running):
        state = 'disabled' if is_running else 'normal'
        self.organize_button.config(state=state)
        self.print_button.config(state=state)
        self.refresh_button.config(state=state)

        if is_running:
            if self.active_process_type == 'organize':
                self.organize_button.grid_forget()
                self.cancel_org_button.grid(row=1, column=0, columnspan=3, sticky="we", pady=(10, 5))
                self.cancel_org_button.config(state='normal')
            elif self.active_process_type == 'print':
                self.print_button.pack_forget()
                self.cancel_print_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
                self.cancel_print_button.config(state='normal')
        else:
            self.cancel_org_button.grid_forget()
            self.cancel_print_button.pack_forget()
            self.organize_button.grid(row=1, column=0, columnspan=3, sticky="we", pady=(10, 5))
            self.print_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

    def log_organizza(self, message, level='INFO'):
        self.master.after(0, self.log_widget, message, level)

    def setup_progress(self, max_value, label_text="Progresso:"):
        self.progress_label['text'] = label_text
        self.progress_frame.pack(fill=tk.X, pady=(10, 5), after=self.print_frame)
        self.progressbar['maximum'] = max_value
        self.progressbar['value'] = 0
        self.percent_label['text'] = "0%"

    def update_progress(self, value):
        self.progressbar['value'] = value
        max_val = self.progressbar['maximum']
        if max_val > 0: percent = (value / max_val) * 100; self.percent_label['text'] = f"{percent:.0f}%"

    def hide_progress(self):
        self.progress_frame.pack_forget()

    def populate_stampa_list(self):
        for widget in self.stampa_checkbox_frame.winfo_children(): widget.destroy()
        self.stampa_checkbox_vars.clear()
        year = self.app_config.canoni_selected_year.get()
        month = self.app_config.canoni_selected_month.get()
        odc_map = self.processor.get_odc_to_canone_map(year, month)
        dest_path = self.app_config.organizza_dest_dir.get()
        if not os.path.isdir(dest_path): return
        try:
            folders = sorted([d for d in os.listdir(dest_path) if os.path.isdir(os.path.join(dest_path, d))])
            for folder_name in folders:
                var = tk.IntVar()
                folder_path = os.path.join(dest_path, folder_name)
                file_count = 0
                try:
                    file_count = len([name for name in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, name))])
                except Exception as e:
                    self.log_organizza(f"Impossibile contare i file nella cartella '{folder_name}': {e}", "WARNING")
                display_text = folder_name
                if folder_name in odc_map: display_text = f"{folder_name} ({odc_map[folder_name]})"
                display_text = f"{display_text} - qt. {file_count}"
                cb = ttk.Checkbutton(self.stampa_checkbox_frame, text=display_text, variable=var)
                cb.pack(anchor="w", padx=5, fill='x')
                self.stampa_checkbox_vars[folder_name] = {"var": var, "path": folder_path}
        except Exception as e:
            self.log_organizza(f"Errore durante la lettura delle cartelle organizzate: {e}", "ERROR")
