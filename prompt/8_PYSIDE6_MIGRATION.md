# Piano di Migrazione a PySide6 (Qt)

Questo documento traccia la procedura dettagliata per migrare l'intera applicazione da `Tkinter` a `PySide6`. La migrazione è suddivisa in fasi sequenziali per garantire che il passaggio avvenga senza perdere funzionalità, mantenendo il disaccoppiamento tra interfaccia e logica di business.

## Obiettivo
Eliminare ogni dipendenza da `tkinter` (`import tkinter`, `import tkinter.ttk`, `tkinter.messagebox`, `tkinter.filedialog`) sostituendo il framework GUI con `PySide6`.

## Fase 1: Setup e Astrazione Variabili (Data Binding)
In Tkinter, il progetto utilizza ampiamente `StringVar`, `BooleanVar`, e `IntVar` (specialmente all'interno di `ConfigManager` in `main_window.py`).
PySide6 non ha queste classi native, ma utilizza il paradigma Signal/Slot.
1. [ ] Aggiungere `PySide6` a `pyproject.toml` o `requirements.txt`.
2. [ ] Creare un modulo `src/utils/qt_vars.py` per simulare il comportamento di `StringVar`, `BooleanVar`, ecc. (es. classi con un segnale `value_changed`, un getter `get()` e un setter `set()`). Questo ridurrà l'impatto sul codice esistente che usa `.get()` e `.set()`.
3. [ ] Sostituire le classi di Tkinter nel `main_window.py` (AppConfig) con queste nuove classi custom.

## Fase 2: Sostituzione delle Utilità UI
Le utilità grafiche di base devono essere mappate ai moduli PySide6.
1. [ ] Modificare `src/utils/ui_utils.py` se presente.
2. [ ] Convertire l'uso di `tkinter.messagebox` in `QMessageBox`.
3. [ ] Convertire l'uso di `tkinter.filedialog` in `QFileDialog`.
4. [ ] Implementare una funzione alternativa a `self.after(0, callback)` di Tkinter. In PySide6, per aggiornare l'interfaccia da thread separati, si devono usare i `Signal` oppure `QTimer.singleShot(0, callback)` oppure `QMetaObject.invokeMethod`. Questo è vitale nei moduli `organization.py`, `signature.py` ecc., dove vengono passate delle callback per gli aggiornamenti della progress bar.

## Fase 3: Entry Point e Main Window
1. [ ] Aggiornare `main.py`. Rimuovere le chiamate a Tkinter, creare `QApplication(sys.argv)`.
2. [ ] Implementare lo Splash Screen tramite `QSplashScreen` o un widget borderless personalizzato in `main.py`.
3. [ ] Convertire `src/gui/main_window.py`. `MainWindow` dovrà ereditare da `QMainWindow`.
4. [ ] Sostituire il gestore dei tab (`ttk.Notebook`) con `QTabWidget`.

## Fase 4: Migrazione dei Tab (Layout e Widget)
Ciascun tab deve essere riscritto passando dai layout `pack`/`grid` di Tkinter a `QVBoxLayout`, `QHBoxLayout`, e `QGridLayout` di Qt.
1. [ ] `src/gui/tabs/organize_tab.py`: Sostituire widget (Frame -> QWidget, Label -> QLabel, Button -> QPushButton, Entry -> QLineEdit).
2. [ ] `src/gui/tabs/signature_tab.py`: Implementare `QProgressBar` invece di `ttk.Progressbar`.
3. [ ] `src/gui/tabs/rename_tab.py`: Sostituire caselle di testo e pulsanti.
4. [ ] `src/gui/tabs/fees_tab.py`: Aggiornare input, dropdown (`ttk.Combobox` -> `QComboBox`).
5. [ ] `src/gui/tabs/settings_tab.py`: È il tab più complesso (usa treeview/grid per la lista modelli). Occorre usare un layout dinamico o un `QTableWidget` in sostituzione dei widget iterati.

## Fase 5: Verifica Disaccoppiamento e Testing
1. [ ] Eseguire una ricerca globale per `tkinter` per assicurarsi che nessun import sia rimasto.
2. [ ] Lanciare i test `pytest` (e adattare eventuali mock nel file `conftest.py` o simili che facevano affidamento ai MagicMock di Tkinter).
3. [ ] Avviare `main.py` e collaudare manualmente i percorsi critici.

---
**Regole di Codifica per la Migrazione:**
- Usare nomi di variabili chiari e idiomatici Qt (es. `layout` invece di chiamare `pack()`).
- Mantenere tutto il resto della logica di business invariata; le modifiche sono relegate al layer `gui/` e alle variabili di configurazione di appoggio.
- Non bloccare l'UI thread: se le callback dai worker chiamano aggiornamenti UI, assicurarsi che i worker emettano Signal collegati a Slot del main thread.
