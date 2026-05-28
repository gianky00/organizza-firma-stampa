# ♾️ PROGETTO: Organizza Firma Stampa (v1.0.0)

## 📝 Descrizione del Progetto
Applicazione desktop Python progettata per l'automazione del workflow documentale tecnico. Il programma gestisce l'organizzazione di schede Excel in cartelle basate sull'ODC (Ordine Di Contratto), l'applicazione di firme/timbri digitali (immagini), la stampa automatizzata e l'invio di notifiche via email.

### 🛠️ Stack Tecnologico
- **Linguaggio:** Python >= 3.12
- **Interfaccia Grafica:** Tkinter (con gestione avanzata di Splash Screen e caricamento pigro)
- **Automazione Documentale:** `pywin32` (COM Interface per Microsoft Excel e Word)
- **Analisi Dati:** `pandas` (caricamento lazy)
- **Sicurezza & Distribuzione:** `PyArmor` (offuscamento codice) e `Inno Setup` (installer Windows)
- **Qualità del Codice:** Ruff (linter/formatter), Mypy (type checking), Bandit (security), Deptry, Pytest.

---

## 🏗️ Architettura del Progetto
Il progetto segue una struttura modulare per separare la presentazione dalla logica di business:

- `main.py`: Entry point dell'applicazione. Gestisce lo Splash Screen e l'inizializzazione delle directory.
- `src/gui/`: Contiene la definizione della finestra principale e dei tab specializzati (`organize`, `signature`, `rename`, `fees`).
- `src/logic/`: Core logic per ogni funzionalità (es. `organization.py` per lo smistamento file, `signature.py` per l'inserimento del timbro).
- `src/utils/`: Utility trasversali. Notevoli i gateway/handler per Excel e Word che incapsulano la complessità delle chiamate COM.
- `src/domain/`: Modelli di dati e definizioni di tipi.
- `prompt/`: Documentazione interna e guide per lo stack di qualità e i processi di refactoring.

---

## 🚀 Comandi Utili

### Sviluppo e Testing
- **Eseguire l'applicazione:** `python main.py`
- **Eseguire i test:** `pytest` (usare `pytest --cov=src` per la copertura)
- **Linting & Formatting (Ruff):** `ruff check .` e `ruff format .`
- **Type Checking (Mypy):** `mypy src`

### Build e Release
Il processo di build è automatizzato e include l'offuscamento:
1. **Generazione build offuscata:** `python build_release.py` (genera la cartella `dist/`)
2. **Creazione Installer:** Compilare il file `installer_setup.iss` tramite Inno Setup Compiler puntando alla cartella `dist/`.

---

## 📏 Convenzioni di Sviluppo

1. **Lazy Loading:** Le dipendenze pesanti (`pandas`, `win32com`, `fpdf`) devono essere importate all'interno delle funzioni o dei metodi dove servono, per mantenere l'avvio della GUI reattivo.
2. **Gestione File:** Utilizzare sempre `src.utils.constants` per i percorsi delle cartelle e i nomi dei file. I path di rete (es. Database Tecnico) sono centralizzati qui.
3. **Automazione Office:** Quando si interagisce con Excel/Word tramite COM, assicurarsi di gestire correttamente la chiusura dei processi anche in caso di errore (usare i contesti definiti negli handler).
4. **Qualità del Codice:** Ogni modifica deve passare i controlli di Ruff e Mypy. Consultare i file in `prompt/` per le procedure di installazione e configurazione dello stack di qualità.

## 🧠 PROJECT MEMORIES
- **Startup:** Ottimizzato tramite Lazy loading `win32com`. Splash screen Tkinter aggiunto per feedback immediato.

---

## 📁 Directory di Lavoro (Runtime)
All'avvio, il programma verifica o crea le seguenti cartelle nella root:
- `FILE EXCEL DA FIRMARE`: Input per la funzionalità di firma.
- `PDF`: Destinazione dei file firmati/esportati.
- `SCHEDE DA ORGANIZZARE`: Input per lo smistamento ODC.
- `SCHEDE ORGANIZZATE`: Destinazione dello smistamento.
- `SCHEDE SENZA DATA`: Cartella di fallback per la rinomina.
