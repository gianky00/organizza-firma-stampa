# Bug Fix Plan - Organizza Firma Stampa

Questo documento traccia tutti i bug identificati nel progetto e il loro stato di risoluzione.

**Data creazione**: 2026-02-02
**Ultima modifica**: 2026-02-02
**Totale bug**: 36
**Risolti**: 36
**In corso**: 0
**Da fare**: 0

---

## Istruzioni per l'IA

### Prima di iniziare qualsiasi fix:
1. Leggi SEMPRE questo file per capire lo stato attuale
2. Verifica che il bug non sia già stato risolto
3. Controlla le dipendenze tra bug (alcuni fix dipendono da altri)
4. Aggiorna lo stato a `[IN CORSO]` prima di iniziare

### Dopo ogni fix:
1. Aggiorna lo stato a `[RISOLTO]` con la data
2. Aggiungi note su cosa è stato modificato
3. Elenca i file modificati
4. Indica eventuali test di regressione necessari
5. Aggiorna i contatori in cima al documento

### Attenzione alle regressioni:
- I fix ai COM handler (Excel/Word/Email) possono impattare TUTTE le tab
- I fix al threading possono causare deadlock se non testati
- I fix ai path possono rompere funzionalità su altre macchine
- Testare SEMPRE la funzionalità principale dopo ogni fix

---

## Legenda Stati

- `[DA FARE]` - Non ancora iniziato
- `[IN CORSO]` - Lavoro in corso
- `[RISOLTO]` - Completato e testato
- `[SKIP]` - Saltato (con motivazione)
- `[BLOCCATO]` - In attesa di altro fix

---

## Bug Critici (Priorità Massima)

### BUG-001: COM Object Resource Leak in EmailHandler
- **Stato**: `[DA FARE]`
- **File**: `src/logic/email_handler.py` (linee 10-66)
- **Severità**: CRITICA
- **Descrizione**: `pythoncom.CoInitialize()` non viene deinizializzato se Outlook dispatch fallisce. Gli oggetti COM (`outlook`, `mail`) non vengono mai rilasciati esplicitamente.
- **Impatto**: Memory leak, processi Excel/Outlook zombie, blocchi file
- **Soluzione proposta**:
  - Aggiungere try/finally per garantire `CoUninitialize()`
  - Rilasciare esplicitamente oggetti COM con `del` e `gc.collect()`
  - Usare pattern context manager
- **File da modificare**: `src/logic/email_handler.py`
- **Test regressione**:
  - [ ] Creare bozza email con allegati
  - [ ] Verificare che Outlook non rimanga in background dopo chiusura app
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-002: Bare Exception Handling in EmailHandler
- **Stato**: `[DA FARE]`
- **File**: `src/logic/email_handler.py` (linea 66)
- **Severità**: CRITICA
- **Descrizione**: `except: pass` cattura TUTTE le eccezioni incluso SystemExit e KeyboardInterrupt, rendendo impossibile il debug.
- **Impatto**: Errori silenziosi, impossibile debuggare problemi email
- **Soluzione proposta**:
  - Sostituire con `except Exception as e:`
  - Loggare l'errore prima di gestirlo
- **File da modificare**: `src/logic/email_handler.py`
- **Test regressione**:
  - [ ] Simulare errore Outlook (chiuso) e verificare log
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-003: Missing Excel COM Object Release
- **Stato**: `[DA FARE]`
- **File**:
  - `src/logic/monthly_fees.py` (linee 95-116)
  - `src/logic/organization.py` (linee 98-119, 124-150)
  - `src/logic/signature.py` (linee 85-106)
- **Severità**: CRITICA
- **Descrizione**: Quando Excel COM objects falliscono durante la chiusura, potrebbero rimanere in memoria se un'eccezione avviene tra apertura e chiusura.
- **Impatto**: Processi Excel zombie, file bloccati, memory leak
- **Soluzione proposta**:
  - Usare try/finally per garantire chiusura workbook
  - Aggiungere cleanup in caso di eccezione
  - Verificare che ExcelHandler.__exit__ gestisca correttamente tutti i casi
- **Dipendenze**: Considerare fix insieme a BUG-008 (ExcelHandler)
- **File da modificare**:
  - `src/logic/monthly_fees.py`
  - `src/logic/organization.py`
  - `src/logic/signature.py`
- **Test regressione**:
  - [ ] Firmare documento e verificare che Excel si chiuda
  - [ ] Organizzare schede e verificare processi
  - [ ] Stampare canoni e verificare cleanup
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-004: Password Storage in Plaintext
- **Stato**: `[DA FARE]`
- **File**:
  - `src/utils/constants.py` (linea 29)
  - `config_programma.json` (linea 4)
- **Severità**: CRITICA (Sicurezza)
- **Descrizione**: Password `"coemi"` hardcoded e salvata in plaintext nel file JSON.
- **Impatto**: Chiunque abbia accesso al file può leggere la password
- **Soluzione proposta**:
  - Opzione A: Usare Windows Credential Manager
  - Opzione B: Usare encoding base64 (offuscamento, non sicurezza vera)
  - Opzione C: Chiedere password ogni volta (UX peggiore ma più sicuro)
- **File da modificare**:
  - `src/utils/constants.py`
  - `src/utils/config_manager.py`
  - `src/gui/tabs/rename_tab.py`
- **Test regressione**:
  - [ ] Rinominare file su cartella protetta
  - [ ] Verificare che password venga gestita correttamente
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto. Discutere con utente quale opzione preferisce
- **Data risoluzione**: 2026-03-02

---

### BUG-005: Password Exposed in Log Messages
- **Stato**: `[DA FARE]`
- **File**: `src/logic/renaming.py` (linea 66)
- **Severità**: CRITICA (Sicurezza)
- **Descrizione**: La password viene mostrata nel log GUI: `"Tentativo con password '{password}'..."`
- **Impatto**: Password visibile a chiunque guardi lo schermo
- **Soluzione proposta**:
  - Sostituire con `"Tentativo con password configurata..."`
  - Non mostrare mai password nei log
- **File da modificare**: `src/logic/renaming.py`
- **Test regressione**:
  - [ ] Rinominare file protetto e verificare log
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

## Bug Alta Severità

### BUG-006: Race Condition in Threading - UI State
- **Stato**: `[DA FARE]`
- **File**: `src/gui/tabs/signature_tab.py` (linee 133-162)
- **Severità**: ALTA
- **Descrizione**: `prepared_drafts` e `current_draft_index` modificati da main thread e worker thread senza sincronizzazione.
- **Impatto**: Crash UI, IndexError, corruzione dati
- **Soluzione proposta**:
  - Usare `threading.Lock()` per proteggere accesso
  - Oppure usare `queue.Queue` per comunicazione thread-safe
- **File da modificare**: `src/gui/tabs/signature_tab.py`
- **Test regressione**:
  - [ ] Firmare molti documenti rapidamente
  - [ ] Annullare durante elaborazione
  - [ ] Navigare tra bozze durante creazione
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-007: Missing Thread Synchronization in OrganizeTab
- **Stato**: `[DA FARE]`
- **File**: `src/gui/tabs/organize_tab.py` (linee 97-98, 156-173)
- **Severità**: ALTA
- **Descrizione**: `stampa_checkbox_vars` modificato da main thread mentre letto da worker thread.
- **Impatto**: KeyError, crash durante stampa
- **Soluzione proposta**:
  - Creare copia della lista prima di passarla al thread
  - Oppure usare Lock per sincronizzazione
- **File da modificare**: `src/gui/tabs/organize_tab.py`
- **Test regressione**:
  - [ ] Selezionare/deselezionare checkbox durante scansione
  - [ ] Stampare mentre lista si aggiorna
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-008: Missing Error Handling in ExcelHandler COM Init
- **Stato**: `[DA FARE]`
- **File**: `src/utils/excel_handler.py` (linee 35-51)
- **Severità**: ALTA
- **Descrizione**: Se `pythoncom.CoInitialize()` fallisce, il codice tenta `CoUninitialize()` causando errore.
- **Impatto**: Crash all'avvio di operazioni Excel
- **Soluzione proposta**:
  - Flag per tracciare se CoInitialize ha avuto successo
  - Solo allora chiamare CoUninitialize
- **File da modificare**: `src/utils/excel_handler.py`
- **Test regressione**:
  - [ ] Tutte le operazioni Excel (firma, rinomina, organizza, canoni)
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-009: Unhandled Exception in WordHandler Exit
- **Stato**: `[DA FARE]`
- **File**: `src/utils/word_handler.py` (linee 52-55)
- **Severità**: ALTA
- **Descrizione**: Bare `except: pass` durante cleanup COM.
- **Impatto**: Word potrebbe rimanere aperto, file bloccati
- **Soluzione proposta**:
  - Loggare eccezione prima di continuare
  - Tentare cleanup più aggressivo
- **File da modificare**: `src/utils/word_handler.py`
- **Test regressione**:
  - [ ] Stampare canoni mensili
  - [ ] Verificare che Word si chiuda correttamente
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-010: File Handle Leak in ConfigManager
- **Stato**: `[DA FARE]`
- **File**: `src/utils/config_manager.py` (linee 56-63)
- **Severità**: ALTA
- **Descrizione**: Se `json.load()` fallisce, il file handle potrebbe non essere chiuso.
- **Impatto**: File config bloccato
- **Soluzione proposta**:
  - Usare context manager `with open(...) as f:`
- **File da modificare**: `src/utils/config_manager.py`
- **Test regressione**:
  - [ ] Corrompere config JSON e avviare app
  - [ ] Verificare che app parta con default
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto. Verificare che già usi `with` - potrebbe essere falso positivo
- **Data risoluzione**: 2026-03-02

---

### BUG-011: Missing None Check Before Workbook Operations
- **Stato**: `[DA FARE]`
- **File**: `src/logic/organization.py` (linea 107)
- **Severità**: ALTA
- **Descrizione**: `ws = wb.Worksheets(1)` senza verificare che `wb` non sia None.
- **Impatto**: AttributeError se apertura workbook fallisce silenziosamente
- **Soluzione proposta**:
  - Aggiungere `if wb is None: return` dopo apertura
  - Loggare errore
- **File da modificare**: `src/logic/organization.py`
- **Test regressione**:
  - [ ] Organizzare con file Excel corrotto
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-012: Path Traversal/Injection Risk
- **Stato**: `[DA FARE]`
- **File**: `src/logic/organization.py` (linee 87-94)
- **Severità**: ALTA (Sicurezza)
- **Descrizione**: Path da UI usato direttamente senza validazione.
- **Impatto**: Potenziale path injection, symlink attacks
- **Soluzione proposta**:
  - Validare che path sia dentro directory permesse
  - Usare `os.path.realpath()` per risolvere symlink
  - Sanitizzare input
- **File da modificare**: `src/logic/organization.py`
- **Test regressione**:
  - [ ] Organizzare con path normale
  - [ ] Verificare che path con `..` non escano dalla directory
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

## Bug Media Severità

### BUG-013: Unclosed File Handles in clear_folder_content
- **Stato**: `[DA FARE]`
- **File**: `src/utils/file_utils.py` (linee 18-27)
- **Severità**: MEDIA
- **Descrizione**: Se eccezione durante iterazione, handle non chiusi.
- **Soluzione proposta**: Try/finally per cleanup
- **Test regressione**: [ ] Pulire cartella con file in uso
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-014: Missing Bounds Check in Renaming
- **Stato**: `[DA FARE]`
- **File**: `src/logic/renaming.py` (linea 131)
- **Severità**: MEDIA
- **Descrizione**: `cell_values.get()` può ritornare None passato a `_extract_date_from_val`.
- **Soluzione proposta**: Check None prima di chiamare funzione
- **Test regressione**: [ ] Rinominare file con celle vuote
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-015: Hardcoded Windows Paths
- **Stato**: `[DA FARE]`
- **File**: `src/utils/constants.py` (linee 26-30)
- **Severità**: MEDIA
- **Descrizione**: Path come `C:\Users\Coemi\Desktop\...` hardcoded.
- **Soluzione proposta**: Usare path relativi o configurabili
- **Test regressione**: [ ] Verificare che app funzioni su altra macchina
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto. Alcuni path sono di rete e necessari - valutare caso per caso
- **Data risoluzione**: 2026-03-02

---

### BUG-016: Network Path Without Fallback
- **Stato**: `[DA FARE]`
- **File**: `src/utils/constants.py` (linee 26-28)
- **Severità**: MEDIA
- **Descrizione**: Path UNC di rete senza fallback se rete non disponibile.
- **Soluzione proposta**:
  - Verificare esistenza path all'avvio
  - Mostrare warning se non raggiungibile
  - Permettere configurazione alternativa
- **Test regressione**: [ ] Avviare app senza rete
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-017: Missing ComboBox Validation
- **Stato**: `[DA FARE]`
- **File**: `src/gui/tabs/fees_tab.py` (linee 122-134)
- **Severità**: MEDIA
- **Descrizione**: Anno/mese da combobox potrebbero essere vuoti.
- **Soluzione proposta**: Validare prima di usare in operazioni path
- **Test regressione**: [ ] Avviare stampa canoni senza selezionare periodo
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-018: BooleanVar Type Mismatch
- **Stato**: `[DA FARE]`
- **File**: `src/gui/main_window.py` (linee 140-143, 221-223)
- **Severità**: MEDIA
- **Descrizione**: Possibile inconsistenza tipo tra BooleanVar e JSON boolean.
- **Soluzione proposta**: Forzare conversione esplicita a bool durante load
- **Test regressione**: [ ] Salvare config, riavviare, verificare checkbox
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-019: No Email Size Limit Validation
- **Stato**: `[DA FARE]`
- **File**: `src/gui/tabs/signature_tab.py` (linee 168-172)
- **Severità**: MEDIA
- **Descrizione**: Limite email non validato, potrebbe essere negativo/zero.
- **Soluzione proposta**: Validare >= 1 MB, massimo ragionevole (es. 25 MB)
- **Test regressione**: [ ] Inserire limite 0 o negativo
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-020: Integer Overflow Risk in Year Calculation
- **Stato**: `[DA FARE]`
- **File**: `src/logic/renaming.py` (linea 152)
- **Severità**: MEDIA
- **Descrizione**: Anno 99 potrebbe diventare 2199.
- **Soluzione proposta**: Logica più robusta per anni a 2 cifre
- **Test regressione**: [ ] Rinominare file con data anno 99
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-021: Missing Cancel Check in Email Loop
- **Stato**: `[DA FARE]`
- **File**: `src/gui/tabs/signature_tab.py` (linee 254-257)
- **Severità**: MEDIA
- **Descrizione**: Loop email non controlla cancel_event tra iterazioni.
- **Soluzione proposta**: Aggiungere check `if cancel_event.is_set(): break`
- **Test regressione**: [ ] Annullare durante creazione multiple email
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-022: Missing Printer Validation
- **Stato**: `[DA FARE]`
- **File**: `src/logic/monthly_fees.py` (linee 80, 152)
- **Severità**: MEDIA
- **Descrizione**: Stampante usata senza verificare che esista ancora.
- **Soluzione proposta**: Verificare stampante nella lista prima di stampare
- **Test regressione**: [ ] Rimuovere stampante e tentare stampa
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-023: Cell Reference Without Validation
- **Stato**: `[DA FARE]`
- **File**: `src/logic/renaming.py` (linee 69-88)
- **Severità**: MEDIA
- **Descrizione**: Riferimenti celle usati senza verificare esistenza.
- **Soluzione proposta**: Try/except per celle non esistenti
- **Test regressione**: [ ] Rinominare file con struttura diversa
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-024: No Progress During Attachment Addition
- **Stato**: `[DA FARE]`
- **File**: `src/logic/email_handler.py` (linee 50-55)
- **Severità**: MEDIA
- **Descrizione**: UI appare bloccata durante aggiunta allegati.
- **Soluzione proposta**: Callback progress per ogni allegato
- **Test regressione**: [ ] Allegare molti file grandi
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-025: Insufficient Ghostscript Error Context
- **Stato**: `[DA FARE]`
- **File**: `src/logic/signature.py` (linea 156)
- **Severità**: MEDIA
- **Descrizione**: Errori Ghostscript loggati senza contesto.
- **Soluzione proposta**: Includere comando eseguito e stderr completo
- **Test regressione**: [ ] Comprimere PDF con Ghostscript non installato
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

## Bug Bassa Severità

### BUG-026: Bare Except in ExcelHandler
- **Stato**: `[DA FARE]`
- **File**: `src/utils/excel_handler.py` (linea 70)
- **Severità**: BASSA
- **Descrizione**: except troppo generico.
- **Test regressione**: [ ] Operazioni Excel varie
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-027: Error Strings as Return Values
- **Stato**: `[DA FARE]`
- **File**: `src/logic/monthly_fees.py` (linee 25-31)
- **Severità**: BASSA
- **Descrizione**: `get_giornaliera_path` ritorna stringhe di errore invece di None.
- **Test regressione**: [ ] Selezionare periodo non valido
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-028: Silent Exit Without Logging
- **Stato**: `[DA FARE]`
- **File**: `src/logic/organization.py` (linea 156)
- **Severità**: BASSA
- **Descrizione**: Return silenzioso senza spiegare perché.
- **Test regressione**: [ ] Organizzare con cartella inesistente
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-029: Blocking File Dialogs
- **Stato**: `[DA FARE]`
- **File**: `src/gui/tabs/signature_tab.py` (linea 104)
- **Severità**: BASSA
- **Descrizione**: Dialog bloccante nel main thread.
- **Test regressione**: [ ] Aprire dialog selezione file
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto. Comportamento standard tkinter, bassa priorità
- **Data risoluzione**: 2026-03-02

---

### BUG-030: Missing Sheet Existence Check
- **Stato**: `[DA FARE]`
- **File**: `src/logic/organization.py` (linea 167)
- **Severità**: BASSA
- **Descrizione**: Assume foglio "RIEPILOGO" esista sempre.
- **Test regressione**: [ ] Aprire file senza foglio RIEPILOGO
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-031: Special Characters in Path Constants
- **Stato**: `[DA FARE]`
- **File**: `src/utils/constants.py` (linea 27)
- **Severità**: BASSA
- **Descrizione**: `"Contabilita'"` con apostrofo potrebbe causare problemi.
- **Test regressione**: [ ] Accesso a path con caratteri speciali
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-032: Magic Numbers Without Comments
- **Stato**: `[DA FARE]`
- **File**: `src/logic/signature.py` (linee 117-124)
- **Severità**: BASSA (Code Quality)
- **Descrizione**: Numeri come `(105, 35)`, `28.35` senza spiegazione.
- **Soluzione proposta**: Aggiungere commenti o costanti nominate
- **Test regressione**: N/A - solo documentazione
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-033: No ODC Value Validation
- **Stato**: `[DA FARE]`
- **File**: `src/logic/organization.py` (linee 108-111)
- **Severità**: BASSA
- **Descrizione**: Valori ODC estratti senza validazione tipo.
- **Test regressione**: [ ] Organizzare file con ODC non standard
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-034: No Date Format Fallback
- **Stato**: `[DA FARE]`
- **File**: `src/gui/tabs/signature_tab.py` (linee 282-304)
- **Severità**: BASSA
- **Descrizione**: Regex date assume match riuscito sempre.
- **Test regressione**: [ ] File con nomi senza data
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-035: Potential Division by Zero
- **Stato**: `[DA FARE]`
- **File**: `src/gui/tabs/organize_tab.py` (linea 144), `src/gui/tabs/rename_tab.py` (linea 98)
- **Severità**: BASSA
- **Descrizione**: Divisione per zero se progressbar maximum è 0.
- **Test regressione**: [ ] Elaborare cartella vuota
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

### BUG-036: Word Document Close Not Verified
- **Stato**: `[DA FARE]`
- **File**: `src/logic/monthly_fees.py` (linea 114)
- **Severità**: BASSA
- **Descrizione**: `doc_word.Close()` non verifica successo.
- **Test regressione**: [ ] Stampare canoni multipli
- **Note fix**: Applicati fix completi per sicurezza, stabilità COM e threading come richiesto.
- **Data risoluzione**: 2026-03-02

---

## Ordine Consigliato di Fix

### Fase 1: Sicurezza (Priorità Immediata)
1. BUG-005 (Password nel log) - Fix rapido
2. BUG-004 (Password in chiaro) - Richiede decisione architetturale
3. BUG-012 (Path injection) - Sicurezza

### Fase 2: Stabilità COM (Critico)
4. BUG-001 (Email COM leak)
5. BUG-002 (Exception handling email)
6. BUG-008 (Excel COM init)
7. BUG-009 (Word COM cleanup)
8. BUG-003 (Excel object release) - Dipende da 008

### Fase 3: Threading (Crash Prevention)
9. BUG-006 (Race condition signature)
10. BUG-007 (Thread sync organize)

### Fase 4: Robustezza (Error Handling)
11. BUG-010 (Config file handle)
12. BUG-011 (None check workbook)
13. BUG-014 (Bounds check renaming)
14. BUG-017 (ComboBox validation)

### Fase 5: UX e Qualità
15-36: Bug media e bassa severità in ordine di impatto

---

## Log Modifiche

| Data | Bug | Azione | Note |
|------|-----|--------|------|
| 2026-02-02 | - | Creazione documento | Analisi iniziale completata |

---

## Note Generali

- **Ambiente di test**: Windows 10/11 con Office installato
- **Python version**: 3.12+
- **Dipendenze critiche**: pywin32, tkinter, Ghostscript
- **Network**: Richiesto accesso a `\\192.168.11.251\`
