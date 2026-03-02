# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Windows desktop application for document management automation. Built with Python/Tkinter, it handles:
- PDF signing (signature stamps on Excel sheets exported to PDF)
- File renaming (add emission dates from Excel cell values)
- File organization (sort by ODC code, print to network printer)
- Monthly fee report generation (Word template population)

## Running the Application

```bash
python main.py
# or double-click avvia.bat (Windows)
```

## Architecture

```
main.py                  # Entry point with splash screen
src/
├── gui/
│   ├── main_window.py   # Main tkinter app with 4 tabs
│   └── tabs/            # One tab per feature (signature, rename, organize, fees)
├── logic/               # Business logic processors (one per tab)
├── utils/
│   ├── constants.py     # Centralized paths, email templates, sheet models
│   ├── config_manager.py# JSON config load/save
│   ├── excel_handler.py # COM wrapper for Excel (context manager)
│   ├── word_handler.py  # COM wrapper for Word (context manager)
│   └── file_utils.py    # Folder operations
└── assets/              # Signature stamp image
```

**Data flow:** GUI tabs → Logic processors → Utils handlers

## Dependencies

- **pywin32** - COM automation for Excel, Word, Outlook, Win32Print
- **Ghostscript** - External executable for PDF compression (path in config)
- **Microsoft Office** - Must be installed (Excel, Word, Outlook)

## Key Configuration

`config_programma.json` stores user settings (paths, email, printer). Auto-loaded on startup.

`src/utils/constants.py` contains:
- `SHEET_MODELS` - Excel cell mappings for different sheet types
- `EMAIL_TEMPLATES` - Formal/informal email body templates
- Network paths, folder names, Italian month mappings

## Excel Sheet Model Detection

The system identifies sheet types by checking specific cells (N1, F3, G3, AY3, T6, etc.). Each model has different:
- Print areas
- Date cell locations
- Signature cell positions

## Windows-Specific Notes

- Uses COM automation (pywin32) - Windows only
- UNC paths for network shares (e.g., `\\192.168.11.251\...`)
- `os.startfile()` for opening folders
- `win32print` for printer enumeration
- Heavy imports are deferred to after splash screen displays

## Threading Pattern

Long operations run in background threads. Use `gui.after()` to marshal callbacks to main thread for UI updates.
