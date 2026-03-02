from dataclasses import dataclass
from datetime import datetime


@dataclass
class DateCandidate:
    cell_ref: str
    status: str = "PENDING"
    value: datetime | None = None


@dataclass
class ModelConfig:
    name: str
    date_cells: list[str]
    # If provided, the value in the first id_cell must match this
    match_value: str | None = None
    id_cell: str | None = None
    tcl_cell: str | None = None


# Configuration for different Excel models used in renaming
RENAME_MODELS: list[ModelConfig] = [
    ModelConfig(
        name="Scheda Tecnica Verifica Disco Calibro",
        id_cell="N1",
        match_value="schedatecnicaverificadiscocalibro",
        date_cells=["AK2"],
        tcl_cell="L50", # Placeholder, da verificare
    ),
    ModelConfig(name="Scheda Valvole", id_cell="F3", match_value="schedavalvole", date_cells=["C54"], tcl_cell="L45"),
    ModelConfig(name="Scheda Valvole (Alt)", id_cell="G3", match_value="schedavalvole", date_cells=["C54"], tcl_cell="L45"),
    ModelConfig(
        name="Scheda Taratura Strumenti Digitali",
        id_cell="Q3",
        match_value="schedataraturastrumentidigitali",
        date_cells=["B50"],
        tcl_cell="L47",
    ),
    ModelConfig(name="Valvole di Regolazione", id_cell="T6", match_value="valvolediregolazione", date_cells=["B108"], tcl_cell="DE105"),
    ModelConfig(
        name="Scheda Controllo Valvole", id_cell="F2", match_value="schedacontrollovalvole", date_cells=["F56"], tcl_cell="L45"
    ),
    ModelConfig(
        name="Scheda Controllo Strumenti Digitali",
        id_cell="F2",
        match_value="schedacontrollostrumentidigitali",
        date_cells=["F44"],
        tcl_cell="L45"
    ),
    ModelConfig(
        name="Scheda Controllo Strumenti Analogici",
        id_cell="F2",
        match_value="schedacontrollostrumentianalogici",
        date_cells=["F49"],
        tcl_cell="L52"
    ),
    ModelConfig(
        name="Scheda Controllo Strumenti",
        id_cell="F2",
        match_value="schedacontrollostrumenti",
        date_cells=["F49", "F44"],
        tcl_cell="L45"
    ),
    ModelConfig(
        name="Scheda Taratura Strumento di Processo",
        id_cell="S3",
        match_value="schedataraturastrumentodiprocesso",
        date_cells=["B99"],
        tcl_cell="L45"
    ),
    ModelConfig(
        name="Scheda Controllo Valvole (E2)",
        id_cell="E2",
        match_value="schedacontrollovalvole",
        date_cells=["L46", "B46", "B108"],
        tcl_cell="L45"
    ),
    ModelConfig(
        name="Scheda Controllo Strumenti Digitali (E2)",
        id_cell="E2",
        match_value="schedacontrollostrumentidigitali",
        date_cells=["B45"],
        tcl_cell="L45"
    ),
    ModelConfig(
        name="Scheda Controllo Strumenti Analogici (E2)",
        id_cell="E2",
        match_value="schedacontrollostrumentianalogici",
        date_cells=["L52", "B45", "B50", "B108", "B99", "B105"],
        tcl_cell="L52"
    ),
    ModelConfig(
        name="Report Manutenzione Correttiva",
        id_cell="E2",
        match_value="schedacontrolloreportmanutenzionecorrettiva",
        date_cells=["B50"],
        tcl_cell="L52"
    ),
]

# Default date cells if no model matches
DEFAULT_DATE_CANDIDATES = ["B45", "B50", "B108", "B99", "B105", "L52", "C54"]
