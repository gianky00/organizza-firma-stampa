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
    match_value: str
    id_cells: list[str]
    date_cells: list[str]
    tcl_cells: list[str]
    print_area: str = "A1:N50"


# Configuration for different Excel models used in renaming, organization and printing
RENAME_MODELS: list[ModelConfig] = [
    ModelConfig(
        name="Valvole Regolazione",
        id_cells=["T3", "T6", "E2"],
        match_value="valvolediregolazione",
        date_cells=["B105", "B108"],
        tcl_cells=["DE108", "DE105"],
        print_area="A1:N115",
    ),
    ModelConfig(
        name="Disco Calibro",
        id_cells=["N1", "F2", "E2"],
        match_value="schedatecnicaverificadiscocalibro",
        date_cells=["AK2"],
        tcl_cells=["L50"],
        print_area="A2:N45",
    ),
    ModelConfig(
        name="Scheda Valvole",
        id_cells=["F3", "G3"],
        match_value="schedavalvole",
        date_cells=["C54"],
        tcl_cells=["L45"],
        print_area="A1:N50",
    ),
    ModelConfig(
        name="Digitali",
        id_cells=["Q3", "F2", "E2"],
        match_value="schedataraturastrumentidigitali",
        date_cells=["B50", "B45"],
        tcl_cells=["L47", "L45"],
        print_area="A2:N50",
    ),
    ModelConfig(
        name="Analogici",
        id_cells=["F2", "E2", "T2"],
        match_value="schedacontrollostrumentianalogici",
        date_cells=["F49", "B45", "L52"],
        tcl_cells=["L52"],
        print_area="A2:N55",
    ),
    ModelConfig(
        name="Correttiva",
        id_cells=["E2", "F2"],
        match_value="schedacontrolloreportmanutenzionecorrettiva",
        date_cells=["B50"],
        tcl_cells=["L52"],
        print_area="A2:N55",
    ),
    ModelConfig(
        name="Scheda Manutenzione",
        id_cells=["T2", "E2", "T3"],
        match_value="schedamanutenzione",
        date_cells=["B108", "B105"],
        tcl_cells=["FO104"],
        print_area="A1:N115",
    ),
]

# Default date cells if no model matches
DEFAULT_DATE_CANDIDATES = ["B45", "B50", "B108", "B99", "B105", "L52", "C54"]
