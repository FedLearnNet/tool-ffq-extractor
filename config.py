from pathlib import Path
from typing import Optional

from pydantic.dataclasses import dataclass
from pyfedappwrap.learning.run_runfig import AppConfig, AppInputConfig, AppOutputConfig


@dataclass
class MyAppConfig(AppConfig):
    workers: int = 4
    sheet: str = ""
    date_cell: str = "A1"
    patient_id_cell: str = "A3"


@dataclass
class MyAppInputConfig(AppInputConfig):
    ffq_archive: Path


@dataclass
class MyAppOutputConfig(AppOutputConfig):
    nutrition_csv: Optional[Path] = None
