"""Application configuration, and where the app keeps its files.

Two locations, deliberately separate:

* **data_dir** — inside the project, holds what the user maintains: the Word
  templates, the legacy supplier-info workbook, the contract details.
* **work_dir** — outside the project, holds everything the pipeline derives.
  The project folder is normally inside OneDrive, and the intermediates are
  rewritten on every run; keeping them here stops a couple of megabytes being
  re-synced each time, and keeps the repository to source.
"""
import json
import os
from pathlib import Path

_HERE = Path(__file__).resolve()
# src/rebate_form_generator/config/settings.py → project root (4 levels up)
_PROJECT_ROOT = _HERE.parent.parent.parent.parent

APP_NAME = "RebateFormGenerator"

CONFIG_FILE = _PROJECT_ROOT / "config.json"

#: User-maintained inputs, kept with the project.
DATA_DIR = _PROJECT_ROOT / "data"


#: Folder name under the user's home directory.
WORK_DIRNAME = ".rebate-form-generator"


def default_work_dir() -> Path:
    """Where derived data lives: ``C:\\Users\\<you>\\.rebate-form-generator``.

    Deliberately *not* ``%LOCALAPPDATA%``. This project's virtualenv is built
    on the Microsoft Store Python, which runs in an app container and silently
    redirects writes to AppData into a private per-package folder. The app
    would see its own files, but Explorer — and so the "Working files" button —
    would open an empty directory. The home directory is not redirected, and
    is not inside OneDrive.
    """
    override = os.environ.get("REBATE_FORM_GENERATOR_HOME")
    if override:
        return Path(override)
    return Path.home() / WORK_DIRNAME


class Settings:
    def __init__(self) -> None:
        self.nb_kb: str = ""
        self.dt_kb: str = ""
        self.peripheral: str = ""
        self.last_fy: str = ""
        self.last_quarter: int = 0
        self.form_numbers: dict[str, str] = {}
        self.icertis_codes: dict[str, str] = {}
        self._load()

    # ── locations ────────────────────────────────────────────────────
    @property
    def data_dir(self) -> Path:
        """User-maintained inputs: templates, supplier info, contract details."""
        return DATA_DIR

    @property
    def work_dir(self) -> Path:
        """Derived data: ingested sheets, rebate raw, contract input, cache."""
        path = default_work_dir()
        path.mkdir(parents=True, exist_ok=True)
        return path

    @property
    def cache_dir(self) -> Path:
        path = self.work_dir / "cache"
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _load(self) -> None:
        if not CONFIG_FILE.exists():
            return
        try:
            data: dict = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            self.nb_kb = data.get("nb_kb", "")
            self.dt_kb = data.get("dt_kb", "")
            self.peripheral = data.get("peripheral", "")
            # output_path is NOT loaded from config — always use DEFAULT_OUTPUT
            # so moving to a new installation folder never picks up a stale path
            self.last_fy = data.get("last_fy", "")
            self.last_quarter = int(data.get("last_quarter", 0))
            self.form_numbers = data.get("form_numbers", {})
            self.icertis_codes = data.get("icertis_codes", {})
        except Exception:
            pass

    def save(self) -> None:
        data = {
            "nb_kb": self.nb_kb,
            "dt_kb": self.dt_kb,
            "peripheral": self.peripheral,
            "last_fy": self.last_fy,
            "last_quarter": self.last_quarter,
            "form_numbers": self.form_numbers,
            "icertis_codes": self.icertis_codes,
        }
        CONFIG_FILE.write_text(
            json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8"
        )
