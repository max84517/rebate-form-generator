"""Load and save application configuration from config.json in the project root."""
import json
from pathlib import Path

_HERE = Path(__file__).resolve()
# src/rebate_form_generator/config/settings.py → project root (4 levels up)
_PROJECT_ROOT = _HERE.parent.parent.parent.parent

CONFIG_FILE = _PROJECT_ROOT / "config.json"
DEFAULT_OUTPUT = _PROJECT_ROOT / "data" / "output"


class Settings:
    def __init__(self) -> None:
        self.nb_kb: str = ""
        self.dt_kb: str = ""
        self.peripheral: str = ""
        self.output_path: str = str(DEFAULT_OUTPUT)
        self._load()

    def _load(self) -> None:
        if not CONFIG_FILE.exists():
            return
        try:
            data: dict = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            self.nb_kb = data.get("nb_kb", "")
            self.dt_kb = data.get("dt_kb", "")
            self.peripheral = data.get("peripheral", "")
            self.output_path = data.get("output_path", str(DEFAULT_OUTPUT))
        except Exception:
            pass

    def save(self) -> None:
        data = {
            "nb_kb": self.nb_kb,
            "dt_kb": self.dt_kb,
            "peripheral": self.peripheral,
            "output_path": self.output_path,
        }
        CONFIG_FILE.write_text(
            json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8"
        )
