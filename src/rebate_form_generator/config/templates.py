"""Word template registry.

Templates used to be picked by "most recently modified file in template/",
which OneDrive can reorder behind your back.  They are now named, dated and
explicitly chosen, with the registry kept in ``data/templates.json``.
"""
from __future__ import annotations

import json
import re
import uuid
from datetime import datetime
from pathlib import Path

REGISTRY_NAME = "templates.json"
PREVIEW_DIRNAME = ".preview"

_ILLEGAL = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _safe_stem(name: str, fallback: str = "template") -> str:
    cleaned = _ILLEGAL.sub("_", str(name or "")).strip().strip(".")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned[:100] or fallback


class TemplateStore:
    """Named, dated Word templates with one marked active."""

    def __init__(self, template_dir: Path) -> None:
        self.dir = template_dir
        self.registry = template_dir / REGISTRY_NAME
        self.preview_dir = template_dir / PREVIEW_DIRNAME

    # ── registry io ──────────────────────────────────────────────────
    def _read(self) -> dict:
        if not self.registry.exists():
            return {"active": "", "items": []}
        try:
            data = json.loads(self.registry.read_text(encoding="utf-8"))
        except (ValueError, OSError):
            return {"active": "", "items": []}
        if not isinstance(data, dict):
            return {"active": "", "items": []}
        data.setdefault("active", "")
        data.setdefault("items", [])
        return data

    def _write(self, data: dict) -> None:
        self.dir.mkdir(parents=True, exist_ok=True)
        self.registry.write_text(
            json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8"
        )

    # ── listing ──────────────────────────────────────────────────────
    def list(self) -> dict:
        """Return ``{"active": id, "items": [...]}``, reconciled with disk."""
        data = self._read()
        known = {i.get("file"): i for i in data["items"] if i.get("file")}

        on_disk = sorted(
            f for f in self.dir.glob("*.docx") if not f.name.startswith("~$")
        ) if self.dir.exists() else []
        on_disk += sorted(
            f for f in self.dir.glob("*.Docx") if not f.name.startswith("~$")
        ) if self.dir.exists() else []

        seen: set[str] = set()
        items: list[dict] = []
        for path in on_disk:
            if path.name.lower() in seen:
                continue
            seen.add(path.name.lower())
            entry = known.get(path.name)
            stat = path.stat()
            if entry is None:
                # A file that was dropped into template/ by hand: adopt it,
                # dating it by its mtime.
                entry = {
                    "id": uuid.uuid4().hex[:12],
                    "name": path.stem,
                    "file": path.name,
                    "added": datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
                }
            entry["size"] = stat.st_size
            entry["missing"] = False
            items.append(entry)

        items.sort(key=lambda i: i.get("added", ""), reverse=True)
        active = data.get("active", "")
        if active not in {i["id"] for i in items}:
            active = items[0]["id"] if items else ""

        result = {"active": active, "items": items}
        self._write(result)
        return result

    # ── mutations ────────────────────────────────────────────────────
    def add(self, filename: str, content: bytes, name: str = "") -> dict:
        """Store an uploaded .docx and return its registry entry."""
        if not content:
            raise ValueError("The uploaded file is empty.")
        if not content.startswith(b"PK"):
            raise ValueError("That is not a .docx file (expected a Word document).")

        self.dir.mkdir(parents=True, exist_ok=True)
        display = _safe_stem(name or Path(filename).stem)
        target = self.dir / f"{_safe_stem(Path(filename).stem)}.docx"
        counter = 2
        while target.exists():
            target = self.dir / f"{_safe_stem(Path(filename).stem)} ({counter}).docx"
            counter += 1
        target.write_bytes(content)

        data = self.list()
        entry = next((i for i in data["items"] if i["file"] == target.name), None)
        if entry is None:                       # should not happen, but be safe
            entry = {"id": uuid.uuid4().hex[:12], "file": target.name}
            data["items"].insert(0, entry)
        entry["name"] = display
        entry["added"] = datetime.now().isoformat(timespec="seconds")
        data["active"] = entry["id"]
        data["items"].sort(key=lambda i: i.get("added", ""), reverse=True)
        self._write(data)
        return entry

    def rename(self, template_id: str, name: str) -> bool:
        data = self.list()
        for item in data["items"]:
            if item["id"] == template_id:
                item["name"] = _safe_stem(name, fallback=item.get("name", "template"))
                self._write(data)
                return True
        return False

    def activate(self, template_id: str) -> bool:
        data = self.list()
        if any(i["id"] == template_id for i in data["items"]):
            data["active"] = template_id
            self._write(data)
            return True
        return False

    def delete(self, template_id: str) -> bool:
        data = self.list()
        entry = next((i for i in data["items"] if i["id"] == template_id), None)
        if entry is None:
            return False
        path = self.dir / entry["file"]
        try:
            if path.exists():
                path.unlink()
        except OSError:
            return False
        preview = self.preview_path(entry)
        try:
            if preview.exists():
                preview.unlink()
        except OSError:
            pass
        data["items"] = [i for i in data["items"] if i["id"] != template_id]
        if data["active"] == template_id:
            data["active"] = data["items"][0]["id"] if data["items"] else ""
        self._write(data)
        return True

    # ── lookup ───────────────────────────────────────────────────────
    def get(self, template_id: str) -> dict | None:
        return next((i for i in self.list()["items"] if i["id"] == template_id), None)

    def active_path(self) -> Path | None:
        data = self.list()
        entry = next((i for i in data["items"] if i["id"] == data["active"]), None)
        if entry is None:
            return None
        path = self.dir / entry["file"]
        return path if path.exists() else None

    def path_of(self, entry: dict) -> Path:
        return self.dir / entry["file"]

    def preview_path(self, entry: dict) -> Path:
        """Where the converted PDF for *entry* is cached."""
        return self.preview_dir / f"{entry['id']}.pdf"
