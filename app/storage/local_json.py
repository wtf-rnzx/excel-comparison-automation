"""Reusable local JSON persistence helpers."""

from __future__ import annotations

from copy import deepcopy
import json
from pathlib import Path


class LocalJSONStorage:
    """Local JSON storage with safe defaults and graceful error handling."""

    def __init__(self, file_path: Path, default_data: dict[str, list[dict[str, str]]]):
        self.file_path = file_path
        self._default_data = deepcopy(default_data)

    def _make_default(self) -> dict[str, list[dict[str, str]]]:
        return deepcopy(self._default_data)

    def load_data(self) -> dict[str, list[dict[str, str]]]:
        """Load JSON data from disk, returning defaults if unreadable."""
        if not self.file_path.exists() or self.file_path.stat().st_size == 0:
            return self._make_default()

        try:
            data = json.loads(self.file_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return self._make_default()

        if not isinstance(data, dict):
            return self._make_default()

        normalized = self._make_default()
        for key in normalized:
            value = data.get(key, [])
            normalized[key] = value if isinstance(value, list) else []

        return normalized

    def save_data(self, data: dict[str, list[dict[str, str]]]) -> None:
        """Persist JSON data to disk."""
        self.file_path.parent.mkdir(parents=True, exist_ok=True)
        self.file_path.write_text(json.dumps(data, indent=2), encoding="utf-8")

    def add_record(self, category: str, source: str, target: str) -> list[dict[str, str]]:
        data = self.load_data()
        records = data.get(category, [])
        records.append({"source": source, "target": target})
        data[category] = records
        self.save_data(data)
        return records

    def update_record(self, category: str, index: int, field: str, value: str) -> list[dict[str, str]]:
        data = self.load_data()
        records = data.get(category, [])
        if 0 <= index < len(records):
            records[index][field] = value
            data[category] = records
            self.save_data(data)
        return records

    def delete_record(self, category: str, index: int) -> list[dict[str, str]]:
        data = self.load_data()
        records = data.get(category, [])
        if 0 <= index < len(records):
            records.pop(index)
            data[category] = records
            self.save_data(data)
        return records

    def replace_records(self, category: str, records: list[dict[str, str]]) -> list[dict[str, str]]:
        data = self.load_data()
        data[category] = records
        self.save_data(data)
        return records
