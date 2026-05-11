"""Application configuration values for the Streamlit UI."""

from __future__ import annotations

from typing import Final


SUPPORTED_UPLOAD_TYPES: Final[list[str]] = [
    "csv",
    "txt",
    "xlsx",
    "xls",
    "xlsm",
    "xlsb",
    "ods",
]

DEFAULT_SHEET_INPUT: Final[str] = "0"

DEFAULT_MAPPING_STATE: Final[dict[str, list[dict[str, str]]]] = {
    "campus_mappings": [],
    "program_mappings": [],
    "college_mappings": [],
}

DEFAULT_MAPPING_TEMPLATES: Final[dict[str, list[dict[str, str]]]] = {
    "campus": [
        {"source": "PABLO BORBON", "target": "MAIN"},
        {"source": "JPLPC-MALVAR", "target": "MALVAR"},
        {"source": "ARASOF-NASUGBU", "target": "NASUGBU"},
    ],
    "college": [
        {"source": "COE", "target": "CEAFA"},
        {"source": "CAFAD", "target": "CEAFA"},
        {"source": "CET", "target": "CIT"},
        {"source": "CHS", "target": "CONAHS"},
        {"source": "CABE", "target": "CABEIHM"},
        {"source": "CCJE", "target": "CAS"},
    ],
}

MAPPING_SECTION_META: Final[list[tuple[str, str, str, str]]] = [
    ("🏫 Campus", "campus", "Campus", "campus_mappings"),
    ("📚 Program", "program", "Program", "program_mappings"),
    ("🎓 College", "college", "College", "college_mappings"),
]
