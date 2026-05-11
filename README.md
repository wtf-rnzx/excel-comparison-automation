# Excel/CSV Comparison Automation Tool

A Python web application built with Streamlit for comparing and validating student or employee records across two Excel or CSV files. Designed for academic institutions managing data across multiple systems, it normalizes values, detects duplicates, and generates a color-coded Excel report.

---

## Table of Contents

- [Project Overview](#project-overview)
- [Key Features](#key-features)
- [Technology Stack](#technology-stack)
- [Project Structure](#project-structure)
- [File and Folder Descriptions](#file-and-folder-descriptions)
- [How the System Works](#how-the-system-works)
- [JSON Local Storage](#json-local-storage)
- [Installation and Setup](#installation-and-setup)
- [Running the Application](#running-the-application)
- [Maintenance Notes](#maintenance-notes)

---

## Project Overview

This tool accepts a **Source file** (the authoritative record) and a **Target file** (the file to validate), then cross-references them by student or employee full name. For every matched record, it validates whether the **campus**, **program**, and **college** values are consistent between both files — accounting for known abbreviation differences through configurable mappings.

Results are presented in-browser with searchable tables and downloadable as a multi-sheet, color-coded Excel report.

---

## Key Features

- Upload two Excel or CSV files and compare them by full name
- Auto-detect columns for fullname, campus, program, and college
- Bidirectional name lookup (Target → Source and Source → Target)
- Duplicate detection in both files
- Campus, program, and college validation with status (`MATCH`, `MISMATCH`, `MISSING`, `REVIEW`)
- Configurable value mappings (e.g., `PABLO BORBON` → `MAIN`) stored persistently in JSON
- Preset equivalency templates for common campus and college abbreviations
- Downloadable multi-sheet Excel report with color-coded cells
- Standalone CLI mode via `automate.py`

---

## Technology Stack

| Technology | Purpose |
|---|---|
| **Python 3.x** | Core application language |
| **Streamlit** | Web-based user interface |
| **Pandas** | File reading, data manipulation, and analysis |
| **OpenPyXL** | Excel report generation with formatting and color coding |
| **JSON** | Local persistent storage for user-defined value mappings |

---

## Project Structure

```
excel-comparison-automation/
├── streamlit_app.py          # Main UI entry point
├── automate.py               # Core comparison engine and report generator
├── mappings.json             # Persistent user-defined value mappings
├── .gitignore
└── app/                      # Application package
    ├── __init__.py
    ├── config/
    │   ├── __init__.py
    │   └── settings.py       # Configuration constants and defaults
    ├── services/
    │   ├── __init__.py
    │   └── comparison.py     # Orchestration service for Streamlit
    ├── state/
    │   ├── __init__.py
    │   └── session.py        # Streamlit session state initialization
    ├── storage/
    │   ├── __init__.py
    │   └── local_json.py     # JSON read/write persistence layer
    └── ui/
        ├── __init__.py
        ├── mapping.py        # Mapping configuration UI components
        └── results.py        # Results display and visualization components
```

---

## File and Folder Descriptions

### Root-Level Files

| File | Description |
|---|---|
| `streamlit_app.py` | Application entry point. Renders the full Streamlit UI: file uploaders, sheet/column inputs, mapping configuration, and the Run Comparison button. Calls into `app/` modules for all logic. |
| `automate.py` | Self-contained comparison engine. Contains the full comparison pipeline, normalization functions, duplicate detection, validation logic, and Excel report writer. Can also be run directly from the command line. |
| `mappings.json` | Persistent local storage for user-defined value equivalencies (campus, program, college). Updated automatically when mappings are added or changed in the UI. |
| `.gitignore` | Excludes `data/`, `output/`, `__pycache__`, and `*.md` from version control. |

### `app/config/`

| File | Description |
|---|---|
| `settings.py` | Defines application-wide constants: accepted file types, default sheet input, default mapping state, preset mapping templates, and UI section metadata (labels, keys, icons). |

### `app/services/`

| File | Description |
|---|---|
| `comparison.py` | Bridges the Streamlit UI and `automate.py`. Saves uploaded files to a temp directory, calls `compare_files()`, generates the Excel report bytes, and stores results in session state. Handles errors with user-facing messages. |

### `app/state/`

| File | Description |
|---|---|
| `session.py` | Initializes Streamlit session state on first load. Seeds mapping data from `mappings.json` and sets default values for comparison result, report bytes, and filenames. |

### `app/storage/`

| File | Description |
|---|---|
| `local_json.py` | `LocalJSONStorage` class that wraps all reads and writes to `mappings.json`. Provides `load_data()`, `save_data()`, `add_record()`, `update_record()`, `delete_record()`, and `replace_records()` methods. Falls back to defaults if the file is missing or malformed. |

### `app/ui/`

| File | Description |
|---|---|
| `mapping.py` | Renders the tabbed mapping configuration interface. Lets users add, edit, delete, and bulk-load preset mappings for campus, program, and college. Changes are immediately persisted to `mappings.json` via `LocalJSONStorage`. |
| `results.py` | Renders the post-comparison results section. Displays summary metrics, data validation table, duplicate records, and dual-tab name comparison views with search and status filtering. Includes the Excel download button. |

---

## How the System Works

### Step 1 — Upload Files

The user uploads a **Source file** and a **Target file** through the Streamlit sidebar. Supported formats: `.csv`, `.txt`, `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.ods`.

Each file requires a sheet name or index (defaults to the first sheet, index `0`). Column names for fullname, campus, program, and college can be specified manually or auto-detected.

**Auto-detected column names:**

| Column | Recognized Headers |
|---|---|
| Full Name | `fullname`, `full name`, `FULLNAME`, `FULL NAME` |
| Campus | `campus`, `CAMPUS` |
| Program (Source) | `program_name`, `program name` |
| Program (Target) | `program`, `program_name`, `program name` |
| College (Source) | `college`, `COLLEGE` |
| College (Target) | `undergraduate_college`, `graduate_college`, `college` |

### Step 2 — Configure Mappings (Optional)

Before running a comparison, users can define value equivalencies. For example, if the Source file uses `PABLO BORBON` but the Target uses `MAIN`, adding this mapping tells the tool to treat them as equivalent during validation.

Mappings are organized into three categories: **Campus**, **Program**, and **College**. Preset templates are available for common abbreviation differences.

### Step 3 — Run Comparison

Clicking **Run Comparison** triggers the following pipeline in `automate.py`:

1. **Read files** — Both files are loaded into Pandas DataFrames.
2. **Normalize names** — Full names are uppercased and whitespace-trimmed.
3. **Detect duplicates** — Duplicate full names are flagged in each file.
4. **Bidirectional name matching** — Each name in Target is checked against Source (and vice versa), producing `FOUND` / `NOT FOUND` status.
5. **Data validation** — For every matched name, campus, program, and college values are compared using normalization rules and the configured mappings.

**Validation statuses:**

| Status | Meaning |
|---|---|
| `MATCH` | Values are identical or mapped as equivalent |
| `MISMATCH` | Values differ and no equivalency mapping covers them |
| `MISSING` | One or both sides have no value for this field |
| `REVIEW` | Match found but some fields are unresolvable |

### Step 4 — Review Results

Results appear in the browser with:

- **Summary metrics** — Total rows, matched names, unmatched names, duplicates
- **Data Validation table** — Per-record campus/program/college status
- **Duplicate records** — Combined list from both files
- **Name comparison tabs** — Target → Source and Source → Target views with search and status filter

### Step 5 — Download Report

The user downloads a timestamped `.xlsx` file containing five sheets:

| Sheet | Contents |
|---|---|
| `Summary` | Dashboard with KPI cards and validation metric counts |
| `Target_to_Source` | All target names with `FOUND` / `NOT FOUND` status, color-coded |
| `Source_to_Target` | All source names with `FOUND` / `NOT FOUND` status, color-coded |
| `Duplicates` | All duplicate full names from both files with occurrence counts |
| `Data_Validation` | Matched names with campus, program, and college validation status, color-coded |

---

## JSON Local Storage

The application uses a single JSON file, `mappings.json`, as its local database for user-configured value mappings.

### File Location

```
excel-comparison-automation/mappings.json
```

### Structure

```json
{
  "campus_mappings": [
    { "source": "PABLO BORBON", "target": "MAIN" },
    { "source": "JPLPC-MALVAR",  "target": "MALVAR" },
    { "source": "ARASOF-NASUGBU", "target": "NASUGBU" }
  ],
  "program_mappings": [],
  "college_mappings": []
}
```

Each entry is a `source` / `target` pair. During validation, if either the source-file value or target-file value matches the `source` key of a mapping, the tool treats both sides as equivalent.

### Behavior

- **On startup** — `LocalJSONStorage.load_data()` reads `mappings.json`. If the file is absent or invalid, it falls back to the defaults defined in `app/config/settings.py`.
- **On change** — Any add, edit, or delete action in the Mapping UI immediately calls `save_data()` to persist the update.
- **Preset templates** — The UI offers a "Load Presets" option that populates campus and college mappings from `DEFAULT_MAPPING_TEMPLATES` in `settings.py`.

---

## Installation and Setup

### Prerequisites

- Python 3.9 or higher
- `pip` package manager

### Install Dependencies

```bash
pip install streamlit pandas openpyxl
```

Or, if a `requirements.txt` file is present:

```bash
pip install -r requirements.txt
```

### Clone or Download the Project

```bash
git clone <repository-url>
cd excel-comparison-automation
```

---

## Running the Application

### Streamlit Web App

```bash
streamlit run streamlit_app.py
```

The app opens in your browser at `http://localhost:8501`.

### Command-Line (Standalone Mode)

`automate.py` can be run directly without Streamlit:

```bash
python automate.py --source path/to/source.xlsx --target path/to/target.xlsx
```

Run `python automate.py --help` for all available arguments, including options to specify sheet names, column names, and output path.

---

## Maintenance Notes

### Adding New Column Auto-Detection Rules

Column name patterns are detected in `automate.py` inside the `resolve_column()` function. Add new recognized header variants to the relevant list within that function.

### Adding New Preset Mapping Templates

Edit `app/config/settings.py` and update the `DEFAULT_MAPPING_TEMPLATES` dictionary. These appear in the UI under "Load Presets" for each mapping category.

### Adding Built-in Equivalency Groups

For campus, program, and college equivalencies that should always apply (without user configuration), edit the preset equivalence groups inside the following functions in `automate.py`:

- `status_from_sets()` — Campus equivalencies
- `status_from_sets_program()` — Program equivalencies
- `status_from_sets_college()` — College equivalencies

### Resetting Mappings

Delete or clear the contents of `mappings.json`. On the next application start, defaults from `app/config/settings.py` will be restored automatically.

### Supported File Formats

Accepted upload types are defined in `app/config/settings.py` under `SUPPORTED_UPLOAD_TYPES`. Add or remove extensions there to change what the file uploader accepts.

### Output Files

The Excel report is generated in-memory and delivered as a browser download. No files are written to disk by default during Streamlit usage. The CLI mode writes the report to the path specified by `--output`.
