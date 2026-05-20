# Rebate Form Generator

A dark-mode desktop tool that reads supplier Master Price Table workbooks, consolidates them into a **Rebate Raw** workbook, and generates **Form Data input** for the Input Device rebate contract process.

## Features

| Stage | Description |
|-------|-------------|
| 1 – Ingest | Reads the latest `.xlsx` from each supplier folder, applies GTK Suppliers fix, splits NB into bNB / cNB |
| 2 – Segment | Consolidates all suppliers per segment (bNB, cNB, DT, Peripheral) |
| 3 – All | Merges the four segment files into one workbook |
| 4 – Rebate Only | Strips HP Cost / ODM Cost columns |
| 5 – Rebate Raw | Filters to a chosen FY sheet, removes blank Platforms/Project rows, outputs `rebate raw.xlsx` |
| 6 – Form Data | Filters to a chosen FY + quarter, expands price-change rows, deduplicates, outputs `input.xlsx` |

## Requirements

- Python ≥ 3.10
- [Poetry](https://python-poetry.org/) ≥ 1.8

## Installation

```bash
git clone https://github.com/max84517/rebate-form-generator.git
cd rebate-form-generator
poetry install
```

## Usage

```bash
poetry run rebate-form-generator
```

### Consolidate Rebate Data (Stages 1–5)

1. Set the three **Source Folder** paths (NB KB, DT KB, Peripheral).
2. Set an **Output Path** (defaults to `data/output`).
3. Click **Consolidate Rebate Data** — stages 1–4 run; on completion a FY selection popup appears.
4. The popup shows per-supplier FY coverage. Select the target FY and click **Generate** — stage 5 runs and writes `rebate raw.xlsx`.

### Generate Form Data (Stage 6)

1. After consolidation, click **Generate Form Data**.
2. A popup lets you:
   - Select the **FY + Quarter** (quarter rules below; defaults to the current quarter)
   - Choose which **feature columns** to include (7 pre-selected by default)
3. Click **Generate** — the tool filters `rebate raw.xlsx` to the three months of the selected quarter, expands rows for price changes within the quarter, deduplicates, and writes `rebate form input/input.xlsx`.

The selected FY is remembered in `config.json` so the quarter dropdown shows only that FY's four quarters on the next launch.

#### Quarter rules

| Quarter | Months |
|---------|--------|
| Q1 | Nov (prev year), Dec (prev year), Jan |
| Q2 | Feb, Mar, Apr |
| Q3 | May, Jun, Jul |
| Q4 | Aug, Sep, Oct |

Example: **FY26 Q1** covers Nov 2025, Dec 2025, Jan 2026.

#### Output columns

Selected feature columns + **GTK Suppliers** (always included) + **Per-Unit Rebate Amount $USD** (currency-formatted) + **Rebate Period Start Date**

Each source row produces 1–3 output rows depending on whether the rebate price changes month-to-month within the quarter. Duplicate rows are dropped automatically.

## Source folder naming convention

Each source folder must contain sub-folders named:

```
Master price table_<Segment>_<Supplier>
```

e.g. `Master price table_NB_CHICONY`, `Master price table_DT_PRIMAX`

## Output layout

```
<output path parent>/
├── source data/
│   ├── NB/
│   │   ├── bNB/            ← processed bNB workbooks per supplier
│   │   └── cNB/            ← processed cNB workbooks per supplier
│   ├── DT/
│   └── Peripheral/
├── rebate raw/
│   └── rebate raw.xlsx     ← stage 5 output
├── rebate form input/
│   └── input.xlsx          ← stage 6 output
└── template/               ← Word / Excel templates (user-managed)
```

`source data`, `rebate raw`, and `rebate form input` are regenerated on each run. Stage 2–4 intermediates go to the system temp folder and are cleaned up automatically.

## Configuration

Settings are stored in `config.json` (project root, git-ignored):

| Key | Description |
|-----|-------------|
| `nb_kb` | Path to NB KB source folder |
| `dt_kb` | Path to DT KB source folder |
| `peripheral` | Path to Peripheral source folder |
| `output_path` | Base output path |
| `last_fy` | Last FY selected (e.g. `FY26`); used to pre-fill the quarter dropdown |

## Project structure

```
src/rebate_form_generator/
├── main.py
├── ui/
│   └── main_window.py        ← CustomTkinter dark-mode UI + dialogs
├── config/
│   └── settings.py           ← load / save config.json
└── consolidation/
    ├── pipeline.py            ← public API + module-level cache
    ├── stage1_ingest.py
    ├── stage2_segment.py
    ├── stage3_all.py
    ├── stage4_rebate.py
    ├── stage5_template.py
    └── stage6_rebate_form.py
```
