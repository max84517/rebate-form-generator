# Rebate Form Generator

A dark-mode desktop tool that reads supplier Master Price Table workbooks and produces a consolidated **Rebate Raw** workbook ready for the Input Device rebate contract process.

## Features

| Stage | Description |
|-------|-------------|
| 1 – Ingest | Reads the latest `.xlsx` from each supplier folder, applies GTK Suppliers fix, splits NB into bNB / cNB |
| 2 – Segment | Consolidates all suppliers per segment (bNB, cNB, DT, Peripheral) |
| 3 – All | Merges the four segment files into one workbook |
| 4 – Rebate Only | Strips HP Cost / ODM Cost columns |
| 5 – Rebate Raw | Filters to a chosen FY sheet, removes blank Platforms/Project rows, outputs `rebate raw.xlsx` |

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

1. Set the three **Source Folder** paths (NB KB, DT KB, Peripheral).
2. Set an **Output Path** (defaults to `data/output`).
3. Click **Consolidate Rebate Data** — stages 1–4 run; on completion a FY selection popup appears.
4. The popup shows which suppliers have data for each FY. Select the target FY and click **Generate** — the dialog closes and stage 5 runs in the background; the output path is logged in the main window.

## Source folder naming convention

Each source folder must contain sub-folders named:

```
Master price table_<Segment>_<Supplier>
```

e.g. `Master price table_NB_CHICONY`, `Master price table_DT_PRIMAX`

## Output layout

```
<output path>/
├── source data/
│   ├── NB/
│   │   ├── bNB/       ← processed bNB workbooks per supplier
│   │   └── cNB/       ← processed cNB workbooks per supplier
│   ├── DT/
│   └── Peripheral/
├── rebate raw/
│   └── rebate raw.xlsx  ← final output (stage 5)
├── rebate form input/   ← user-managed input files
└── template/            ← Word / Excel templates
```

`source data` and `rebate raw` are regenerated on every run. Stage 2–4 intermediates are written to the system temp folder and cleaned up automatically.

## Configuration

Settings are stored in `config.json` (project root, git-ignored):

| Key | Description |
|-----|-------------|
| `nb_kb` | Path to NB KB source folder |
| `dt_kb` | Path to DT KB source folder |
| `peripheral` | Path to Peripheral source folder |
| `output_path` | Base output path |

## Project structure

```
src/rebate_form_generator/
├── main.py
├── ui/
│   └── main_window.py        ← CustomTkinter dark-mode UI + FY dialog
├── config/
│   └── settings.py           ← load / save config.json
└── consolidation/
    ├── pipeline.py            ← public API + module-level cache
    ├── stage1_ingest.py
    ├── stage2_segment.py
    ├── stage3_all.py
    ├── stage4_rebate.py
    └── stage5_template.py
```
