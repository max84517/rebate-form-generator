# Rebate Form Generator

A dark-mode desktop tool that reads supplier Master Price Table workbooks and produces a consolidated **Pricing Template** ready for the Input Device rebate contract process.

## Features

| Stage | Description |
|-------|-------------|
| 1 – Ingest | Reads the latest `.xlsx` from each supplier folder, applies GTK Suppliers fix, splits NB into bNB / cNB |
| 2 – Segment | Consolidates all suppliers per segment (bNB, cNB, DT, Peripheral) |
| 3 – All | Merges the four segment files into one workbook |
| 4 – Rebate Only | Strips HP Cost / ODM Cost columns |
| 5 – Template | Filters to a chosen FY sheet, removes blank Platforms/Project rows, outputs `Pricing_Template_InputDevices.xlsx` |

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
3. Click **Build Pricing Data** — stages 1–4 run; on completion a popup appears.
4. Select the FY sheet and click **Write Pricing Template** — the dialog closes and stage 5 runs in the background; the result path is logged in the main window.

## Source folder naming convention

Each source folder must contain a sub-folder named:

```
Master price table_<Segment>_<Supplier>
```

e.g. `Master price table_NB_CHICONY`, `Master price table_DT_PRIMAX`

## Output layout

```
data/
├── output/            ← stage 2-4 intermediate files (tmp/)
├── pricing data/      ← Pricing_Template_InputDevices.xlsx
└── source data/
    ├── NB/
    │   ├── bNB/       ← processed bNB workbooks per supplier
    │   └── cNB/       ← processed cNB workbooks per supplier
    ├── DT/
    └── Peripheral/
```

`source data` and `pricing data` are regenerated on every build run.

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
│   └── main_window.py        ← CustomTkinter dark-mode UI
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
