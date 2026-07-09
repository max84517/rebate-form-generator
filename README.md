# Rebate Form Generator

A dark-mode desktop tool that reads supplier Master Price Table workbooks, consolidates them into a **Rebate Raw** workbook, generates **Form Data input**, and produces per-supplier **Word contract update forms** for the Input Device rebate process.

## Features

| Stage | Description |
|-------|-------------|
| 1 – Ingest | Reads the latest `.xlsx` from each supplier folder, applies GTK Suppliers fix, combines NB bNB + cNB into one workbook per supplier (`FY## bNB` / `FY## cNB` sheets) |
| 2 – Segment | Consolidates all suppliers per segment (bNB, cNB, DT, Peripheral) |
| 3 – All | Merges the four segment files into one workbook |
| 4 – Rebate Only | Strips HP Cost / ODM Cost columns |
| 5 – Rebate Raw | Filters to a chosen FY sheet, removes blank Platforms/Project rows, outputs `rebate raw.xlsx` |
| 6 – Form Data | Filters to a chosen FY + quarter, expands price-change rows, deduplicates, outputs per-supplier `contract input - <Supplier>.xlsx` |
| 7 – Report | Fills a Word template with supplier info and product rebate table, outputs per-supplier `.docx` contracts |

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
2. Click **Consolidate Rebate Data** — stages 1–4 run; on completion a FY selection popup appears.
3. The popup shows per-supplier FY coverage. Select the target FY and click **Generate** — stage 5 runs and writes `rebate raw.xlsx`. Output is saved to `data/output` by default.

### Generate Form Data (Stage 6)

1. After consolidation, click **Generate Form Data**.
2. A popup lets you:
   - Select the **FY + Quarter** (quarter rules below; defaults to the current quarter)
   - Choose which **feature columns** to include (7 pre-selected by default)
3. Click **Generate** — outputs `rebate form input/contract input - <Supplier>.xlsx` per supplier.

The selected FY + Quarter is remembered in `config.json`, so the next time you open **Generate Form Data** the dropdown defaults to the same FY's current quarter automatically.

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

Default checked columns: `Segment`, `Color`, `HP/ODM Part#`, `Platforms/Project`, `Product`, `Size`, `ODM (Regional Site)`

Each source row produces 1–3 output rows depending on whether the rebate price changes month-to-month within the quarter. Prices are rounded to 2 decimal places before deduplication. Duplicate rows are dropped automatically.

> **Note:** Columns where every row contains the same value are automatically removed from each supplier’s output file.

### Generate Report (Stage 7)

1. Place the Word template (`Rebate Agreement Update Form_*.docx`) in the `template/` folder under your output parent directory.
2. Place the supplier info Excel (`Contract Source info.xlsx`) in the `supplier info/` folder.
3. Click **Generate Report**, select the suppliers and enter the Form # for each, then click **Generate**.
   - Form numbers are saved to `config.json` and pre-filled automatically on the next launch.
4. The tool fills keyword placeholders in the template (`<Supplier Name>`, `<Contract Number>`, `<Version>`, `<Name of Entity>`, `<Address>`, `<Signer>`, `<Title>`, `<SUPPLIER-Sign>`, `<Effective Date>`), inserts product rebate data into the table, and saves one `.docx` per supplier.

| Placeholder | Value |
|-------------|-------|
| `<Supplier Name>` | From supplier info Excel |
| `<Contract Number>` | From supplier info Excel |
| `<Version>` | Form # entered by user (saved per-supplier in `config.json`) |
| `<Name of Entity>` | From supplier info Excel |
| `<Address>` | From supplier info Excel |
| `<Signer>` | From supplier info Excel |
| `<Title>` | From supplier info Excel |
| `<SUPPLIER-Sign>` | From supplier info Excel (bold) |
| `<Effective Date>` | First month of the selected quarter, e.g. FY26 Q3 → `May 2026` (auto-filled) |

> **Template note:** The footer's page-number field (`PAGE`) must be a real Word field (not static text). Keyword placeholders in the footer are replaced using run-by-run substitution to preserve the field structure.

#### Product table formatting

| Column | Format |
|--------|--------|
| `Size` | Integer string, e.g. `14` (not `14.00`) |
| `Per-Unit Rebate Amount $USD` | Currency with `$`, e.g. `$0.40` |
| Date columns | `YYYY-MM-DD` |
| Other numbers | `1,234.56` |

### Run All

Click **Run All** (green button, far right) to execute all three stages in sequence. Each stage still opens its own confirmation dialog so you can review settings before proceeding.

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
│   │   └── <Supplier>.xlsx ← one workbook per supplier; sheets: "FY## bNB", "FY## cNB"
│   ├── DT/
│   └── Peripheral/
├── rebate raw/
│   └── rebate raw.xlsx           ← stage 5 output
├── rebate form input/
│   └── contract input - <Supplier>.xlsx  ← stage 6 output (one per supplier)
├── supplier info/
│   └── Contract Source info.xlsx ← user-managed; required for stage 7
├── template/
│   └── Rebate Agreement Update Form_*.docx  ← user-managed Word template
└── output/
    └── <YYYYMMDD>/
        └── Rebate Agreement Update Form#<N>_<Supplier>.docx  ← stage 7 output
```

## Configuration

Settings are stored in `config.json` (project root, git-ignored):

| Key | Description |
|-----|-------------|
| `nb_kb` | Path to NB KB source folder |
| `dt_kb` | Path to DT KB source folder |
| `peripheral` | Path to Peripheral source folder |
| `last_fy` | Last FY selected (e.g. `FY26`) |
| `last_quarter` | Last quarter selected (1–4); used to compute `<Effective Date>` in Stage 7 |
| `form_numbers` | Per-supplier Form # map (e.g. `{"CHICONY": "9"}`) |

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
    ├── stage6_rebate_form.py
    └── stage7_report.py      ← Word contract generation
```
