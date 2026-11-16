# Rebate Form Generator

A dark-mode **local web app** for the Input Device rebate process. It reads the
supplier Master Price Table workbooks, consolidates them for one fiscal year,
and produces the per-supplier Word contract update forms — delivered as a single
zip download.

Everything runs on `127.0.0.1`. No data leaves the machine.

```
start.bat  →  local server on 127.0.0.1:8765  →  Chrome app window
```

---

## Quick start

**Double-click `start.bat`.**

It starts the server and opens the UI in a Chrome app-mode window with its own
profile. Closing that window stops the server; launching again while a window is
open reuses it rather than starting a second server.

From a shell instead:

```bash
poetry run rebate-form-generator                      # Chrome app window (default)
poetry run rebate-form-generator --browser default    # system default browser
poetry run rebate-form-generator --browser none --port 9000
```

| Flag | Default | Description |
|------|---------|-------------|
| `--port` | `8765` | Preferred port; the next free one is used if taken |
| `--browser` | `chrome` | `chrome` \| `default` \| `none` |

### Requirements

- Python ≥ 3.10
- [Poetry](https://python-poetry.org/) ≥ 1.8 (for development)
- Google Chrome — optional, falls back to the default browser
- Microsoft Word — optional, only for rendering the template preview

No web framework is needed: the server is built on the Python standard library
(`http.server`), so `openpyxl` and `python-docx` are the only runtime
dependencies.

```bash
git clone https://github.com/max84517/rebate-form-generator.git
cd rebate-form-generator
poetry install
```

---

## The pipeline

| Stage | Description |
|-------|-------------|
| 1 – Ingest | Reads the latest `.xlsx` from each supplier folder, applies the GTK Suppliers fix, combines NB bNB + cNB into one workbook per supplier |
| 2 – Segment | Consolidates all suppliers per segment (bNB, cNB, DT, Peripheral) |
| 3 – All | Merges the four segment files into one workbook |
| 4 – Rebate Only | Strips HP Cost / ODM Cost columns |
| 5 – Rebate Raw | Removes blank Platforms/Project rows, writes `rebate raw.xlsx` |
| 6 – Form Data | Filters to a chosen quarter, expands price-change rows, deduplicates, writes `contract input - <Supplier>.xlsx` |
| 7 – Report | Fills the Word template per supplier and packages the results as a zip |

**Stages 1–5 run for one fiscal year only.** The FY is chosen *before* any
processing starts, so a build for FY27 never reads, processes or reports on an
FY26 sheet.

---

## Using it

### 1. Source folders

Set the three paths (NB KB, DT KB, Peripheral) — type them, or click **Browse**
for the built-in folder picker. Paths are saved to `config.json` as you type.

Each one must be the folder that **contains** a `Master price table_<Segment>`
sub-folder. A live check under each field confirms what was found:

```
NB KB      [ C:/…/05. Cost/1 NB KB                        ] [ Browse ]
Notebook     ✓ Master price table_NB → 6 supplier(s):
keyboards      CHICONY, DARFON, LITEON, PRIMAX, SUNREX, TRANSIMAGE
```

### 2. File Tracker

The **File Tracker** button in the top bar lists the newest workbook in every
supplier folder with its date, and flags anything not updated in the current
month — a green tick or an orange warning per supplier, out-of-date first.
**Refresh** re-scans at any time.

### 3. Consolidate Rebate Data

Click it and the source folders are **scanned** — sheet names only, a second or
two — then the FY dialog opens showing which FY each supplier has. Pick the
target FY and click **Generate**: stages 1–5 run for that FY and write
`rebate raw.xlsx`.

### 4. Generate Form Data

Pick the **FY + Quarter** and tick which **feature columns** to include, then
**Generate**. Writes `contract input - <Supplier>.xlsx` per supplier.

The selection is remembered in `config.json` and pre-filled next time.

#### Quarter rules

| Quarter | Months |
|---------|--------|
| Q1 | Nov (prev year), Dec (prev year), Jan |
| Q2 | Feb, Mar, Apr |
| Q3 | May, Jun, Jul |
| Q4 | Aug, Sep, Oct |

Example: **FY26 Q1** covers Nov 2025, Dec 2025, Jan 2026.

#### Output columns

Selected feature columns + **GTK Suppliers** (always included) +
**Per-Unit Rebate Amount $USD** (currency-formatted) +
**Rebate Period Start Date**.

Default checked: `Segment`, `Color`, `HP/ODM Part#`, `Platforms/Project`,
`Product`, `Size`, `ODM (Regional Site)`.

Each source row produces 1–3 output rows depending on whether the rebate price
changes month-to-month within the quarter. Prices are rounded to 2 decimal
places before deduplication, and duplicate rows are dropped. A blank rebate cell
is written as `$0.00`; a cell holding an Excel error drops the whole row and is
reported.

### 5. Generate Report

Select the suppliers, enter **iCertis Code** and **Form #** for each, then
**Generate**. Both values are saved per-supplier and pre-filled next time.

The contracts are built in a temporary folder, packaged as
`Rebate Agreement Update Forms_FY##Q#_<YYYYMMDD>.zip` and **downloaded by the
browser** to your Downloads folder. Nothing is written to `data/output`.

### Run All

Runs all three stages in sequence. Each stage still opens its own dialog so you
can review before proceeding; closing a dialog cancels the rest of the chain.

---

## Settings

The gear in the top right opens **Contract Info**, **Contract Template** and
**Exit**.

### Contract Info

Everything substituted into the Word template, in two parts.

**Shared values** — the HP-side fields that are identical on every contract: the
internal signer and their title.

A shared value is **required** when the template still shows the keyword itself
where the value belongs. If the signature block reads `ICMInternalSignatory`,
somebody has to say who signs: Generate Report stops before producing anything
and names the value to set, and the Report dialog checks the same thing when it
opens. Where the template already spells a value out — HP's legal entity is
written as `HP Inc.` — nothing is demanded and the template's own text is kept.

**Per-supplier columns** — a table: add a supplier, change a value, delete a row.
**Each column name is a placeholder keyword**, so a new keyword only needs a new
column: use **+ Add column** and type the exact keyword from the template.
Columns can be renamed and removed, and the table scrolls sideways once there
are more columns than fit.

`ICMAgreementCode`, `ICMAgreementNumber1` and `ICMEffectiveDate` are refused as
columns — they come from the Generate Report dialog and the selected quarter.

Stored in `data/contract_info.json`. The first time it opens, the values are
imported from `supplier info/Contract Source info.xlsx`; after that the
spreadsheet is no longer read.

### Contract Template

Upload a `.docx`, give it a name, and see when it was added. One template is
marked **In use** and that is the one Generate Report fills in; the rest are
kept so you can switch back, and any can be deleted.

Below the list is a preview: first a check of every placeholder the current
settings will substitute — instant, and it follows your columns and shared
values, so a column with no matching placeholder in the template shows up as
missing — then the document itself rendered through the installed Word. The
first render takes roughly fifteen seconds and is cached, so reopening is
immediate. **Expand** hides the list so the document fills the window. Without
Word on the machine the placeholder check is still shown.

### Exit, and idle shutdown

**Exit** stops the server and closes the window. The app also closes itself
after **30 minutes without interaction**, warning a minute beforehand; a running
job postpones it.

---

## Template placeholders

Keywords are plain text — no brackets. Placeholders inside Word Content Controls
(`w:sdt`) are handled, and are filled **by the keyword the control is bound to**,
not only by matching its visible text. That matters for fields the template
already shows a real value for: the internal signer and HP legal entity are
literal text inside their control, so text matching alone would never replace
them.

| Placeholder | Value |
|-------------|-------|
| `ICMAgreementCode` | iCertis Code, entered per supplier |
| `ICMAgreementNumber1` | Form #, entered per supplier |
| `ICMEffectiveDate` | First day of the selected quarter, e.g. FY26 Q3 → `May 1, 2026` |
| `ICMInternalSignatory` | Shared value — internal signer |
| `ICMHPSignatoryTitle` | Shared value — internal signer title |
| `ICMRebateAgreementCode` | Per-supplier column |
| `CW#` | Per-supplier column |
| `ICMPartyName1` | Per-supplier column |
| `ICMExternalSignatory` | Per-supplier column |
| `ICMExternalSignatoryTitle` | Per-supplier column |
| *any other column* | Its value for that supplier |

### Product table formatting

| Column | Format |
|--------|--------|
| `Size` | Integer string, e.g. `14` (not `14.00`) |
| `Per-Unit Rebate Amount $USD` | Currency with `$`, e.g. `$0.40` |
| Date columns | `YYYY-MM-DD` |
| Other numbers | `1,234.56` |

---

## Source file validation

Supplier workbooks are hand-maintained, so the pipeline never trusts their
layout:

- The **header row is located per sheet** by finding the row that starts with
  `Segment`, scanning up to 15 rows down. Rows inserted above the header cost
  nothing.
- **Columns are matched by name, never by position.** The workbooks differ —
  50 to 68 columns in practice — and a positional stack would silently file one
  supplier's numbers under another's headings.
- **Month names are folded**: `Rebate June 2027` and `Rebate Jun 2027` are the
  same column.

Anything the pipeline cannot handle on its own is reported against the file that
caused it, in the log and in a **Source File Problems** pop-up with the full
path.

| Reported | Why it matters |
|---|---|
| **Header row not found** | No row starting with `Segment` in the first 15 rows — the sheet cannot be used at all |
| **File could not be read** | A workbook is open in Excel. That supplier is left out; the rest of the run continues |
| **Wrong months for this FY** | An `FY27` sheet whose rebate columns are FY26 months — copied from last year without updating the headers |
| **Missing month column** | The sheet lacks one or more of this FY's 12 rebate months |
| **Excel error value** | `#VALUE!`, `#REF!`, `#N/A` … in a rebate cell. The affected row is dropped, never written as a number |
| **Supplier skipped** | No usable rebate value for any month of the quarter, so no contract input is produced for them |

Handled silently, as a log line only: a shifted header row, a duplicate header
row inside the data, and month names written in full. Nothing is lost in any of
those cases, so there is nothing to act on.

---

## Folder conventions

Each source folder must contain a segment folder, which in turn holds one folder
per supplier:

```
<source folder>/
└── Master price table_<Segment>/
    ├── Master price table_<Segment>_CHICONY/
    │   └── <newest .xlsx wins>
    └── Master price table_<Segment>_DARFON/
```

e.g. `Master price table_NB_CHICONY`, `Master price table_DT_PRIMAX`.
Folders that do not match the pattern are skipped with a warning.

### Data layout

```
data/
├── contract_info.json            ← supplier contract details (Settings; git-ignored)
├── source data/                  ← ingested supplier workbooks, one FY at a time
├── rebate raw/rebate raw.xlsx    ← stage 5 output
├── rebate form input/            ← stage 6 output, one file per supplier
├── supplier info/                ← legacy; read once to seed contract_info.json
└── template/
    ├── templates.json            ← names, dates and which one is in use
    ├── .preview/                 ← cached PDF renders
    └── *.docx                    ← uploaded Word templates
```

Stage 7 output is not stored on disk: the contracts are zipped in a temporary
folder and handed to the browser as a download.

### Configuration

`config.json` (project root, git-ignored):

| Key | Description |
|-----|-------------|
| `nb_kb`, `dt_kb`, `peripheral` | Source folder paths |
| `last_fy` / `last_quarter` | Last selection; pre-fills the dialogs and sets `ICMEffectiveDate` |
| `form_numbers` | Per-supplier Form # |
| `icertis_codes` | Per-supplier iCertis Code |

On load the page writes a one-line self-report to `data/last-client.json` (local
only) recording which dialogs were visible and whether the stylesheet applied —
there to make a bad launch diagnosable after the fact.

---

## Project structure

```
start.bat                       ← launcher: starts the server, opens Chrome
src/rebate_form_generator/
├── main.py                     ← entry point, single-instance check, Chrome launcher
├── web/
│   ├── server.py               ← http.server routes → pipeline calls
│   ├── jobs.py                 ← single-slot job runner + SSE event hub
│   ├── docx_preview.py         ← placeholder check + Word-driven PDF render
│   └── static/                 ← index.html, style.css, app.js
├── config/
│   ├── settings.py             ← config.json
│   ├── contract_info.py        ← contract details + shared values store
│   └── templates.py            ← Word template registry
└── consolidation/
    ├── pipeline.py             ← public API: scan → build → form data → report
    ├── headers.py              ← header location, name folding, safe filenames
    ├── issues.py               ← structured source-file problems for the pop-up
    ├── stage1_ingest.py … stage7_report.py
```

The web layer only translates HTTP into pipeline calls; all business logic stays
in `consolidation/`. The UI talks to it over JSON, and the log streams back over
Server-Sent Events, so reloading the page replays the current run instead of
losing it.

Asset URLs are stamped with the server's start id, so a restarted server can
never be paired with a script or stylesheet the browser kept from an earlier run.

## Packaging

`build_exe.bat` and `RebateFormGenerator.spec` still target the old
CustomTkinter desktop build and **do not work with the web UI** — the `.spec`
needs `web/static/` added to its `datas`. Use `start.bat`.
