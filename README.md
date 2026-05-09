# KPI Dictionary Generator

Python tool for generating a formal Thai/English KPI Dictionary document from Excel input. The output is a print-ready `.docx` with:

- cover page
- table of contents
- multiple search index sections
- one structured KPI dictionary page per KPI
- page number footer
- Thai-first layout with English subtitles
- validation report workbook

The layout is designed to resemble a hospital KPI manual / THIP-style reference document.

## Repository

GitHub repository:

`https://github.com/Burinboo256/siriraj-kpi-dictionary-builder.git`

Clone the project with:

```bash
git clone https://github.com/Burinboo256/siriraj-kpi-dictionary-builder.git
cd siriraj-kpi-dictionary-builder
```

## Folder Structure

```text
docx_kpi_dictionary_codex/
├── assets/
│   └── logo.png
├── input/
│   └── kpi_dictionary_template.example.xlsx
├── output/
│   ├── KPI_Dictionary.docx
│   ├── KPI_Dictionary.pdf
│   └── validation_report.xlsx
├── src/
│   ├── __init__.py
│   ├── generate_document.py
│   ├── generate_template.py
│   ├── styles.py
│   └── utils.py
├── temp/
├── tests/
│   ├── test_document.py
│   ├── test_template.py
│   └── test_utils.py
├── .gitignore
├── requirements.txt
└── README.md
```

## Excel Template

The repository keeps the example workbook as `input/kpi_dictionary_template.example.xlsx`.
Your working file should be `input/kpi_dictionary_template.xlsx`, which is ignored by Git.

Prepare a local working copy with:

```bash
cp input/kpi_dictionary_template.example.xlsx input/kpi_dictionary_template.xlsx
```

Or regenerate the working file from code:

```bash
python3 src/generate_template.py
```

The workbook structure includes:

- `01_Instruction`: how to complete the workbook
- `02_KPI_Input_Form`: the only sheet general users need to fill in
- `03_KPI_Input_Dictionary`: visible dictionary for every column in `02_KPI_Input_Form`
- `04_KPI_Master`: optional hidden admin metadata
- `05_KPI_Logic`: optional hidden analyst logic
- `06_KPI_CodeSet`: optional hidden structured ICD / procedure mapping
- `07_KPI_Reference`: optional hidden reference rows
- `08_KPI_Version`: optional hidden version metadata
- `09_KPI_Owner`: optional hidden owner details
- `10_Validation_List`: dropdown source values
- `11_Config`: document metadata such as title, subtitle, fonts, and logo path

The input form is aligned to the KPI detail page and keeps only the fields needed to generate that page directly:

- required highlighted fields for category, type, code, Thai/English names, definition, objective, formula, numerator, denominator, and related code sets
- optional non-highlighted fields for frequency, benchmark, interpretation, team/reference, data source, version dates, revision reason, and note
- a companion input-dictionary sheet that explains each column, whether it is required, a sample value, and any allowed dropdown values
- dropdown validation for category, type, and frequency on the user-facing sheet
- hidden advanced sheets for teams that want to maintain implementation logic without forcing all users to fill it

The generator also supports a simpler public Google Sheet shape like the shared example, as long as it exposes a `KPIs` sheet with flat KPI columns.

## Git Preparation

Recommended tracked files:

- `src/`
- `tests/`
- `assets/`
- `input/kpi_dictionary_template.example.xlsx`
- `requirements.txt`
- `README.md`

Recommended untracked files:

- `input/kpi_dictionary_template.xlsx`
- `output/` generated `.docx`, `.pdf`, and validation workbooks
- `temp/` downloaded Google Sheet exports
- `.venv/`
- Python cache files
- macOS Finder files

## Setup

This repository standard prefers `uv`. A typical setup is:

```bash
cd docx_kpi_dictionary_codex
uv venv
source .venv/bin/activate
uv pip install -r requirements.txt
```

If `uv` is unavailable, a standard Python virtual environment also works:

```bash
cd docx_kpi_dictionary_codex
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Quick Start

Create a local working workbook from the tracked example:

```bash
cp input/kpi_dictionary_template.example.xlsx input/kpi_dictionary_template.xlsx
```

Then generate the Word document:

```bash
python3 src/generate_document.py input/kpi_dictionary_template.xlsx --output output/KPI_Dictionary.docx --validation-report output/validation_report.xlsx
```

## Usage

Generate the Excel template:

```bash
python3 src/generate_template.py
```

This writes the local working file to `input/kpi_dictionary_template.xlsx`.

Generate the Word document:

```bash
python3 src/generate_document.py input/kpi_dictionary_template.xlsx --output output/KPI_Dictionary.docx --validation-report output/validation_report.xlsx
```

Generate from a public Google Sheet URL or sheet ID:

```bash
python3 src/generate_document.py --google-sheet "https://docs.google.com/spreadsheets/d/1s6LB0a_j4k-RhBTKAfZwQVDda25xzgmbGFDTLRC1LvM/edit?usp=sharing" --output output/KPI_Dictionary.docx --validation-report output/validation_report.xlsx
```

When using `--google-sheet`, the workbook is exported to `temp/` as `.xlsx` first and then processed through the same pipeline.

Attempt Word + PDF generation:

```bash
python3 src/generate_document.py input/kpi_dictionary_template.xlsx --output output/KPI_Dictionary.docx --validation-report output/validation_report.xlsx --pdf
```

## Generated Output

The Word document includes:

- cover page and TOC
- index by KPI code
- index by KPI category
- index by KPI type
- index by frequency
- index by reference team
- index by ICD / procedure code
- one KPI detail table per KPI aligned to the section reference template
- required detail rows for category, type, code, Thai/English names, definition, objective, formula, numerator, denominator, and related code sets
- optional detail rows for frequency, benchmark, interpretation, team/reference, effective date, last update, revision reason, and note
- footer with document version and page number

The validation workbook includes findings for:

- missing required fields
- duplicate KPI codes
- invalid dropdown values
- missing formula
- ICD format warning
- related KPI code not found

## Notes on Thai Language Support

- The document config defaults to `TH Sarabun New` for Thai and `Arial` for English.
- Thai labels and KPI names are rendered first.
- English KPI names are italicized beneath the Thai names.
- For best results, install the configured fonts on the machine that opens the `.docx`.

## Notes on Page Numbers and TOC

- Footer page numbers are added as Word fields.
- The footer also shows the document version.
- The table of contents is inserted as a Word TOC field.
- After opening the `.docx` in Microsoft Word, update fields so the TOC reflects the latest pagination.
- The search-index page numbers are generated from document order as a practical default.

## Testing

Run the test suite with:

```bash
python3 -m unittest discover -s tests -v
```

## Typical Commit Workflow

Run the main checks before commit:

```bash
python3 -m unittest discover -s tests -v
python3 src/generate_template.py
python3 src/generate_document.py input/kpi_dictionary_template.xlsx --output output/KPI_Dictionary.docx --validation-report output/validation_report.xlsx
```

## Implementation Details

- `src/generate_template.py`: builds the simplified KPI detail-first Excel template with required-field highlighting
- `src/utils.py`: resolves local or Google Sheet sources, joins normalized sheets, and generates validation findings
- `src/styles.py`: document formatting helpers for borders, shading, TOC, and page numbers
- `src/generate_document.py`: composes the final `.docx`, search indexes, and KPI detail tables

## Current PDF Export Behavior

PDF export is optional and depends on `docx2pdf`, which usually requires Microsoft Word on Windows or macOS. If PDF export is unavailable, the `.docx` output is still generated normally.
