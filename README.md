# DonyatiDocs-sow

A Claude Code plugin for generating SOW pricing Excel workbooks and converting them to professional Word documents with Donyati branding.

**Current Version:** donsow v1.0.0 | sowdoc v1.0.0

## Features

- **SOW Pricing** (`donsow`): Create 11-sheet Excel workbooks with resource allocations, timelines, and deliverables
- **SOW Documents** (`sowdoc`): Convert pricing workbooks to professional Word documents
- **Three Input Modes**: JSON input, conversational interview, or clone-from-existing
- **Brand Consistency**: Automatic Donyati colors and formatting

## Installation

### Option 1: Install from GitHub (Recommended)

```
/plugin install github:matt-janecek/DonyatiDocs-sow
```

### Option 2: Install from Local Path

```
/plugin install /path/to/DonyatiDocs-sow
```

### Dependencies

```bash
pip install openpyxl python-docx
```

## Quick Start

### Generate SOW Pricing Excel

```bash
python3 scripts/create_sow_workbook.py sow-data.json output.xlsx
```

### Convert to Word Document

```bash
python3 scripts/create_sow_document.py pricing.xlsx output.docx
```

### Full Pipeline

```bash
# Step 1: Create pricing workbook
python3 scripts/create_sow_workbook.py sow-data.json pricing.xlsx

# Step 2: Convert to Word document
python3 scripts/create_sow_document.py pricing.xlsx sow-document.docx
```

## Skills Reference

| Skill | Description |
|-------|-------------|
| `donsow` | SOW pricing Excel generation with 11 standardized sheets |
| `sowdoc` | SOW Excel to Word conversion with cover page template |

## Project Structure

```
DonyatiDocs-sow/
├── .claude-plugin/plugin.json
├── skills/
│   ├── donsow/SKILL.md
│   └── sowdoc/SKILL.md
├── scripts/
│   ├── create_sow_workbook.py
│   ├── create_sow_document.py
│   └── create_document.py
├── data/sow-reference-data.json
├── templates/
│   ├── excel/ieee-excel/          # Example pricing workbooks
│   └── word/
│       ├── donyati-word-template-2025.docx
│       ├── donyati-word-template-cover-2025.docx
│       └── sow/                   # Example SOW Word documents
├── CLAUDE.md
├── DEPENDENCIES.md
└── README.md
```

## Related Plugins

- **[DonyatiDocs](https://github.com/matt-janecek/DonyatiDocs)** - Core Donyati document generation (donppt, donword)
- **[DonyatiDocs-ieeeppt](https://github.com/matt-janecek/DonyatiDocs-ieeeppt)** - IEEE PowerPoint presentations (ieeeppt)

## License

MIT License
