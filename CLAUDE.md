# CLAUDE.md

This file provides guidance to Claude Code when working with the DonyatiDocs-sow plugin.

## Project Overview

DonyatiDocs-sow is a Claude Code plugin for generating SOW pricing Excel workbooks and converting them to professional Word documents with Donyati branding.

## Project Structure

```
DonyatiDocs-sow/
├── .claude-plugin/
│   └── plugin.json          # Plugin manifest
├── skills/
│   ├── donsow/
│   │   └── SKILL.md         # SOW pricing Excel skill
│   └── sowdoc/
│       └── SKILL.md         # SOW Excel to Word converter
├── scripts/
│   ├── create_sow_workbook.py      # SOW pricing Excel generator
│   ├── create_sow_document.py      # SOW Excel to Word converter
│   └── create_document.py          # Word document generator (donword engine)
├── data/
│   └── sow-reference-data.json     # SOW picklist/rate data
├── templates/
│   ├── excel/ieee-excel/           # Example SOW pricing sheets
│   └── word/
│       ├── donyati-word-template-2025.docx
│       ├── donyati-word-template-cover-2025.docx
│       └── sow/                    # Example SOW Word documents
├── CLAUDE.md
├── DEPENDENCIES.md
└── README.md
```

## Available Skills

### SOW Pricing Generation (`donsow`)

Creates SOW pricing Excel workbooks (11 sheets):
- **Workflow**: Create sow-data.json -> run create_sow_workbook.py
- **Modes**: JSON input, conversational interview, or clone-from-existing
- **Reference**: `data/sow-reference-data.json` (clients, practices, roles, rates)

### SOW Excel to Word (`sowdoc`)

Converts SOW pricing Excel workbooks to Word documents:
- **Workflow**: pricing.xlsx -> run create_sow_document.py -> output.docx
- **Uses**: donword template system with cover page
- **Custom**: Optional scope.json for custom scope/deliverables content

## Commands

```bash
SCRIPTS="/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/scripts"

# SOW Pricing Excel
python3 "$SCRIPTS/create_sow_workbook.py" sow-data.json output.xlsx

# SOW Excel to Word
python3 "$SCRIPTS/create_sow_document.py" pricing.xlsx output.docx
```

## Dependencies

See `DEPENDENCIES.md` for required Python packages.

```bash
pip install openpyxl python-docx
```

## Installation

### Plugin Install (per-project)
```
/plugin install github:matt-janecek/DonyatiDocs-sow
```

### Global Skill Registration (all projects)

```bash
mkdir -p ~/.claude/skills/{donsow,sowdoc}
cp DonyatiDocs-sow/skills/donsow/SKILL.md ~/.claude/skills/donsow/
cp DonyatiDocs-sow/skills/sowdoc/SKILL.md ~/.claude/skills/sowdoc/
```
