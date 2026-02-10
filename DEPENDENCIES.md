# DonyatiDocs-sow Dependencies

## Python Requirements

### donsow Skill (SOW Pricing Excel Generation)

| Package | Version | Purpose |
|---------|---------|---------|
| `openpyxl` | >=3.1.0 | Excel workbook creation and manipulation |

**Install:**
```bash
pip install openpyxl
```

### sowdoc Skill (SOW Excel to Word Conversion)

| Package | Version | Purpose |
|---------|---------|---------|
| `openpyxl` | >=3.1.0 | Reading SOW pricing Excel workbooks |
| `python-docx` | >=0.8.11 | Word document creation (via inlined donword engine) |

**Install:**
```bash
pip install openpyxl python-docx
```

## Combined Installation

```bash
pip install openpyxl python-docx
```

## System Requirements

| Requirement | Details |
|-------------|---------|
| Python | 3.8+ recommended |
| OS | Windows, macOS, Linux |
| Disk Space | ~30 MB for packages + templates |

## Template Files Required

| Skill | Template File | Size |
|-------|---------------|------|
| sowdoc | `templates/word/donyati-word-template-2025.docx` | 2.5 MB |
| sowdoc | `templates/word/donyati-word-template-cover-2025.docx` | 3.9 MB |
| donsow | `data/sow-reference-data.json` | ~50 KB |

## Verification

```bash
python3 -c "import openpyxl; print('openpyxl OK')"
python3 -c "from docx import Document; print('python-docx OK')"
```

## Troubleshooting

### "ModuleNotFoundError: No module named 'openpyxl'"
```bash
pip install openpyxl
```

### "ModuleNotFoundError: No module named 'docx'"
```bash
pip install python-docx
```

### Permission errors on Linux/WSL
```bash
pip install --user openpyxl python-docx
```
