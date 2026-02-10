---
name: sowdoc
description: Convert SOW pricing Excel workbooks to professional Word documents. Reads resource allocations, fees, and project info from Excel and generates a complete Statement of Work document using the donword template system.
---

# SOW Excel to Word Converter

Convert SOW pricing Excel workbooks into professional Statement of Work Word documents with Donyati branding.

## Quick Start

1. **Have a SOW pricing Excel workbook** (from donsow or manual creation)
2. **Run create_sow_document.py** to generate the Word document
3. **Optionally provide a scope.json** for custom content

## Files

- **Script**: `/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/scripts/create_sow_document.py`
- **Example SOWs**: `/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/templates/word/sow/`
- **Example Excel**: `/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/templates/excel/ieee-excel/`

## Command

```bash
SCRIPTS="/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/scripts"

# Basic usage - Excel to Word
python3 "$SCRIPTS/create_sow_document.py" pricing.xlsx output.docx

# With custom scope content
python3 "$SCRIPTS/create_sow_document.py" pricing.xlsx output.docx --scope scope.json

# Output JSON only (for inspection/editing)
python3 "$SCRIPTS/create_sow_document.py" pricing.xlsx output.json --json-only
```

---

## Excel Workbook Requirements

The script reads data from these sheets:

### Pricing Details Sheet (Required)

| Row | Column A | Column B | Column C | Column D | Column E | Column F |
|-----|----------|----------|----------|----------|----------|----------|
| 1 | Client Name: | *value* | Pricing Date: | *value* | | |
| 2 | Capability: | *value* | Sales Rep: | *value* | Type: | *value* |
| 3 | Service Type: | *value* | Presales: | *value* | Risk: | *value* |
| 4 | Contract Type: | *value* | Engagement Lead: | *value* | Start Date: | *value* |
| 6 | **Headers:** Practice | Role | Project Role | Resource | Location | Rate |
| 7+ | *resource data* | | | | | |

### Summary Sheet (Optional)

Used for project name and totals if not in Pricing Details.

### Deliverables Sheet (Optional)

Phase names in row 1, deliverables listed below each phase.

---

## Generated Document Structure

The Word document includes these sections:

1. **Header** - Client name, project name, date
2. **About this Statement of Work** - Standard intro paragraph
3. **Scope of Services** - From scope.json or generated from resources
4. **Deliverables** - From Excel Deliverables sheet or scope.json
5. **Client Responsibilities** - Standard list (customizable)
6. **Engagement Fees** - Resource table with hours, rates, totals
7. **Assumptions** - Standard assumptions (customizable)
8. **Outside the Scope of this SOW** - Exclusions
9. **Miscellaneous Provisions** - Terms table

---

## Scope JSON Format (Optional)

Provide custom content via `--scope scope.json`:

```json
{
  "msa_date": "January 1, 2025",

  "about": "Custom about section text with {date}, {client_name}, {project_name} placeholders...",

  "scope_intro": "Under this SOW, Donyati will provide the following services:",

  "scope_items": [
    "Provide project management oversight",
    "Conduct requirements workshops",
    "Design and implement solution architecture",
    "Perform system integration testing"
  ],

  "deliverables": [
    "Project Charter and Schedule",
    "Requirements Traceability Matrix",
    "Design Document",
    "Test Plan and Results",
    "Training Materials"
  ],

  "client_responsibilities": [
    "Provide dedicated project sponsor",
    "Make SMEs available for workshops",
    "Approve deliverables within 5 business days",
    "Provide system access and credentials"
  ],

  "assumptions": [
    "All services provided in English",
    "Remote work unless travel approved",
    "Client provides test environment access",
    "Standard 8-hour business day"
  ],

  "out_of_scope": [
    "Data migration from legacy systems",
    "Custom report development",
    "End-user training delivery"
  ],

  "misc_provisions": [
    {
      "provision": "Change Orders",
      "narrative": "Changes require written approval..."
    },
    {
      "provision": "Intellectual Property",
      "narrative": "Deliverables become Client property..."
    }
  ]
}
```

---

## Workflow Examples

### Example 1: Basic Conversion

```bash
SCRIPTS="/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/scripts"

# Convert IEEE pricing Excel to Word
python3 "$SCRIPTS/create_sow_document.py" \
  "templates/excel/ieee-excel/SOW30 Pricing IEEE IT Change Management -20251222.xlsx" \
  "IEEE-SOW30.docx"
```

### Example 2: With Custom Scope

1. Create scope.json with project-specific content:

```json
{
  "msa_date": "March 15, 2024",
  "scope_items": [
    "Lead change management strategy development",
    "Conduct stakeholder impact assessments",
    "Develop and execute communication plans",
    "Support training and adoption activities"
  ],
  "deliverables": [
    "Change Strategy Document",
    "Stakeholder Analysis Matrix",
    "Communications Plan",
    "Training Readiness Report"
  ]
}
```

2. Generate document:

```bash
SCRIPTS="/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/scripts"
python3 "$SCRIPTS/create_sow_document.py" pricing.xlsx output.docx --scope scope.json
```

### Example 3: Review Content Before Generation

```bash
SCRIPTS="/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/scripts"

# Output JSON for review
python3 "$SCRIPTS/create_sow_document.py" pricing.xlsx content.json --json-only

# Edit content.json as needed, then use create_document.py directly
python3 "$SCRIPTS/create_document.py" content.json output.docx --template cover
```

---

## Integration with donsow

Use together with the donsow skill:

```bash
SCRIPTS="/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/scripts"

# Step 1: Create pricing workbook from JSON
python3 "$SCRIPTS/create_sow_workbook.py" sow-data.json pricing.xlsx

# Step 2: Convert to Word document
python3 "$SCRIPTS/create_sow_document.py" pricing.xlsx sow-document.docx
```

Or with scope customization:

```bash
SCRIPTS="/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/scripts"

# Generate both pricing sheet and SOW document
python3 "$SCRIPTS/create_sow_workbook.py" project-data.json pricing.xlsx
python3 "$SCRIPTS/create_sow_document.py" pricing.xlsx sow.docx --scope project-scope.json
```

---

## Data Mapping

| Excel Field | Word Document Location |
|-------------|------------------------|
| Client Name (B1) | Header, About section |
| Pricing Date (D1) | Header, Date field |
| Capability (B2) | About section |
| Service Type (B3) | About section |
| Contract Type (B4) | Engagement Fees intro |
| Start Date (F4) | Engagement Fees section |
| Resources (rows 7+) | Fees table |
| Deliverables sheet | Deliverables section |

---

## Tips

1. **Resource Roles**: The script uses `Project Role` if available, otherwise falls back to `Resource Role` from Excel.

2. **Deliverables**: If no Deliverables sheet exists and no scope.json provided, a placeholder is inserted.

3. **Totals**: Calculated from resource data if not found in Summary sheet.

4. **Custom Content**: Use scope.json to override any default section content.

5. **Template Selection**: The script always uses the cover page template for professional SOW appearance.

---

## Troubleshooting

### "Excel file not found"
Verify the path to the pricing workbook is correct.

### "No resources found"
Check that the Pricing Details sheet has resource data starting at row 7.

### "Missing project info"
Ensure rows 1-4 of Pricing Details have client name, dates, and project info.

### Word document formatting issues
Use `--json-only` to inspect the generated content, then manually adjust and run create_document.py directly.
