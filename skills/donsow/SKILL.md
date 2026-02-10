---
name: donsow
description: Generate SOW pricing Excel workbooks with Donyati branding. Supports JSON input, conversational interview, and clone-from-existing modes. Creates 11-sheet workbooks matching Donyati SOW pricing standards.
---

# Donyati SOW Pricing Generator

Create professional SOW (Statement of Work) pricing Excel workbooks with Donyati branding and standard structure.

## Quick Start

1. **Create sow-data.json** with project, sales team, and resource details
2. **Run create_sow_workbook.py** to generate the .xlsx file
3. Open the generated workbook in Excel

## Files

- **Script**: `/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/scripts/create_sow_workbook.py`
- **Reference Data**: `/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/data/sow-reference-data.json`
- **Example SOWs**: `/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/templates/excel/ieee-excel/`

## Three Input Modes

### Mode 1: JSON Input (Recommended)

Create a JSON file with full SOW specification, then generate:

```bash
SCRIPTS="/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/scripts"
python3 "$SCRIPTS/create_sow_workbook.py" sow-data.json output.xlsx
```

### Mode 2: Conversational (Claude builds JSON)

Claude guides you through data collection using questions:
1. Project basics: client, name, dates, duration, contract type
2. Sales team: relationship owner, reps, leader
3. Resources: iterate to add each resource with role, location, rate, hours
4. Timeline: phases and durations
5. Deliverables: use defaults or customize

After collecting info, Claude generates the JSON and creates the workbook.

### Mode 3: Clone & Modify

Start from an existing SOW and apply modifications:

```bash
SCRIPTS="/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/scripts"

# Clone without modifications
python3 "$SCRIPTS/create_sow_workbook.py" --clone existing.xlsx output.xlsx

# Clone with modifications
python3 "$SCRIPTS/create_sow_workbook.py" --clone existing.xlsx output.xlsx --json changes.json
```

---

## JSON Schema

### Complete Example

```json
{
  "project": {
    "client_name": "IEEE",
    "project_name": "AI Chatbot Implementation",
    "capability_area": "Cloud Engineering",
    "service_type": "Implementation",
    "contract_type": "T&M",
    "project_type": "New",
    "risk_profile": "Low",
    "pricing_date": "2026-01-29",
    "project_start_date": "2026-02-15",
    "duration_months": 6
  },
  "sales_team": {
    "relationship_owner": "Jai Jain",
    "sales_rep": "Camryn Barry",
    "inside_sales": "Melissa O'Blenis",
    "sales_team_leader": "Moe Gohary"
  },
  "resources": [
    {
      "practice": "EPM_Practice_USA",
      "resource_role": "Project Manager",
      "project_role": "AI Scrum Master",
      "potential_resource": "TBD",
      "location": "USA",
      "hourly_rate": 230,
      "monthly_hours": [24, 168, 160, 176, 168, 168]
    },
    {
      "practice": "TS_Practice_India",
      "resource_role": "Developer",
      "project_role": "AI Developer",
      "potential_resource": "TBD",
      "location": "India",
      "hourly_rate": 55,
      "monthly_hours": [0, 168, 168, 176, 168, 80]
    }
  ],
  "phases": [
    {"name": "Mobilize", "start_week": 1, "end_week": 2},
    {"name": "Requirements", "start_week": 2, "end_week": 5},
    {"name": "Design", "start_week": 4, "end_week": 8},
    {"name": "Build", "start_week": 7, "end_week": 18},
    {"name": "Testing", "start_week": 16, "end_week": 22},
    {"name": "Training", "start_week": 20, "end_week": 23},
    {"name": "HyperCare", "start_week": 23, "end_week": 26}
  ],
  "deliverables": "default"
}
```

### Field Reference

#### project (required)

| Field | Type | Description |
|-------|------|-------------|
| `client_name` | string | Client organization name |
| `project_name` | string | SOW project name |
| `capability_area` | string | Practice area (see Capabilities list) |
| `service_type` | string | Type of engagement (see Service Types list) |
| `contract_type` | string | T&M, Fixed Fee, Milestone, T&M - On Demand |
| `project_type` | string | New, Renewal, Extension |
| `risk_profile` | string | Low, Medium, High |
| `pricing_date` | string | Date pricing was created (YYYY-MM-DD) |
| `project_start_date` | string | Project start date (YYYY-MM-DD) |
| `duration_months` | integer | Project duration in months |

#### sales_team (required)

| Field | Type | Description |
|-------|------|-------------|
| `relationship_owner` | string | Account relationship owner |
| `sales_rep` | string | Sales representative |
| `inside_sales` | string | Inside sales contact |
| `sales_team_leader` | string | Sales team leader |

#### resources (required)

Array of resource objects:

| Field | Type | Description |
|-------|------|-------------|
| `practice` | string | Practice code (see Practices list) |
| `resource_role` | string | Certinia role (see Roles list) |
| `project_role` | string | Role description for this project |
| `potential_resource` | string | Named resource or "TBD" |
| `location` | string | USA, India, etc. |
| `hourly_rate` | number | Billing rate per hour (auto-looked up if omitted) |
| `monthly_hours` | array | Hours per month (e.g., [24, 168, 160, 176]) |

#### phases (optional)

Array of phase objects. If omitted, uses default phases.

| Field | Type | Description |
|-------|------|-------------|
| `name` | string | Phase name |
| `start_week` | integer | Starting week number |
| `end_week` | integer | Ending week number |

#### deliverables (optional)

- `"default"` - Use standard deliverables matrix
- Custom object - Define deliverables by phase

---

## Generated Workbook Structure

The workbook contains 11 sheets:

| Sheet | Purpose |
|-------|---------|
| **Summary** | Project overview and fee totals |
| **Pricing Details** | Resource allocation and monthly hours |
| **Timeline 36 Weeks** | Gantt-style phase timeline |
| **Deliverables** | Phase-by-phase deliverable matrix |
| **Complexity Considerations** | Factors affecting estimate |
| **Formulas** | Calculation helpers |
| **Picklist** | Reference data for dropdowns |
| **INSTRUCTIONS** | How to use the workbook |
| **Update Resource Instructions** | Admin instructions |
| **Certinia Resource List** | Employee/contractor reference |
| **Customer Capability Service** | Capability/service matrix |

---

## Reference Data

### Contract Types

- T&M
- T&M - On Demand
- Fixed Fee
- Milestone

### Practices (Common)

| Practice | Location |
|----------|----------|
| `EPM_Practice_USA` | USA |
| `EPM_Practice_India` | India |
| `TS_Practice_USA` | USA |
| `TS_Practice_India` | India |
| `Cloud_Practice_USA` | USA |
| `ERP_Practice_USA` | USA |
| `ERP_Practice_India` | India |

### Common Roles & USA Rates

| Role | Rate |
|------|------|
| Application Developer | $215 |
| Project Manager | $230 |
| Sr. Application Developer | $240 |
| Sr. Project Manager | $250 |
| Application Lead | $265 |
| Functional Lead | $315 |
| Solution Architect | $365 |
| Sr. Solution Architect | $390 |
| Principal | $415 |
| Principal Architect | $450 |

India rates are typically $45-75 depending on role.

### Service Types

- Implementation
- Assessment
- Migration
- Development
- Support
- Enhancement
- Project Management
- Change Management
- Advisory
- Strategy

### Capabilities

- EPM PM (Performance Management)
- Cloud Engineering
- Data Transformation
- Power Apps
- Oracle ERP Cloud
- BI
- AWS
- Azure

---

## Commands

```bash
SCRIPTS="/mnt/c/Users/MattJanecek/Claude/DonyatiDocs-sow/scripts"

# Generate from JSON
python3 "$SCRIPTS/create_sow_workbook.py" sow-data.json output.xlsx

# Clone existing workbook
python3 "$SCRIPTS/create_sow_workbook.py" --clone templates/excel/ieee-excel/SOW19*.xlsx new-sow.xlsx

# Clone and modify
python3 "$SCRIPTS/create_sow_workbook.py" --clone existing.xlsx modified.xlsx --json changes.json
```

---

## Conversational Interview Questions

When using conversational mode, Claude collects:

### Project Information
1. Client name?
2. Project name?
3. Capability/practice area? (EPM PM, Cloud Engineering, etc.)
4. Service type? (Implementation, Assessment, etc.)
5. Contract type? (T&M, Fixed Fee, Milestone)
6. Risk profile? (Low, Medium, High)
7. Project start date?
8. Duration in months?

### Sales Team
1. Relationship owner?
2. Sales rep?
3. Inside sales contact?
4. Sales team leader?

### Resources (repeat for each)
1. Practice (EPM_Practice_USA, TS_Practice_India, etc.)?
2. Certinia role (Project Manager, Developer, etc.)?
3. Project-specific role description?
4. Named resource or TBD?
5. Location (USA/India)?
6. Monthly hours array (e.g., 24, 168, 160, 176)?
7. Add another resource?

### Timeline (optional)
1. Use default phases or customize?
2. If custom: phase name, start week, end week?

### Deliverables
1. Use default deliverables or customize?

---

## Tips

1. **Rate Lookup**: If `hourly_rate` is omitted from a resource, it's automatically looked up from reference data based on practice + role combination.

2. **Monthly Hours**: Typical values:
   - Full-time: 168 hours/month
   - Part-time: 80-120 hours/month
   - Light involvement: 20-40 hours/month

3. **Duration**: The `monthly_hours` array length should match or be less than `duration_months`.

4. **Formulas**: The Pricing Details sheet uses Excel formulas for Total Hours and Total Fees.

5. **Clone Mode**: Use clone mode when you want to preserve complex formatting and formulas from an existing SOW.

---

## Troubleshooting

### "Reference data not found"
Ensure `data/sow-reference-data.json` exists. Regenerate from IEEE examples if needed.

### Rate lookup returns default
The practice + role combination must exactly match entries in reference data. Check spelling and underscores.

### Missing sheets
The script creates all 11 sheets. If some are missing, verify the JSON structure is valid.

---

## Brand Reference

| Element | Value |
|---------|-------|
| Donyati Purple | `#4A4778` |
| Donyati Black | `#12002A` |
| Light Purple (editable cells) | `#E8E6F0` |
| Alternate Row | `#F5F0FA` |
