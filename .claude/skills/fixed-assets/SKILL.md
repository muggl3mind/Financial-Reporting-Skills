---
name: fixed-assets
description: Manage fixed assets including capitalization, depreciation, and disposal. Use when user wants to (1) process asset transactions from source files, (2) run the interactive workflow with guided questions, (3) execute scripts directly, or (4) understand depreciation calculations and account mappings.
---

# Fixed Assets Module

Manage the complete lifecycle of fixed assets: capitalization, depreciation, and disposal.

---

## Interactive Workflow

When user invokes /fixed-assets or asks to process assets:

1. **First, check for source files:**
   > "Let me check what source data is available in data/source/fixed-assets/"

   Run: `ls -la data/source/fixed-assets/`

2. **If source files exist, ask:**
   > "I found these files: [list]. Which file contains the data you want to process?"

3. **If no source files, ask:**
   > "No source files found in data/source/fixed-assets/. Would you like to:
   > - Place a CSV/PDF file there and try again
   > - Enter transaction details manually"

4. **Ask for report period end date (REQUIRED):**
   > "What is the reporting period end date? (Must be Mar 31, Jun 30, Sep 30, or Dec 31)"

5. **Then ask operation:**
   > What would you like to do?
   > 1. **Capitalize** - Record a new asset purchase
   > 2. **Depreciate** - Generate depreciation entries
   > 3. **Dispose** - Record asset disposal
   > 4. **Status** - View current asset register

---

## Source Data Location

Users should place their asset data files in:
```
data/source/fixed-assets/
```

Supported formats: CSV, PDF

**View current asset register:**
```bash
cat data/output/fixed-assets/asset_register.json | python3 -m json.tool
```

---

## Required User Inputs

Before generating any schedules or journal entries, you MUST obtain these from the user:

1. **Report Period End Date** - Must be quarter-end (Mar 31, Jun 30, Sep 30, or Dec 31)
2. **Transaction Dates** - Must be provided for each capitalization or disposal

**Period Start:** Defaults to January 1st of the report year. JEs always cover full YTD.

---

## Auto-Run Workflow (CRITICAL)

When user selects an operation or says "do all", Claude MUST automatically run ALL applicable steps in sequence:

### For Purchase + Disposal (full lifecycle):
1. `capitalize.py` → Record purchase
2. `depreciate.py` → Generate schedule + JEs (Jan 1 through **disposal date**)
   - **CRITICAL:** Include `"disposal_date"` in asset JSON to stop depreciation at disposal
3. `dispose.py` → Record disposal with gain/loss

### For Purchase only:
1. `capitalize.py` → Record purchase
2. `depreciate.py` → Generate schedule + JEs (Jan 1 through report date)

### Required Outputs Checklist

- [ ] Capitalization journal entry (if new purchase)
- [ ] Depreciation schedule (CSV + XLSX)
- [ ] Depreciation journal entries (from Jan 1 of report year)
- [ ] PPE rollforward (Jan 1 through report date)
- [ ] Disposal journal entry (if sold/retired)
- [ ] Updated asset register
- [ ] **Recalculate all XLSX formulas** (see below)

**DO NOT skip any steps. DO NOT complete without ALL outputs.**

### XLSX Post-Processing (MANDATORY)

After generating ANY XLSX files, Claude MUST run recalc.py to recalculate formulas and verify zero errors:

```bash
.venv/bin/python lib/recalc.py <output_file.xlsx>
```

**Example:**
```bash
.venv/bin/python lib/recalc.py data/output/fixed-assets/ppe_rollforward_2025-09-30.xlsx
```

**Report results to user:**
- If `status: success` → "Generated [file] with X formulas, 0 errors"
- If `status: errors_found` → Show error details and fix before delivering

**DO NOT deliver XLSX files without running recalc.py first.**

---

## Accounting Policies

- **Depreciation Method:** Straight-line only
- **First/Final Period:** Prorated based on actual days in service
- **Salvage Value:** Deducted from depreciable base before calculating depreciation
- **Disposal:** Gain/loss recognized immediately at disposal date

---

## Reference Documents

| Topic | Reference |
|-------|-----------|
| Default useful lives (IRS MACRS) | [references/asset_lives.md](references/asset_lives.md) |
| GL account codes | [references/account_mapping.md](references/account_mapping.md) |
| Excel formulas & proration | [references/depreciation_math.md](references/depreciation_math.md) |
| Partial disposal, batch processing | [references/advanced_usage.md](references/advanced_usage.md) |

---

## Scripts

```
scripts/
├── capitalize.py    # Record asset purchases
├── depreciate.py    # Generate depreciation schedules
├── dispose.py       # Record disposals with gain/loss
└── depreciation.py  # Core calculation functions
```

---

## 1. Capitalize (Purchase)

Record the initial asset purchase.

```bash
.venv/bin/python .claude/skills/fixed-assets/scripts/capitalize.py \
  --name "CNC Milling Machine" \
  --category Equipment \
  --date "2025-03-15" \
  --cost 180000 \
  --credit-account cash \
  --output-dir "data/output/fixed-assets"
```

**Parameters:**
- `--name`: Asset name
- `--category`: Equipment, Vehicle, Furniture, Computer, Building, Software, Patent
- `--date`: Purchase date (YYYY-MM-DD)
- `--cost`: Purchase cost
- `--credit-account`: `cash` or `ap` (accounts payable)
- `--output-dir`: Output directory (default: data/output/fixed-assets)

**Journal Entry:**
```
Dr 1500 Property Plant & Equipment    $180,000
   Cr 1000 Cash                                  $180,000
```

---

## 2. Depreciate

Generate depreciation schedules and periodic journal entries.

```bash
.venv/bin/python .claude/skills/fixed-assets/scripts/depreciate.py \
  --assets '[{"name": "CNC Milling Machine", "category": "Equipment", "purchase_date": "2025-03-15", "cost": 180000, "salvage": 15000, "useful_life": 10}]' \
  --period monthly \
  --report-date "2025-09-30" \
  --output-dir "data/output/fixed-assets"
```

**Parameters:**
- `--assets`: JSON array of assets or path to JSON file
- `--period`: Depreciation frequency - monthly, quarterly, or annual
- `--report-date`: Period end date (YYYY-MM-DD)
- `--period-start`: Period begin date (YYYY-MM-DD) - alternative to --period-type
- `--period-type`: `quarterly`, `half-year`, or `annual` - auto-calculates period-start
- `--output-dir`: Output directory (default: data/output/fixed-assets)

**Asset JSON Format:**
```json
{
  "name": "Asset Name",
  "category": "Equipment",
  "purchase_date": "2025-01-15",
  "cost": 50000,
  "salvage": 5000,
  "useful_life": 5,
  "disposal_date": "2025-05-15"  // Optional - stops depreciation at this date
}
```

**Output Files:**
- `schedule_<asset>.csv/.xlsx` - Period-by-period schedule
- `consolidated_schedule.csv/.xlsx` - All assets combined
- `ppe_rollforward.xlsx` - PPE summary with formulas
- `consolidated_journal.csv/.xlsx` - All journal entries

---

## 3. Dispose

Record asset disposal with gain/loss calculation.

```bash
.venv/bin/python .claude/skills/fixed-assets/scripts/dispose.py \
  --name "CNC Milling Machine" \
  --category Equipment \
  --purchase-date "2025-03-15" \
  --cost 180000 \
  --salvage 15000 \
  --useful-life 10 \
  --disposal-date "2025-07-08" \
  --proceeds 180000 \
  --output-dir "data/output/fixed-assets"
```

**Journal Entry (Gain):**
```
Dr 1000 Cash                          $180,000
Dr 1600 Accumulated Depreciation        $5,259
   Cr 1500 Property Plant & Equipment           $180,000
   Cr 7200 Gain on Disposal                       $5,259
```

---

## Output Files

| File | Description |
|------|-------------|
| `capitalize_<asset>.csv/.xlsx` | Purchase journal entry |
| `schedule_<asset>.csv/.xlsx` | Depreciation schedule |
| `disposal_<asset>.csv/.xlsx` | Disposal journal entry |
| `consolidated_schedule.csv/.xlsx` | All assets combined |
| `ppe_rollforward.xlsx` | PPE summary with formulas |
| `consolidated_journal.csv/.xlsx` | All depreciation entries |
| `asset_register.json` | Fixed asset tracking |
