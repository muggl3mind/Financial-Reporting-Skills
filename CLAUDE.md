# Financial Reporting Skills

> **Type**: Claude Code Agent Skills
> **Purpose**: Generate financial statement reports requiring special accounting treatment

---

## Project Summary

This project implements **Claude Code Agent Skills** for financial statement accounts requiring special treatment:

- **Fixed Assets** - Capitalization, depreciation, disposal with gain/loss
- **Investments** - Securities purchase, mark-to-market, sale
- **Capital Accounts** - Equity transactions (placeholder)

Each module outputs CSV (for accounting software) and XLSX (with formulas for audit trails).

---

## Project Structure

```
Financial-Reporting-Skills/
├── .claude/skills/
│   ├── fixed-assets/        # See SKILL.md for details
│   ├── investments/         # See SKILL.md for details
│   └── capital-accounts/    # Placeholder
│
├── lib/                     # Shared Python modules
│   ├── accounts.py          # Chart of Accounts (single source of truth)
│   ├── utils.py             # CSV/XLSX utilities
│   ├── xlsx_styles.py       # Color-coding standards
│   └── recalc.py            # Formula recalculation
│
├── data/
│   ├── source/              # Input files (CSV, PDF)
│   └── output/              # Generated reports
│
└── .venv/                   # Python virtual environment
```

---

## Skills

| Skill | Invoke | Description |
|-------|--------|-------------|
| Fixed Assets | `/fixed-assets` | PP&E lifecycle management |
| Investments | `/investments` | Securities lifecycle management |
| Capital Accounts | `/capital-accounts` | Equity (not yet implemented) |

Each skill has a `SKILL.md` with full documentation, workflows, and script usage.

---

## XLSX Standards

### Color Coding

| Usage | Color |
|-------|-------|
| Hardcoded inputs | Blue |
| Formulas | Black |
| Internal links | Green |
| External links | Red |

### Formula Recalculation

After generating XLSX files, recalculate and validate:

```bash
.venv/bin/python lib/recalc.py <file.xlsx>
```

Returns JSON with status and any errors (`#REF!`, `#DIV/0!`, etc.).

**Requires:** LibreOffice (`brew install --cask libreoffice`)

---

## File Locations

| What | Path |
|------|------|
| Fixed Assets Skill | `.claude/skills/fixed-assets/SKILL.md` |
| Investments Skill | `.claude/skills/investments/SKILL.md` |
| Chart of Accounts | `lib/accounts.py` |
| Fixed Assets Output | `data/output/fixed-assets/` |
| Investments Output | `data/output/investments/` |

---

## JSON Registers

### Asset Register (`data/output/fixed-assets/asset_register.json`)

```json
{
  "assets": [{
    "id": "asset_xxx",
    "name": "Asset Name",
    "category": "Equipment",
    "purchase_date": "YYYY-MM-DD",
    "cost": 0.00,
    "salvage_value": 0.00,
    "useful_life_years": 0,
    "status": "active|disposed"
  }],
  "metadata": { "last_updated": "ISO timestamp" }
}
```

### Securities Register (`data/output/investments/securities_register.json`)

```json
{
  "transactions": [{
    "id": "txn_xxx",
    "type": "purchase|sale",
    "security": "Name",
    "date": "YYYY-MM-DD",
    "shares": 0,
    "security_type": "AFS|Trading"
  }],
  "holdings": {
    "Security Name": { "shares": 0, "cost_basis": 0.00 }
  },
  "fmv_adjustments": [{
    "security": "Name",
    "date": "YYYY-MM-DD",
    "adjustment": 0.00
  }]
}
```

---

## Dependencies

- Python 3.x
- openpyxl (XLSX output)
- LibreOffice (formula recalculation)
