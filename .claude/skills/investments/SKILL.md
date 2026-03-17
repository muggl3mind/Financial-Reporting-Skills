---
name: investments
description: Manage investment securities lifecycle including purchase, mark-to-market, and sale. Use when user wants to (1) process investment transactions from source files, (2) run the interactive workflow with guided questions, (3) execute scripts directly, or (4) understand AFS vs Trading accounting treatment and account mappings.
---

# Investments Module

Manage investment securities lifecycle: purchase, mark-to-market, and sale.

---

## Interactive Workflow

When user invokes /investments or asks to process securities:

1. **First, check for source files:**
   > "Let me check what source data is available in data/source/investments/"

   Run: `ls -la data/source/investments/`

2. **If source files exist, ask:**
   > "I found these files: [list]. Which file contains the data you want to process?"

3. **If no source files, ask:**
   > "No source files found in data/source/investments/. Would you like to:
   > - Place a CSV/PDF file there and try again
   > - Enter transaction details manually"

4. **Then ask operation:**
   > What would you like to do?
   > 1. **Purchase** - Record securities acquisition
   > 2. **Mark-to-Market** - Fair value adjustment
   > 3. **Sell** - Record securities sale
   > 4. **Rollforward** - Generate unrealized G/L rollforward
   > 5. **Status** - View current holdings

---

## Source Data Location

Users should place their securities data files in:
```
data/source/investments/
```

Supported formats: CSV, PDF (broker statements, trade confirmations)

**View current securities register:**
```bash
cat data/output/investments/securities_register.json | python3 -m json.tool
```

---

## Required User Inputs

Before generating any journal entries or reports, you MUST obtain these from the user:

1. **Report Period End Date** - Always ask the user; never assume a date
2. **Transaction Dates** - Must be provided for each purchase, sale, or FMV adjustment
3. **Security Classification** - AFS or Trading (affects accounting treatment)

---

## Accounting Policies

- **AFS Securities:** Unrealized G/L recorded in OCI (equity), not P&L
- **Trading Securities:** Unrealized G/L recorded directly in P&L
- **Realized G/L:** Recognized immediately at sale date
- **Cost Basis:** User must provide cost basis on sales (FIFO, LIFO, or specific ID)

---

## Security Classifications

### Available for Sale (AFS)
- Unrealized gains/losses go to OCI (Other Comprehensive Income)
- Only affects equity, not P&L
- When sold, reclassified to realized gain/loss

### Trading Securities
- Unrealized gains/losses go directly to P&L
- Mark-to-market affects income statement

---

## Reference Documents

| Topic | Reference |
|-------|-----------|
| GL account codes | [references/account_mapping.md](references/account_mapping.md) |
| Rollforward report format | [references/rollforward_format.md](references/rollforward_format.md) |

---

## Scripts

```
scripts/
├── purchase.py         # Record securities acquisition
├── sale.py             # Record sale with realized gain/loss
├── mark_to_market.py   # Fair value adjustments
└── rollforward.py      # Unrealized G/L rollforward report
```

---

## XLSX Post-Processing (MANDATORY)

After generating ANY XLSX files, Claude MUST run recalc.py to recalculate formulas and verify zero errors:

```bash
.venv/bin/python lib/recalc.py <output_file.xlsx>
```

**Example:**
```bash
.venv/bin/python lib/recalc.py data/output/investments/securities_rollforward_2025-09-30.xlsx
```

**Report results to user:**
- If `status: success` → "Generated [file] with X formulas, 0 errors"
- If `status: errors_found` → Show error details and fix before delivering

**DO NOT deliver XLSX files without running recalc.py first.**

---

## 1. Purchase

Record securities acquisition.

```bash
.venv/bin/python .claude/skills/investments/scripts/purchase.py \
  --name "TechCorp Inc." \
  --date "2025-04-08" \
  --shares 500 \
  --price 75.00 \
  --security-type AFS \
  --output-dir "data/output/investments"
```

**Parameters:**
- `--name`: Security name
- `--date`: Purchase date (YYYY-MM-DD)
- `--shares`: Number of shares
- `--price`: Price per share
- `--security-type`: AFS (Available for Sale) or Trading
- `--output-dir`: Output directory (default: data/output/investments)

**Journal Entry:**
```
Dr 1750 Investments - AFS    $37,500
   Cr 1000 Cash                       $37,500
```

---

## 2. Mark to Market (FMV Adjustment)

Record period-end fair value adjustment.

```bash
.venv/bin/python .claude/skills/investments/scripts/mark_to_market.py \
  --name "TechCorp Inc." \
  --date "2025-09-30" \
  --fair-value 45000 \
  --output-dir "data/output/investments"
```

**Parameters:**
- `--name`: Security name (must exist in holdings)
- `--date`: Valuation date (YYYY-MM-DD)
- `--fair-value`: Current fair market value
- `--output-dir`: Output directory (default: data/output/investments)

**Journal Entry (AFS Gain):**
```
Dr 1750 Investments - AFS             $3,000
   Cr 3600 Accumulated OCI                     $3,000
```

**Journal Entry (Trading Gain):**
```
Dr 1760 Investments - Trading         $3,000
   Cr 7400 Unrealized Gain                     $3,000
```

---

## 3. Sale

Record securities sale with realized gain/loss.

```bash
.venv/bin/python .claude/skills/investments/scripts/sale.py \
  --name "TechCorp Inc." \
  --date "2025-09-15" \
  --shares 500 \
  --cost-basis 37500 \
  --proceeds 46000 \
  --security-type AFS \
  --output-dir "data/output/investments"
```

**Parameters:**
- `--name`: Security name
- `--date`: Sale date (YYYY-MM-DD)
- `--shares`: Number of shares sold
- `--cost-basis`: Total cost basis of shares sold
- `--proceeds`: Total sale proceeds
- `--security-type`: AFS or Trading
- `--output-dir`: Output directory (default: data/output/investments)

**Journal Entry (Gain):**
```
Dr 1000 Cash                          $46,000
   Cr 1750 Investments - AFS                   $37,500
   Cr 7300 Realized Gain                        $8,500
```

**XLSX Output Structure:**

The sale XLSX includes a full gain/loss calculation with formulas:

```
SECURITIES SALE - GAIN/LOSS CALCULATION

Security:       TechCorp Inc.
Sale Date:      2025-09-15
Classification: AFS

COST BASIS CALCULATION
Total Shares (before sale):    500             (input)
Total Cost Basis:              $37,500.00      (input)
Cost Per Share:                =TotalCost/TotalShares  (formula)

Shares Sold:                   500             (input)
Cost Basis of Shares Sold:     =SharesSold*CostPerShare  (formula)

GAIN/LOSS CALCULATION
Sale Proceeds:                 $46,000.00      (input)
Less: Cost Basis:              =CostBasisCalc  (formula)
REALIZED GAIN:                 =Proceeds-CostBasis  (formula)

JOURNAL ENTRIES
[Standard journal entry format with DR/CR totals]
```

---

## 4. Unrealized G/L Rollforward

Generate a rollforward report showing movement in cost basis, unrealized gains/losses, and fair value. Also generates a consolidated journal with all period transactions.

```bash
.venv/bin/python .claude/skills/investments/scripts/rollforward.py \
  --report-date "2025-09-30" \
  --period-type quarterly \
  --output-dir "data/output/investments"
```

**Parameters:**
- `--report-date`: Period end date (YYYY-MM-DD) - **required**
- `--period-begin`: Period begin date (YYYY-MM-DD) - alternative to --period-type
- `--period-type`: `quarterly`, `half-year`, or `annual` - auto-calculates begin date
- `--output-dir`: Output directory (default: data/output/investments)

**Output Files:**
- `securities_rollforward_<date>.csv/.xlsx` - Rollforward report
- `journal_<date>.csv/.xlsx` - Consolidated journal entries (all purchases, sales, FMV adjustments)

See [references/rollforward_format.md](references/rollforward_format.md) for the industry-standard report format.

---

## Output Files

| File | Description |
|------|-------------|
| `securities_purchase_<security>_<date>.csv/.xlsx` | Purchase entry |
| `securities_sale_<security>_<date>.csv/.xlsx` | Sale entry |
| `fmv_adjustment_<security>_<date>.csv/.xlsx` | FMV adjustment |
| `securities_rollforward_<date>.csv/.xlsx` | Unrealized G/L rollforward |
| `journal_<date>.csv/.xlsx` | Consolidated journal entries |
| `securities_register.json` | Securities tracking |
