# Straight-Line Depreciation Calculations

Depreciation schedules are **always generated monthly** with day-based proration. This ensures:
- Consistent calculation between schedule and disposal files
- Excel formulas that exactly match Python calculations
- No hardcoded values in XLSX output

## Formulas

```
Annual Depreciation = (Cost - Salvage Value) / Useful Life
Monthly Depreciation = Annual / 12
Period Depreciation = Monthly × (Days in Period / Days in Month)
```

## Excel Formula (same for all periods)

```
=MIN(ROUND(Summary!$D$13*D{row}/E{row},2), F{row}-Summary!$D$7)
```

## Schedule XLSX Structure

### Summary Sheet

| Row | Cell | Label | Value/Formula |
|-----|------|-------|---------------|
| 6 | D6 | Cost | (input) |
| 7 | D7 | Salvage Value | (input) |
| 8 | D8 | Useful Life | (input) |
| 11 | D11 | Depreciable Amount | `=D6-D7` |
| 12 | D12 | Annual Depreciation | `=D11/D8` |
| 13 | D13 | Monthly Depreciation | `=D12/12` |

### Schedule Sheet Columns

| Col | Header | Formula |
|-----|--------|---------|
| A | Period | (number) |
| B | Start Date | (date) |
| C | End Date | (date) |
| D | Days | `=C{row}-B{row}+1` |
| E | Days in Month | `=DAY(EOMONTH(B{row},0))` |
| F | Beginning Book Value | `=Summary!$D$6` or `=I{row-1}` |
| G | Depreciation | `=MIN(ROUND(Summary!$D$13*D{row}/E{row},2),F{row}-Summary!$D$7)` |
| H | Accumulated Depreciation | `=G{row}` or `=H{row-1}+G{row}` |
| I | Ending Book Value | `=F{row}-G{row}` |

## Proration Examples

- Purchase Jan 15 → Jan: Monthly × (17/31)
- Disposal Jul 8 → Jul: Monthly × (8/31)
- Full month → Monthly × (days/days) = Monthly
