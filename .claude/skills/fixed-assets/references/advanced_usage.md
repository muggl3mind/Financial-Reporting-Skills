# Advanced Usage

## Partial Disposal (US GAAP)

The system handles **full asset disposals only**. For partial disposals (e.g., selling 1 of 3 machines), capitalize assets individually at purchase:

**Example - 3 CNC Machines at $60,000 each:**
```bash
# Capitalize each machine separately at purchase
capitalize.py --name "CNC Machine #1" --cost 60000 --category Equipment --date "2025-01-15" ...
capitalize.py --name "CNC Machine #2" --cost 60000 --category Equipment --date "2025-01-15" ...
capitalize.py --name "CNC Machine #3" --cost 60000 --category Equipment --date "2025-01-15" ...

# Later, dispose of only Machine #1
dispose.py --name "CNC Machine #1" --cost 60000 --proceeds 45000 --disposal-date "2025-07-08" ...
```

This ensures proper cost basis and accumulated depreciation allocation per ASC 360.

---

## Batch Processing

When processing multiple transactions, use `--run-timestamp` to keep all outputs in the same folder:

```bash
# Generate timestamp once
TIMESTAMP=$(date +"%Y-%m-%d_%H-%M")

# Use same timestamp for all scripts
capitalize.py --run-timestamp $TIMESTAMP ...
depreciate.py --run-timestamp $TIMESTAMP ...
dispose.py --run-timestamp $TIMESTAMP ...
```

All scripts accept `--run-timestamp YYYY-MM-DD_HH-MM` to specify the output folder.
If not provided, each script generates its own timestamp.
