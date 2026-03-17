# Financial Reporting Skills

Claude Code skills for generating financial statement reports that require special accounting treatment.

## What This Does

- **Fixed Assets** - Capitalization, straight-line depreciation, disposal with gain/loss
- **Investments** - Securities purchase, mark-to-market (AFS/Trading), sale with realized G/L
- **Capital Accounts** - Equity transactions (placeholder)

Outputs both CSV (for accounting software import) and XLSX (with working Excel formulas for audit trails).

---

## Prerequisites

Before you begin, you'll need:

| Requirement | How to Get It |
|-------------|---------------|
| Mac or PC with a terminal | Built-in on both (Terminal on Mac, Command Prompt or PowerShell on Windows) |
| Claude Pro or Max subscription | $20/month or $100/month at [claude.ai](https://claude.ai) |
| Claude Code | Free CLI tool from Anthropic |
| Python 3.x | Usually pre-installed on Mac. [Download here](https://python.org) if needed |
| LibreOffice | For XLSX formula validation (optional but recommended) |

---

## Quick Start (3 Steps)

### 1. Install Claude Code

```bash
npm install -g @anthropic-ai/claude-code
```

(Need npm? Install Node.js first from [nodejs.org](https://nodejs.org))

### 2. Clone and Open

```bash
git clone https://github.com/muggl3mind/Financial-Reporting-Skills.git
cd Financial-Reporting-Skills
claude
```

### 3. Run Setup

Once Claude Code opens, type:

```
/setup
```

Claude will automatically:
- Create the Python virtual environment
- Install all dependencies
- Create the necessary folders
- Verify everything works

That's it. You're ready to go.

---

## How to Use the Skills

### Step 1: Launch Claude Code

In your terminal, make sure you're in the project folder:

```bash
cd ~/Documents/Financial-Reporting-Skills
```

Then launch Claude Code:

```bash
claude
```

You'll see a prompt where you can type commands.

### Step 2: Invoke a Skill

Type a slash command to load a skill:

```
/fixed-assets
```

or

```
/investments
```

Claude will read the skill instructions and start an interactive workflow.

### Step 3: Follow the Prompts

Claude will ask you questions in plain English:

1. "What source files are available?"
2. "What's the reporting period end date?"
3. "What operation? (Capitalize, Depreciate, Dispose, or Status)"

Answer in plain English. Claude runs the Python scripts behind the scenes.

### Step 4: Review Your Outputs

Generated files appear in:

```
data/output/fixed-assets/     # Fixed asset reports
data/output/investments/      # Investment reports
```

Each operation generates both CSV and XLSX files.

---

## Slash Command Tips & Tricks

### What Slash Commands Are

Slash commands like `/fixed-assets` are shortcuts that tell Claude Code which skill to load. The name after the slash matches the folder name inside `.claude/skills/`.

```
/fixed-assets  →  loads  .claude/skills/fixed-assets/SKILL.md
/investments   →  loads  .claude/skills/investments/SKILL.md
```

### Available Commands

| Command | What It Does |
|---------|--------------|
| `/setup` | First-time setup: installs dependencies, creates folders |
| `/fixed-assets` | Manage PP&E: capitalize, depreciate, dispose |
| `/investments` | Manage securities: purchase, mark-to-market, sell |
| `/capital-accounts` | Equity transactions (placeholder - not yet implemented) |

### You Don't Have to Use Slash Commands

You can also just describe what you want in plain English:

```
"Capitalize a Dell laptop purchased January 15, 2025 for $1,850"

"Generate depreciation schedule through June 30, 2025"

"Record disposal of the laptop on May 15, 2025 for $1,200"
```

Claude will figure out which skill to use based on your request.

### When to Use Slash Commands vs. Plain English

| Use Slash Commands When... | Use Plain English When... |
|---------------------------|---------------------------|
| You want the full guided workflow | You know exactly what you need |
| You're new to the skill | You've used it before |
| You want Claude to ask clarifying questions | You want to skip the Q&A |

### Pro Tips

**Chain operations:** After running `/fixed-assets`, you can say "now depreciate it through September 30" without re-invoking the skill. Claude remembers the context.

**Check status first:** Type `/fixed-assets` then choose "Status" to see what's already in your asset register before making changes.

**Batch processing:** You can describe multiple assets at once: "Capitalize these three purchases: [list details]"

---

## Project Structure

```
Financial-Reporting-Skills/
├── .claude/
│   └── skills/
│       ├── setup/                 # First-time setup
│       │   └── SKILL.md
│       ├── fixed-assets/          # PP&E skill
│       │   ├── SKILL.md           # Main instructions
│       │   ├── references/        # Account mappings, useful lives, formulas
│       │   └── scripts/           # Python scripts
│       ├── investments/           # Securities skill
│       │   ├── SKILL.md
│       │   ├── references/
│       │   └── scripts/
│       └── capital-accounts/      # Placeholder
│
├── lib/                           # Shared Python modules
│   ├── accounts.py                # Chart of Accounts (single source of truth)
│   ├── utils.py                   # CSV/XLSX utilities
│   └── xlsx_styles.py             # Color-coding standards
│
├── data/
│   ├── source/                    # Put your input files here
│   │   ├── fixed-assets/
│   │   └── investments/
│   └── output/                    # Generated reports appear here
│       ├── fixed-assets/
│       └── investments/
│
└── README.md
```

---

## Customizing for Your Organization

### Change the Chart of Accounts

Edit `lib/accounts.py` to match your GL codes:

```python
CHART_OF_ACCOUNTS = {
    '1500': 'Property Plant & Equipment',  # Change codes as needed
    '1600': 'Accumulated Depreciation',
    '6300': 'Depreciation Expense',
    # Add your accounts...
}
```

All skills import from this file, so changes apply everywhere.

### Modify Accounting Policies

Edit the `SKILL.md` files to reflect your policies:

- `.claude/skills/fixed-assets/SKILL.md` - Depreciation methods, proration rules
- `.claude/skills/investments/SKILL.md` - Classification rules, valuation policies

### Add Reference Documents

Place additional documentation in the `references/` folder of each skill. Claude can consult these when needed.

---

## Sample Data

The repository includes sample input files in `data/source/` so you can test immediately.

When ready to use real data:
1. Place your files in the appropriate `data/source/` subfolder
2. Invoke the skill
3. Claude will ask which file to process

Supported formats: CSV, PDF (for broker statements, invoices, etc.)

---

## Troubleshooting

### "Command not found: claude"

Claude Code isn't installed. Run:
```bash
npm install -g @anthropic-ai/claude-code
```

### "No module named 'openpyxl'" or other dependency errors

Run `/setup` again—it will fix missing dependencies.

### Claude doesn't recognize the skill

Make sure you're in the project folder when you launch Claude Code:
```bash
cd ~/Documents/Financial-Reporting-Skills
claude
```

### Something else broke

Just ask Claude: "Help me fix this error: [paste the error]"

---

## License

MIT
