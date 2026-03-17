---
name: setup
description: Set up the project environment for first-time users. Use when user wants to install dependencies, create virtual environment, or prepare the project for use.
---

# Project Setup

Automatically prepare Financial Reporting Skills for first-time use.

---

## When to Use

Run this skill when a user:
- Just cloned the repository
- Says "set up the project" or "get started"
- Gets errors about missing modules or dependencies

---

## Setup Steps

When invoked, run these steps in order. **Detect the operating system first** to use the correct commands.

### Step 0: Detect Operating System

```bash
uname -s 2>/dev/null || echo "Windows"
```

- If output contains "Darwin" → **Mac**
- If output contains "Linux" → **Linux**
- If output is "Windows" or command fails → **Windows**

Set commands accordingly:

| OS | Python command | Pip path | Activate path |
|----|---------------|----------|---------------|
| Mac/Linux | `python3` | `.venv/bin/pip` | `.venv/bin/python` |
| Windows | `python` | `.venv\Scripts\pip` | `.venv\Scripts\python` |

### Step 1: Check Python

**Mac/Linux:**
```bash
python3 --version
```

**Windows:**
```bash
python --version
```

If Python is not installed, tell the user:
> "Python 3 is required but not installed. Download it from https://python.org"

### Step 2: Create Virtual Environment

Check if `.venv` already exists:

```bash
ls .venv 2>/dev/null || dir .venv 2>nul || echo "No venv found"
```

If no venv exists, create one:

**Mac/Linux:**
```bash
python3 -m venv .venv
```

**Windows:**
```bash
python -m venv .venv
```

### Step 3: Install Dependencies

**Mac/Linux:**
```bash
.venv/bin/pip install --upgrade pip -q
.venv/bin/pip install openpyxl -q
```

**Windows:**
```bash
.venv\Scripts\pip install --upgrade pip -q
.venv\Scripts\pip install openpyxl -q
```

### Step 4: Verify Installation

**Mac/Linux:**
```bash
.venv/bin/python -c "import openpyxl; print('openpyxl installed successfully')"
```

**Windows:**
```bash
.venv\Scripts\python -c "import openpyxl; print('openpyxl installed successfully')"
```

### Step 5: Check LibreOffice (Optional)

**Mac/Linux:**
```bash
which soffice || echo "LibreOffice not found"
```

**Windows:**
```bash
where soffice 2>nul || echo "LibreOffice not found"
```

If LibreOffice is not found, inform the user:
> "LibreOffice is optional but recommended for validating Excel formulas."
> - **Mac:** `brew install --cask libreoffice`
> - **Windows:** Download from https://www.libreoffice.org/download/download/

### Step 6: Create Output Directories

**Mac/Linux:**
```bash
mkdir -p data/output/fixed-assets
mkdir -p data/output/investments
mkdir -p data/source/fixed-assets
mkdir -p data/source/investments
```

**Windows:**
```bash
if not exist "data\output\fixed-assets" mkdir "data\output\fixed-assets"
if not exist "data\output\investments" mkdir "data\output\investments"
if not exist "data\source\fixed-assets" mkdir "data\source\fixed-assets"
if not exist "data\source\investments" mkdir "data\source\investments"
```

### Step 7: Confirm Success

After all steps complete, tell the user:

> **Setup complete!** You're ready to use the skills:
> - Type `/fixed-assets` to manage fixed assets
> - Type `/investments` to manage investment securities
>
> Place your source files in `data/source/` when ready.

---

## Troubleshooting

### If pip install fails

Try running pip as a module:

**Mac/Linux:**
```bash
.venv/bin/python -m pip install openpyxl
```

**Windows:**
```bash
.venv\Scripts\python -m pip install openpyxl
```

### If venv creation fails on Linux

Install the venv module:
```bash
sudo apt install python3-venv
```

### If permissions errors occur

Make sure you're running from a directory you own (not system folders).

### If Python isn't found on Windows

Python might not be in your PATH. Reinstall Python and check "Add Python to PATH" during installation.
