# Variance Commentary Generator

**Finance × AI: In Practice** — [alloutoftokens.com](https://alloutoftokens.com)

A Python tool that reads two periods of GL data, identifies material variances, and generates CFO-grade management commentary using the Claude AI API.

**Outputs:** 9-tab Excel workbook + interactive HTML dashboard with charts.

Built in public as part of the [Finance × AI: In Practice](https://alloutoftokens.com/blog-variance-1.html) blog series.

---

## What it produces

| Tab | Content |
|---|---|
| 01 Executive Summary | KPI tiles, top P&L variances, management input questions |
| 02 P&L Commentary | Full P&L with inline drivers and narrative |
| 03 P&L Bridge | Waterfall chart from Net Revenue to EBITDA |
| 04 Material Variances | All accounts above the materiality threshold |
| 05 Balance Sheet Flux | BS movements with one-off events flagged |
| 06 Transaction Trail | GL evidence for every material account |
| 07 Opex Breakdown | Cost categories with formula totals |
| 08 Stable Accounts | Below-threshold accounts summarised |
| 09 Methodology | Thresholds, AI engine, sign-off checklist |

---

## Requirements

### 1. Python 3.9+
Download: **https://www.python.org/downloads/**

After installing, verify: open Terminal (Mac) or Command Prompt (Windows) and type:
```
python --version
```

### 2. VS Code (recommended)
Download: **https://code.visualstudio.com/**

This is where you open the project and run the tool. No coding knowledge needed.

### 3. Python libraries

These are packages that extend Python. Install all four:
```
pip install anthropic openpyxl pandas numpy
```

What each one does:
- **anthropic** — lets Python call the Claude AI API (from Anthropic, makers of Claude)
- **openpyxl** — reads and writes `.xlsx` Excel files
- **pandas** — handles tables of data, like Excel but inside Python
- **numpy** — maths library used by pandas automatically

### 4. Anthropic API key

The tool uses Claude AI to write the commentary. You need an account:

1. Go to: **https://console.anthropic.com**
2. Sign up → click **API Keys** → **Create Key**
3. Copy the key immediately (starts with `sk-ant-...`) — you only see it once
4. Each run costs approximately **$0.05–0.15** of API credit

---

## Quick start

### Step 1 — Get the code

```
git clone https://github.com/alloutofxps/variance-commentary-generator.git
cd variance-commentary-generator
```
Or: click the green **Code** button → **Download ZIP** → unzip it.

### Step 2 — Open in VS Code

File → Open Folder → select `variance-commentary-generator`.

### Step 3 — Set your API key

Open `analyse.py`. Find:
```python
API_KEY = os.environ.get("ANTHROPIC_API_KEY", "paste-your-api-key-here")
```
Replace `paste-your-api-key-here` with your key. Save with Ctrl+S (or Cmd+S on Mac).

> ⚠️ Never share `analyse.py` after adding your key. Never commit it to GitHub.
> The `.gitignore` in this repo already excludes `analyse.py` to protect you.

### Step 4 — Set your period labels

In `analyse.py`, update:
```python
Y0_LABEL = "FY 2020"
Y1_LABEL = "FY 2021"
```

### Step 5 — Add your data files

Files must be named exactly:
- `period_0.xlsx` — prior year (Y0)
- `period_1.xlsx` — current year (Y1)

Sample files are included — use them to try the tool immediately.

### Step 6 — Run the tool

In VS Code: open the Terminal (Terminal → New Terminal), type:
```
python analyse.py
```

~20-30 seconds. Two output files appear:
- `variance_analysis.xlsx` — open in Excel
- `variance_dashboard.html` — open in Chrome or Firefox

---

## Your data requirements

**Sheet `GL` must have:** `EntryNo`, `Date`, `Account_key`, `Details`, `Amount`
- Costs must be **negative**, revenue **positive** (standard double-entry)
- Optional: `Territory_key` for the territory chart

**Sheet `Chart of Accounts` must have:** `Account_key`, `Report` ("Profit and Loss" or "Balance Sheet"), `Class`, `SubClass`, `SubClass2`, `Account`
- Optional: `SubAccount` (shown in Transaction Trail sub-category column)

**Year filter:** The tool filters Y0 data to year 2020. To change this, find in `analyse.py`:
```python
y0 = gl[gl["Date"].dt.year == 2020].copy()
```

---

## Understanding the output colours

| Colour | Meaning |
|---|---|
| **Green** | Favourable — revenue up OR cost down vs prior year |
| **Red** | Adverse — revenue down OR cost up vs prior year |
| No colour | Below materiality threshold — not material enough to comment |
| **Blue text** (Excel) | Input value — hardcoded GL total |
| **Black text** (Excel) | Formula — calculated automatically |

The F/A rule is: `Var = Y1 - Y0`. If `Var > 0` = Favourable for ALL accounts.
This works because costs are negative: `Var > 0` means cost became less negative = lower cost = good.

---

## Adding management context

After your first run, Tab 01 lists 5 specific questions in **Areas Requiring Management Input**.

To get better commentary:
1. Open `context.txt`
2. Add your answers (headcount changes, one-off costs, known reasons)
3. Re-run: `python analyse.py`

---

## Materiality thresholds

Both conditions must be met for an account to appear in the variance tables:

| Category | Abs | Pct |
|---|---|---|
| Sales | ≥ 10,000 | ≥ 3% |
| Cost of Sales | ≥ 10,000 | ≥ 3% |
| Administration | ≥ 5,000 | ≥ 5% |
| Balance Sheet | ≥ 50,000 | ≥ 10% |

Adjust in the `THRESHOLDS` section at the top of `analyse.py`.

---

## Troubleshooting

| Error | Fix |
|---|---|
| `No such file: period_0.xlsx` | Rename your files to `period_0.xlsx` and `period_1.xlsx` |
| `AuthenticationError` | Re-check your API key in Step 3 |
| `ModuleNotFoundError` | Run `pip install anthropic openpyxl pandas numpy` |
| Commentary blank in Excel | Text is there — drag the row height down to see it |
| Charts blank in HTML | Use Chrome or Firefox, not Internet Explorer |

---

## Limitations

- Actuals vs prior year only — no budget/forecast comparison
- Designed for this specific GL column structure
- AI commentary is a first draft — always review before sharing
- Portfolio/learning project — not production-ready

---

*Built in public · Finance × AI: In Practice · [alloutoftokens.com](https://alloutoftokens.com)*
