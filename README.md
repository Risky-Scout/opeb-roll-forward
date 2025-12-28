# OPEB Roll-Forward Model

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![GASB 75 Compliant](https://img.shields.io/badge/GASB%2075-Compliant-green.svg)](https://www.gasb.org/)

**Production-ready GASB 75 roll-forward valuation system with Excel automation.**

Version 2.1.0 - West Florida Planning Corrections (2025-12-28)

---

## Quick Start

```bash
# Clone and install
git clone https://github.com/Risky-Scout/opeb-roll-forward.git
cd opeb-roll-forward
pip install -e .

# Run interactive mode
python run_roll_forward.py --interactive

# Or command line
python run_roll_forward.py \
    --input prior_year.xlsx \
    --output current_year.xlsx \
    --prior-date 2024-09-30 \
    --new-date 2025-09-30 \
    --prior-rate 0.0381 \
    --new-rate 0.0502 \
    --verify
```

---

## What is a Roll-Forward?

A roll-forward valuation projects the prior year's OPEB liability to the current measurement date without new census data. It's used when:

- A full valuation was performed in the prior year
- Census hasn't materially changed
- GASB 75 allows biennial full valuations with roll-forwards in between

**Roll-Forward Equation:**
```
EOY_TOL = BOY_TOL + Service_Cost + Interest + Assumption_Change + Experience
```

For pure roll-forwards, **Experience = $0** (no census data to generate gains/losses).

---

## Features

### 1. Actuarial Calculations
- **Interest Cost**: Mid-year approximation `(BOY + 0.5×SC) × rate`
- **Assumption Change**: Duration approximation `BOY × (1 - D × Δr) - BOY`
- **Sensitivities**: ±1% discount rate and healthcare trend

### 2. Excel Template Automation
- Updates existing GASB 75 Excel templates
- Handles all 10 worksheet tabs
- Preserves formula structures
- Copies cell formatting correctly

### 3. Quality Verification
- Automated verification checks
- Validates output formulas evaluate correctly

---

## Critical Fixes (2025-12-28)

This version incorporates all corrections from production debugging:

| # | Fix | Why It Matters |
|---|-----|----------------|
| 1 | Clear OPEB Exp & Def C6:C28 | Prevents leftover data in reports |
| 2 | Net OPEB D22:D25 No Fill | Consistent formatting |
| 3 | RSI I23 = "None" | Indicates no benefit changes |
| 4 | RSI I26 = VALUE | Must be number, not formula |
| 5 | B14/B26 copy FULL style | Formatting matches adjacent cells |
| 6 | B13 forces $0 | `=IF('Net OPEB'!D22<1,0,'Net OPEB'!D22)` |
| 7 | Skip Table7AmortDeferred2 | Not currently used |

---

## Usage

### Python API

```python
from datetime import date
from opeb_rollforward import run_roll_forward, print_roll_forward_summary

# Run complete roll-forward
output, results = run_roll_forward(
    input_path='GASB_75_2024.xlsx',
    output_path='GASB_75_2025.xlsx',
    prior_measurement_date=date(2024, 9, 30),
    new_measurement_date=date(2025, 9, 30),
    prior_discount_rate=0.0381,
    new_discount_rate=0.0502,
    duration=10.0,           # Liability duration
    trend_duration=5.0,      # Healthcare trend duration
    payroll_growth_rate=0.03, # 3% annual payroll growth
    benefit_changes="None",  # Or description of changes
)

# Print summary
print_roll_forward_summary(results, inputs)
```

### Command Line

```bash
# Interactive (prompts for all inputs)
python run_roll_forward.py --interactive

# Command line with all parameters
python run_roll_forward.py \
    --input "prior.xlsx" \
    --output "current.xlsx" \
    --prior-date 2024-09-30 \
    --new-date 2025-09-30 \
    --prior-rate 0.0381 \
    --new-rate 0.0502 \
    --duration 10.0 \
    --trend-duration 5.0 \
    --payroll-growth 0.03 \
    --benefit-changes "None" \
    --verify
```

---

## Module Structure

```
opeb-roll-forward/
├── run_roll_forward.py          # Main entry point
├── src/opeb_rollforward/
│   ├── __init__.py              # Package exports
│   ├── engine.py                # Core roll-forward engine
│   └── excel_updater.py         # Excel automation (NEW)
├── tests/
├── README.md
└── pyproject.toml
```

---

## Excel Template Structure

The module expects GASB 75 Excel templates with these worksheets:

| Sheet | Purpose |
|-------|---------|
| Model Inputs | Primary data input (TOL, rates, dates) |
| Net OPEB | TOL roll-forward reconciliation |
| RSI | Required Supplementary Information |
| OPEB Exp & Def | OPEB Expense and Deferred items |
| Table7AmortDeferred | Deferred inflows/outflows amortization |
| AmortDeferredOutsIns | ARSL tracking |
| 1%-+ | Sensitivity analysis |
| Assumptions | Actuarial assumptions |

---

## Verification Checklist

After running, open the output in Excel and verify:

| Check | Expected |
|-------|----------|
| Net OPEB D22 | $0 (for roll-forward) |
| Net OPEB D22:D25 | No background fill |
| Table7AmortDeferred AI49 | "GOOD" |
| OPEB Exp & Def H40 | "GOOD" |
| RSI I23 | "None" |
| RSI I26 | Discount rate (e.g., 5.02%) |
| B14, B26 formatting | Matches adjacent cells |

---

## Dependencies

- Python 3.10+
- openpyxl

```bash
pip install openpyxl
```

---

## License

MIT License - See LICENSE file

---

## Author

**Joseph Shackelford** - Actuarial Pipeline Project

---

## Changelog

### v2.1.0 (2025-12-28)
- Added `excel_updater.py` with production-ready Excel automation
- Incorporated all West Florida Planning corrections
- Added verification functions
- Added interactive mode runner
- Updated documentation

### v2.0.0
- Initial production release
- Core roll-forward engine
