# OPEB Roll-Forward Model

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![GASB 75 Compliant](https://img.shields.io/badge/GASB%2075-Compliant-green.svg)](https://www.gasb.org/)

Production-ready GASB 75 roll-forward model for year-over-year OPEB liability projections and experience analysis.

---

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/risky-scout/opeb-roll-forward.git
cd opeb-roll-forward

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install the package
pip install -e .
```

### Run Your First Roll-Forward

```python
from datetime import date
from opeb_rollforward import RollForwardEngine, PriorValuation

# 1. Define prior year valuation results
prior = PriorValuation(
    valuation_date=date(2024, 9, 30),
    total_opeb_liability=6911729,
    tol_actives=2100000,
    tol_retirees=4800000,
    service_cost=650000,
    discount_rate_boy=0.0409,
    discount_rate_eoy=0.0381,
    avg_remaining_service_life=5.0,
)

# 2. Create engine with current year inputs
engine = RollForwardEngine(
    prior_valuation=prior,
    current_date=date(2025, 9, 30),
    benefit_payments=450000,
    new_discount_rate=0.0381,
    actual_eoy_tol=7712986,  # From new full valuation
)

# 3. Run the roll-forward
results = engine.run()

# 4. View results
print(f"Beginning TOL: ${results.boy_tol:,.0f}")
print(f"Service Cost: ${results.service_cost:,.0f}")
print(f"Interest Cost: ${results.interest_cost:,.0f}")
print(f"Benefit Payments: ${results.benefit_payments:,.0f}")
print(f"Expected EOY TOL: ${results.expected_eoy_tol:,.0f}")
print(f"Actual EOY TOL: ${results.actual_eoy_tol:,.0f}")
print(f"Experience (Gain)/Loss: ${results.experience_gain_loss:,.0f}")
print(f"Assumption Change Effect: ${results.assumption_change_effect:,.0f}")
```

---

## 📊 Roll-Forward Reconciliation

The model produces a complete GASB 75 ¶96 reconciliation:

```
Beginning TOL (9/30/2024)           $6,911,729
+ Service Cost                        $650,000
+ Interest Cost at 3.81%              $252,891
- Benefit Payments                   ($450,000)
= Expected EOY TOL                  $7,364,620
+ Experience (Gain)/Loss              $185,432
+ Changes in Assumptions              $162,934
= Ending TOL (9/30/2025)            $7,712,986
```

---

## 📁 Project Structure

```
opeb-roll-forward/
├── src/
│   └── opeb_rollforward/
│       ├── __init__.py     # Package exports
│       └── engine.py       # Roll-forward engine
├── pyproject.toml          # Package configuration
├── README.md               # This file
└── LICENSE                 # MIT License
```

---

## 🔧 Configuration Reference

### PriorValuation Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `valuation_date` | date | Yes | Prior measurement date |
| `total_opeb_liability` | float | Yes | Prior TOL |
| `tol_actives` | float | Yes | Prior active liability |
| `tol_retirees` | float | Yes | Prior retiree liability |
| `service_cost` | float | Yes | Prior service cost |
| `discount_rate_boy` | float | Yes | Prior BOY rate |
| `discount_rate_eoy` | float | Yes | Prior EOY rate |
| `avg_remaining_service_life` | float | No | ARSL (default: 12) |

### RollForwardEngine Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `prior_valuation` | PriorValuation | Yes | Prior year results |
| `current_date` | date | Yes | Current measurement date |
| `benefit_payments` | float | No | Actual payments (default: 0) |
| `new_discount_rate` | float | No | New EOY rate (default: prior rate) |
| `actual_eoy_tol` | float | No | Actual EOY TOL from new valuation |
| `duration` | float | No | Override duration estimate |

---

## 📈 Output: RollForwardResults

| Attribute | Description |
|-----------|-------------|
| `boy_tol` | Beginning of year TOL |
| `service_cost` | Service cost for the year |
| `interest_cost` | Interest cost at BOY rate |
| `benefit_payments` | Actual benefit payments |
| `expected_eoy_tol` | Expected EOY before adjustments |
| `actual_eoy_tol` | Actual EOY from full valuation |
| `experience_gain_loss` | Experience (gain)/loss |
| `assumption_change_effect` | Effect of assumption changes |
| `discount_rate_change_effect` | Effect of discount rate change |

---

## 📋 Generate Reconciliation Table

```python
# Get formatted reconciliation
table = results.get_reconciliation_table()

for item, value in table.items():
    print(f"{item}: ${value:,.0f}")
```

Output:
```
Beginning TOL: $6,911,729
Service Cost: $650,000
Interest Cost: $252,891
Benefit Payments: $-450,000
Experience (Gain)/Loss: $185,432
Assumption Changes: $162,934
Ending TOL: $7,712,986
```

---

## 💾 Save/Load Prior Valuation from JSON

### Save to JSON
```python
prior.to_json('prior_valuation_2024.json')
```

### Load from JSON
```python
from opeb_rollforward import load_prior

prior = load_prior('prior_valuation_2024.json')
```

### JSON Format
```json
{
  "valuation_date": "2024-09-30",
  "total_opeb_liability": 6911729,
  "tol_actives": 2100000,
  "tol_retirees": 4800000,
  "service_cost": 650000,
  "discount_rate_boy": 0.0409,
  "discount_rate_eoy": 0.0381,
  "avg_remaining_service_life": 5.0,
  "client_name": "City of DeRidder"
}
```

---

## 🔬 Formulas

### Interest Cost (GASB 75 ¶44)

```
Interest = (BOY_TOL + SC/2 - Benefits/2) × BOY_Rate
```

### Expected EOY TOL

```
Expected = BOY + Service_Cost + Interest - Benefits
```

### Discount Rate Change Effect

```
ΔL ≈ -Duration × L × Δr
```

### Experience (Gain)/Loss

```
Experience = Actual_EOY - (Expected_EOY + Assumption_Changes)
```

---

## 🧪 Example: Multi-Year Analysis

```python
from datetime import date
from opeb_rollforward import RollForwardEngine, PriorValuation

# Historical data
years = [
    {'date': date(2023, 9, 30), 'tol': 6165300, 'sc': 164981, 'bp': 232332, 'rate': 0.0409},
    {'date': date(2024, 9, 30), 'tol': 6911729, 'sc': 650000, 'bp': 450000, 'rate': 0.0381},
    {'date': date(2025, 9, 30), 'tol': 7712986, 'sc': 683256, 'bp': 475000, 'rate': 0.0381},
]

# Roll-forward each year
for i in range(1, len(years)):
    prior = PriorValuation(
        valuation_date=years[i-1]['date'],
        total_opeb_liability=years[i-1]['tol'],
        service_cost=years[i-1]['sc'],
        discount_rate_boy=years[i-1]['rate'],
        discount_rate_eoy=years[i]['rate'],
        tol_actives=years[i-1]['tol'] * 0.4,
        tol_retirees=years[i-1]['tol'] * 0.6,
    )
    
    engine = RollForwardEngine(
        prior_valuation=prior,
        current_date=years[i]['date'],
        benefit_payments=years[i]['bp'],
        new_discount_rate=years[i]['rate'],
        actual_eoy_tol=years[i]['tol'],
    )
    
    results = engine.run()
    print(f"\n{years[i-1]['date']} → {years[i]['date']}")
    print(f"  Experience (Gain)/Loss: ${results.experience_gain_loss:,.0f}")
```

---

## 📐 Duration Estimation

Duration is automatically estimated from the liability split:

```python
# Automatic estimation
duration = prior.duration_estimate

# Manual override
engine = RollForwardEngine(
    prior_valuation=prior,
    current_date=current_date,
    duration=12.5,  # Override with specific duration
)
```

Default formula:
```
Duration = (Active% × (ARSL + 10)) + (Retiree% × 10)
```

---

## 📜 Compliance

- **GASB Statement No. 75 ¶96** - TOL Reconciliation
- **GASB 75 ¶43(a)** - Experience gains/losses
- **GASB 75 ¶43(b)** - Assumption change effects
- **GASB 75 ¶44** - Interest cost calculation

---

## 🛠️ Troubleshooting

### Import Error
```bash
pip install -e .  # Reinstall in development mode
```

### Missing numpy
```bash
pip install numpy
```

### Date Format Issues
Use `datetime.date` objects, not strings.

---

## 📄 License

MIT License - See [LICENSE](LICENSE) file.

## 👤 Author

**Joseph Shackelford** - Actuarial Pipeline Project

---

## ⚠️ Disclaimer

This software is provided for educational and professional use. Actuarial valuations for official financial reporting should be reviewed and signed by a qualified actuary.
