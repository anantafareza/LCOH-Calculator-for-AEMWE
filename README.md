# LCOH Calculator for AEMWE

A desktop techno-economic analysis (TEA) tool for computing the **Levelised Cost of Hydrogen (LCOH)** from Anion Exchange Membrane Water Electrolysis (AEMWE) systems. The calculator evaluates LCOH across a range of operating current densities, identifying the optimal operating point that minimises hydrogen production cost.

The tool integrates capital expenditure (CAPEX), operating expenditure (OPEX), and stack degradation modelling, with support for custom Bill of Materials (BOM) and MEA cost estimation. Life cycle emissions (LCE) analysis and probabilistic Monte Carlo uncertainty quantification are included for comprehensive techno-economic assessment.

---

## Features

- LCOH calculation across a full polarisation curve operating range
- Bill of Materials / MEA cost breakdown with custom catalyst synthesis cost estimation
- Stack lifetime degradation modelling as a function of current density
- LCE (Life Cycle Emissions) analysis with break-even carbon cost visualisation
- Sensitivity analysis with ±20% parameter sliders
- Monte Carlo uncertainty analysis (Normal, Log-normal, Gamma distributions)
- Export results to CSV, TXT, JSON, or Excel

---

## Requirements

- Python 3.10 or later
- Dependencies listed in `requirements.txt`

---

## Installation

**1. Clone the repository**

```bash
git clone https://github.com/yourusername/lcoh-aemwe-calculator.git
cd lcoh-aemwe-calculator
```

**2. (Recommended) Create a virtual environment**

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate
```

**3. Install dependencies**

```bash
pip install -r requirements.txt
```

**4. Run the app**

```bash
python LCOH_AEMWE.py
```

---

## Usage

1. **Inputs tab** — set plant design targets, financial parameters, and the polarisation curve
2. **BOM / MEA Materials tab** — select materials and costs for each MEA component
3. **Anode / Cathode Catalyst Synthesis Cost tabs** — unlock by setting Anode or Cathode Catalyst to "Custom" in BOM
4. **LCOH Results tab** — click **Calculate LCOH** to run the model and view results
5. **LCE Analysis tab** — click **Link from LCOH Results**, set grey H₂ reference values, and visualise break-even carbon pricing
6. **Sensitivity Analysis tab** — click **Set Base Case**, adjust sliders, and run Monte Carlo uncertainty analysis

---

## File structure

```
lcoh-aemwe-calculator/
├── LCOH_AEMWE.py        # Main application
├── requirements.txt     # Python dependencies
└── README.md            # This file
```

---

## Licence

MIT Licence. See [LICENSE](LICENSE) for details.

---

## Contact

Ananta Fareza
