# Monte Carlo Schedule Risk Simulation

A Python-based Monte Carlo schedule risk simulator that matches **Primavera Risk Analysis (OPRA)** methodology. Runs three-point triangular duration distributions on a CPM activity network and produces standard risk analysis outputs.

## What It Does

1. Reads a deterministic CPM schedule (activity network with durations and FS predecessors)
2. Renders a Gantt chart of the deterministic baseline
3. Applies triangular duration distributions (optimistic / most likely / pessimistic) per activity
4. Runs N iterations of the CPM forward pass with randomly sampled durations
5. Produces: histogram, S-curve (CDF), tornado sensitivity chart, and a formatted Excel workbook with results

## Quick Start

```bash
pip install -r requirements.txt

# Step 1: Render the deterministic Gantt chart
python3 scripts/render_gantt.py \
  --schedule data/example/summit-tower-activities.csv

# Step 2: Create the Excel workbook from your schedule + risk parameters
python3 scripts/monte_carlo.py \
  --schedule data/example/summit-tower-activities.csv \
  --risk-params data/example/summit-tower-risk-params.csv \
  --init

# Step 3: Review/edit the "Duration Uncertainty" sheet in Excel
# Adjust optimistic/pessimistic factors based on your risk register

# Step 4: Run the simulation
python3 scripts/monte_carlo.py \
  --schedule data/example/summit-tower-activities.csv \
  --workbook data/example/summit-tower-activities-monte-carlo.xlsx
```

All outputs land in `outputs/`. The Excel workbook's Results and Sensitivity sheets are updated in place.

## Input Format

### Schedule CSV

Standard CPM activity network. All relationships are finish-to-start, zero lag.

```csv
wbs,wbs_name,activity_id,activity_name,duration_days,predecessors
1.0,Pre-Construction,A101,Project kickoff,5,
1.0,Pre-Construction,A102,Design coordination,60,A101
1.0,Pre-Construction,A103,Permit submission,30,A102
```

- `predecessors` — comma-separated activity IDs (FS zero-lag)
- Activities with no predecessors are project starts

### Risk Parameters CSV

Three-point estimates as multipliers on the baseline duration.

```csv
activity_id,optimistic_factor,most_likely_factor,pessimistic_factor,notes
A101,0.80,1.00,1.40,Kickoff — low variance
A102,0.85,1.00,1.50,Design coordination — scope creep risk
A103,0.90,1.00,1.30,Permit submission — mostly administrative
```

- `optimistic_factor` — best case (e.g., 0.80 = 80% of baseline)
- `most_likely_factor` — typically 1.00
- `pessimistic_factor` — worst case (e.g., 2.00 = 200% of baseline)
- `notes` — trace which risks from your register justify the distribution width

## Outputs

| Output | Description |
|---|---|
| `gantt.png` | Deterministic CPM schedule Gantt chart |
| `monte-carlo-histogram.png` | Distribution of project finish dates with P50/P80/P90 markers |
| `monte-carlo-scurve.png` | Cumulative probability curve (CDF) with confidence annotations |
| `monte-carlo-tornado.png` | Top 15 activities by Spearman rank correlation with finish |
| Excel "Results" sheet | Full percentile table (P5 through P95), mean, std dev, min/max |
| Excel "Sensitivity" sheet | All activities ranked by correlation with interpretation labels |

## CLI — Gantt Chart

```
python3 scripts/render_gantt.py --schedule PATH [--output PATH]
```

## CLI — Monte Carlo Simulation

```
python3 scripts/monte_carlo.py --schedule PATH --risk-params PATH --init
python3 scripts/monte_carlo.py --schedule PATH [--workbook PATH] [-n N] [--seed N] [--output DIR]
```

| Flag | Description |
|---|---|
| `--schedule` | Schedule CSV — activity network (always required) |
| `--risk-params` | Risk parameters CSV (required with `--init`) |
| `--workbook` | Excel workbook path (auto-generated if omitted) |
| `--output` | Output directory (default: `outputs/`) |
| `--init` | Create workbook from CSVs |
| `-n, --iterations` | Number of iterations (default: 10,000) |
| `--seed` | Random seed (default: 42) |

## Workflow: Edit Assumptions → Rerun

The Excel workbook is both input and output:

1. **Duration Uncertainty** sheet (blue tab) — edit optimistic/pessimistic factors here
2. Rerun `monte_carlo.py` (without `--init`)
3. **Results** sheet (green tab) and **Sensitivity** sheet (orange tab) are overwritten with new results
4. Charts in `outputs/` are regenerated

This is the same loop you'd use in OPRA: adjust three-point estimates, re-analyze, review outputs.

## How It Matches OPRA

This tool replicates the core methodology of Oracle Primavera Risk Analysis:

- **Duration uncertainty** modeled as three-point triangular distributions per activity
- **CPM forward pass** computed each iteration (same as OPRA's schedule engine)
- **No discrete risk events** — risk register findings are baked into the distribution parameters (wider pessimistic tails on risk-exposed activities)
- **Standard outputs** — histogram, S-curve, tornado chart are the same outputs OPRA produces

The key difference: assumptions are fully transparent in a single Excel workbook rather than buried in software dialogs.

## Example: Summit Tower

The `data/example/` directory contains a complete worked example — a 41-story, 489-ft residential tower in Newark, NJ with 44 activities and 15 identified risks across 9 categories. The risk register findings are reflected in the duration distributions (see the `notes` column in `summit-tower-risk-params.csv`).

## Dependencies

- Python 3.10+
- numpy
- matplotlib
- openpyxl
