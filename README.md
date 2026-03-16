# PLAXIS 2D Automation — Python Scripting for Retaining Wall Analysis

Automates the full numerical modelling workflow for a braced excavation retaining wall in PLAXIS 2D: launching the solver, running staged construction calculations, extracting bending moment envelopes, and exporting comparison charts to Excel — all from a single Python environment.

Built as part of MSc geotechnical engineering research. Designed to replace repetitive manual interaction with PLAXIS for parametric studies and phase-by-phase result extraction.

---

## What this does

| Step | Script | What happens |
|---|---|---|
| 1 | `01_launch_and_calculate.py` | Launches PLAXIS 2D, opens a `.p2dx` model, polls until loaded, lists phases, runs full calculation |
| 2 | `02_extract_results.py` | Connects to PLAXIS Output, extracts Y-coordinates and bending moments (M2D) per phase, writes Excel with per-phase sheets + combined sheet + scatter chart |

The two scripts connect via the **PLAXIS scripting server** (`plxscripting`) over localhost — no GUI clicks required after the initial launch.

---

## Geotechnical context

The workflow targets a **braced excavation with a retaining wall plate**, modelled using staged construction in PLAXIS 2D. Typical use case:

- Sheet pile or diaphragm wall supporting a multi-stage excavation
- Strut or anchor installed at intermediate stages
- Bending moment envelope extracted at each excavation stage to check structural demand vs. capacity

The bending moment diagram (M vs. RL) is the key design output. This script automates its extraction across all phases and overlays them in a single Excel chart for envelope assessment.

**Model geometry** — two-layer soil profile (sand over clay), retaining wall plate with fixed-end anchors, water table at excavation level:

![PLAXIS 2D model](results/plaxis_model.jpg)

---

## Requirements

- **PLAXIS 2D 2024** (or compatible version) — must be installed and licensed
- **Python 3.9+**
- `plxscripting` — bundled with PLAXIS, no separate install needed
- Python packages:

```
pip install pandas openpyxl xlsxwriter
```

---

## Setup

1. Clone or download this repository.
2. Open `01_launch_and_calculate.py` and update the **USER SETTINGS** block:

```python
PLAXIS_PATH = r'C:\Program Files\Seequent\PLAXIS 2D 2024\Plaxis2DXInput.exe'
FILE_PATH   = r'C:\path\to\your\model.p2dx'
PASSWORD    = 'your-server-password'
```

3. Open `02_extract_results.py` and update the **USER SETTINGS** block:

```python
OUTPUT_DIR  = r'C:\path\to\output\folder\\'
PLATE_NAME  = 'Plate_1'           # must match your model's plate name
PHASE_NAMES = [
    'Installation of strut [Phase_3]',
    'Second (submerged) excavation stage [Phase_4]',
    'Third excavation stage [Phase_5]',
]
```

> **Tip:** Phase names must match exactly as they appear in PLAXIS. Run script 1 first — it prints all phase names so you can copy them.

---

## Usage

Run the scripts in order. PLAXIS must remain open between them.

```bash
# Step 1: launch PLAXIS, load model, calculate
python 01_launch_and_calculate.py

# Step 2: extract results and write Excel (while PLAXIS is still open)
python 02_extract_results.py
```

### Example output

`Plate_y_Results.xlsx` contains:
- One sheet per phase (e.g. `Plate_1_Phase_3`, `Plate_1_Phase_4`)
- A `combined_Plate_1` sheet with all phases side by side
- A scatter chart on the combined sheet: **Bending Moment (kNm/m) vs. RL (m)** for all phases overlaid

The script produces a multi-sheet Excel workbook. The `combined_Plate_1` sheet contains all phases side by side and an auto-generated scatter chart:

![Bending moment envelope — all phases](results/bending_moment_chart.jpg)

Three series are plotted: strut installation (blue), second submerged excavation (red), and third excavation stage (green). The chart immediately shows how the moment distribution evolves and where peak demand occurs — useful for checking the wall against ULS bending capacity at each construction stage.

---

## Project structure

```
plaxis-automation/
├── 01_launch_and_calculate.py   # Launch PLAXIS, open model, run calculation
├── 02_extract_results.py        # Extract bending moments, export to Excel
├── requirements.txt
└── results/
    ├── plaxis_model.jpg          # PLAXIS 2D model screenshot
    └── bending_moment_chart.jpg  # Excel output chart (all phases)
```

---

## Key scripting API calls

| Call | Purpose |
|---|---|
| `new_server(host, port, password)` | Connect to PLAXIS scripting server |
| `s_i.open(filepath)` | Open a `.p2dx` project file |
| `g_i.gotostages()` | Switch to staged construction mode |
| `g_i.calculate()` | Run the full calculation queue |
| `g_i.view(phase)` | Open Output viewer at a specific phase |
| `g_o.getresults(obj, phase, ResultType, "node")` | Extract nodal results for a structural element |

---

## Extending this workflow

Some directions this can be taken further:

- **Parametric study:** loop over a range of strut stiffness or excavation depths, run `g_i.calculate()` after each change, and collect bending moment peaks into a summary table
- **Shear force & deflection:** add `g_o.ResultTypes.Plate.Q2D` and `g_o.ResultTypes.Plate.Utot` to the extraction function
- **Multiple plates:** pass a list of plate names and loop `get_plate_results()` across all of them
- **Automated PDF report:** pipe the Excel chart into a `reportlab` or `matplotlib` PDF summary

---

## Author

MSc Geotechnical Engineering  
Python scripting for numerical geotechnical modelling (PLAXIS 2D)
email: chalalbaraa@gmail.com

---

## License

MIT — free to use and adapt. Attribution appreciated.
