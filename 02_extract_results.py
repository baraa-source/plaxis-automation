"""
02_extract_results.py
---------------------
Connects to a running PLAXIS 2D Output session, extracts bending moment
diagrams for a retaining wall plate across multiple excavation phases,
exports the data to Excel, and auto-generates comparison scatter charts.

Run AFTER 01_launch_and_calculate.py (PLAXIS must still be open).

Usage:
    Configure USER SETTINGS below, then run:
        python 02_extract_results.py

Requirements:
    - plxscripting, pandas, openpyxl, xlsxwriter
    - pip install pandas openpyxl xlsxwriter
"""

from plxscripting.easy import *
import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import ScatterChart, Reference, Series

# ── USER SETTINGS ────────────────────────────────────────────────────────────
PORT_i      = 10000
PORT_o      = 10001
PASSWORD    = 'SxDBR<TYKRAX834~'

OUTPUT_DIR  = r'C:\Users\rufus\Desktop\from plan to plaxis\\'
OUTPUT_NAME = 'Plate_y_Results.xlsx'

# Names must match exactly how they appear in your PLAXIS model
PLATE_NAME  = 'Plate_1'
PHASE_NAMES = [
    'Installation of strut [Phase_3]',
    'Second (submerged) excavation stage [Phase_4]',
    'Third excavation stage [Phase_5]',
]
# ─────────────────────────────────────────────────────────────────────────────

FILEPATH = OUTPUT_DIR + OUTPUT_NAME


def connect_output():
    """Connect to the PLAXIS Output scripting server."""
    print("[1/3] Connecting to PLAXIS Output server...")
    s_i, g_i = new_server('localhost', PORT_i, password=PASSWORD)
    s_o, g_o = new_server('localhost', PORT_o, password=PASSWORD)
    return g_i, g_o


def get_plate_results(g_o, plate_obj, phase_obj):
    """
    Extract Y-coordinates and bending moment (M2D) for a plate at a given phase.

    Returns a DataFrame sorted top-to-bottom (descending Y).
    """
    y_coords = g_o.getresults(plate_obj, phase_obj, g_o.ResultTypes.Plate.Y,   "node")
    moments  = g_o.getresults(plate_obj, phase_obj, g_o.ResultTypes.Plate.M2D, "node")

    phase_label = str(phase_obj.Identification).split('[')[0].strip()
    col_name    = f'Bending Moment [kNm/m]_{phase_label}'

    df = pd.DataFrame({'Y': y_coords, col_name: moments})
    return df.sort_values(by='Y', ascending=False).reset_index(drop=True)


def export_to_excel(g_o, filepath):
    """
    Extract results for all target phases, write per-phase sheets,
    and write a combined sheet with all phases side-by-side.
    """
    plates = list(g_o.Plates)
    phases = list(g_o.Phases)

    # Find the target plate object
    plate_obj = next((p for p in plates if p.Name == PLATE_NAME), None)
    if plate_obj is None:
        raise ValueError(
            f"Plate '{PLATE_NAME}' not found. "
            f"Available: {[p.Name for p in plates]}"
        )

    writer   = pd.ExcelWriter(filepath, engine='xlsxwriter')
    combined = []

    for phase_obj in phases:
        if phase_obj.Identification not in PHASE_NAMES:
            continue

        # Build a short sheet name from the phase tag, e.g. "Plate_1_Phase_4"
        tag = str(phase_obj.Identification).split('[')[-1].rstrip(']')
        sheet_name = f"{PLATE_NAME}_{tag}"

        df = get_plate_results(g_o, plate_obj, phase_obj)
        combined.append(df)
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"      Sheet written: {sheet_name}  ({len(df)} rows)")

    if not combined:
        raise RuntimeError(
            "No matching phases found. Check PHASE_NAMES against your model."
        )

    combined_sheet = f'combined_{PLATE_NAME}'
    pd.concat(combined, axis=1).to_excel(writer, sheet_name=combined_sheet, index=False)
    writer.close()
    print(f"      Combined sheet written: {combined_sheet}")
    return combined_sheet


def add_bm_chart(filepath, sheet_name):
    """
    Add a bending moment vs. RL (elevation) scatter chart to the combined sheet.
    All phases are plotted as separate series on a single chart for comparison.
    """
    df = pd.read_excel(filepath, sheet_name=sheet_name, engine='openpyxl')
    wb = load_workbook(filepath)
    ws = wb[sheet_name]
    n_rows = len(df) + 1  # +1 for header

    chart = ScatterChart()
    chart.title            = 'Bending moment envelope — all phases'
    chart.x_axis.title     = 'Bending Moment (kNm/m)'
    chart.y_axis.title     = 'RL (m)'
    chart.y_axis.tickLblPos = 'low'
    chart.legend.position  = 'b'
    chart.height           = 15
    chart.width            = 20

    # Identify column indices: Y columns provide x-axis, BM columns provide series
    y_col_idx = None
    for col_idx, col_name in enumerate(df.columns, start=1):
        if col_name == 'Y' or (col_name.startswith('Y') and '.' not in col_name):
            y_col_idx = col_idx

        if col_name.startswith('Bending Moment') and y_col_idx is not None:
            x_ref = Reference(ws, min_col=y_col_idx, min_row=2, max_row=n_rows)
            y_ref = Reference(ws, min_col=col_idx,   min_row=2, max_row=n_rows)
            series = Series(x_ref, y_ref, title=col_name)
            chart.series.append(series)

    ws.add_chart(chart, 'G1')
    wb.save(filepath)
    print(f"      Chart added to sheet '{sheet_name}' at cell G1.")


if __name__ == '__main__':
    g_i, g_o = connect_output()

    print("[2/3] Extracting results and writing Excel...")
    combined_sheet = export_to_excel(g_o, FILEPATH)

    print("[3/3] Generating comparison chart...")
    add_bm_chart(FILEPATH, combined_sheet)

    print(f"\nDone! Results saved to:\n  {FILEPATH}")
