"""
01_launch_and_calculate.py
--------------------------
Launches PLAXIS 2D, opens a project file, waits for it to load,
lists all phases, runs the full calculation, and switches to Output.

Usage:
    Configure the USER SETTINGS section below, then run:
        python 01_launch_and_calculate.py

Requirements:
    - PLAXIS 2D 2024 installed
    - plxscripting (comes with PLAXIS installation)
    - A valid .p2dx project file
"""

from plxscripting.easy import *
import subprocess
import time
import os

# ── USER SETTINGS ────────────────────────────────────────────────────────────
PLAXIS_PATH = r'C:\Program Files\Seequent\PLAXIS 2D 2024\Plaxis2DXInput.exe'
FILE_PATH   = r'C:\Users\rufus\Desktop\from plan to plaxis\ff.p2dx'
PORT_i      = 10000   # Input server port
PORT_o      = 10001   # Output server port
PASSWORD    = 'SxDBR<TYKRAX834~'
LOAD_TIMEOUT_S = 60   # Max seconds to wait for the model to load
POLL_INTERVAL  = 2    # Seconds between readiness checks
# ─────────────────────────────────────────────────────────────────────────────


def launch_plaxis():
    """Start PLAXIS 2D with the scripting server enabled."""
    if not os.path.exists(FILE_PATH):
        raise FileNotFoundError(f"Project file not found: {FILE_PATH}")

    print(f"[1/4] Launching PLAXIS 2D...")
    subprocess.Popen(
        [PLAXIS_PATH,
         f'--AppServerPassword={PASSWORD}',
         f'--AppServerPort={PORT_i}'],
        shell=False
    )
    time.sleep(10)  # Allow the application to initialise


def connect_servers():
    """Connect to PLAXIS Input and Output scripting servers."""
    print("[2/4] Connecting to scripting servers...")
    s_i, g_i = new_server('localhost', PORT_i, password=PASSWORD)
    s_o, g_o = new_server('localhost', PORT_o, password=PASSWORD)
    return s_i, g_i, s_o, g_o


def open_and_wait(s_i, g_i):
    """Open the project file and poll until phases are accessible."""
    print(f"[3/4] Opening: {FILE_PATH}")
    s_i.open(FILE_PATH)

    print(f"      Waiting for model to load (timeout: {LOAD_TIMEOUT_S}s)...")
    attempts = LOAD_TIMEOUT_S // POLL_INTERVAL
    for attempt in range(attempts):
        time.sleep(POLL_INTERVAL)
        try:
            n_phases = len(list(g_i.Phases))
            elapsed = (attempt + 1) * POLL_INTERVAL
            print(f"      Ready — {n_phases} phase(s) detected after {elapsed}s")
            return
        except Exception as e:
            print(f"      Attempt {attempt + 1}/{attempts}: not ready ({e})")

    raise RuntimeError(
        f"Model did not load within {LOAD_TIMEOUT_S}s. "
        "Check that PLAXIS opened correctly and the file path is valid."
    )


def list_phases(g_i):
    """Print all phases available in the model."""
    g_i.gotostages()
    phases = list(g_i.Phases)
    print(f"\n      Phases found ({len(phases)}):")
    for i, p in enumerate(phases):
        print(f"        [{i}] {p.Identification.value}")
    return phases


def calculate_and_view(g_i):
    """Run the full calculation and open the Output viewer."""
    print("\n[4/4] Calculating all phases...")
    g_i.calculate()
    print("      Calculation complete.")

    g_i.view(g_i.Phases[1])
    print("      Output viewer opened on Phase[1].")


if __name__ == '__main__':
    launch_plaxis()
    s_i, g_i, s_o, g_o = connect_servers()
    open_and_wait(s_i, g_i)
    phases = list_phases(g_i)
    calculate_and_view(g_i)
    print("\nDone. Proceed to 02_extract_results.py to export data.")
