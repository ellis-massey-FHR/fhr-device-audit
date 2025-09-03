# service_runner.py
import os
import time
import datetime as dt
import subprocess
import traceback

# =================== SETTINGS ===================
# Times to run (24h). Change as needed.
RUN_TIMES = os.getenv("RUN_TIMES", "08:00,13:00").split(",")  # e.g., "08:00,13:00"
RUN_TIMES = [t.strip() for t in RUN_TIMES if t.strip()]

# If True: if the service starts after a scheduled time today,
# it will run the missed job immediately (once).
CATCH_UP = os.getenv("CATCH_UP", "true").lower() == "true"

# Choose how to execute your job:
#   "python" -> call your .py file with python
#   "exe"    -> call your built .exe (no Python needed)
RUN_MODE = os.getenv("RUN_MODE", "python").lower()

# Paths (edit these to match your environment)
PYTHON_EXE = os.getenv("PYTHON_EXE", "python")  # or full path to python.exe
PY_SCRIPT  = os.getenv("PY_SCRIPT", r"C:\Users\ellism3\fhr-device-audit\scripts\get_servicenow_data.py")
EXE_PATH   = os.getenv("EXE_PATH",  r"C:\Users\ellism3\fhr-device-audit\scripts\get_servicenow_data.exe")

# Where to write a tiny state file so we don't double-run
STATE_DIR  = os.getenv("STATE_DIR", r"C:\ProgramData\FHRReportService")
STATE_FILE = os.path.join(STATE_DIR, "last_runs.json")

# Minimal log (NSSM can also capture stdout/err; this is a backup)
LOG_FILE   = os.getenv("LOG_FILE",  r"C:\ProgramData\FHRReportService\runner.log")
# ================================================

import json
from pathlib import Path

def log(msg: str):
    Path(STATE_DIR).mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{stamp}] {msg}\n")

def parse_time(tstr: str) -> dt.time:
    return dt.datetime.strptime(tstr, "%H:%M").time()

def run_job():
    start = dt.datetime.now()
    log("Starting job...")
    try:
        if RUN_MODE == "exe":
            proc = subprocess.run([EXE_PATH], capture_output=True, text=True)
        else:
            proc = subprocess.run([PYTHON_EXE, PY_SCRIPT], capture_output=True, text=True)

        # Log outputs (useful if NSSM logs aren’t set)
        if proc.stdout:
            log("STDOUT:\n" + proc.stdout)
        if proc.stderr:
            log("STDERR:\n" + proc.stderr)

        if proc.returncode == 0:
            log(f"Job finished OK in {(dt.datetime.now()-start).total_seconds():.1f}s")
        else:
            log(f"Job FAILED with code {proc.returncode}")
    except Exception:
        log("Job crashed with exception:\n" + traceback.format_exc())

def load_state() -> dict:
    if not os.path.exists(STATE_FILE):
        return {}
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_state(state: dict):
    Path(STATE_DIR).mkdir(parents=True, exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f)

def today_key():
    return dt.date.today().isoformat()

def main_loop():
    schedule_times = [parse_time(t) for t in RUN_TIMES]
    state = load_state()
    already_ran = set(state.get(today_key(), []))

    # Optional catch-up: if service starts after a scheduled time today, run it now
    now = dt.datetime.now()
    if CATCH_UP:
        for t in schedule_times:
            if t <= now.time() and t.strftime("%H:%M") not in already_ran:
                log(f"Catch-up run for {t.strftime('%H:%M')}")
                run_job()
                already_ran.add(t.strftime("%H:%M"))
        state[today_key()] = sorted(list(already_ran))
        save_state(state)

    while True:
        now = dt.datetime.now()

        # New day -> reset
        if today_key() not in state:
            already_ran = set()
            state[today_key()] = []
            save_state(state)

        # Run if current time is within the current minute of any target time
        for t in schedule_times:
            key = t.strftime("%H:%M")
            if key in already_ran:
                continue

            # Is it that minute?
            should_run_now = (now.hour == t.hour and now.minute == t.minute)
            if should_run_now:
                log(f"Scheduled run triggered for {key}")
                run_job()
                already_ran.add(key)
                state[today_key()] = sorted(list(already_ran))
                save_state(state)
                # Avoid re-running within the same minute
                time.sleep(65)

        # Midnight reset (lightweight)
        if dt.datetime.now().date() != now.date():
            already_ran = set()
            state = load_state()  # reload, in case anything changed

        time.sleep(20)  # check ~3x per minute

if __name__ == "__main__":
    log("Service runner starting up")
    try:
        main_loop()
    except KeyboardInterrupt:
        log("Service runner received KeyboardInterrupt, exiting...")
    except Exception:
        log("Service runner crashed:\n" + traceback.format_exc())
