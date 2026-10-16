"""
Background runner for the dashboard.

Usage (started by app.py, not by hand):
    python runner.py <run_dir> <script> [script args...]

It runs the RPA script in <run_dir>, writes the script's output to <run_dir>/log.txt and,
when the script ends, updates <run_dir>/meta.json. It does not depend on Streamlit, so the run
keeps going and is recorded even if the dashboard page is refreshed, you switch tabs, or the
Streamlit server restarts.
"""
import json
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path


def read_meta(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def write_meta(path: Path, meta: dict) -> None:
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(meta), encoding="utf-8")
    os.replace(tmp, path)                      # atomic: the dashboard never reads a half-written file


def main() -> None:
    run_dir = Path(sys.argv[1])
    script = sys.argv[2]
    args = sys.argv[3:]
    meta_path = run_dir / "meta.json"

    meta = read_meta(meta_path)
    meta["pid"] = os.getpid()                  # lets the dashboard check this run is alive and stop it
    write_meta(meta_path, meta)

    with open(run_dir / "log.txt", "ab") as log:
        proc = subprocess.Popen(
            [sys.executable, "-u", script, *args],
            cwd=run_dir, stdout=log, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
        )
        code = proc.wait()

    meta = read_meta(meta_path)
    if meta.get("status") == "running":        # if the dashboard marked it 'stopped', keep that
        started = datetime.fromisoformat(meta["started"])
        meta.update(
            status="finished" if code == 0 else "error",
            exit_code=code,
            finished=datetime.now().isoformat(timespec="seconds"),
            seconds=round((datetime.now() - started).total_seconds()),
        )
        write_meta(meta_path, meta)


if __name__ == "__main__":
    main()
