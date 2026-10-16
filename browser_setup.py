"""
Starts Chromium for the RPA scripts, and installs it first if the server does not have it yet.

Why: on hosting (for example Streamlit Community Cloud) `pip install playwright` installs the Python
package only - the Chromium browser is a separate download. This module (1) keeps the browser in a
folder next to the app, (2) installs it automatically the first time it is missing, and
(3) always runs headless on a Linux server, which has no screen.
"""
import os
import subprocess
import sys
from pathlib import Path

ON_LINUX = sys.platform.startswith("linux")

# Same folder the dashboard (app.py) uses. A value that is already set is kept.
if ON_LINUX:
    os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(Path(__file__).resolve().parent / ".pw-browsers"))


def _install_chromium() -> None:
    print("Chromium is not installed on this machine - installing it now (first run only, about a minute)...",
          flush=True)
    r = subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"],
                       capture_output=True, text=True, timeout=900)
    tail = ((r.stdout or "") + (r.stderr or "")).strip()[-400:]
    if tail:
        print(tail, flush=True)
    if r.returncode != 0:
        raise RuntimeError("Chromium could not be installed automatically. "
                           "Run:  python -m playwright install chromium")


def launch_browser(p):
    """Launch Chromium. HEADLESS=1 in the environment (set by the dashboard) hides the window."""
    headless = os.environ.get("HEADLESS", "0") == "1"
    kwargs = {}
    if ON_LINUX:
        if not os.environ.get("DISPLAY"):
            headless = True                      # a server has no screen to show a window on
        kwargs["args"] = ["--no-sandbox", "--disable-dev-shm-usage"]   # needed inside most containers
    try:
        return p.chromium.launch(headless=headless, **kwargs)
    except Exception as e:                       # noqa: BLE001
        msg = str(e)
        if "Executable doesn't exist" in msg or "playwright install" in msg:
            _install_chromium()
            return p.chromium.launch(headless=headless, **kwargs)
        if "shared libraries" in msg or "missing dependencies" in msg.lower():
            raise RuntimeError("The server is missing system libraries Chromium needs. "
                               "On Streamlit Cloud, list them in packages.txt and reboot the app.") from e
        raise
