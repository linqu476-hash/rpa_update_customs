"""
Starts a browser for the RPA scripts - as fast as possible, installing something only if it must.

Order tried by launch_browser():
  1. Playwright's own Chromium, if it is already installed (instant).
  2. A browser already on the machine (no download):
       Linux            -> system Chromium (listed in packages.txt on Streamlit Cloud, or CHROMIUM_PATH)
       Windows / macOS  -> Microsoft Edge, then Google Chrome
  3. Download Chromium. On a Linux server only the small headless build is fetched.

On Linux the browser folder is kept next to the app, and the browser always runs headless
(a server has no screen).
"""
import os
import shutil
import subprocess
import sys
from pathlib import Path

ON_LINUX = sys.platform.startswith("linux")

# Same folder the dashboard (app.py) uses. A value that is already set is kept.
if ON_LINUX:
    os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(Path(__file__).resolve().parent / ".pw-browsers"))


def find_system_chromium():
    """Path of a Chromium / Chrome already installed on this Linux machine, or None."""
    env_path = os.environ.get("CHROMIUM_PATH")
    if env_path and Path(env_path).exists():
        return env_path
    for name in ("chromium", "chromium-browser", "google-chrome", "google-chrome-stable"):
        found = shutil.which(name)
        if found:
            return found
    return None


def _launch_system(p, headless, kwargs):
    """Launch a browser that is already installed. Returns the browser, or None if there is none."""
    if ON_LINUX:
        path = find_system_chromium()
        if path:
            try:
                b = p.chromium.launch(headless=headless, executable_path=path, **kwargs)
                print(f"Using the browser already installed on this machine: {path}", flush=True)
                return b
            except Exception as e:                       # noqa: BLE001
                print(f"System browser {path} could not be started: {str(e).splitlines()[0]}", flush=True)
        return None
    for channel in ("msedge", "chrome"):                 # Windows / macOS
        try:
            b = p.chromium.launch(headless=headless, channel=channel, **kwargs)
            print(f"Using the installed browser: {channel}", flush=True)
            return b
        except Exception:                                # noqa: BLE001
            continue
    return None


def _install_chromium(headless):
    # On a server only the headless build is needed - about half the download.
    target = "chromium-headless-shell" if (ON_LINUX and headless) else "chromium"
    print(f"No browser found - downloading {target} (first run only)...", flush=True)
    r = subprocess.run([sys.executable, "-m", "playwright", "install", target],
                       capture_output=True, text=True, timeout=900)
    if r.returncode != 0 and target != "chromium":       # older Playwright without the shell-only target
        print("Smaller download not available - downloading full Chromium...", flush=True)
        r = subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"],
                           capture_output=True, text=True, timeout=900)
    tail = ((r.stdout or "") + (r.stderr or "")).strip()[-300:]
    if tail:
        print(tail, flush=True)
    if r.returncode != 0:
        raise RuntimeError("The browser could not be installed automatically. "
                           "Run:  python -m playwright install chromium")


def launch_browser(p, headless=None):
    """Launch a browser. HEADLESS=1 in the environment (set by the dashboard) hides the window."""
    if headless is None:
        headless = os.environ.get("HEADLESS", "0") == "1"
    kwargs = {}
    if ON_LINUX:
        if not os.environ.get("DISPLAY"):
            headless = True                              # a server has no screen to show a window on
        kwargs["args"] = ["--no-sandbox", "--disable-dev-shm-usage"]   # needed inside most containers
    try:
        return p.chromium.launch(headless=headless, **kwargs)           # 1. Playwright's Chromium
    except Exception as e:                               # noqa: BLE001
        msg = str(e)
        if "shared libraries" in msg or "missing dependencies" in msg.lower():
            raise RuntimeError("The server is missing system libraries Chromium needs. "
                               "On Streamlit Cloud, list them in packages.txt and reboot the app.") from e
        if "Executable doesn't exist" not in msg and "playwright install" not in msg:
            raise
    browser = _launch_system(p, headless, kwargs)        # 2. a browser that is already installed
    if browser:
        return browser
    _install_chromium(headless)                          # 3. download
    return p.chromium.launch(headless=headless, **kwargs)


def prepare_browser():
    """Make sure a browser can start (installing one if needed), then close it. Returns a short note.
    The dashboard calls this once in the background when it starts, so the first run does not wait."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = launch_browser(p, headless=True)
        browser.close()
    return "ready"
