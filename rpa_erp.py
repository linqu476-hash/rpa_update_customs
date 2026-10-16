"""
RPA_Invoice_SMS_ERP - Step 1: Sign in to D365 Production (ADFS)
Step 2: Open Export Invoice (Bulk), filter by invoice no., open Invoice Header tab

Usage:  python rpa_erp.py SC260309

Setup:
    pip install playwright python-dotenv
    playwright install chromium

Credentials are read from environment variables / .env file (see .env.example).
Do NOT hardcode them in this script or in the README.
"""
import json
import os
import re
import sys
from dotenv import load_dotenv

from browser_setup import launch_browser   # sets the browser folder + installs Chromium if missing
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

load_dotenv()

# Use the base app URL, NOT the long ADFS URL with state/nonce
# (that one expires and will fail on the next run). The app redirects to ADFS itself.
ERP_URL = "https://ax.d365ffo.onprem.libgroup.com/namespaces/AXSF/"
# ---- Step 3: values to fill on the Invoice Header tab (label as shown in D365, value) ----
HEADER_FIELDS = [
    # (label as shown in D365, text to type, kind, [dropdown row texts to click - optional])
    # kind: "text" = type + Tab;  "lookup" = type, then click the matching dropdown row
    ("Invoice Date",      "1/11/2026",  "text"),
    ("Booking Date",      "1/11/2026",  "text"),
    ("Exfact Date",       "1/11/2026",  "text"),
    ("FGR Date",          "1/11/2026",  "text"),     # value not given in the request - change/remove if different
    ("ConsignId",         "400447-SMS", "lookup"),   # must be number-SMS format
    ("From",              "Pnh-air",    "lookup"),
    ("Ship To",           "DNK",        "lookup", ["DNK", "Danmark"]),        # Name fills in as Danmark
    ("SHIP MODE",         "MO-010",     "lookup", ["MO-010", "Courier-CC"]),
    ("Forwarder",         "DHL-ex",     "lookup"),
    ("Term of delivery",  "FOB",        "lookup"),
    ("Customer",          "AD",         "lookup"),
    ("Bank ID",           "CIMB",       "lookup"),
    ("Custom Type",       "Clearance",  "lookup"),
    ("Shipping",          "DHL-ex",     "lookup"),
    ("CO Type",           "Rex",        "lookup"),
]

# The dashboard can override the list above (JSON in the HEADER_FIELDS_JSON environment variable)
if os.environ.get("HEADER_FIELDS_JSON"):
    HEADER_FIELDS = json.loads(os.environ["HEADER_FIELDS_JSON"])

USERNAME = os.environ["ERP_USER"]
PASSWORD = os.environ["ERP_PASSWORD"]


def login(page):
    page.goto(ERP_URL, wait_until="domcontentloaded", timeout=90_000)

    # Screenshot step 1: username field (standard ADFS ids shown first, generic fallback second)
    user_box = page.locator("#userNameInput, input[type='text'], input[type='email']").first
    user_box.wait_for(state="visible", timeout=60_000)
    user_box.fill(USERNAME)

    # Screenshot step 2: password field
    page.locator("#passwordInput, input[type='password']").first.fill(PASSWORD)

    # Screenshot step 3: Sign in button
    page.locator("#submitButton, span:has-text('Sign in'), input[type='submit']").first.click()

    # Wait up to 2 minutes for ONE of: the D365 app opens (success) - or ADFS rejects the sign-in.
    for _ in range(240):
        url = page.url
        if "/namespaces/AXSF" in url and "/adfs/" not in url.lower():
            break                                                        # signed in
        err = page.locator("#errorText:visible")
        if err.count() > 0 and err.first.inner_text().strip():
            raise RuntimeError("Sign-in was rejected by D365: " + err.first.inner_text().strip() +
                               " - check the user and password (repeated wrong attempts can lock the account).")
        page.wait_for_timeout(500)
    else:
        page.screenshot(path="login_timeout.png")
        try:
            text = page.inner_text("body")[:200].replace("\n", " ")
        except Exception:                                                # noqa: BLE001
            text = "(page text not readable)"
        raise RuntimeError(f"Sign-in did not finish within 2 minutes. The browser is on: {url} | "
                           f"the page says: {text}")

    # D365 keeps background requests running, so 'networkidle' may never come - do not depend on it.
    for state, secs in (("load", 60), ("networkidle", 8)):
        try:
            page.wait_for_load_state(state, timeout=secs * 1000)
        except PWTimeout:
            pass


def dismiss_overlays(page, wait_seconds=15):
    """Close the 'About Finance and Operations' pop-up shown on the dashboard.
    Waits up to `wait_seconds` for it to appear, closes it, confirms it is gone."""
    popup_marker = page.get_by_text("Installed product version", exact=False)
    try:
        popup_marker.first.wait_for(state="visible", timeout=wait_seconds * 1000)
    except PWTimeout:
        print("No pop-up shown - continuing")
        return

    print("Pop-up detected - closing it")
    close_selectors = [
        "button:has-text('Close')",
        "[role='button']:has-text('Close')",
        "[data-dyn-controlname*='Close' i]",
    ]
    for sel in close_selectors:
        btn = page.locator(sel).last      # the pop-up's Close is the last one on the page
        if btn.count() > 0 and btn.is_visible():
            try:
                btn.click(timeout=5_000)
                break
            except PWTimeout:
                continue
    else:
        page.keyboard.press("Escape")      # fallback 1
        page.wait_for_timeout(500)
        if popup_marker.first.is_visible():
            page.mouse.click(1869, 1035)   # fallback 2: Close button position (1920x1080)

    popup_marker.first.wait_for(state="hidden", timeout=10_000)
    page.wait_for_timeout(500)
    print("Pop-up closed")


def open_invoice_header(page, invoice_no):
    # Optional shortcut: open the form directly by URL (most reliable).
    # In the browser, open Export Invoice (Bulk), copy the address bar URL
    # into ERP_FORM_URL in .env. Then steps 1-2 are skipped.
    form_url = os.environ.get("ERP_FORM_URL", "").strip()
    if form_url:
        page.goto(form_url, wait_until="domcontentloaded")
    else:
        page.wait_for_timeout(1500)
        target = page.get_by_text("Export Invoice (Bulk)", exact=True)

        # 1. Open the navigation pane (hamburger) - try visible selectors, then position
        print("Step 2.1: opening navigation pane")
        candidates = [
            "[data-dyn-controlname='NavPaneButton']:visible",
            "button.navigationBar-toggleButton:visible",
            "button[aria-label*='avigation' i]:visible",
            "button[title*='avigation' i]:visible",
        ]
        opened = False
        for sel in candidates:
            loc = page.locator(sel)
            if loc.count() == 0:
                continue
            try:
                loc.first.click(timeout=3_000)
                target.first.wait_for(state="visible", timeout=4_000)
                print("  opened with:", sel)
                opened = True
                break
            except PWTimeout:
                continue
        if not opened:
            print("  selectors failed - clicking hamburger by position")
            page.mouse.click(24, 69)
            try:
                target.first.wait_for(state="visible", timeout=5_000)
                opened = True
            except PWTimeout:
                page.screenshot(path="nav_debug.png")
                raise RuntimeError("Navigation pane did not open - see nav_debug.png")

        # 2. Favorites -> "Export Invoice (Bulk)"
        print("Step 2.2: clicking Export Invoice (Bulk)")
        target.first.click()
    print("Step 2.3: waiting for Export Invoice list")
    page.wait_for_selector("text=Export Invoice List >> visible=true", timeout=60_000)

    # 3. Click the "Export Invoice Id" column header to open the filter flyout
    print("Step 2.4: opening column filter")
    page.get_by_role("columnheader", name="Export Invoic").first.click()

    # 4. Type invoice number in the filter flyout, then Apply
    #    (operator is already "begins with" by default)
    apply_btn = page.get_by_role("button", name="Apply", exact=True).first
    try:
        apply_btn.wait_for(state="visible", timeout=15_000)
    except PWTimeout:
        page.screenshot(path="filter_debug.png")
        raise RuntimeError("Filter flyout did not open - see filter_debug.png")

    print("Step 2.5: filtering", invoice_no)
    # Find the input that lives in the same flyout as the Apply button
    flyout_input = apply_btn.locator("xpath=ancestor::*[.//input][1]//input").first
    try:
        flyout_input.click(timeout=3_000)
        flyout_input.fill(invoice_no)
    except PWTimeout:
        # Fallback: the box is auto-focused when the flyout opens, so just type
        page.keyboard.press("Control+A")
        page.keyboard.type(invoice_no, delay=50)
    page.wait_for_timeout(500)
    apply_btn.click()
    # Wait for the filter flyout to close and the grid to reload.
    # (D365 grid cells are <input> boxes, so get_by_text() cannot see the value.)
    apply_btn.wait_for(state="hidden", timeout=15_000)
    page.wait_for_timeout(1500)
    try:
        page.wait_for_load_state("networkidle", timeout=10_000)
    except PWTimeout:
        pass
    shown = page.locator(f"input[value='{invoice_no}' i], :text-is('{invoice_no}')")
    try:
        shown.first.wait_for(state="visible", timeout=10_000)
        print("  invoice found in grid")
    except PWTimeout:
        page.screenshot(path="after_filter.png")
        print("  WARNING: could not confirm invoice in grid - see after_filter.png; continuing")

    # 5. Click the "Invoice Header" tab and confirm it opened
    print("Step 2.6: clicking Invoice Header tab")
    # Fields that only exist on the Invoice Header tab
    header_field = (
        page.get_by_text("Shipment Detail", exact=True)
        .or_(page.get_by_text("Customs No", exact=True))
        .or_(page.get_by_text("Notify Party 1", exact=True))
    ).first

    attempts = [
        page.get_by_role("tab", name="Invoice Header"),
        page.get_by_text("Invoice Header", exact=True),
    ]
    opened = False
    for tab in attempts:
        if tab.count() == 0:
            continue
        try:
            tab.first.click(timeout=5_000)
            header_field.wait_for(state="visible", timeout=10_000)
            opened = True
            break
        except PWTimeout:
            continue
    if not opened:
        page.screenshot(path="header_debug.png")
        raise RuntimeError("Invoice Header tab did not open - see header_debug.png")
    try:
        page.wait_for_load_state("networkidle", timeout=10_000)   # D365 never fully goes quiet - do not insist
    except PWTimeout:
        pass
    print("Step 2.6 done: Invoice Header tab is open")


def find_input(page, label):
    """Return the visible <input> that belongs to the on-screen label."""
    low = "abcdefghijklmnopqrstuvwxyz"
    up = low.upper()
    xp = (f"//*[translate(normalize-space(text()),'{up}','{low}')='{label.lower()}']"
          f"/ancestor::*[.//input][1]//input")
    box = page.locator(f"xpath={xp} >> visible=true")
    if box.count() == 0:
        box = page.get_by_label(label, exact=True).locator("visible=true")
    if box.count() == 0:
        raise RuntimeError(f"Field '{label}' not found on the page")
    return box.first


def click_pencil(page, box, label):
    """Fields are read-only until you hover and click the little pencil icon at the field's right end."""
    box.scroll_into_view_if_needed()
    box.hover()
    page.wait_for_timeout(500)

    # (a) pencil element inside the same field, if D365 exposes one
    field = box.locator("xpath=ancestor::*[.//input][1]")
    pencil = field.locator(
        "[class*='pencil' i]:visible, [class*='edit' i]:visible, "
        "[title*='Edit' i]:visible, [aria-label*='Edit' i]:visible"
    )
    if pencil.count() > 0:
        pencil.first.click()
        print("    pencil clicked (icon element)")
        page.wait_for_timeout(500)
        return

    # (b) otherwise click where the pencil appears: right end of the field, vertically centred
    bb = box.bounding_box()
    y = bb["y"] + bb["height"] / 2
    x = bb["x"] + bb["width"] - 12
    page.mouse.move(bb["x"] + bb["width"] / 2, y)   # make sure the pencil is shown
    page.wait_for_timeout(300)
    page.mouse.move(x, y)
    page.wait_for_timeout(300)
    page.mouse.click(x, y)
    print(f"    pencil clicked (position {int(x)},{int(y)})")
    page.wait_for_timeout(500)


# Collects every visible text / input value near the field (D365 dropdown cells are <input> boxes,
# so normal text search cannot see them).
_NEAR_JS = """
(inp) => {
  const out = [];
  const els = document.querySelectorAll('input, [role="gridcell"], [role="option"], td, li, span, div, a, label');
  for (const e of els) {
    if (e.tagName !== 'INPUT' && e.children.length > 0) continue;   // leaf elements only
    const r = e.getBoundingClientRect();
    if (r.width < 2 || r.height < 2) continue;
    const st = getComputedStyle(e);
    if (st.visibility === 'hidden' || st.display === 'none') continue;
    const overlap = r.left < inp.x + inp.width && r.right > inp.x && r.top < inp.y + inp.height && r.bottom > inp.y;
    if (overlap) continue;                                           // skip the field itself
    if (Math.abs(r.top - inp.y) > 600 || Math.abs(r.left - inp.x) > 800) continue;
    const t = (e.tagName === 'INPUT' ? e.value : e.textContent || '').trim();
    if (!t || t.length > 80) continue;
    out.push({t: t, x: r.left + r.width / 2, y: r.top + r.height / 2,
              d: Math.hypot(r.left - inp.x, r.top - inp.y)});
  }
  return out;
}
"""


def pick_dropdown_row(page, box, candidates, wait_s=7):
    """Return {t, x, y} of the dropdown row to click, or None.
    Exact (case-insensitive) match first; 'contains' only for values of 5+ characters."""
    inp = box.bounding_box()
    items = []
    for _ in range(int(wait_s / 0.4)):
        items = page.evaluate(_NEAR_JS, inp)
        for mode in ("exact", "contains"):
            for cand in candidates:
                c = cand.lower()
                if mode == "exact":
                    hits = [i for i in items if i["t"].lower() == c]
                elif len(c) >= 5:
                    hits = [i for i in items if c in i["t"].lower()]
                else:
                    hits = []
                if hits:
                    return min(hits, key=lambda i: i["d"])
        page.wait_for_timeout(400)
    print("    texts seen near the field:", sorted({i["t"] for i in items})[:25])
    return None


def fill_field(page, label, value, kind="text", picks=None):
    box = find_input(page, label)
    handle = box.element_handle()           # keep THIS element, so later checks cannot hit another field
    click_pencil(page, box, label)          # make the field editable first
    try:
        box.click(timeout=5_000)
    except PWTimeout:
        box.click(force=True)
    page.keyboard.press("Control+A")
    page.keyboard.type(value, delay=60)

    if kind == "lookup":
        row = pick_dropdown_row(page, box, picks or [value])
        if row:
            print(f"    dropdown row found: {row['t']!r} - clicking it")
            page.mouse.click(row["x"], row["y"])
        else:
            page.screenshot(path=f"lookup_{label.replace(' ', '_')}.png")
            print("    NO dropdown row matched - pressing Tab (saved lookup_*.png)")
            page.keyboard.press("Tab")
    else:
        page.keyboard.press("Tab")        # commit the typed value

    page.wait_for_timeout(1500)
    try:
        print(f"    {label} now = {handle.input_value()!r}")
    except Exception:
        print(f"    {label} saved (field left edit mode - check step3_filled.png)")


def fill_invoice_header(page):
    print("Step 3.0: filling Invoice Header (hover + click pencil on every field)")
    for i, item in enumerate(HEADER_FIELDS, start=1):
        label, value, kind = item[:3]
        picks = item[3] if len(item) > 3 else None
        print(f"Step 3.{i}: {label} -> {value}")
        try:
            fill_field(page, label, value, kind, picks)
        except Exception as e:
            page.screenshot(path=f"step3_error_{i}.png")
            print(f"  FAILED: {str(e).splitlines()[0]}  (see step3_error_{i}.png)")
            raise

    page.screenshot(path="step3_filled.png")
    print("Step 3 done: fields filled (see step3_filled.png)")


def main():
    invoice_no = (sys.argv[1] if len(sys.argv) > 1 else os.environ.get("INVOICE_NO", "")).strip()
    if not invoice_no:
        raise SystemExit("Usage: python rpa_erp.py SC260309")

    with sync_playwright() as p:
        browser = launch_browser(p)      # installs Chromium on first use if it is missing
        context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page = context.new_page()
        try:
            login(page)
            print("Login successful:", page.url)
            dismiss_overlays(page)   # close dashboard pop-up before step 2
            # Reuse this session in later steps without logging in again
            context.storage_state(path="session.json")
            open_invoice_header(page, invoice_no)
            print("Invoice Header opened for", invoice_no)
            page.screenshot(path="invoice_header.png")
            fill_invoice_header(page)
            # --- Step 4 goes here ---
        except (PWTimeout, RuntimeError) as e:
            page.screenshot(path="login_error.png")
            print("Login error:", e)
            raise
        finally:
            browser.close()


if __name__ == "__main__":
    main()
