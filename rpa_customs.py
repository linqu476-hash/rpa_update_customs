"""
RPA_Customs_No_Multi_Invoice
  1. Login -> Export Invoice (Bulk) -> tab "Update CustomsNo Multi Invoice"
  2. Fill Customs No. ONCE
  3. For each invoice: fill Invoice No -> click "Add Invoice to List" -> next invoice ...
  4. After the last invoice: click "Update Customs No"

Usage:
    python rpa_customs.py                      uses the DATA block below
    python rpa_customs.py --file mydata.txt    file with lines:  CustomsNo  date  InvoiceNo  (tab/space separated)
    python rpa_customs.py E274120 C265825 C265240      customs no. followed by invoice numbers
    add  --no-update  to fill + add invoices but NOT click "Update Customs No"

Credentials come from .env (ERP_USER / ERP_PASSWORD). Optional ERP_FORM_URL = URL of the
Export Invoice (Bulk) page, which skips the menu clicks.
"""
import argparse
import json
import os
import re
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

load_dotenv()

ERP_URL = "https://ax.d365ffo.onprem.libgroup.com/namespaces/AXSF/"
# Paste rows here:  CustomsNo   date(ignored)   InvoiceNo   (tab or space separated)
DATA = """
E274120	10/Sep/26	C265825
E274120	10/Sep/26	C265825-1
E274120	10/Sep/26	C265240
E274120	10/Sep/26	C265236
E274120	10/Sep/26	C265236-1
E274120	10/Sep/26	C265357
"""

USERNAME = os.environ["ERP_USER"]
PASSWORD = os.environ["ERP_PASSWORD"]


def login(page):
    page.goto(ERP_URL, wait_until="domcontentloaded")

    # Screenshot step 1: username field (standard ADFS ids shown first, generic fallback second)
    user_box = page.locator("#userNameInput, input[type='text'], input[type='email']").first
    user_box.wait_for(state="visible", timeout=30_000)
    user_box.fill(USERNAME)

    # Screenshot step 2: password field
    page.locator("#passwordInput, input[type='password']").first.fill(PASSWORD)

    # Screenshot step 3: Sign in button
    page.locator("#submitButton, span:has-text('Sign in'), input[type='submit']").first.click()

    # Wait until we are back on the AX app (not on the ADFS page anymore)
    page.wait_for_url("**/namespaces/AXSF/**", timeout=60_000)
    page.wait_for_load_state("networkidle", timeout=60_000)

    # Detect failed login (ADFS shows an error message and stays on the login page)
    if page.locator("#errorText:visible").count() > 0:
        raise RuntimeError("Login failed: " + page.locator("#errorText").inner_text())


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


def open_bulk_form(page):
    """Step 2: navigation pane -> Favorites -> Export Invoice (Bulk)."""
    form_url = os.environ.get("ERP_FORM_URL", "").strip()
    if form_url:
        page.goto(form_url, wait_until="domcontentloaded")
    else:
        page.wait_for_timeout(1500)
        target = page.get_by_text("Export Invoice (Bulk)", exact=True)
        print("Step 2.1: opening navigation pane")
        opened = False
        for sel in ["[data-dyn-controlname='NavPaneButton']:visible",
                    "button.navigationBar-toggleButton:visible",
                    "button[aria-label*='avigation' i]:visible",
                    "button[title*='avigation' i]:visible"]:
            loc = page.locator(sel)
            if loc.count() == 0:
                continue
            try:
                loc.first.click(timeout=3_000)
                target.first.wait_for(state="visible", timeout=4_000)
                opened = True
                break
            except PWTimeout:
                continue
        if not opened:
            print("  selectors failed - clicking hamburger by position")
            page.mouse.click(24, 69)
            try:
                target.first.wait_for(state="visible", timeout=5_000)
            except PWTimeout:
                page.screenshot(path="nav_debug.png")
                raise RuntimeError("Navigation pane did not open - see nav_debug.png")
        print("Step 2.2: clicking Export Invoice (Bulk)")
        target.first.click()
    print("Step 2.3: waiting for the form")
    page.wait_for_selector("text=Export Invoice List >> visible=true", timeout=60_000)


def open_update_customs_tab(page):
    """Step 3: click tab 'Update CustomsNo Multi Invoice' and confirm it opened."""
    print("Step 3: opening tab 'Update CustomsNo Multi Invoice'")
    marker = page.get_by_text("Add Invoice to List", exact=True).locator("visible=true")
    tabs = [page.get_by_role("tab", name="Update CustomsNo Multi Invoice"),
            page.get_by_text("Update CustomsNo Multi Invoice", exact=True)]
    for tab in tabs:
        if tab.count() == 0:
            continue
        try:
            tab.first.click(timeout=5_000)
            marker.first.wait_for(state="visible", timeout=10_000)
            print("  tab is open")
            return
        except PWTimeout:
            continue
    page.screenshot(path="tab_debug.png")
    raise RuntimeError("Tab 'Update CustomsNo Multi Invoice' did not open - see tab_debug.png")


def click_pencil(page, box, label):
    """Fields are read-only until you hover and click the little pencil icon at the field's right end."""
    box.scroll_into_view_if_needed(timeout=5_000)
    box.hover(timeout=5_000)
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


def _read_value(box):
    """Current value of the input, re-resolved every time (never a stale element handle)."""
    try:
        return (box.input_value(timeout=2_000) or "").strip()
    except Exception:
        return None


def _dump_inputs(page, label):
    """Debug aid: print every visible input (aria-label / readonly / value) so a failure is diagnosable."""
    try:
        rows = page.evaluate("""() => [...document.querySelectorAll('input')].filter(e => {
            const r = e.getBoundingClientRect(); return r.width > 2 && r.height > 2; })
          .map(e => ({aria: e.getAttribute('aria-label'), ro: e.readOnly, dis: e.disabled,
                      val: e.value, name: e.name || e.id}))""")
        print(f"    visible inputs while filling '{label}':")
        for r in rows:
            print("      ", r)
    except Exception as e:
        print("    (could not dump inputs:", e, ")")


def _enter_value(page, box, label, value, mode):
    """Try one way of getting `value` into the box. Returns True only if the box really holds it.
      pencil : hover + click the pencil (field is read-only until then), click, type
      direct : click the field and type (field already editable, e.g. focus moved there after Tab)
      focus  : focus the field with no mouse click and type (keyboard only)
      js     : set the value with JS and fire input/change events
    """
    try:
        box.scroll_into_view_if_needed(timeout=5_000)
        if mode == "pencil":
            click_pencil(page, box, label)
            try:
                box.click(timeout=5_000)
            except PWTimeout:
                box.click(force=True, timeout=5_000)
        elif mode == "direct":
            try:
                box.click(timeout=5_000)
            except PWTimeout:
                box.click(force=True, timeout=5_000)
        elif mode == "focus":
            box.focus(timeout=5_000)

        if mode == "js":
            box.evaluate(
                """(el, v) => { el.focus(); el.value = v;
                       el.dispatchEvent(new Event('input', {bubbles: true}));
                       el.dispatchEvent(new Event('change', {bubbles: true})); }""", value)
        else:
            page.keyboard.press("Control+A")
            page.keyboard.type(value, delay=50)
        page.wait_for_timeout(400)

        got = _read_value(box)
        if got is not None and got.lower() == value.lower():
            return True
        print(f"    [{mode}] box holds {got!r}, expected {value!r}")
    except Exception as e:
        print(f"    [{mode}] failed: {str(e).splitlines()[0]}")
    try:
        page.screenshot(path=f"field_debug_{label.strip('. ').replace(' ', '_')}_{mode}.png")
    except Exception:
        pass
    return False


def type_into(page, label, value):
    """Find the visible input for `label`, type the value, verify it is in the box, then press Tab."""
    low = "abcdefghijklmnopqrstuvwxyz"
    t = lambda expr: f"translate(normalize-space({expr}),'{low.upper()}','{low}')='{label.lower()}'"
    finders = [
        f"input[aria-label='{label}' i] >> visible=true",
        f"xpath=//*[{t('text()')}]/ancestor::*[.//input][1]//input >> visible=true",
        f"xpath=//*[{t('.')}]/ancestor::*[.//input][1]//input >> visible=true",
        f"xpath=//*[{t('.')}]/following::input[1] >> visible=true",
    ]
    for attempt in (1, 2):                              # D365 may still be re-rendering: retry once
        for n, f in enumerate(finders, start=1):
            loc = page.locator(f)
            if loc.count() == 0:
                continue
            box = loc.first
            try:
                box.wait_for(state="visible", timeout=5_000)
            except PWTimeout:
                continue
            try:
                editable = box.is_editable(timeout=2_000)
            except Exception:
                editable = False
            # read-only field -> pencil first (this is what worked for Customs No.);
            # already-editable field (focus moved here after Tab) -> type straight in, pencil last
            order = ["direct", "focus", "pencil", "js"] if editable else ["pencil", "direct", "focus", "js"]
            print(f"    '{label}' finder {n}: editable={editable}, trying {order}")
            for mode in order:
                if _enter_value(page, box, label, value, mode):
                    page.keyboard.press("Tab")           # commit
                    page.wait_for_timeout(800)
                    print(f"    {label} = {value!r}  (finder {n}, {mode}, attempt {attempt})")
                    return
            print(f"    could not fill '{label}' with finder {n} - trying next way to find the field")
        page.wait_for_timeout(1_500)
    _dump_inputs(page, label)
    page.screenshot(path=f"field_error_{label.strip('. ').replace(' ', '_')}.png")
    raise RuntimeError(f"Could not fill '{label}' with {value!r}")


_IN_GRID_JS = """
([val, minY]) => {
  for (const e of document.querySelectorAll('input, [role="gridcell"], td, span, div, a')) {
    if (e.tagName !== 'INPUT' && e.children.length > 0) continue;
    const r = e.getBoundingClientRect();
    if (r.width < 2 || r.height < 2 || r.top < minY) continue;
    const t = (e.tagName === 'INPUT' ? e.value : e.textContent || '').trim().toLowerCase();
    if (t === val.toLowerCase()) return true;
  }
  return false;
}
"""


def in_grid(page, invoice, wait_s=0):
    """True if the invoice is shown in the list below the input fields."""
    box = page.locator("xpath=//*[translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
                       "'abcdefghijklmnopqrstuvwxyz')='invoice no']/following::input[1] >> visible=true")
    min_y = (box.first.bounding_box()["y"] + 40) if box.count() else 330
    for _ in range(max(1, int(wait_s / 0.5))):
        if page.evaluate(_IN_GRID_JS, [invoice, min_y]):
            return True
        page.wait_for_timeout(500)
    return False


TOP_BAR_PX = 90          # buttons above this y position belong to the D365 top bar / form header


def popup_close_buttons(page):
    """Visible 'Close' buttons that belong to a dialog. The form's OWN Close (X) in the top-right
    corner is also a button named 'Close' - clicking it closes the whole Export Invoice form - so
    anything in the top bar is ignored."""
    btns = page.get_by_role("button", name="Close", exact=True).locator("visible=true")
    keep = []
    for i in range(btns.count()):
        b = btns.nth(i)
        try:
            bb = b.bounding_box()
        except Exception:
            continue
        if bb and bb["y"] > TOP_BAR_PX:
            keep.append(b)
    return keep


def close_popup_if_any(page, tag):
    """If a D365 message dialog (e.g. 'Invoice X not found') is open: read it, click its Close.
    Returns the message text, or None when no dialog is showing. Never clicks the form's own Close (X)."""
    buttons = popup_close_buttons(page)
    if not buttons:
        return None
    msg = None
    nf = page.get_by_text(re.compile(r"not found|error|already|cannot|invalid", re.I)).locator("visible=true")
    if nf.count() > 0:
        msg = nf.first.inner_text().strip()
    page.screenshot(path=f"popup_{tag}.png")
    btn = buttons[-1]
    bb = btn.bounding_box()
    print(f"    dialog Close button clicked at ({int(bb['x'])},{int(bb['y'])})" if bb else "    dialog Close clicked")
    btn.click()
    try:
        btn.wait_for(state="hidden", timeout=5_000)
    except PWTimeout:
        pass                                   # no Escape here: it can close the whole form in D365
    page.wait_for_timeout(500)
    return msg or "(popup message not read - see popup_%s.png)" % tag


_EMPTY_MSG = re.compile(r"didn.t find anything to show", re.I)


def list_is_empty(page, wait_s=4):
    """True when the list below the fields shows D365's 'We didn't find anything to show here.'"""
    msg = page.get_by_text(_EMPTY_MSG).locator("visible=true")
    try:
        msg.first.wait_for(state="visible", timeout=int(wait_s * 1000))
        return True
    except PWTimeout:
        return False


def clear_list(page):
    """OPT-IN (--clear-list): select every row in the list and click Delete. Best effort - the
    selectors are untested against your D365, so the result is always re-checked with list_is_empty()."""
    print("    trying to clear the list (select all + Delete)")
    for sel in ["[role='columnheader'] [role='checkbox']:visible",
                "[role='checkbox'][aria-label*='all' i]:visible",
                "[role='radio'][aria-label*='all' i]:visible",
                "[class*='selectAll' i]:visible", "[class*='select-all' i]:visible"]:
        loc = page.locator(sel)
        if loc.count() > 0:
            try:
                loc.first.click(timeout=3_000)
                break
            except PWTimeout:
                continue
    page.wait_for_timeout(500)
    delete = page.get_by_role("button", name="Delete", exact=True).locator("visible=true")
    if delete.count() == 0:
        delete = page.get_by_text("Delete", exact=True).locator("visible=true")
    if delete.count() > 0:
        delete.first.click()
        page.wait_for_timeout(800)
        yes = page.get_by_role("button", name="Yes", exact=True).locator("visible=true")
        if yes.count() > 0:
            yes.first.click()
            page.wait_for_timeout(1_000)
    return list_is_empty(page)


def ensure_list_empty(page, customs_no, clear):
    """Update Customs No updates EVERY row in the list, so each group must start from an empty list.
    Rows left over from an earlier failed group/run would get this group's Customs No by mistake."""
    if list_is_empty(page):
        return
    page.screenshot(path=f"list_not_empty_{customs_no}.png")
    if clear and clear_list(page):
        print("    list cleared")
        return
    raise RuntimeError(
        f"The invoice list is not empty before starting {customs_no} (rows left from an earlier run). "
        f"Update Customs No would change them too, so this group was not started. "
        f"Empty the list in D365 (see list_not_empty_{customs_no}.png), or run with --clear-list.")


def add_invoices(page, customs_no, invoices, res, clear=False):
    """Fill Customs No. once, then add each invoice. A 'not found' popup is closed and the bot moves on.
    Fills `res` in place ({'added': [...], 'skipped': [...], 'failed': {inv: msg}}) so progress is kept
    even if something goes wrong half way. One invoice that cannot be typed is marked failed and the
    next one is tried; two in a row stop the group (the exception goes up to main())."""
    ensure_list_empty(page, customs_no, clear)
    print(f"Step 2: Customs No. = {customs_no}  (filled once)")
    type_into(page, "Customs No.", customs_no)
    page.wait_for_timeout(1_000)                         # let D365 settle after Tab

    add_btn = page.get_by_role("button", name="Add Invoice to List")
    if add_btn.count() == 0:
        add_btn = page.get_by_text("Add Invoice to List", exact=True)

    errors_in_row = 0
    for n, inv in enumerate(invoices, start=1):
        try:
            close_popup_if_any(page, f"before_{n}")      # make sure no old dialog is blocking
            if in_grid(page, inv):
                print(f"Step 3.{n}: {inv} is already in the list - skipped")
                res["skipped"].append(inv)
                errors_in_row = 0
                continue
            print(f"Step 3.{n}: Invoice No = {inv} -> Add Invoice to List")
            type_into(page, "Invoice No", inv)
            add_btn.locator("visible=true").first.click()

            # wait for the outcome: either a popup (e.g. 'not found') or the invoice showing in the list
            outcome = None
            for _ in range(16):                          # up to ~8 s
                msg = close_popup_if_any(page, str(n))
                if msg:
                    outcome = ("failed", msg)
                    break
                if in_grid(page, inv):
                    outcome = ("added", None)
                    break
                page.wait_for_timeout(500)

            if outcome and outcome[0] == "failed":
                print(f"    NOT ADDED: {outcome[1]}  -> popup closed, continuing")
                res["failed"][inv] = outcome[1]
            elif outcome and outcome[0] == "added":
                print(f"    {inv} added to the list")
                res["added"].append(inv)
            else:
                page.screenshot(path=f"add_check_{n}.png")
                print(f"    WARNING: no popup and {inv} not visible in the list (see add_check_{n}.png)")
                res["failed"][inv] = "not visible in list after adding"
            errors_in_row = 0
        except Exception as e:
            errors_in_row += 1
            why = str(e).splitlines()[0] if str(e) else type(e).__name__
            print(f"    ERROR on {inv}: {why}")
            res["failed"][inv] = f"error: {why}"
            try:
                page.screenshot(path=f"invoice_error_{customs_no}_{n}.png")
            except Exception:
                pass
            if errors_in_row >= 2:
                raise RuntimeError(f"2 invoices in a row could not be processed (last: {inv}: {why})")
    page.screenshot(path="customs_added.png")


def confirm_update_dialog(page, wait_s=15):
    """D365 asks 'Do you want to update Customs No ?' (Yes / No, with No pre-selected). Click Yes.
    Yes is only clicked when this exact question is on screen, so no other dialog is ever confirmed."""
    question = page.get_by_text(re.compile(r"Do you want to update Customs No", re.I)).locator("visible=true")
    try:
        question.first.wait_for(state="visible", timeout=wait_s * 1000)
    except PWTimeout:
        return False
    page.wait_for_timeout(500)
    page.screenshot(path="update_confirm.png")
    yes = page.get_by_role("button", name="Yes", exact=True).locator("visible=true")
    if yes.count() == 0:
        yes = page.get_by_text("Yes", exact=True).locator("visible=true")
    yes.last.click()
    print("    confirmation dialog: clicked Yes")
    try:
        question.first.wait_for(state="hidden", timeout=10_000)
    except PWTimeout:
        print("    WARNING: the confirmation dialog is still showing after clicking Yes")
    return True


def click_update(page, customs_no):
    """Click 'Update Customs No', confirm the 'Do you want to update Customs No ?' dialog with Yes,
    then wait for D365's 'completed' message. Returns that message, or None if not seen."""
    print("Step 4: clicking 'Update Customs No'")
    btn = page.get_by_role("button", name="Update Customs No", exact=True)
    if btn.count() == 0:
        btn = page.get_by_text("Update Customs No", exact=True)
    btn.locator("visible=true").first.click()
    if not confirm_update_dialog(page):
        page.screenshot(path="update_no_dialog.png")
        print("    WARNING: the 'Do you want to update Customs No ?' dialog did not appear "
              "(see update_no_dialog.png) - checking for the completed message anyway")
    done = page.get_by_text(f"Update Customs No : {customs_no} completed", exact=False)
    msg = None
    try:
        done.first.wait_for(state="visible", timeout=120_000)
        msg = done.first.inner_text().strip()
        print("    D365 message:", msg)
    except PWTimeout:
        print("    WARNING: no 'completed' message seen within 2 min - check customs_updated.png")
    page.screenshot(path="customs_updated.png")
    if msg:
        if list_is_empty(page, wait_s=15):
            print("    list is empty - ready for the next Customs No")
        else:
            print("    WARNING: the list is not empty after the update (the next group will be blocked until it is)")
    return msg


def parse_rows(text):
    """'CustomsNo date InvoiceNo' lines -> {customs_no: [invoice, ...]} (order kept, no duplicates)."""
    groups = {}
    for line in text.splitlines():
        parts = line.split()
        if len(parts) < 2:
            continue
        cust, inv = parts[0], parts[-1]
        if inv not in groups.setdefault(cust, []):
            groups[cust].append(inv)
    return groups


def read_args():
    p = argparse.ArgumentParser()
    p.add_argument("items", nargs="*", help="optional: customs_no followed by invoice numbers")
    p.add_argument("--file", help="text file with lines: CustomsNo date InvoiceNo")
    p.add_argument("--no-update", action="store_true", help="do not click 'Update Customs No'")
    p.add_argument("--clear-list", action="store_true",
                   help="if the invoice list is not empty at the start of a group, try to empty it (default: stop that group)")
    a = p.parse_args()
    if a.items:
        groups = {a.items[0]: a.items[1:]}
    elif a.file:
        with open(a.file, encoding="utf-8") as f:
            groups = parse_rows(f.read())
    else:
        groups = parse_rows(DATA)
    if not groups or not any(groups.values()):
        p.error("no invoice data found")
    clear = a.clear_list or os.environ.get("RPA_CLEAR_LIST", "0") == "1"
    return groups, not a.no_update, clear


MAX_GROUPS_IN_A_ROW = 3          # stop the whole run after this many failed Customs No groups in a row


def recover(page):
    """After a failed group: close any dialog and make sure the 'Update CustomsNo Multi Invoice' tab is open."""
    try:
        close_popup_if_any(page, "recover")
        marker = page.get_by_text("Add Invoice to List", exact=True).locator("visible=true")
        if marker.count() == 0:
            print("    form is not on the Update CustomsNo tab - reopening it")
            open_bulk_form(page)
            open_update_customs_tab(page)
        return True
    except Exception as e:
        print("    recovery failed:", str(e).splitlines()[0])
        return False


def main():
    groups, do_update, clear = read_args()
    for c, inv in groups.items():
        print(f"Plan: Customs No {c} <- {len(inv)} invoice(s): {', '.join(inv)}")
    results, error = {}, None
    with sync_playwright() as p:
        # HEADLESS=1 (set by the dashboard) runs without showing the browser
        browser = p.chromium.launch(headless=os.environ.get("HEADLESS", "0") == "1")
        context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page = context.new_page()
        try:
            login(page)
            print("Step 1: login successful:", page.url)
            dismiss_overlays(page)
            open_bulk_form(page)
            open_update_customs_tab(page)
            failed_in_row = 0
            for customs_no, invoices in groups.items():
                res = {"added": [], "skipped": [], "failed": {}, "updated": None, "update_message": None}
                results[customs_no] = res
                try:
                    add_invoices(page, customs_no, invoices, res, clear)
                    print(f"\nSummary for {customs_no}: added {len(res['added'])}, "
                          f"already in list {len(res['skipped'])}, failed {len(res['failed'])}")
                    for inv, why in res["failed"].items():
                        print(f"   - {inv}: {why}")
                    if do_update and (res["added"] or res["skipped"]):
                        msg = click_update(page, customs_no)
                        res["updated"], res["update_message"] = bool(msg), msg
                    elif do_update:
                        print("Nothing in the list - 'Update Customs No' NOT clicked")
                    failed_in_row = 0
                except Exception as e:
                    failed_in_row += 1
                    why = str(e).splitlines()[0] if str(e) else type(e).__name__
                    print(f"\nGROUP FAILED {customs_no}: {why}")
                    for inv in invoices:                 # everything not finished is reported as failed
                        if inv not in res["added"] and inv not in res["skipped"] and inv not in res["failed"]:
                            res["failed"][inv] = f"group stopped: {why}"
                    if res["added"] and not res["updated"]:
                        print(f"   NOTE: {len(res['added'])} invoice(s) were added to the list but "
                              f"'Update Customs No' was NOT clicked for {customs_no}.")
                    try:
                        page.screenshot(path=f"group_error_{customs_no}.png")
                    except Exception:
                        pass
                    if failed_in_row >= MAX_GROUPS_IN_A_ROW:
                        raise RuntimeError(f"Stopped: {failed_in_row} Customs No groups failed in a row "
                                           f"(last: {customs_no}: {why})")
                    if not recover(page):
                        raise RuntimeError(f"Stopped: could not recover after {customs_no} failed ({why})")
        except Exception as e:
            error = str(e)
            try:
                page.screenshot(path="error.png")
            except Exception:
                pass
            print("ERROR:", e, "(see error.png)")
            raise
        finally:
            browser.close()
            out = os.environ.get("RPA_RESULT_JSON")
            if out:
                with open(out, "w", encoding="utf-8") as f:
                    json.dump({"results": results, "error": error}, f, indent=2)


if __name__ == "__main__":
    main()
