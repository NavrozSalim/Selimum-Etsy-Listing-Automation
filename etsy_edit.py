"""
Shared Etsy edit session for the price/stock and SKU replace bots.

Finds an existing listing by its SKU in Shop Manager, opens it in the standard listing editor
(listing-editor/edit/<id>), changes fields and saves it. It never opens the create page or Quick listing.
"""
from __future__ import annotations

import re
import subprocess
import time
from pathlib import Path
from typing import Callable, Optional

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement

import create_listing as cl
import etsy_output

LISTINGS_URL = "https://www.etsy.com/your/shops/me/tools/listings"
EDIT_URL = "https://www.etsy.com/your/shops/me/listing-editor/edit/{}"
# Shop Manager listing filters. Draft is kept apart so a draft is never published by these bots.
STATES = ("active", "draft", "sold_out", "inactive", "expired")
MAX_CANDIDATES = 8
EDITOR_TIMEOUT = 25
ROW_DELAY = 3

_SEARCH_BOX_JS = """
function visible(el) {
    const r = el.getBoundingClientRect();
    const s = getComputedStyle(el);
    return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none';
}
return Array.from(document.querySelectorAll('input[type="search"], input[type="text"], input:not([type])'))
    .filter(visible)
    .find(el => /search/i.test([el.placeholder, el.name, el.id, el.getAttribute('aria-label')].join(' '))) || null;
"""

_LISTING_IDS_JS = """
function visible(el) {
    const r = el.getBoundingClientRect();
    return r.width > 1 && r.height > 1;
}
const ids = [];
const add = (id) => { if (id && !ids.includes(id)) ids.push(id); };
for (const a of document.querySelectorAll('a[href]')) {
    if (!visible(a)) continue;
    const m = a.href.match(/listing-editor\\/edit\\/(\\d+)|\\/listing\\/(\\d+)/);
    if (m) add(m[1] || m[2]);
}
for (const el of document.querySelectorAll('[data-listing-id]')) {
    if (visible(el)) add(el.getAttribute('data-listing-id'));
}
return ids;
"""

_SAVE_LABELS_LIVE = ("Publish changes", "Save changes", "Publish", "Update", "Save")
_SAVE_LABELS_DRAFT = ("Save as draft", "Save draft", "Save")


def _visible_input(driver: WebDriver, word: str) -> Optional[WebElement]:
    """The single listing-level price/quantity/sku box on Pricing & Shipping.

    Only exact names are used, so a variation table row is never picked. Two or more
    visible matches means the listing has per-variation boxes and is not changed.
    """
    selector = ", ".join(
        [
            f"input[name='{word}']",
            f"input[name='variations.configuration.{word}']",
            f"input#{word}-input",
            f"input#listing-{word}-input",
        ]
    )
    try:
        found = [el for el in driver.find_elements(By.CSS_SELECTOR, selector) if el.is_displayed() and el.is_enabled()]
    except Exception:
        return None
    unique = list({el.id: el for el in found}.values())
    if len(unique) > 1:
        raise RuntimeError(f"This listing has more than one {word} box (variations). Nothing was changed.")
    return unique[0] if unique else None


def _wait_input(driver: WebDriver, word: str, timeout: float = 8) -> Optional[WebElement]:
    deadline = time.time() + timeout
    while time.time() < deadline:
        el = _visible_input(driver, word)
        if el is not None:
            return el
        time.sleep(0.5)
    return None


def _search_state(driver: WebDriver, state: str, sku: str) -> list[str]:
    driver.get(f"{LISTINGS_URL}/state:{state}")
    time.sleep(3)
    cl.wait_for_human_or_login(driver)
    box = None
    deadline = time.time() + 15
    while time.time() < deadline and box is None:
        box = driver.execute_script(_SEARCH_BOX_JS)
        if box is None:
            time.sleep(0.5)
    if box is None:
        raise RuntimeError("Could not find the Shop Manager listing search box")
    cl.set_input_value(driver, box, sku)
    box.send_keys(Keys.ENTER)
    time.sleep(3)
    cl.wait_for_human_or_login(driver)
    body = (driver.execute_script("return document.body ? document.body.innerText.slice(0, 4000) : ''") or "").lower()
    if re.search(r"no listings|didn't find|did not find|no results|0 listings", body):
        return []
    return [str(x) for x in (driver.execute_script(_LISTING_IDS_JS) or [])][:MAX_CANDIDATES]


def _open_editor(driver: WebDriver, listing_id: str) -> None:
    driver.get(EDIT_URL.format(listing_id))
    time.sleep(3)
    cl.wait_for_human_or_login(driver)
    deadline = time.time() + EDITOR_TIMEOUT
    while time.time() < deadline:
        screen = cl.listing_screen(driver)
        url = driver.current_url or ""
        if screen.get("quick") and not screen.get("editor"):
            raise RuntimeError("Etsy opened Quick listing instead of the standard editor. Nothing was changed.")
        if "listing-editor/edit" in url and screen.get("editor"):
            cl.dismiss_discard_modal(driver)
            return
        time.sleep(0.8)
    raise RuntimeError(f"The standard listing editor did not open for listing {listing_id}")


def _editor_sku(driver: WebDriver) -> Optional[str]:
    """SKU value on Pricing & Shipping, or None when there is no single SKU box (variations)."""
    cl.click_listing_tab(driver, "Pricing & Shipping")
    time.sleep(0.6)
    cl.dismiss_discard_modal(driver)
    cl.scroll_to_section_heading(driver, "price")
    el = _wait_input(driver, "sku")
    if el is None:
        return None
    return (el.get_attribute("value") or "").strip()


def open_listing_by_sku(driver: WebDriver, sku: str) -> str:
    """Opens the existing listing whose SKU box equals sku. Returns the Shop Manager state it was found in."""
    seen: set[str] = set()
    no_sku_box = 0
    for state in STATES:
        cl.log.info("Searching %s listings for SKU %s", state, sku)
        ids = [listing_id for listing_id in _search_state(driver, state, sku) if listing_id not in seen]
        for listing_id in ids:
            seen.add(listing_id)
            _open_editor(driver, listing_id)
            try:
                found = _editor_sku(driver)
            except RuntimeError as exc:
                cl.log.info("Listing %s skipped: %s", listing_id, exc)
                found = None
            if found is None:
                no_sku_box += 1
                continue
            if found.lower() == sku.strip().lower():
                cl.log.info("Found SKU %s on listing %s (%s)", sku, listing_id, state)
                return state
    if no_sku_box:
        raise RuntimeError(
            f"No listing with SKU {sku} found. {no_sku_box} matching listing(s) had no single SKU box (variations?)."
        )
    raise RuntimeError(f"No listing with SKU {sku} found")


def _set_field(driver: WebDriver, word: str, value: str, label: str) -> None:
    el = _wait_input(driver, word, timeout=6)
    if el is None:
        raise RuntimeError(f"Could not find the {label} box. The listing may use variations.")
    if not cl.set_input_value(driver, el, value):
        raise RuntimeError(f"Etsy did not keep the new {label} ({value})")
    error = cl.field_error_text(driver, el)
    if error:
        raise RuntimeError(f"Etsy rejected the {label}: {error}")
    cl.log.info("Set %s: %s", label, value)


def _no_unsaved_changes(driver: WebDriver) -> bool:
    text = driver.execute_script("return document.body ? document.body.innerText : ''") or ""
    return "you have no unsaved changes" in text.lower()


def save_listing(driver: WebDriver, state: str) -> str:
    """Saves the edit. A draft stays a draft; any other listing is published so the live listing changes."""
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)
    cl.dismiss_discard_modal(driver)
    labels = _SAVE_LABELS_DRAFT if state == "draft" else _SAVE_LABELS_LIVE
    start_url = driver.current_url
    clicked = next((label for label in labels if cl._click_footer_button(driver, label)), "")
    if not clicked and _no_unsaved_changes(driver):
        # Etsy greys out Publish changes when the listing already has these values.
        cl.log.info("Etsy shows no unsaved changes. The listing already has these values.")
        return "Already up to date"
    if not clicked:
        raise RuntimeError("Could not find the Publish or Save button on the listing editor")
    cl.log.info("Clicked %s", clicked)
    time.sleep(1.5)
    if clicked.lower().startswith("publish"):
        cl._confirm_publish_dialog(driver)
    cl.wait_for_human_or_login(driver)
    outcome = cl._wait_for_save(driver, start_url, timeout=20)
    if outcome == "saved":
        return "Saved as draft" if state == "draft" else "Published"
    fields = cl.blocking_fields(driver)
    if fields:
        raise RuntimeError(f"Etsy blocked the save: {', '.join(fields)}")
    raise RuntimeError("Etsy did not confirm the save. Check the screenshot.")


def update_price_and_stock(driver: WebDriver, sku: str, price: str, quantity: str) -> str:
    state = open_listing_by_sku(driver, sku)
    if price:
        _set_field(driver, "price", price, "price")
    if quantity:
        _set_field(driver, "quantity", quantity, "quantity")
    return save_listing(driver, state)


_UPC_JS = """
function visible(el) {
    const r = el.getBoundingClientRect();
    const s = getComputedStyle(el);
    return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none';
}
return Array.from(document.querySelectorAll('input')).filter(visible).find(el => {
    const lab = el.id ? document.querySelector('label[for="' + CSS.escape(el.id) + '"]') : null;
    const blob = [el.name, el.id, el.placeholder, el.getAttribute('aria-label'), lab ? lab.innerText : ''].join(' ');
    return /\\bupc\\b|gtin|\\bean\\b|barcode|product identifier/i.test(blob);
}) || null;
"""


def update_sku(driver: WebDriver, sku: str, new_sku: str, upc: str) -> str:
    state = open_listing_by_sku(driver, sku)
    upc_box = None
    if upc:
        # Look for the UPC box before changing anything, so a missing box leaves the listing untouched.
        upc_box = driver.execute_script(_UPC_JS)
        if upc_box is None:
            raise RuntimeError("This Etsy listing editor has no UPC/GTIN box. Nothing was changed.")
    _set_field(driver, "sku", new_sku, "SKU")
    if upc_box is not None:
        if not cl.set_input_value(driver, upc_box, upc):
            raise RuntimeError(f"Etsy did not keep the new UPC ({upc})")
        cl.log.info("Set UPC: %s", upc)
    return save_listing(driver, state)


def run_job(args, excel_path: Path, rows: list, apply_row: Callable, folder: Path) -> int:
    """Runs every selected row store by store and always writes Updated and remain-to-update."""
    headers, source = etsy_output.read_source(excel_path)
    all_rows = rows

    def save_split() -> tuple[Path, Path]:
        return etsy_output.write_split(folder, "Updated", "remain-to-update", headers, source, etsy_output.finished(all_rows))

    store_filter = (args.store or "").strip().upper()
    selected = [row for row in rows if not store_filter or row.store == store_filter]
    grouped: dict[str, list] = {}
    for row in selected:
        grouped.setdefault(row.store, []).append(row)
    if args.limit and args.limit > 0:
        grouped = {key: value[: args.limit] for key, value in grouped.items()}

    # Written now, after every row, and in finally, so a captcha, crash, Ctrl+C or the
    # worker killing python still leaves both files.
    save_split()
    code = 0
    try:
        if not grouped:
            cl.log.error("No rows to process%s", f" for store {store_filter}" if store_filter else "")
            return 1
        cl.log.info("Loaded %s row(s) across %s store(s)", sum(len(v) for v in grouped.values()), len(grouped))
        cfg = cl._load_config()
        chrome, user_data, vpncmd = cl.resolve_paths(cfg)
        for store_name, store_rows in grouped.items():
            info = cl.store_lookup(cfg, store_name)
            if not info:
                cl.log.error("Store %s is in Excel but not in stores.json. Skipping.", store_name)
                cl.mark(store_rows, "SKIPPED", "Not in stores.json")
                continue
            if all(row.status for row in store_rows):
                continue
            run_store_rows(
                cfg,
                chrome,
                user_data,
                vpncmd,
                info,
                store_rows,
                apply_row,
                use_vpn=not args.no_vpn,
                check_ip=not args.skip_ip_check,
                on_progress=save_split,
            )
    except cl.HumanCheckTimeout as exc:
        cl.log.error("Stopped: %s", exc)
        code = 3
    finally:
        done_path, remain_path = save_split()
        ok = sum(1 for row in all_rows if row.status == "OK")
        failed = [row for row in all_rows if row.status and row.status != "OK"]
        for row in failed:
            cl.log.info("Not done: row %s %s | %s", row.row_number, row.sku, row.detail)
        cl.log.info("Done. OK=%s  NOT OK=%s  NOT REACHED=%s", ok, len(failed), sum(1 for r in all_rows if not r.status))
        cl.log.info("Updated: %s", done_path)
        cl.log.info("Remaining: %s", remain_path)
    if code == 0 and any(row.status == "FAILED" for row in all_rows):
        code = 2
    return code


def run_store_rows(
    cfg: dict,
    chrome: Path,
    user_data: Path,
    vpncmd: Path,
    store_info: dict,
    rows: list,
    apply_row: Callable,
    *,
    use_vpn: bool,
    check_ip: bool,
    on_progress: Optional[Callable] = None,
) -> None:
    """Same Chrome profile, VPN and IP steps as create_listing.process_store, then apply_row per row."""
    store_name = store_info["name"]
    profile = store_info["chrome_profile"]
    expected_ip = (store_info.get("ip") or "").strip()
    vpn_account = (store_info.get("vpn") or store_name).strip()
    port = int(cfg.get("debug_port") or 9222)
    driver: Optional[WebDriver] = None
    chrome_proc: Optional[subprocess.Popen] = None
    keep_chrome_open = False

    cl.log.info("=" * 65)
    cl.log.info("Store: %s | Chrome: %s | VPN: %s | IP: %s", store_name, profile, vpn_account, expected_ip or "(none)")
    cl.log.info("=" * 65)

    if use_vpn:
        cl.disconnect_all_vpns(vpncmd, cfg)
        time.sleep(2)
        if not cl.connect_store_vpn(vpncmd, vpn_account) or not cl.wait_for_vpn(vpncmd, vpn_account):
            cl.mark(rows, "FAILED", f"VPN {vpn_account} did not connect")
            return

    try:
        cl.close_chrome(port)
        profile_dir = cl.automation_user_data(user_data, profile)
        chrome_proc = cl.start_chrome(chrome, profile_dir, profile, port)
        driver = cl.connect_selenium(port)

        if check_ip:
            ok, shown = cl.wait_for_expected_ip(driver, expected_ip) if expected_ip else (False, cl.read_ip(driver))
            if expected_ip and not ok:
                msg = f"IP mismatch: expected {expected_ip}, detected {shown or 'unknown'}"
                cl.log.error("%s: %s", store_name, msg)
                cl.mark(rows, "FAILED", msg)
                return

        for idx, row in enumerate(rows, start=1):
            if row.status:
                continue
            cl.log.info("--- Store %s | Row %d/%d | SKU: %s ---", store_name, idx, len(rows), row.sku)
            try:
                row.detail = apply_row(driver, row) or "Updated"
                row.status = "OK"
                cl.log.info("OK %s: %s", row.sku, row.detail)
            except cl.HumanCheckTimeout:
                raise
            except Exception as exc:
                cl.log.error("FAIL %s: %s", row.sku, exc)
                cl.screenshot(driver, f"fail_{row.sku}")
                row.status = "FAILED"
                row.detail = str(exc)
            if on_progress:
                on_progress()
            time.sleep(ROW_DELAY)

    except cl.HumanCheckTimeout:
        keep_chrome_open = True
        raise
    except Exception as exc:
        cl.log.exception("Chrome launch failed for %s", store_name)
        cl.mark(rows, "FAILED", f"Chrome launch failed: {exc}")
    finally:
        if keep_chrome_open:
            cl.log.error("Chrome was left open on the Etsy check. Finish it there, then send the remain-to-update file again.")
        else:
            try:
                if driver:
                    driver.quit()
            except Exception:
                pass
            cl.close_chrome(port, chrome_proc)
            if use_vpn:
                cl.disconnect_all_vpns(vpncmd, cfg)
