#!/usr/bin/env python3
"""
Create Etsy listings in Chrome (Selenium), one store at a time.
Supports SoftEther VPN / Dedicated Store IP binding, Chrome Profile isolation,
and 10-step listing automation with the strict 20-Rule Etsy SEO Tag Engine.

Usage:
  python create_listing.py
  python create_listing.py "Create.xlsx"
  python create_listing.py --store AMH --limit 1
  python create_listing.py --tags "Burgundy Velvet Ribbon 2.5 Inch Wired Ribbon for Christmas Tree Bow Making, 10 Yards"
  python create_listing.py --no-vpn --skip-ip-check
  python create_listing.py --publish
"""

from __future__ import annotations

import argparse
import html
import json
import logging
import math
import os
import re
import shutil
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Optional, List, Dict

import requests
from openpyxl import Workbook, load_workbook
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

from core.tag_engine import generate_etsy_tags, validate_tags, format_tags_output

ROOT = Path(__file__).resolve().parent
STORES_PATH = ROOT / "stores.json"
LOG_PATH = ROOT / "create_listing.log"
OUTPUT_DIR = ROOT / "output"
FAILURES_DIR = OUTPUT_DIR / "failures"
PHOTO_DIR = OUTPUT_DIR / "photos"
DEFAULT_EXCEL = ROOT / "Create.xlsx"

IP_CHECK_URL = "https://nordvpn.com/what-is-my-ip/"
IPIFY_URL = "https://api.ipify.org"
ETSY_SHOP_URL = "https://www.etsy.com/your/shops/me"
ETSY_CREATE_URL = "https://www.etsy.com/your/shops/me/listing-editor/create"
IPV4_RE = re.compile(r"\b(\d{1,3}(?:\.\d{1,3}){3})\b")
URL_RE = re.compile(r"https?://[^\s|,]+")

VPN_CONNECT_TIMEOUT = 45
VPN_STATUS_POLL = 2
IP_SETTLE_TIMEOUT = 60
IP_SETTLE_POLL = 4
ROW_DELAY = 3
CHROME_CONNECT_TIMEOUT = 35
SIGN_IN_TIMEOUT = 300
MAX_PHOTOS = 10

HEADER_MAP = {
    "store name": "store",
    "store_name": "store",
    "storename": "store",
    "store": "store",
    "title": "title",
    "description": "description",
    "price": "price",
    "quantity": "quantity",
    "sku": "sku",
    "photos": "photos",
    "images": "photos",
    "category": "category",
    "tags": "tags",
    "shipping_profile": "shipping_profile",
    "shipping": "shipping_profile",
    "item_weight": "item_weight",
    "weight": "item_weight",
    "package_length": "package_length",
    "length": "package_length",
    "package_width": "package_width",
    "width": "package_width",
    "package_height": "package_height",
    "height": "package_height",
    "return_policy": "return_policy",
    "who_made_it": "who_made_it",
    "what_is_it": "what_is_it",
    "when_made": "when_made",
    "shop_section": "shop_section",
    "section": "shop_section",
    "allow_offers": "allow_offers",
    "featured": "featured",
    "renewal": "renewal",
    "status": "status",
    "publish": "publish",
    "remove_branding": "remove_branding",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("etsy_listing")


@dataclass
class ListingRow:
    row_number: int
    store: str
    title: str
    description: str
    price: str
    quantity: str
    sku: str
    photos: list[str]
    category: str
    tags: list[str]
    shipping_profile: str
    item_weight: str
    package_length: str
    package_width: str
    package_height: str
    return_policy: str
    who_made_it: str
    what_is_it: str
    when_made: str
    shop_section: str
    allow_offers: bool
    featured: bool
    renewal: str
    publish: bool
    remove_branding: list[str] = field(default_factory=list)
    status: str = ""
    detail: str = ""
    photos_done: bool = False


def _load_config() -> dict:
    if not STORES_PATH.exists():
        raise SystemExit(f"Missing {STORES_PATH}")
    with STORES_PATH.open(encoding="utf-8") as f:
        return json.load(f)


def _chrome_candidates() -> list[Path]:
    pf = Path(os.environ.get("PROGRAMFILES", r"C:\Program Files"))
    pf86 = Path(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"))
    return [
        pf / "Google" / "Chrome" / "Application" / "chrome.exe",
        pf86 / "Google" / "Chrome" / "Application" / "chrome.exe",
        Path.home() / "AppData" / "Local" / "Google" / "Chrome" / "Application" / "chrome.exe",
    ]


def _vpncmd_candidates(configured: str) -> list[Path]:
    paths = []
    if configured:
        paths.append(Path(configured))
    env = os.environ.get("VPNCMD_PATH", "").strip()
    if env:
        paths.append(Path(env))
    pf = Path(os.environ.get("PROGRAMFILES", r"C:\Program Files"))
    paths.append(pf / "SoftEther VPN Client" / "vpncmd.exe")
    return paths


def resolve_paths(cfg: dict) -> tuple[Path, Path, Path]:
    chrome_raw = (cfg.get("chrome_path") or "").strip()
    chrome = Path(chrome_raw) if chrome_raw else Path()
    if not chrome.is_file():
        chrome = next((p for p in _chrome_candidates() if p.is_file()), chrome)
    if not chrome.is_file():
        raise SystemExit("Chrome not found. Set chrome_path in stores.json")

    user_raw = (cfg.get("chrome_user_data_dir") or "").strip()
    user_data = Path(user_raw) if user_raw else Path.home() / "AppData" / "Local" / "Google" / "Chrome" / "User Data"
    if not user_data.is_dir():
        raise SystemExit("Chrome User Data folder not found. Set chrome_user_data_dir in stores.json")

    vpncmd = Path("vpncmd")
    for candidate in _vpncmd_candidates(cfg.get("vpncmd_path") or ""):
        if candidate.is_file():
            vpncmd = candidate
            break
    return chrome, user_data, vpncmd


def store_lookup(cfg: dict, name: str) -> Optional[dict]:
    stores = cfg.get("stores") or {}
    key = (name or "").strip().upper()
    if key not in stores:
        # Try case-insensitive
        for k, v in stores.items():
            if k.upper() == key:
                row = dict(v)
                row["name"] = k
                return row
        return None
    row = dict(stores[key])
    row["name"] = key
    return row


def coerce_code(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none"}:
        return ""
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".", 1)[0]
    return text


def coerce_money(value: Any) -> str:
    text = coerce_code(value)
    if not text:
        return ""
    try:
        return f"{float(text):.2f}"
    except ValueError:
        return text


def parse_bool(value: Any, default: bool = False) -> bool:
    text = coerce_code(value).lower()
    if not text:
        return default
    return text in {"true", "1", "yes", "y", "on"}


def _plain(fragment: str) -> str:
    text = re.sub(r"<[^>]+>", "", fragment or "")
    return re.sub(r"[ \t]+", " ", html.unescape(text)).strip()


def html_to_text(raw: str) -> str:
    text = re.sub(r"<!--.*?-->", "", raw or "", flags=re.S)
    text = re.sub(r"(?is)<h[1-6][^>]*>(.*?)</h[1-6]>", lambda m: f"\n{_plain(m.group(1))}\n", text)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</p>", "\n\n", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def strip_brands(text: str, brands: list[str]) -> str:
    out = text or ""
    for brand in brands:
        if not brand:
            continue
        out = re.sub(rf"(?i)\b{re.escape(brand)}\b", "", out)
    out = out.replace("\r\n", "\n").replace("\r", "\n")
    out = re.sub(r"[ \t]{2,}", " ", out)
    out = re.sub(r" *\n *", "\n", out)
    return re.sub(r"\n{3,}", "\n\n", out).strip()


def to_bullet_description(text: str) -> str:
    """Turn * markers and list items into one bullet per line for Etsy's textarea."""
    raw = text or ""
    raw = re.sub(r"(?i)</li\s*>", "\n", raw)
    raw = re.sub(r"(?i)<li[^>]*>", "\n* ", raw)
    raw = re.sub(r"<[^>]+>", "", raw)
    raw = html.unescape(raw).replace("\r\n", "\n").replace("\r", "\n")
    has_markers = "*" in raw or "•" in raw
    if has_markers:
        parts = re.split(r"\s*[*•]\s*", raw)
    else:
        parts = re.split(r"\n+", raw)
        if len([part for part in parts if part.strip()]) < 2:
            return re.sub(r"[ \t]+", " ", raw).strip()
    lines: list[str] = []
    for part in parts:
        line = re.sub(r"\s+", " ", part).strip(" \t-*•")
        if line:
            lines.append(f"• {line}")
    return "\n".join(lines) if lines else raw.strip()


def enforce_rules_on_tags(raw_tags: list[str], title: str = "", description: str = "") -> list[str]:
    """Ensures any tags (from Excel or user) strictly follow all 20 Etsy SEO rules."""
    from core.tag_engine import sanitize_tag, contains_banned_terms
    cleaned = []
    for t in raw_tags:
        s = sanitize_tag(t)
        if s and not contains_banned_terms(s) and len(s) <= 20:
            if s.lower() not in ["party decor", "party supplies"] and s.lower() not in [x.lower() for x in cleaned]:
                cleaned.append(s)
    first_11 = cleaned[:11]
    if len(first_11) < 11:
        auto_tags = generate_etsy_tags(title, description)
        for at in auto_tags[:11]:
            if at.lower() not in [x.lower() for x in first_11] and len(first_11) < 11:
                first_11.append(at)
    return first_11 + ["party decor", "party supplies"]


def parse_photos(value: Any) -> list[str]:
    text = "" if value is None else str(value)
    if not text or text.lower() == "nan":
        return []
    found: list[str] = []
    for part in re.split(r"[\n|;]+", text):
        part = part.strip().strip(",")
        if not part:
            continue
        if part.lower().startswith("http"):
            found.extend(URL_RE.findall(part))
        elif Path(part).is_file():
            found.append(str(Path(part)))
    return found or URL_RE.findall(text)


def read_excel(path: Path) -> list[ListingRow]:
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    header_cells = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if not header_cells:
        raise SystemExit("Excel has no header row")
    index: dict[str, int] = {}
    for i, col in enumerate(header_cells):
        mapped = HEADER_MAP.get(str(col or "").strip().lower())
        if mapped and mapped not in index:
            index[mapped] = i

    missing = [c for c in ("store", "title") if c not in index]
    if missing:
        raise SystemExit(f"Excel is missing mandatory columns: {', '.join(missing)}")

    def cell(values: tuple, key: str) -> Any:
        pos = index.get(key)
        if pos is None or pos >= len(values):
            return None
        return values[pos]

    rows: list[ListingRow] = []
    for row_number, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not values:
            continue
        store = coerce_code(cell(values, "store")).upper()
        title = coerce_code(cell(values, "title"))
        if not store and not title:
            continue

        brands = [b.strip() for b in coerce_code(cell(values, "remove_branding")).split(",") if b.strip()]
        description = html_to_text(coerce_code(cell(values, "description"))) or title
        title = strip_brands(title, brands) or title or "Untitled Listing"
        description = strip_brands(description, brands) or title

        status_text = coerce_code(cell(values, "status")).lower()
        publish_cell = coerce_code(cell(values, "publish")).lower()
        if status_text in {"draft"}:
            publish = False
        elif status_text in {"live", "published", "publish"}:
            publish = True
        elif publish_cell:
            publish = publish_cell in {"true", "1", "yes", "y", "on"}
        else:
            publish = False

        qty = coerce_code(cell(values, "quantity")) or "100"
        price = coerce_money(cell(values, "price")) or "19.99"
        sku = coerce_code(cell(values, "sku")) or f"SKU-{row_number}"
        photos = parse_photos(cell(values, "photos"))[:MAX_PHOTOS]

        category = coerce_code(cell(values, "category")) or "Craft Supplies"
        tags_raw = coerce_code(cell(values, "tags"))
        if tags_raw:
            raw_list = [t.strip() for t in tags_raw.split(",") if t.strip()]
            tags = enforce_rules_on_tags(raw_list, title, description)
        else:
            # Auto-generate 13 SEO tags strictly following 20 rules
            tags = generate_etsy_tags(title, description)

        shipping_profile = coerce_code(cell(values, "shipping_profile"))
        weight = coerce_code(cell(values, "item_weight")) or "0.5"
        length = coerce_code(cell(values, "package_length")) or "6"
        width = coerce_code(cell(values, "package_width")) or "4"
        height = coerce_code(cell(values, "package_height")) or "2"
        return_policy = coerce_code(cell(values, "return_policy")) or "30 days"
        who_made_it = coerce_code(cell(values, "who_made_it")) or "Another company"
        what_is_it = coerce_code(cell(values, "what_is_it")) or "A supply or tool"
        raw_when = coerce_code(cell(values, "when_made"))
        when_made = "2020–2026" if ("2020" in raw_when or not raw_when) else raw_when
        shop_section = coerce_code(cell(values, "shop_section"))
        allow_offers = parse_bool(cell(values, "allow_offers"), default=False)
        featured = parse_bool(cell(values, "featured"), default=False)
        renewal = coerce_code(cell(values, "renewal")) or "Automatic"

        rows.append(
            ListingRow(
                row_number=row_number,
                store=store,
                title=title[:140],
                description=description[:8000],
                price=price,
                quantity=qty.split(".", 1)[0] or "100",
                sku=sku,
                photos=photos,
                category=category,
                tags=tags,
                shipping_profile=shipping_profile,
                item_weight=weight,
                package_length=length,
                package_width=width,
                package_height=height,
                return_policy=return_policy,
                who_made_it=who_made_it,
                what_is_it=what_is_it,
                when_made=when_made,
                shop_section=shop_section,
                allow_offers=allow_offers,
                featured=featured,
                renewal=renewal,
                publish=publish,
                remove_branding=brands,
            )
        )
    return rows


def write_results(rows: list[ListingRow], dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "results"
    ws.append(["STORE NAME", "SKU", "TITLE", "STATUS", "DETAIL", "TIME"])
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for r in rows:
        ws.append([r.store, r.sku, r.title, r.status, r.detail, now_str])
    wb.save(dest)
    log.info("Results saved to %s", dest)


# -------------------------------------------------------------
# SoftEther VPN & IP Management
# -------------------------------------------------------------
def run_vpncmd(vpncmd: Path, parts: list[str], timeout: int = 30) -> subprocess.CompletedProcess:
    return subprocess.run(
        [str(vpncmd), "localhost", "/CLIENT", "/CMD", *parts],
        capture_output=True,
        text=True,
        timeout=timeout,
        shell=False,
    )


def vpn_account_status(vpncmd: Path, account: str) -> str:
    try:
        result = run_vpncmd(vpncmd, ["AccountStatusGet", account], timeout=20)
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return "error"
    out = ((result.stdout or "") + "\n" + (result.stderr or "")).lower()
    if "not connected" in out or "error code: 37" in out:
        return "down"
    if "does not exist" in out or "error code: 36" in out:
        return "error"
    if "connection completed" in out or "session established" in out or "established" in out:
        return "up"
    if "connecting" in out:
        return "connecting"
    return "error" if result.returncode != 0 else "down"


def disconnect_all_vpns(vpncmd: Path, cfg: dict) -> None:
    seen = set()
    for info in (cfg.get("stores") or {}).values():
        account = (info.get("vpn") or "").strip()
        if not account or account in seen:
            continue
        seen.add(account)
        try:
            run_vpncmd(vpncmd, ["AccountDisconnect", account], timeout=20)
        except Exception:
            pass


def connect_store_vpn(vpncmd: Path, account: str) -> bool:
    try:
        result = run_vpncmd(vpncmd, ["AccountConnect", account], timeout=30)
    except FileNotFoundError:
        log.error("vpncmd not found: %s", vpncmd)
        return False
    except subprocess.TimeoutExpired:
        log.error("vpncmd timed out connecting %s", account)
        return False
    out = ((result.stdout or "") + "\n" + (result.stderr or "")).lower()
    if result.returncode == 0 or "command completed successfully" in out:
        log.info("VPN connect issued for %s", account)
        return True
    log.error("VPN connect failed for %s: %s", account, (result.stderr or result.stdout or "").strip())
    return False


def wait_for_vpn(vpncmd: Path, account: str) -> bool:
    start = time.time()
    while time.time() - start < VPN_CONNECT_TIMEOUT:
        state = vpn_account_status(vpncmd, account)
        if state == "up":
            log.info("VPN session established for %s", account)
            return True
        if state == "error":
            log.error("VPN status error for account %s", account)
            return False
        time.sleep(VPN_STATUS_POLL)
    log.error("VPN did not come up for %s", account)
    return False


# -------------------------------------------------------------
# Chrome Launch & Debugger Attachment
# -------------------------------------------------------------
LOCK_PATH = OUTPUT_DIR / "create_listing.lock"


def _pid_running(pid: int) -> bool:
    if pid <= 0:
        return False
    try:
        import ctypes
        handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)
        if handle:
            ctypes.windll.kernel32.CloseHandle(handle)
            return True
    except Exception:
        return False
    return False


def acquire_run_lock() -> None:
    """One listing run at a time. A second copy used to kill the Chrome the first copy was using."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for _ in range(2):
        try:
            fd = os.open(str(LOCK_PATH), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
            return
        except FileExistsError:
            holder = ""
            try:
                holder = LOCK_PATH.read_text(encoding="utf-8", errors="ignore").strip()
                pid = int(holder)
            except Exception:
                pid = 0
            if _pid_running(pid):
                raise SystemExit(
                    f"Another listing run is already active (PID {holder}). "
                    "Close that window and wait for it to finish. "
                    "Starting a second run closes Chrome and breaks the first one."
                )
            try:
                LOCK_PATH.unlink()
            except OSError:
                pass
    raise SystemExit("Could not start because another listing run still holds the lock.")


def release_run_lock() -> None:
    try:
        if LOCK_PATH.exists() and LOCK_PATH.read_text(encoding="utf-8", errors="ignore").strip() == str(os.getpid()):
            LOCK_PATH.unlink()
    except OSError:
        pass


def close_chrome(port: int, proc: Optional[subprocess.Popen] = None) -> None:
    """Stop the automation Chrome only. Never taskkill every chrome.exe on the PC."""
    if proc is not None and proc.poll() is None:
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True, text=True, shell=False)
        time.sleep(1)
    script = (
        "Get-CimInstance Win32_Process | "
        "Where-Object { ($_.Name -eq 'chrome.exe' -or $_.Name -eq 'chromedriver.exe') -and "
        "$_.CommandLine -and ($_.CommandLine -like '*remote-debugging-port="
        + str(port)
        + "*' -or $_.CommandLine -like '*chrome_debug_data*') } | "
        "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
    )
    subprocess.run(["powershell", "-NoProfile", "-Command", script], capture_output=True, text=True, shell=False)
    time.sleep(1)


def _disable_driver_retries(driver: WebDriver) -> None:
    """A dead Chrome should fail at once instead of retrying the closed port for minutes."""
    import urllib3
    from urllib3.util.retry import Retry

    executor = driver.command_executor
    executor._conn = urllib3.PoolManager(
        timeout=20,
        retries=Retry(total=0, connect=0, read=0, redirect=0, status=0),
    )
    try:
        executor._client_config.timeout = 20
    except Exception:
        pass


def _browser_gone(exc: BaseException) -> bool:
    text = str(exc).lower()
    return any(
        bit in text
        for bit in (
            "connection refused",
            "failed to establish a new connection",
            "invalid session",
            "session deleted",
            "not connected to devtools",
            "chrome not reachable",
            "disconnected",
            "winerror 10061",
            "max retries exceeded",
        )
    )


DEBUG_USER_DATA = ROOT / "chrome_debug_data"
_SKIP_PROFILE_DIRS = {
    "Cache", "Code Cache", "GPUCache", "DawnGraphiteCache", "DawnWebGPUCache",
    "GrShaderCache", "ShaderCache", "Service Worker", "CacheStorage", "Crashpad",
    "BrowserMetrics", "blob_storage",
}


def _ignore_profile_junk(_directory: str, names: list[str]) -> list[str]:
    return [name for name in names if name in _SKIP_PROFILE_DIRS]


def automation_user_data(user_data: Path, profile: str) -> Path:
    """Prepares isolated copy of profile to enable remote-debugging on Chrome 153+."""
    default = Path.home() / "AppData" / "Local" / "Google" / "Chrome" / "User Data"
    try:
        is_default = user_data.resolve().samefile(default)
    except OSError:
        is_default = os.path.normcase(str(user_data)) == os.path.normcase(str(default))
    if not is_default:
        return user_data

    source = user_data / profile
    if not source.is_dir():
        log.warning("Chrome profile %s not found in %s. Using default target.", profile, user_data)
        return DEBUG_USER_DATA

    target = DEBUG_USER_DATA / profile
    if not target.is_dir():
        log.info("Copying profile %s to %s for automation session...", profile, DEBUG_USER_DATA)
        DEBUG_USER_DATA.mkdir(parents=True, exist_ok=True)
        shutil.copytree(source, target, ignore=_ignore_profile_junk, dirs_exist_ok=True)
        for root_file in user_data.glob("Local State*"):
            try:
                shutil.copy2(root_file, DEBUG_USER_DATA / root_file.name)
            except Exception:
                pass
    return DEBUG_USER_DATA


def start_chrome(chrome: Path, user_data: Path, profile: str, port: int) -> subprocess.Popen:
    args = [
        str(chrome),
        f"--remote-debugging-port={port}",
        "--remote-allow-origins=*",
        "--disable-features=DevToolsDebuggingRestrictions",
        f"--user-data-dir={user_data}",
        f"--profile-directory={profile}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-session-crashed-bubble",
        "--hide-crash-restore-bubble",
        IP_CHECK_URL,
    ]
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    err_path = OUTPUT_DIR / "chrome_stderr.log"
    err_file = err_path.open("w", encoding="utf-8", errors="replace")
    log.info("Launching Chrome %s from %s (port %s)", profile, user_data, port)
    proc = subprocess.Popen(args, stdout=err_file, stderr=subprocess.STDOUT)
    proc.stderr_log = err_file  # type: ignore[attr-defined]
    return proc


def wait_for_debug_port(port: int) -> None:
    url = f"http://127.0.0.1:{port}/json/version"
    deadline = time.time() + CHROME_CONNECT_TIMEOUT
    last = "port closed"
    while time.time() < deadline:
        try:
            resp = requests.get(url, timeout=2)
            if resp.ok and "webSocketDebuggerUrl" in resp.text:
                log.info("Chrome debug port %s is open", port)
                return
            last = f"HTTP {resp.status_code}"
        except Exception as exc:
            last = str(exc)
        time.sleep(1)
    raise RuntimeError(f"Chrome did not open debug port {port} ({last})")


def local_chromedriver(browser_version: str) -> Path:
    cache = Path.home() / ".cache" / "selenium" / "chromedriver" / "win64"
    drivers = [path for path in cache.glob("*/chromedriver.exe") if path.is_file()]
    if not drivers:
        raise RuntimeError("ChromeDriver is not installed in the Selenium cache.")
    major = (browser_version or "").split(".", 1)[0]
    same_major = [path for path in drivers if path.parent.name.startswith(f"{major}.")]
    pool = same_major or drivers
    return max(pool, key=lambda path: [int(part) if part.isdigit() else 0 for part in path.parent.name.split(".")])


def connect_selenium(port: int) -> WebDriver:
    wait_for_debug_port(port)
    browser = ""
    try:
        meta = requests.get(f"http://127.0.0.1:{port}/json/version", timeout=2).json()
        browser = str(meta.get("Browser") or "").split("/", 1)[-1]
    except Exception:
        browser = ""
    driver_path = local_chromedriver(browser)
    log.info("Attaching Selenium with ChromeDriver %s", driver_path.parent.name)
    options = Options()
    options.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
    service = Service(executable_path=str(driver_path))
    holder: dict[str, Any] = {}

    def _attach() -> None:
        try:
            holder["driver"] = webdriver.Chrome(service=service, options=options)
        except Exception as exc:
            holder["error"] = exc

    worker = threading.Thread(target=_attach, daemon=True)
    worker.start()
    worker.join(45)
    if worker.is_alive():
        service.stop()
        raise RuntimeError(
            "Attaching Selenium timed out. Close the automation Chrome window and run the script again."
        )
    if "error" in holder:
        raise RuntimeError(f"Could not attach Selenium to Chrome on port {port}: {holder['error']}") from holder["error"]
    log.info("Selenium attached successfully.")
    _disable_driver_retries(holder["driver"])
    return holder["driver"]


def screenshot(driver: WebDriver, name: str) -> None:
    FAILURES_DIR.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", name)[:80]
    dest = FAILURES_DIR / f"{datetime.now().strftime('%H%M%S')}_{safe}.png"
    try:
        driver.save_screenshot(str(dest))
        log.info("Saved screenshot: %s", dest.name)
    except Exception as exc:
        log.debug("Screenshot failed: %s", exc)


def dump_form(driver: WebDriver, name: str) -> None:
    """Writes every visible field on the editor with its label, value and invalid flag."""
    script = """
    function visible(el) {
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 1 && r.height > 1 && s.visibility !== 'hidden' && s.display !== 'none';
    }
    function labelFor(el) {
        if (el.id) {
            const lab = document.querySelector('label[for="' + CSS.escape(el.id) + '"]');
            if (lab) return lab.innerText.trim();
        }
        const wrap = el.closest('label');
        if (wrap) return wrap.innerText.trim();
        let box = el.parentElement;
        for (let i = 0; i < 3 && box; i++, box = box.parentElement) {
            const t = (box.innerText || '').trim();
            if (t && t.length < 120) return t;
        }
        return '';
    }
    const out = [];
    for (const el of document.querySelectorAll('input, textarea, select, [role="combobox"]')) {
        if (!visible(el) || (el.type || '') === 'hidden') continue;
        out.push({
            tag: el.tagName, type: el.type || '', name: el.name || '', id: el.id || '',
            label: labelFor(el).slice(0, 120),
            aria: el.getAttribute('aria-label') || '', placeholder: el.placeholder || '',
            value: el.tagName === 'SELECT' ? (el.options[el.selectedIndex] || {}).text : (el.value || '').slice(0, 80),
            invalid: el.getAttribute('aria-invalid') === 'true',
        });
    }
    const tabs = Array.from(document.querySelectorAll('button, a, [role="tab"]'))
        .filter(visible).map(b => (b.innerText || '').trim()).filter(t => t && t.length < 30);
    return {url: location.href, fields: out, buttons: tabs.slice(0, 80)};
    """
    try:
        data = driver.execute_script(script)
    except Exception as exc:
        log.debug("Form dump failed: %s", exc)
        return
    FAILURES_DIR.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", name)[:60]
    dest = FAILURES_DIR / f"{datetime.now().strftime('%H%M%S')}_form_{safe}.json"
    dest.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    log.info("Saved form snapshot: %s", dest.name)


def read_ip(driver: WebDriver) -> str:
    # 1. Fast check with lightweight ipify API
    try:
        driver.get(IPIFY_URL)
        time.sleep(1.5)
        text = driver.find_element(By.TAG_NAME, "body").text
        m = IPV4_RE.search(text)
        if m:
            return m.group(1)
    except Exception:
        pass

    # 2. NordVPN what is my ip
    try:
        driver.get(IP_CHECK_URL)
        time.sleep(3)
        for css in ("p.heading-xl", "p[data-hk]", "[class*='heading-xl']"):
            found = driver.find_elements(By.CSS_SELECTOR, css)
            if found:
                m = IPV4_RE.search(found[0].text)
                if m:
                    return m.group(1)
        m = IPV4_RE.search(driver.find_element(By.TAG_NAME, "body").text)
        if m:
            return m.group(1)
    except Exception:
        pass
    return ""


def wait_for_expected_ip(driver: WebDriver, expected: str) -> tuple[bool, str]:
    shown = ""
    deadline = time.time() + IP_SETTLE_TIMEOUT
    while True:
        shown = read_ip(driver)
        if expected and shown == expected:
            log.info("Store IP verified: %s", shown)
            return True, shown
        if time.time() >= deadline:
            break
        log.info("Current IP: %s, waiting for assigned IP %s ...", shown or "(unknown)", expected)
        time.sleep(IP_SETTLE_POLL)
    return False, shown


# -------------------------------------------------------------
# Photos Download & Upload Helpers
# -------------------------------------------------------------
def download_photos(row: ListingRow) -> list[str]:
    PHOTO_DIR.mkdir(parents=True, exist_ok=True)
    paths: list[str] = []
    for i, url in enumerate(row.photos, start=1):
        if Path(url).is_file():
            paths.append(str(Path(url)))
            continue
        try:
            resp = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
            resp.raise_for_status()
            ext = "jpg"
            ctype = (resp.headers.get("Content-Type") or "").lower()
            if "png" in ctype:
                ext = "png"
            elif "webp" in ctype:
                ext = "webp"
            dest = PHOTO_DIR / f"{re.sub(r'[^A-Za-z0-9._-]+', '_', row.sku)[:30]}_{i}.{ext}"
            dest.write_bytes(resp.content)
            paths.append(str(dest))
        except Exception as exc:
            log.warning("Photo download failed %s: %s", url[:80], exc)
    return paths


def upload_photos(driver: WebDriver, paths: list[str]) -> bool:
    if not paths:
        return False
    abs_paths = [str(Path(p).resolve()) for p in paths if Path(p).is_file()]
    if not abs_paths:
        return False

    for _ in range(3):
        inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
        for el in inputs:
            try:
                el.send_keys("\n".join(abs_paths))
                log.info("Uploaded %d photos.", len(abs_paths))
                time.sleep(3)
                return True
            except Exception:
                continue
        time.sleep(1)
    return False


# -------------------------------------------------------------
# Etsy Listing Form Automation: Steps 1 through 10
# -------------------------------------------------------------
def wait_for_human_or_login(driver: WebDriver, timeout: int = 180) -> bool:
    """Detects if Etsy requires login or human verification, and pauses for user."""
    start = time.time()
    last_msg = ""
    while time.time() - start < timeout:
        url = (driver.current_url or "").lower()
        title = (driver.title or "").lower()
        if "signin" not in url and "login" not in url and "captcha" not in title and "are you human" not in title and "security check" not in title:
            return True
        msg = "Etsy login or Cloudflare check detected! Please complete login/verification in the Chrome window..."
        if msg != last_msg:
            log.warning(msg)
            last_msg = msg
        time.sleep(3)
    return False


def click_listing_tab(driver: WebDriver, tab_name: str) -> bool:
    """Clicks one listing-editor tab. Skips the click when that section is already open."""
    log.info("Navigating to section: '%s'...", tab_name)
    script = """
    const target = arguments[0].replace(/[^a-z0-9]/gi, '').toLowerCase();
    function visible(el) {
        if (!el) return false;
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none';
    }
    function same(el) {
        const text = (el.innerText || el.getAttribute('aria-label') || '').trim();
        if (!text || text.length > 40 || !visible(el)) return false;
        return text.replace(/[^a-z0-9]/gi, '').toLowerCase() === target;
    }
    let match = Array.from(document.querySelectorAll('[role="tab"]')).find(same);
    if (!match) {
        // The editor's section strip uses plain buttons/links under the "New listing" heading.
        match = Array.from(document.querySelectorAll('button, a, [role="link"], li'))
            .filter(el => !el.closest('nav[aria-label*="shop" i], aside, [data-sidebar]'))
            .find(same);
    }
    if (!match) return 'missing';
    if (match.getAttribute('aria-selected') === 'true' || match.getAttribute('aria-current')) return 'already';
    match.scrollIntoView({block: 'nearest'});
    match.click();
    return 'clicked';
    """
    try:
        state = driver.execute_script(script, tab_name)
    except Exception as exc:
        if _browser_gone(exc):
            raise RuntimeError(
                "Chrome closed while opening the listing form. Only one listing script can run at a time."
            ) from exc
        log.warning("Tab click note: %s", exc)
        return False
    if state == "already":
        log.info("Section '%s' is already open.", tab_name)
        return True
    if state == "clicked":
        time.sleep(0.8)
        return True
    log.warning("Could not find the '%s' tab.", tab_name)
    return False


_SET_VALUE_JS = """
const el = arguments[0];
const val = arguments[1];
el.scrollIntoView({block: 'center'});
el.focus();
const proto = el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype
            : el instanceof HTMLSelectElement ? HTMLSelectElement.prototype
            : HTMLInputElement.prototype;
Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, val);
el.dispatchEvent(new Event('input', {bubbles: true}));
el.dispatchEvent(new Event('change', {bubbles: true}));
el.dispatchEvent(new FocusEvent('blur', {bubbles: false}));
el.dispatchEvent(new FocusEvent('focusout', {bubbles: true}));
return el.value;
"""


def _same_value(current: str, wanted: str) -> bool:
    a = (current or "").strip()
    b = (wanted or "").strip()
    if a == b:
        return True
    try:
        return float(a.replace(",", "")) == float(b.replace(",", ""))
    except ValueError:
        return False


def set_input_value(driver: WebDriver, element: WebElement, value: str) -> bool:
    """Sets a field so Etsy's React form keeps it, then leaves the field so Etsy validates it."""
    text = str(value)
    current = ""
    try:
        current = driver.execute_script(_SET_VALUE_JS, element, text) or ""
    except Exception as exc:
        if _browser_gone(exc):
            raise RuntimeError("Chrome closed while entering listing fields.") from exc
    if not _same_value(current, text):
        try:
            element.click()
            element.send_keys(Keys.CONTROL, "a")
            element.send_keys(Keys.BACKSPACE)
            element.send_keys(text)
            current = element.get_attribute("value") or ""
        except Exception as exc:
            if _browser_gone(exc):
                raise RuntimeError("Chrome closed while entering listing fields.") from exc
    try:
        driver.execute_script("arguments[0].blur();", element)
    except Exception:
        pass
    return _same_value(current, text)


def fill_input_by_label(driver: WebDriver, label: str, value: str) -> bool:
    """Finds the input that belongs to a visible label such as SKU, Length, Width, or Height."""
    if not str(value).strip():
        return False
    script = """
    const wanted = arguments[0].toLowerCase().replace(/[^a-z]/g, '');
    function visible(el) {
        if (!el) return false;
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 1 && r.height > 1 && s.visibility !== 'hidden' && s.display !== 'none';
    }
    function leaf(el) {
        const direct = Array.from(el.childNodes).filter(n => n.nodeType === 3).map(n => n.textContent).join(' ');
        const raw = direct.trim() ? direct : (el.children.length === 0 ? (el.textContent || '') : '');
        return raw.toLowerCase().replace(/[^a-z]/g, '');
    }
    function matches(clean) {
        return clean === wanted || (clean.startsWith(wanted) && clean.length <= wanted.length + 6);
    }
    function badInput(inp) {
        if (!inp || inp.tagName !== 'INPUT') return true;
        const kind = (inp.type || 'text').toLowerCase();
        if (['hidden', 'file', 'checkbox', 'radio', 'search', 'submit', 'button'].includes(kind)) return true;
        if (inp.closest('nav, header, [role="navigation"]')) return true;
        const blob = ((inp.getAttribute('aria-label') || '') + ' ' + (inp.placeholder || '') + ' ' + (inp.name || '') + ' ' + (inp.id || '')).toLowerCase();
        return ['title', 'description', 'tag', 'email', 'password'].some(word => blob.includes(word));
    }
    const stopWords = new Set(['length', 'width', 'height', 'weight', 'sku', 'price', 'quantity']);
    const nodes = Array.from(document.querySelectorAll('label, span, p, div, legend, h3, h4, dt, strong'));
    for (const el of nodes) {
        if (!visible(el) || el.closest('nav, header, [role="navigation"]')) continue;
        const shown = (el.innerText || '').trim();
        if (!shown || shown.length > 40) continue;
        const clean = leaf(el);
        if (!matches(clean)) continue;
        if (el.tagName === 'LABEL' && el.htmlFor) {
            const byId = document.getElementById(el.htmlFor);
            if (byId && visible(byId) && !badInput(byId)) return byId;
        }
        const nested = el.querySelector('input');
        if (nested && visible(nested) && !badInput(nested)) return nested;
        let parent = el.parentElement;
        for (let i = 0; i < 3 && parent; i++) {
            const inputs = Array.from(parent.querySelectorAll('input')).filter(inp => visible(inp) && !badInput(inp));
            if (inputs.length === 1) return inputs[0];
            parent = parent.parentElement;
        }
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT);
        walker.currentNode = el;
        let steps = 0;
        while (walker.nextNode() && steps < 30) {
            steps++;
            const node = walker.currentNode;
            const nodeLeaf = leaf(node);
            if (node !== el && stopWords.has(nodeLeaf) && !matches(nodeLeaf)) break;
            if (node.tagName === 'INPUT' && visible(node) && !badInput(node)) return node;
        }
    }
    return null;
    """
    try:
        element = driver.execute_script(script, label)
    except Exception as exc:
        if _browser_gone(exc):
            raise RuntimeError("Chrome closed while entering listing fields.") from exc
        return False
    if element is None:
        return False
    if set_input_value(driver, element, value):
        log.info("Set %s: %s", label, value)
        return True
    log.warning("Typed %s but the form did not keep %r.", label, value)
    return False


def choose_unit_dropdowns(driver: WebDriver) -> list[str]:
    """
    Etsy's item size and weight boxes each have a unit dropdown that starts on
    "Choose a unit". Size and weight stay invalid until a unit is chosen.
    Size uses Inches and weight uses Pounds, matching the Excel columns.
    """
    script = """
    function visible(el) {
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 1 && r.height > 1 && s.visibility !== 'hidden' && s.display !== 'none';
    }
    const picks = [];
    for (const sel of document.querySelectorAll('select')) {
        if (!visible(sel)) continue;
        const opts = Array.from(sel.options);
        const texts = opts.map(o => (o.text || '').trim().toLowerCase());
        let want = null;
        if (texts.includes('inches')) want = 'inches';
        else if (texts.some(t => t === 'pounds' || t === 'lb' || t === 'lbs')) want = texts.find(t => t === 'pounds' || t === 'lb' || t === 'lbs');
        if (!want) continue;
        const current = (sel.options[sel.selectedIndex] && sel.options[sel.selectedIndex].text || '').trim().toLowerCase();
        if (current === want) continue;
        const opt = opts[texts.indexOf(want)];
        const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set;
        setter.call(sel, opt.value);
        sel.dispatchEvent(new Event('input', {bubbles: true}));
        sel.dispatchEvent(new Event('change', {bubbles: true}));
        picks.push(opt.text.trim());
    }
    return picks;
    """
    try:
        return driver.execute_script(script) or []
    except Exception as exc:
        if _browser_gone(exc):
            raise RuntimeError("Chrome closed while choosing size units.") from exc
        return []


def clean_etsy_title(raw: str) -> str:
    """
    Etsy rejects titles over 140 characters, more than one '&', '%' or ':', and some
    symbols. Cut at a word boundary instead of mid-word.
    """
    text = re.sub(r"\s+", " ", raw or "").strip()
    text = re.sub(r"[^\w\s\-',.&%:/()$°™©®!?\"+]", "", text)
    for symbol in ("&", "%", ":"):
        first = text.find(symbol)
        if first != -1:
            text = text[: first + 1] + text[first + 1 :].replace(symbol, " ")
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > 140:
        cut = text[:140]
        if " " in cut:
            cut = cut[: cut.rfind(" ")]
        text = cut.rstrip(" ,-&:/")
    return text


def _input_for_label(driver: WebDriver, label: str) -> Optional[WebElement]:
    script = """
    const wanted = arguments[0].toLowerCase().replace(/[^a-z]/g, '');
    for (const lab of document.querySelectorAll('label')) {
        const text = (lab.innerText || '').toLowerCase().replace(/[^a-z]/g, '');
        if (!text.startsWith(wanted) || text.length > wanted.length + 12) continue;
        const target = lab.htmlFor ? document.getElementById(lab.htmlFor) : lab.querySelector('input, textarea');
        if (target) return target;
    }
    return null;
    """
    try:
        return driver.execute_script(script, label)
    except Exception:
        return None


def field_error_text(driver: WebDriver, element: WebElement) -> str:
    """Returns Etsy's inline error message for one field, or '' when the field is valid."""
    script = """
    const el = arguments[0];
    const bits = [];
    for (const id of (el.getAttribute('aria-describedby') || '').split(/\\s+/)) {
        const node = id && document.getElementById(id);
        if (node && node.innerText) bits.push(node.innerText.trim());
    }
    if (el.getAttribute('aria-invalid') !== 'true' && !bits.some(b => /error|must|can't|cannot|invalid|only|required/i.test(b))) return '';
    let box = el.parentElement;
    for (let i = 0; i < 4 && box; i++, box = box.parentElement) {
        for (const node of box.querySelectorAll('[class*="error" i], [class*="invalid" i], [role="alert"]')) {
            const t = (node.innerText || '').trim();
            if (t && t.length < 300) bits.push(t);
        }
    }
    return [...new Set(bits)].join(' | ').slice(0, 300);
    """
    try:
        return str(driver.execute_script(script, element) or "")
    except Exception:
        return ""


def _click_control_named(driver: WebDriver, label: str) -> bool:
    """Clicks a short button or link whose own text matches, such as 'Add SKU'."""
    script = """
    const wanted = arguments[0].toLowerCase().replace(/[^a-z]/g, '');
    function visible(el) {
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none';
    }
    for (const el of document.querySelectorAll('button, a, [role="button"]')) {
        if (!visible(el) || el.closest('nav, header')) continue;
        const text = (el.innerText || el.getAttribute('aria-label') || '').trim();
        if (!text || text.length > 40) continue;
        const clean = text.toLowerCase().replace(/[^a-z]/g, '');
        if (clean === wanted) {
            el.click();
            return true;
        }
    }
    return false;
    """
    try:
        return bool(driver.execute_script(script, label))
    except Exception as exc:
        if _browser_gone(exc):
            raise RuntimeError("Chrome closed while entering listing fields.") from exc
        return False


def dismiss_discard_modal(driver: WebDriver) -> bool:
    """Dismisses the 'Discard changes?' modal by clicking 'Keep editing' if it appears."""
    script = """
    const buttons = Array.from(document.querySelectorAll('button, [role="button"]'));
    for (const btn of buttons) {
        const r = btn.getBoundingClientRect();
        const s = getComputedStyle(btn);
        if (r.width < 2 || r.height < 2 || s.visibility === 'hidden' || s.display === 'none') continue;
        const txt = (btn.innerText || '').trim().toLowerCase();
        if (txt === 'keep editing') {
            btn.click();
            return true;
        }
    }
    return false;
    """
    try:
        dismissed = driver.execute_script(script)
        if dismissed:
            log.info("Dismissed 'Discard changes?' modal -> Keep editing")
            time.sleep(0.8)
            return True
    except Exception:
        pass
    return False


def scroll_to_section_heading(driver: WebDriver, heading_text: str, timeout: float = 5.0) -> bool:
    """Scrolls to a visible section heading on the listing form without clicking tab links."""
    script = """
    const target = arguments[0].replace(/[^a-z0-9]/gi, '').toLowerCase();
    const candidates = Array.from(document.querySelectorAll('h2, h3, h4, legend, [class*="section"] h2, [class*="section"] h3, [class*="heading"]'));
    for (const el of candidates) {
        const txt = (el.innerText || el.textContent || '').replace(/[^a-z0-9]/gi, '').toLowerCase();
        if (txt.includes(target)) {
            el.scrollIntoView({block: 'center', behavior: 'smooth'});
            return true;
        }
    }
    return false;
    """
    end = time.time() + timeout
    while time.time() < end:
        try:
            found = driver.execute_script(script, heading_text)
            if found:
                time.sleep(1.2)
                dismiss_discard_modal(driver)
                return True
        except Exception:
            pass
        time.sleep(0.5)
    return False


def click_matching_element(driver: WebDriver, patterns: list[str], timeout: float = 3) -> bool:
    """Finds and clicks any clickable element (label, button, radio, option, div, span) matching any text pattern."""
    end = time.time() + timeout
    script = """
    const patterns = arguments[0].map(p => p.replace(/[^a-z0-9]/gi, '').toLowerCase()).filter(Boolean);
    function visible(el) {
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none';
    }
    const elements = Array.from(document.querySelectorAll(
        'label, button, [role="radio"], [role="option"], [role="menuitem"], [role="button"], input[type="radio"], div, span'
    ));
    let best = null, bestLen = 1e9;
    for (const el of elements) {
        if (!visible(el) || el.closest('nav, header, aside, [role="navigation"]')) continue;
        const text = (el.innerText || el.value || el.getAttribute('aria-label') || '').trim();
        // Big containers also "contain" the words. Only short, control-sized text counts.
        if (!text || text.length > 90) continue;
        const clean = text.replace(/[^a-z0-9]/gi, '').toLowerCase();
        if (!patterns.some(p => clean.includes(p))) continue;
        if (clean.length < bestLen) { best = el; bestLen = clean.length; }
    }
    if (!best) return false;
    const target = best.closest('label, button, [role="radio"], [role="option"], [role="menuitem"], [role="button"]') || best;
    target.scrollIntoView({block: 'center', behavior: 'instant'});
    target.click();
    return true;
    """
    while time.time() < end:
        try:
            clicked = driver.execute_script(script, patterns)
            if clicked:
                time.sleep(0.5)
                return True
        except Exception:
            pass
        time.sleep(0.3)
    return False


def step_1_open_add_listing(driver: WebDriver) -> None:
    log.info("Step 1: Opening Etsy Shop Manager -> Add a Listing...")
    driver.get(ETSY_CREATE_URL)
    time.sleep(4)
    wait_for_human_or_login(driver)

    if "listing-editor" not in driver.current_url:
        driver.get("https://www.etsy.com/your/shops/me/listings")
        time.sleep(3)
        for css in ("a[href*='listing-editor']", "button[data-action*='add-listing']", "//a[contains(., 'Add a listing')]"):
            try:
                if css.startswith("//"):
                    btn = driver.find_element(By.XPATH, css)
                else:
                    btn = driver.find_element(By.CSS_SELECTOR, css)
                btn.click()
                time.sleep(4)
                break
            except Exception:
                pass


def step_2_fill_item_details(driver: WebDriver, row: ListingRow) -> None:
    log.info("Step 2: Item Details (Photos, SEO Title, Description)...")

    # 2a. Photos
    click_listing_tab(driver, "Photo & Video")
    time.sleep(0.5)
    dismiss_discard_modal(driver)
    if row.photos and not row.photos_done:
        local_paths = download_photos(row)
        if upload_photos(driver, local_paths):
            row.photos_done = True
            time.sleep(4)

    # 2b. Title
    click_listing_tab(driver, "Item Details")
    time.sleep(0.5)
    dismiss_discard_modal(driver)
    fill_title(driver, row)
    fill_description(driver, row)


def fill_title(driver: WebDriver, row: ListingRow) -> bool:
    title = clean_etsy_title(row.title)
    title_el = None
    for loc in [
        (By.NAME, "title"),
        (By.ID, "listing-edit-title-input"),
        (By.CSS_SELECTOR, "input[data-field='title'], textarea[data-field='title']"),
        (By.XPATH, "//*[self::input or self::textarea][contains(@aria-label, 'Title') or contains(@placeholder, 'Title') or contains(@id, 'title')]"),
    ]:
        try:
            el = driver.find_element(*loc)
            if el.is_displayed():
                title_el = el
                break
        except Exception:
            continue
    if title_el is None:
        title_el = _input_for_label(driver, "Title")
    if title_el is None:
        log.warning("Title box not found on Item Details.")
        return False
    if not set_input_value(driver, title_el, title):
        log.warning("Title was typed but Etsy did not keep it.")
        return False
    log.info("Entered SEO Title (%d chars): %s", len(title), title[:60])
    time.sleep(0.4)
    problem = field_error_text(driver, title_el)
    if problem:
        log.warning("Etsy says the title is invalid: %s", problem)
        return False
    return True


def fill_description(driver: WebDriver, row: ListingRow) -> None:
    desc = to_bullet_description(row.description.strip())
    for loc in [
        (By.NAME, "description"),
        (By.ID, "description-text-area"),
        (By.CSS_SELECTOR, "textarea[data-field='description']"),
        (By.XPATH, "//textarea[contains(@aria-label, 'Description') or contains(@id, 'description')]")
    ]:
        try:
            el = driver.find_element(*loc)
            if el.is_displayed():
                set_input_value(driver, el, desc)
                log.info("Entered Description.")
                break
        except Exception:
            continue


def _category_queries(category: str, title: str) -> list[str]:
    """Search phrases to try in Etsy's category picker, most specific first."""
    queries: list[str] = []

    def add(value: str) -> None:
        text = re.sub(r"\s+", " ", value or "").strip()
        if len(text) < 3:
            return
        if text.lower() not in {q.lower() for q in queries}:
            queries.append(text)

    add(category)
    if category.lower().endswith("s") and len(category) > 4:
        add(category[:-1])
    for word in ("banner", "garland", "ribbon", "balloon"):
        blob = f"{category} {title}".lower()
        if word in blob:
            add(word)
    return queries or ["party supplies"]


def select_etsy_category(driver: WebDriver, category: str, title: str = "") -> bool:
    """
    Opens Etsy's category control and chooses a real taxonomy result.

    The listing editor does not use a plain input named "category". The field is a
    button/combobox labeled Category. Clicking it opens a search list, and the
    result has to be clicked or Etsy keeps the required-field error.
    """
    queries = _category_queries(category, title)
    click_listing_tab(driver, "Item Details")
    time.sleep(0.5)
    if dismiss_discard_modal(driver):
        # The first tab click can be swallowed by the unsaved-changes prompt.
        time.sleep(0.3)
        click_listing_tab(driver, "Item Details")
        time.sleep(0.5)

    open_js = """
    const query = (arguments[0] || '').toLowerCase();
    function visible(el) {
        if (!el) return false;
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none';
    }
    function clean(s) {
        return (s || '').replace(/\\s+/g, ' ').trim();
    }
    function leafText(el) {
        const direct = Array.from(el.childNodes)
            .filter(n => n.nodeType === 3)
            .map(n => n.textContent)
            .join(' ');
        const raw = direct.trim() ? direct : (el.children.length === 0 ? (el.textContent || '') : '');
        return clean(raw).toLowerCase().replace(/[^a-z]/g, '');
    }
    function blocked(el) {
        const box = el.closest('[role="alert"]');
        const around = ((box && box.innerText) || (el.parentElement && el.parentElement.innerText) || '').toLowerCase();
        return around.includes('please fix');
    }
    function skipControl(el) {
        const blob = clean((el.innerText || '') + ' ' + (el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('placeholder') || '') + ' ' + (el.id || '') + ' ' + (el.name || '')).toLowerCase();
        return ['title', 'description', 'tag', 'price', 'quantity', 'sku', 'search shop'].some(word => blob.includes(word));
    }
    const labels = Array.from(document.querySelectorAll('label, h2, h3, h4, h5, legend, span, p, div'));
    let label = null;
    for (const el of labels) {
        if (!visible(el) || blocked(el)) continue;
        if (leafText(el) === 'category') { label = el; break; }
    }
    let root = label ? label.parentElement : null;
    if (label) {
        let node = label.parentElement;
        for (let i = 0; i < 6 && node; i++) {
            const controls = Array.from(node.querySelectorAll('button, [role="combobox"], input, [role="option"]')).filter(el => visible(el) && !skipControl(el));
            if (controls.length && controls.length < 12) { root = node; break; }
            node = node.parentElement;
        }
    }
    function scoreText(text) {
        const t = clean(text).toLowerCase();
        if (!t || t.length > 220) return 0;
        const flat = t.replace(/[^a-z]/g, '');
        if (['done', 'apply', 'cancel', 'search', 'category', 'selectacategory'].includes(flat)) return 0;
        const stem = query.endsWith('s') ? query.slice(0, -1) : query;
        let score = 0;
        if (query && t.includes(query)) score += 10;
        if (stem.length > 3 && t.includes(stem)) score += 6;
        const leaf = t.split(/\\n|>/).map(s => s.trim()).filter(Boolean).pop() || t;
        if (query && leaf.includes(query)) score += 8;
        if (t.includes('>')) score += 2;
        return score;
    }
    if (root) {
        const suggestions = Array.from(root.querySelectorAll('[role="option"], [role="menuitem"], button, li, a'))
            .filter(el => visible(el) && !skipControl(el) && clean(el.innerText).length < 220);
        let best = null, bestScore = 0, bestLen = 9999;
        for (const el of suggestions) {
            const score = scoreText(el.innerText);
            const len = clean(el.innerText).length;
            if (score >= 8 && (score > bestScore || (score === bestScore && len < bestLen))) {
                best = el; bestScore = score; bestLen = len;
            }
        }
        if (best) {
            best.scrollIntoView({block: 'center', behavior: 'instant'});
            best.click();
            return {action: 'suggestion', text: clean(best.innerText).slice(0, 160)};
        }
        const fields = Array.from(root.querySelectorAll('input, button, [role="combobox"], [aria-haspopup]'))
            .filter(el => visible(el) && !skipControl(el) && !el.closest('nav, header, [role="navigation"]') && leafText(el) !== 'category');
        const field = fields.find(el => /select a category|add a category|choose a category|search/.test(clean(el.innerText || el.getAttribute('placeholder') || el.getAttribute('aria-label') || '').toLowerCase()))
            || fields.find(el => el.tagName === 'INPUT' || el.getAttribute('role') === 'combobox')
            || fields.find(el => clean(el.innerText || el.getAttribute('aria-label') || '').length > 0);
        if (field) {
            field.scrollIntoView({block: 'center', behavior: 'instant'});
            field.setAttribute('data-etsy-bot-category', '1');
            field.click();
            field.focus();
            return {action: field.tagName === 'INPUT' ? 'input' : 'opened', text: clean(field.innerText || field.getAttribute('aria-label') || field.getAttribute('placeholder') || '').slice(0, 80)};
        }
    }
    const fallbacks = Array.from(document.querySelectorAll('button, [role="combobox"], [aria-haspopup]'));
    for (const el of fallbacks) {
        if (!visible(el) || blocked(el)) continue;
        const text = clean((el.innerText || '') + ' ' + (el.getAttribute('aria-label') || '')).toLowerCase();
        if (text.length > 80) continue;
        if (text.includes('select a category') || text.includes('add a category') || text.includes('choose a category') || text === 'category') {
            el.scrollIntoView({block: 'center', behavior: 'instant'});
            el.click();
            return {action: 'opened', text: text.slice(0, 80)};
        }
    }
    const nearby = [];
    for (const el of document.querySelectorAll('button, [role="combobox"], label, h2, h3, input')) {
        const text = clean((el.innerText || el.getAttribute('aria-label') || el.getAttribute('placeholder') || '')).slice(0, 60);
        const blob = (text + ' ' + (el.id || '')).toLowerCase();
        if (blob.includes('categor') || blob.includes('taxonom')) nearby.push(el.tagName + ': ' + text);
        if (nearby.length >= 8) break;
    }
    return {action: 'missing', text: nearby.join(' | ')};
    """

    find_search_js = """
    function visible(el) {
        if (!el) return false;
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none';
    }
    function skip(el) {
        const blob = ((el.getAttribute('aria-label') || '') + ' ' + (el.placeholder || '') + ' ' + (el.id || '') + ' ' + (el.name || '')).toLowerCase();
        return ['title', 'description', 'tag', 'price', 'quantity', 'sku'].some(word => blob.includes(word));
    }
    function inChrome(el) {
        return !el.closest('nav, header, [role="navigation"], [data-sidebar]');
    }
    const marked = document.querySelector('input[data-etsy-bot-category="1"]');
    if (marked && visible(marked) && !skip(marked)) return marked;
    const active = document.activeElement;
    if (active && active.tagName === 'INPUT' && visible(active) && !skip(active) && inChrome(active)) return active;
    const inputs = Array.from(document.querySelectorAll('input')).filter(el => visible(el) && !skip(el) && inChrome(el));
    const inDialog = inputs.find(el => el.closest('[role="dialog"], [role="listbox"]'));
    if (inDialog) return inDialog;
    return inputs.find(el => {
        const blob = ((el.placeholder || '') + ' ' + (el.getAttribute('aria-label') || '') + ' ' + (el.id || '')).toLowerCase();
        const near = ((el.closest('section, fieldset, div') || el.parentElement || document.body).innerText || '').toLowerCase();
        return (blob.includes('categor') || blob.includes('taxonom') || near.includes('category')) && !blob.includes('shop');
    }) || null;
    """

    click_option_js = """
    const query = (arguments[0] || '').toLowerCase();
    function visible(el) {
        if (!el) return false;
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none';
    }
    function clean(s) { return (s || '').replace(/\\s+/g, ' ').trim(); }
    function scoreText(text) {
        const t = clean(text).toLowerCase();
        if (!t || t.length > 220 || t.length < 3) return 0;
        const flat = t.replace(/[^a-z]/g, '');
        if (['done', 'apply', 'cancel', 'search', 'category', 'selectacategory', 'keepediting', 'saveasdraft', 'publish'].includes(flat)) return 0;
        const stem = query.endsWith('s') ? query.slice(0, -1) : query;
        let score = 0;
        if (query && t.includes(query)) score += 10;
        if (stem.length > 3 && t.includes(stem)) score += 6;
        const leaf = (t.split(/\\n|>/).map(s => s.trim()).filter(Boolean).pop()) || t;
        if (query && leaf.includes(query)) score += 8;
        if (stem.length > 3 && leaf.includes(stem)) score += 4;
        if (t.includes('>')) score += 2;
        return score;
    }
    const pools = [
        '[role="dialog"] [role="option"]',
        '[role="listbox"] [role="option"]',
        '[role="dialog"] [role="menuitem"]',
        '[role="listbox"] [role="menuitem"]',
        '[role="option"]',
        '[role="menuitem"]',
        '[role="dialog"] li',
        '[role="dialog"] button',
        '[role="dialog"] a'
    ];
    const seen = new Set();
    const options = [];
    for (const sel of pools) {
        for (const el of document.querySelectorAll(sel)) {
            if (seen.has(el) || !visible(el) || el.closest('nav, header, [role="navigation"]')) continue;
            seen.add(el);
            options.push(el);
        }
    }
    for (const el of document.querySelectorAll('li, a, button, [role="option"]')) {
        if (seen.has(el) || !visible(el) || el.closest('nav, header, [role="navigation"]')) continue;
        const text = clean(el.innerText);
        if (text.includes('>') && text.length < 220) {
            seen.add(el);
            options.push(el);
        }
    }
    let best = null, bestScore = 0, bestLen = 9999;
    const preview = [];
    for (const el of options) {
        const text = clean(el.innerText);
        if (!text || text.length > 220) continue;
        if (preview.length < 8) preview.push(text.slice(0, 90));
        const score = scoreText(text);
        if (score >= 8 && (score > bestScore || (score === bestScore && text.length < bestLen))) {
            best = el; bestScore = score; bestLen = text.length;
        }
    }
    if (!best) return {picked: '', options: preview};
    best.scrollIntoView({block: 'center', behavior: 'instant'});
    best.dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));
    best.click();
    return {picked: clean(best.innerText).slice(0, 160), options: preview};
    """

    confirm_js = """
    const dialog = document.querySelector('[role="dialog"]');
    const scope = dialog || document;
    if (!dialog) return false;
    for (const btn of scope.querySelectorAll('button')) {
        const text = (btn.innerText || '').replace(/\\s+/g, ' ').trim().toLowerCase();
        if (['done', 'apply', 'select', 'save', 'continue'].includes(text)) {
            btn.click();
            return text;
        }
    }
    return false;
    """

    try:
        opened = driver.execute_script(open_js, queries[0]) or {}
    except Exception as exc:
        log.warning("Category picker could not be opened: %s", exc)
        return False

    action = str(opened.get("action") or "")
    if action == "suggestion":
        time.sleep(0.6)
        driver.execute_script(confirm_js)
        log.info("Selected category: %s", opened.get("text"))
        return True
    if action == "missing":
        log.warning("Category control not found on Item Details. Nearby: %s", opened.get("text") or "(none)")
        return False

    log.info("Opened category picker (%s).", opened.get("text") or action)
    time.sleep(0.8)
    chosen = ""
    for query in queries:
        try:
            search = driver.execute_script(find_search_js)
        except Exception:
            search = None
        if search:
            try:
                search.click()
                search.send_keys(Keys.CONTROL, "a")
                search.send_keys(Keys.BACKSPACE)
                search.send_keys(query)
            except Exception as exc:
                log.warning("Could not type category search %r: %s", query, exc)
        result: dict = {}
        deadline = time.time() + 8
        while time.time() < deadline:
            time.sleep(0.7)
            try:
                result = driver.execute_script(click_option_js, query) or {}
            except Exception as exc:
                if _browser_gone(exc):
                    raise RuntimeError("Chrome closed while choosing a category.") from exc
                result = {}
            if result.get("picked"):
                break
        picked = str(result.get("picked") or "")
        if picked:
            chosen = picked
            log.info("Category result for %r: %s", query, picked.replace("\n", " ")[:140])
            break
        preview = result.get("options") or []
        if preview:
            log.info("No category row matched %r. Visible options: %s", query, " | ".join(preview[:6]))

    if not chosen:
        log.warning("Category %r was not selected. Etsy will block the save until one is chosen.", category)
        return False

    time.sleep(0.4)
    try:
        confirmed = driver.execute_script(confirm_js)
        if confirmed:
            log.info("Confirmed category dialog via '%s'.", confirmed)
    except Exception:
        pass
    time.sleep(0.5)
    return True


def step_3_fill_attributes_and_tags(driver: WebDriver, row: ListingRow) -> None:
    log.info("Step 3: Attributes (Category & 13 Etsy SEO Tags)...")

    # 3a. Category — the editor opens a picker; typing into a plain input never selects one.
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(0.4)
    dismiss_discard_modal(driver)

    if row.category:
        if not select_etsy_category(driver, row.category, row.title):
            log.warning("Category is still empty. Save will fail until Etsy has a category.")

    # 3b. 13 Etsy SEO Tags — scroll down to tags section (no tab clicks)
    driver.execute_script("window.scrollBy(0, 1500);")
    time.sleep(0.8)
    dismiss_discard_modal(driver)
    tags = row.tags
    log.info("Applying 13 SEO Tags (Rules: <=20 chars, ending in 'party decor', 'party supplies')...")
    tag_input = None
    for loc in [
        (By.CSS_SELECTOR, "input[placeholder*='tag' i]"),
        (By.CSS_SELECTOR, "input[aria-label*='tag' i]"),
        (By.ID, "tags-input"),
        (By.XPATH, "//input[contains(@placeholder, 'tag') or contains(@id, 'tag') or contains(@aria-label, 'tag')]")
    ]:
        try:
            el = driver.find_element(*loc)
            if el.is_displayed():
                tag_input = el
                break
        except Exception:
            continue

    if tag_input:
        for tag_str in tags:
            clean_tag = tag_str.strip()[:20]
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", tag_input)
                tag_input.clear()
                tag_input.send_keys(clean_tag)
                tag_input.send_keys(Keys.ENTER)
                time.sleep(0.35)
                try:
                    add_btn = driver.find_element(By.XPATH, "//button[text()='Add' or contains(., 'Add tag')]")
                    add_btn.click()
                except Exception:
                    pass
            except Exception:
                continue
        log.info("Applied 13 SEO Tags: %s", ", ".join(tags))
    else:
        log.warning("Tags input element not found in Item Options.")


def step_4_fill_price_and_inventory(driver: WebDriver, row: ListingRow) -> None:
    log.info("Step 4: Price & Inventory (Price, Quantity, SKU, Offers)...")
    click_listing_tab(driver, "Pricing & Shipping")
    time.sleep(0.4)
    dismiss_discard_modal(driver)
    scroll_to_section_heading(driver, "price")
    dismiss_discard_modal(driver)

    # Price
    for loc in [
        (By.NAME, "price"),
        (By.ID, "price-input"),
        (By.CSS_SELECTOR, "input[data-field='price']"),
        (By.XPATH, "//input[contains(@name, 'price') or contains(@aria-label, 'Price') or contains(@id, 'price')]")
    ]:
        try:
            el = driver.find_element(*loc)
            if el.is_displayed():
                set_input_value(driver, el, row.price)
                log.info("Set Price: $%s", row.price)
                break
        except Exception:
            continue

    # Quantity
    for loc in [
        (By.NAME, "quantity"),
        (By.ID, "quantity-input"),
        (By.CSS_SELECTOR, "input[data-field='quantity']"),
        (By.XPATH, "//input[contains(@name, 'quantity') or contains(@aria-label, 'Quantity') or contains(@id, 'quantity')]")
    ]:
        try:
            el = driver.find_element(*loc)
            if el.is_displayed():
                set_input_value(driver, el, row.quantity)
                log.info("Set Quantity: %s", row.quantity)
                break
        except Exception:
            continue

    # SKU box is listing-sku-input / name="sku". It can render after price and quantity,
    # so wait for it instead of skipping when it is not visible yet.
    if row.sku:
        sku_set = False
        deadline = time.time() + 6
        while time.time() < deadline and not sku_set:
            sku_set = _fill_by_name(driver, "sku", row.sku, "SKU")
            if not sku_set:
                time.sleep(0.5)
        if not sku_set and (_click_control_named(driver, "Add SKU") or _click_control_named(driver, "Add a SKU")):
            time.sleep(0.5)
            sku_set = _fill_by_name(driver, "sku", row.sku, "SKU")
        if not sku_set:
            log.warning("SKU was not entered. The SKU box was not found on Pricing & Shipping.")

    # Offers toggle
    try:
        offers_toggle = driver.find_element(By.XPATH, "//label[contains(., 'offer') or contains(., 'make offers')]//input | //button[@role='switch' and contains(., 'offer')]")
        if offers_toggle.is_displayed():
            is_checked = offers_toggle.is_selected() if offers_toggle.tag_name == "input" else (offers_toggle.get_attribute("aria-checked") == "true")
            if is_checked != row.allow_offers:
                driver.execute_script("arguments[0].click();", offers_toggle)
                log.info("Let buyers make offers set to: %s", row.allow_offers)
    except Exception:
        pass


def _fill_by_name(driver: WebDriver, name: str, value: str, label: str) -> bool:
    if not str(value).strip():
        return False
    try:
        element = driver.find_element(
            By.CSS_SELECTOR,
            f"input[name='{name}'], input[id='shipping_item_dimension-{name}'], input[id='listing-{name}-input']",
        )
    except NoSuchElementException:
        return False
    except Exception as exc:
        if _browser_gone(exc):
            raise RuntimeError("Chrome closed while entering item size.") from exc
        return False
    if set_input_value(driver, element, value):
        log.info("Set %s: %s", label, value)
        return True
    log.warning("Typed %s but the form did not keep %r.", label, value)
    return False


def _fill_item_size_group(driver: WebDriver, length: str, width: str, height: str) -> bool:
    """Fills Length, Width, Height when they sit together under an Item size heading."""
    script = """
    function visible(el) {
        if (!el) return false;
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 1 && r.height > 1 && s.visibility !== 'hidden' && s.display !== 'none';
    }
    function badInput(inp) {
        const kind = (inp.type || 'text').toLowerCase();
        if (['hidden', 'file', 'checkbox', 'radio', 'search'].includes(kind)) return true;
        const blob = ((inp.getAttribute('aria-label') || '') + ' ' + (inp.name || '') + ' ' + (inp.id || '')).toLowerCase();
        return ['title', 'price', 'quantity', 'sku', 'tag'].some(word => blob.includes(word));
    }
    const headings = Array.from(document.querySelectorAll('h2, h3, h4, legend, label, span, p, div'));
    let heading = null;
    for (const el of headings) {
        if (!visible(el)) continue;
        const text = (el.innerText || '').trim().toLowerCase().replace(/[^a-z]/g, '');
        if ((text === 'itemsize' || text === 'size') && (el.innerText || '').trim().length < 30) {
            heading = el;
            break;
        }
    }
    if (!heading) return [];
    let node = heading.parentElement;
    for (let i = 0; i < 5 && node; i++) {
        const inputs = Array.from(node.querySelectorAll('input')).filter(inp => visible(inp) && !badInput(inp));
        if (inputs.length >= 3 && inputs.length <= 6) return inputs.slice(0, 3);
        node = node.parentElement;
    }
    return [];
    """
    try:
        fields = driver.execute_script(script) or []
    except Exception as exc:
        if _browser_gone(exc):
            raise RuntimeError("Chrome closed while entering item size.") from exc
        return False
    if len(fields) < 3:
        return False
    for element, value, name in zip(fields, (length, width, height), ("Length", "Width", "Height")):
        set_input_value(driver, element, value)
        log.info("Set %s: %s", name, value)
    return True


def step_5_fill_shipping_and_returns(driver: WebDriver, row: ListingRow) -> None:
    log.info("Step 5: Shipping, Processing, and Returns...")
    click_listing_tab(driver, "Pricing & Shipping")
    time.sleep(0.4)
    dismiss_discard_modal(driver)
    scroll_to_section_heading(driver, "shipping")
    dismiss_discard_modal(driver)

    # Shipping profile
    if row.shipping_profile:
        try:
            sel_el = driver.find_element(By.XPATH, "//select[contains(@name, 'shipping_profile')]")
            if sel_el.is_displayed():
                Select(sel_el).select_by_visible_text(row.shipping_profile)
                log.info("Shipping profile selected: %s", row.shipping_profile)
        except Exception:
            click_matching_element(driver, [row.shipping_profile, "Standard Shipping", "Free Shipping"])

    # Item weight and item size (when packed). Item Details also has product attributes
    # labeled Length/Width, so the shipping boxes must be found by their own names.
    scroll_to_section_heading(driver, "item size", timeout=2) or scroll_to_section_heading(driver, "weight", timeout=2)

    weight_set = False
    if row.item_weight:
        weight_set = _fill_by_name(driver, "itemWeightPrimary", row.item_weight, "Item weight (lb)")
        if not weight_set:
            log.warning("Item weight was not entered.")
    if not _fill_by_name(driver, "itemWeightSecondary", "1", "Item weight (oz)"):
        log.warning("Item weight ounces were not entered.")

    length_set = _fill_by_name(driver, "itemLength", row.package_length, "Length")
    width_set = _fill_by_name(driver, "itemWidth", row.package_width, "Width")
    height_set = _fill_by_name(driver, "itemHeight", row.package_height, "Height")
    if not (length_set and width_set and height_set):
        grouped = _fill_item_size_group(driver, row.package_length, row.package_width, row.package_height)
        length_set = length_set or grouped
        width_set = width_set or grouped
        height_set = height_set or grouped

    if length_set and width_set and height_set:
        log.info(
            "Item size set: %s x %s x %s (weight %s)",
            row.package_length,
            row.package_width,
            row.package_height,
            row.item_weight if weight_set else "not set",
        )
    else:
        log.warning(
            "Item size was not fully entered (length=%s, width=%s, height=%s).",
            "ok" if length_set else "missing",
            "ok" if width_set else "missing",
            "ok" if height_set else "missing",
        )


def step_6_fill_returns_policy(driver: WebDriver, row: ListingRow) -> None:
    log.info("Step 6: Setting Returns Policy (30 Days)...")
    # Already on same continuous page - just scroll down a bit and click
    dismiss_discard_modal(driver)
    clicked = click_matching_element(driver, ["30 days", "30-day", "accept returns", "standard 30 days"])
    if clicked:
        log.info("Returns policy set to: 30 days.")


def _click_radio_option(driver: WebDriver, question_hint: str, answer_hint: str) -> bool:
    """
    Precisely selects a radio-card / radio-button option on Etsy's listing form.
    Finds the containing group by question text, then picks the matching answer.
    Falls back to page-wide label search if group detection fails.
    """
    script = """
    const qHint = arguments[0].replace(/[^a-z0-9]/gi, '').toLowerCase();
    const aHint = arguments[1].replace(/[^a-z0-9]/gi, '').toLowerCase();

    // Strategy 1: find fieldset / section containing the question heading
    const containers = Array.from(document.querySelectorAll('fieldset, [class*="section"], [class*="group"], [class*="card"], form > div'));
    for (const container of containers) {
        const heading = (container.querySelector('legend, h2, h3, h4, label, p') || container);
        const headTxt = (heading.innerText || heading.textContent || '').replace(/[^a-z0-9]/gi, '').toLowerCase();
        if (!headTxt.includes(qHint)) continue;
        // Found the container — now pick the matching radio/label inside
        const options = Array.from(container.querySelectorAll('label, [role="radio"], [role="option"], button, input[type="radio"]'));
        for (const opt of options) {
            const optTxt = (opt.innerText || opt.textContent || opt.value || opt.getAttribute('aria-label') || '').replace(/[^a-z0-9]/gi, '').toLowerCase();
            if (optTxt.includes(aHint)) {
                opt.scrollIntoView({block: 'center', behavior: 'instant'});
                opt.click();
                return 'container:' + optTxt.slice(0, 40);
            }
        }
    }

    // Strategy 2: page-wide label / radio search restricted to visible elements
    const allOpts = Array.from(document.querySelectorAll('label, [role="radio"], input[type="radio"]'));
    for (const opt of allOpts) {
        if (opt.offsetParent === null) continue;
        const optTxt = (opt.innerText || opt.textContent || opt.value || opt.getAttribute('aria-label') || '').replace(/[^a-z0-9]/gi, '').toLowerCase();
        if (optTxt.includes(aHint)) {
            opt.scrollIntoView({block: 'center', behavior: 'instant'});
            opt.click();
            return 'global:' + optTxt.slice(0, 40);
        }
    }
    return null;
    """
    try:
        result = driver.execute_script(script, question_hint, answer_hint)
        if result:
            time.sleep(0.6)
            log.info("  Radio clicked [%s -> %s]: %s", question_hint, answer_hint, result)
            return True
    except Exception as exc:
        log.warning("  Radio click error [%s -> %s]: %s", question_hint, answer_hint, exc)
    return False


def step_7_fill_how_its_made(driver: WebDriver, row: ListingRow) -> None:
    log.info("Step 7: Production Details (Who made it / What is it / When was it made)...")

    # No page-level scroll here — _click_radio_option handles scrollIntoView internally.
    # A large scrollBy triggers Etsy's SPA router and crashes ChromeDriver.
    time.sleep(0.5)
    dismiss_discard_modal(driver)

    # 1. Who made it? -> Another company
    who_clicked = (
        _click_radio_option(driver, "who made", "another company")
        or _click_radio_option(driver, "who", "another company")
        or click_matching_element(driver, ["Another company"], timeout=3)
    )
    log.info("Who made it -> Another company: %s", "OK" if who_clicked else "NOT SELECTED")
    time.sleep(0.6)
    dismiss_discard_modal(driver)

    # 2. What is it? -> A supply or tool
    what_clicked = (
        _click_radio_option(driver, "what is it", "supply or tool")
        or _click_radio_option(driver, "what", "supply")
        or click_matching_element(driver, ["A supply or tool", "supply or tool"], timeout=3)
    )
    log.info("What is it -> A supply or tool: %s", "OK" if what_clicked else "NOT SELECTED")
    time.sleep(0.6)
    dismiss_discard_modal(driver)

    # 3. When was it made? -> 2020-2026
    # Strategy A: look for a <select> element near the 'when' section
    when_clicked = False
    try:
        when_select_script = """
        // Find any select whose options include '2020'
        const selects = Array.from(document.querySelectorAll('select'));
        for (const sel of selects) {
            const opts = Array.from(sel.options);
            for (const opt of opts) {
                if (opt.text.includes('2020') || opt.value.includes('2020')) {
                    sel.scrollIntoView({block: 'center', behavior: 'instant'});
                    sel.value = opt.value;
                    sel.dispatchEvent(new Event('change', {bubbles: true}));
                    return 'select:' + opt.text;
                }
            }
        }
        return null;
        """
        sel_result = driver.execute_script(when_select_script)
        if sel_result:
            log.info("  When made via <select>: %s", sel_result)
            when_clicked = True
    except Exception as exc:
        log.warning("  When made select error: %s", exc)

    if not when_clicked:
        # Strategy B: radio-card / label with '2020'
        when_clicked = (
            _click_radio_option(driver, "when was it made", "2020 to 2026")
            or _click_radio_option(driver, "when", "made to order")
            or click_matching_element(driver, ["2020\u20132026", "2020 to 2026", "2020 - 2026"], timeout=3)
        )

    if not when_clicked:
        # Strategy C: find any radio input whose aria-label or value mentions 2020
        try:
            when_radio_script = """
            const radios = Array.from(document.querySelectorAll('input[type="radio"], [role="radio"]'));
            for (const r of radios) {
                const lbl = document.querySelector('label[for="' + r.id + '"]');
                const txt = (lbl ? lbl.innerText : '') + (r.getAttribute('aria-label') || '') + (r.value || '');
                if (txt.includes('2020')) {
                    r.scrollIntoView({block: 'center', behavior: 'instant'});
                    r.click();
                    if (lbl) lbl.click();
                    return 'radio:' + txt.slice(0, 40);
                }
            }
            return null;
            """
            res = driver.execute_script(when_radio_script)
            if res:
                log.info("  When made via radio: %s", res)
                when_clicked = True
        except Exception as exc:
            log.warning("  When made radio error: %s", exc)

    log.info("When was it made -> 2020-2026: %s", "OK" if when_clicked else "NOT SELECTED")
    time.sleep(0.5)
    dismiss_discard_modal(driver)


def step_8_legal_and_compliance(driver: WebDriver, row: ListingRow) -> None:
    log.info("Step 8: Legal & Compliance reviewed (Skipping optional fields).")
    driver.execute_script("window.scrollBy(0, 2000);")
    time.sleep(0.5)
    dismiss_discard_modal(driver)


def step_9_fill_settings(driver: WebDriver, row: ListingRow) -> None:
    log.info("Step 9: Listing Settings (Shop section, Feature, Renewal)...")
    driver.execute_script("window.scrollBy(0, 2000);")
    time.sleep(1)
    dismiss_discard_modal(driver)

    # 9a. Shop section
    if row.shop_section and row.shop_section.lower() != "none":
        try:
            for loc in [
                (By.XPATH, "//button[contains(., 'None') or contains(@id, 'section') or contains(@aria-label, 'Shop section')]"),
                (By.XPATH, "//select[contains(@name, 'section')]")
            ]:
                try:
                    el = driver.find_element(*loc)
                    if el.is_displayed():
                        if el.tag_name == "select":
                            Select(el).select_by_visible_text(row.shop_section)
                        else:
                            driver.execute_script("arguments[0].click();", el)
                            time.sleep(0.8)
                            click_matching_element(driver, [row.shop_section])
                        log.info("Shop section set: %s", row.shop_section)
                        break
                except Exception:
                    continue
        except Exception:
            pass

    # 9b. Feature this listing
    try:
        feat_switch = driver.find_element(
            By.XPATH,
            "//*[contains(normalize-space(.), 'Feature this listing') and string-length(normalize-space(.)) < 200]"
            "//*[self::button[@role='switch'] or self::input[@type='checkbox']]",
        )
        if feat_switch.is_displayed():
            is_active = feat_switch.get_attribute("aria-checked") == "true" if feat_switch.tag_name == "button" else feat_switch.is_selected()
            if is_active != row.featured:
                driver.execute_script("arguments[0].click();", feat_switch)
                log.info("Feature this listing set to: %s", row.featured)
    except Exception:
        pass

    # 9c. Renewal options (Automatic vs Manual)
    renewal = row.renewal.strip() or "Automatic"
    ren_clicked = click_matching_element(driver, [
        f"{renewal}",
        "Automatic" if renewal.lower() == "automatic" else "Manual",
        "I'll renew" if renewal.lower() == "manual" else "renew automatically"
    ])
    log.info("Renewal option (%s): %s", renewal, "Selected" if ren_clicked else "Checked")
    if not ren_clicked:
        try:
            for radio in driver.find_elements(By.XPATH, f"//input[@type='radio' and contains(@value, '{renewal.lower()}')]"):
                driver.execute_script("arguments[0].click();", radio)
                log.info("Renewal radio input clicked: %s", renewal)
                break
        except Exception:
            pass


def blocking_fields(driver: WebDriver) -> list[str]:
    """
    Reads Etsy's "Please fix the following errors" bar and returns the field names it
    lists, e.g. ['Title', 'Item size', 'Item weight']. Empty list means nothing is blocking.
    """
    script = """
    function visible(el) {
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none';
    }
    let banner = null;
    for (const el of document.querySelectorAll('div, section, p, span')) {
        if (!visible(el)) continue;
        const text = (el.innerText || '').trim();
        if (!/^please fix the following error/i.test(text) || text.length > 400) continue;
        if (!banner || text.length < (banner.innerText || '').length) banner = el;
    }
    if (!banner) return [];
    let box = banner;
    for (let i = 0; i < 3 && box && !box.querySelector('a, button'); i++) box = box.parentElement;
    const names = Array.from((box || banner).querySelectorAll('a, button'))
        .filter(visible)
        .map(el => (el.innerText || '').trim())
        .filter(t => t && t.length < 40 && !/cancel|preview|save|publish/i.test(t));
    if (names.length) return names;
    const tail = (banner.innerText || '').split(':').slice(1).join(':');
    return tail.split(/,|\\n/).map(s => s.trim()).filter(s => s && s.length < 40);
    """
    try:
        return [str(name) for name in (driver.execute_script(script) or [])]
    except Exception as exc:
        if _browser_gone(exc):
            raise RuntimeError("Chrome closed while saving the listing.") from exc
        return []


def step_10_finalize_and_publish(driver: WebDriver, row: ListingRow) -> str:
    log.info("Step 10: Final Checklist Before Publishing...")
    # Scroll to top so the Save as draft button is visible
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1.5)
    dismiss_discard_modal(driver)
    screenshot(driver, f"checklist_{row.sku}")

    label = "Save as draft" if not row.publish else "Publish"
    log.info("%s: %s...", "Safe Mode (Draft)" if not row.publish else "Publish mode", label)
    start_url = driver.current_url

    for attempt in range(1, 4):
        if not _click_footer_button(driver, label):
            raise RuntimeError(f"'{label}' button not found on the listing editor.")
        time.sleep(1.5)
        if row.publish:
            _confirm_publish_dialog(driver)
        outcome = _wait_for_save(driver, start_url)
        screenshot(driver, f"result_{row.sku}")
        if outcome == "saved":
            log.info("Listing %s successfully.", "saved as Draft" if not row.publish else "published")
            return "DRAFT SAVED" if not row.publish else "PUBLISHED"

        fields = blocking_fields(driver)
        if not fields:
            raise RuntimeError("Etsy did not confirm the save and showed no field errors. Check the result screenshot.")
        if attempt == 3:
            break
        log.warning("Etsy blocked the save. Fields to fix: %s. Fixing them and saving again.", ", ".join(fields))
        fix_blocking_fields(driver, row, fields)
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(0.8)

    fields = blocking_fields(driver)
    raise RuntimeError(f"Etsy still rejects these fields after retries: {', '.join(fields)}")


def _click_footer_button(driver: WebDriver, label: str) -> bool:
    script = """
    const wanted = arguments[0].toLowerCase();
    function visible(el) {
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none';
    }
    const btn = Array.from(document.querySelectorAll('button'))
        .filter(b => visible(b) && !b.disabled && !b.closest('[role="dialog"]'))
        .find(b => (b.innerText || '').trim().toLowerCase() === wanted);
    if (!btn) return false;
    btn.scrollIntoView({block: 'center'});
    btn.click();
    return true;
    """
    try:
        return bool(driver.execute_script(script, label))
    except Exception as exc:
        if _browser_gone(exc):
            raise RuntimeError("Chrome closed while saving the listing.") from exc
        return False


def _confirm_publish_dialog(driver: WebDriver) -> None:
    """Publishing opens a $0.20 fee confirmation. Click Publish inside that dialog."""
    script = """
    const dialog = Array.from(document.querySelectorAll('[role="dialog"], [aria-modal="true"]'))
        .find(d => d.getBoundingClientRect().height > 10 && /publish/i.test(d.innerText || ''));
    if (!dialog) return false;
    const btn = Array.from(dialog.querySelectorAll('button')).find(b => /^publish/i.test((b.innerText || '').trim()));
    if (!btn) return false;
    btn.click();
    return true;
    """
    for _ in range(6):
        try:
            if driver.execute_script(script):
                log.info("Confirmed the publish fee dialog.")
                return
        except Exception as exc:
            if _browser_gone(exc):
                raise RuntimeError("Chrome closed while publishing.") from exc
        time.sleep(0.5)


def _wait_for_save(driver: WebDriver, start_url: str, timeout: float = 15) -> str:
    """'saved' when Etsy leaves the editor or shows a saved message, 'blocked' when the fix-errors bar shows."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(1)
        try:
            url = driver.current_url
            if url != start_url and "/create" not in url:
                return "saved"
            text = (driver.execute_script("return document.body ? document.body.innerText.slice(0, 6000) : ''") or "").lower()
        except Exception as exc:
            if _browser_gone(exc):
                raise RuntimeError("Chrome closed while saving the listing.") from exc
            continue
        if "please fix the following error" in text:
            return "blocked"
        if any(bit in text for bit in ("draft saved", "saved as a draft", "listing saved", "changes saved", "your listing is live", "listing published")):
            return "saved"
    return "unknown"


def fix_blocking_fields(driver: WebDriver, row: ListingRow, fields: list[str]) -> None:
    names = " ".join(fields).lower()
    if "category" in names:
        select_etsy_category(driver, row.category, row.title)
    if "title" in names:
        click_listing_tab(driver, "Item Details")
        time.sleep(0.5)
        fill_title(driver, row)
    if "description" in names:
        click_listing_tab(driver, "Item Details")
        time.sleep(0.5)
        fill_description(driver, row)
    if any(word in names for word in ("price", "quantity", "sku")):
        step_4_fill_price_and_inventory(driver, row)
    if any(word in names for word in ("size", "weight", "length", "width", "height", "shipping")):
        step_5_fill_shipping_and_returns(driver, row)
    if any(word in names for word in ("who made", "what is it", "when")):
        step_7_fill_how_its_made(driver, row)


def create_etsy_listing(driver: WebDriver, row: ListingRow) -> str:
    """Executes the full 10-step sequence for a single listing."""
    step_1_open_add_listing(driver)
    step_2_fill_item_details(driver, row)
    step_3_fill_attributes_and_tags(driver, row)
    step_4_fill_price_and_inventory(driver, row)
    step_5_fill_shipping_and_returns(driver, row)
    step_6_fill_returns_policy(driver, row)
    step_7_fill_how_its_made(driver, row)
    step_8_legal_and_compliance(driver, row)
    step_9_fill_settings(driver, row)
    res = step_10_finalize_and_publish(driver, row)
    return res


# -------------------------------------------------------------
# Store Batch Processing
# -------------------------------------------------------------
def mark(rows: list[ListingRow], status: str, detail: str) -> None:
    for row in rows:
        if not row.status:
            row.status = status
            row.detail = detail


def process_store(
    cfg: dict,
    chrome: Path,
    user_data: Path,
    vpncmd: Path,
    store_info: dict,
    rows: list[ListingRow],
    *,
    use_vpn: bool,
    check_ip: bool,
) -> None:
    store_name = store_info["name"]
    profile = store_info["chrome_profile"]
    expected_ip = (store_info.get("ip") or "").strip()
    vpn_account = (store_info.get("vpn") or store_name).strip()
    port = int(cfg.get("debug_port") or 9222)
    driver: Optional[WebDriver] = None
    chrome_proc: Optional[subprocess.Popen] = None

    log.info("=" * 65)
    log.info("Processing Store: %s | Chrome: %s | VPN: %s | IP: %s", store_name, profile, vpn_account, expected_ip or "(none)")
    log.info("=" * 65)

    if use_vpn:
        disconnect_all_vpns(vpncmd, cfg)
        time.sleep(2)
        if not connect_store_vpn(vpncmd, vpn_account) or not wait_for_vpn(vpncmd, vpn_account):
            mark(rows, "FAILED", f"VPN {vpn_account} did not connect")
            return

    try:
        close_chrome(port)
        profile_dir = automation_user_data(user_data, profile)
        chrome_proc = start_chrome(chrome, profile_dir, profile, port)
        driver = connect_selenium(port)

        if check_ip:
            ok, shown = wait_for_expected_ip(driver, expected_ip) if expected_ip else (False, read_ip(driver))
            if expected_ip and not ok:
                msg = f"IP mismatch: expected {expected_ip}, detected {shown or 'unknown'}"
                log.error("%s: %s", store_name, msg)
                mark(rows, "FAILED", msg)
                return

        for idx, row in enumerate(rows, start=1):
            log.info("--- Store %s | Row %d/%d | SKU: %s ---", store_name, idx, len(rows), row.sku)
            try:
                res_status = create_etsy_listing(driver, row)
                row.status = "SUCCESS"
                row.detail = res_status
                time.sleep(ROW_DELAY)
            except Exception as exc:
                log.exception("Failed row %d (%s): %s", row.row_number, row.sku, exc)
                screenshot(driver, f"fail_{row.sku}")
                dump_form(driver, row.sku)
                row.status = "FAILED"
                row.detail = str(exc)

    except Exception as exc:
        log.exception("Chrome launch failed for %s", store_name)
        mark(rows, "FAILED", f"Chrome launch failed: {exc}")
    finally:
        try:
            if driver:
                driver.quit()
        except Exception:
            pass
        close_chrome(port, chrome_proc)
        if use_vpn:
            disconnect_all_vpns(vpncmd, cfg)


# -------------------------------------------------------------
# CLI Arguments & Entry Point
# -------------------------------------------------------------
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create Etsy listings via Selenium with dedicated Store IP & Chrome Profile isolation.")
    parser.add_argument("excel", nargs="?", default=str(DEFAULT_EXCEL), help="Path to Create.xlsx")
    parser.add_argument("--store", "-s", help="Run only this store name (e.g. AMH, TSS, GG)")
    parser.add_argument("--limit", "-l", type=int, help="Limit number of listings per store")
    parser.add_argument("--tags", "-t", help="Generate 13 Etsy SEO tags strictly following 20 rules for this title and exit")
    parser.add_argument("--no-vpn", action="store_true", help="Do not connect/disconnect SoftEther VPN")
    parser.add_argument("--skip-ip-check", action="store_true", help="Skip checking outgoing IP against stores.json")
    parser.add_argument("--publish", action="store_true", help="Publish listings directly (default is Save as Draft)")
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    # If --tags requested, run tag generator immediately and exit
    if args.tags:
        tags = generate_etsy_tags(args.tags)
        print(format_tags_output(tags))
        return 0

    excel_path = Path(args.excel)
    if not excel_path.is_file():
        # Check if sample exists, otherwise generate
        if not DEFAULT_EXCEL.is_file():
            from core.data_manager import create_sample_excel_template
            create_sample_excel_template(str(DEFAULT_EXCEL))
            log.info("Generated sample template at %s", DEFAULT_EXCEL)
        excel_path = DEFAULT_EXCEL

    acquire_run_lock()
    try:
        return _run_listings(args, excel_path)
    finally:
        release_run_lock()


def _run_listings(args: argparse.Namespace, excel_path: Path) -> int:
    cfg = _load_config()
    chrome, user_data, vpncmd = resolve_paths(cfg)

    log.info("Loading listings from: %s", excel_path)
    rows = read_excel(excel_path)
    if not rows:
        log.warning("No listings found in %s", excel_path)
        return 0

    if args.publish:
        for r in rows:
            r.publish = True

    # Group by store
    by_store: dict[str, list[ListingRow]] = {}
    for r in rows:
        by_store.setdefault(r.store, []).append(r)

    target_store = (args.store or "").strip().upper()
    if target_store:
        if target_store not in by_store:
            log.error("Store %s not found in Excel rows (%s)", target_store, ", ".join(by_store.keys()))
            return 1
        by_store = {target_store: by_store[target_store]}

    for store_name, store_rows in by_store.items():
        info = store_lookup(cfg, store_name)
        if not info:
            log.error("Store %s is in Excel but not in stores.json. Skipping.", store_name)
            mark(store_rows, "SKIPPED", "Not in stores.json")
            continue

        active_rows = store_rows[:args.limit] if args.limit else store_rows
        process_store(
            cfg=cfg,
            chrome=chrome,
            user_data=user_data,
            vpncmd=vpncmd,
            store_info=info,
            rows=active_rows,
            use_vpn=not args.no_vpn,
            check_ip=not args.skip_ip_check,
        )

    # Write output summary
    results_path = OUTPUT_DIR / f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    write_results(rows, results_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
