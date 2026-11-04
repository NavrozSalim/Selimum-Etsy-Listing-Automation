#!/usr/bin/env python3
"""
Replace the SKU on existing Etsy listings.

Excel columns:
  store_name, sku, new_sku, upc

sku is the SKU on the listing now. new_sku replaces it. A blank upc is left unchanged.
The listing is opened in the standard editor (never Quick listing) and saved. No listing is created.

Output, written on every stop (finish, captcha, crash, user stop):
  output/sku_replace/Updated(YYYY-MM-DD).xlsx
  output/sku_replace/remain-to-update(YYYY-MM-DD).xlsx

Usage:
  python sku_replace/sku_replace.py sku_replace/SkuReplace.xlsx
  python sku_replace/sku_replace.py --create-sample
  python sku_replace/sku_replace.py SkuReplace.xlsx --store AMH --limit 1 --no-vpn --skip-ip-check
"""
from __future__ import annotations

import argparse
import logging
import sys
from dataclasses import dataclass
from pathlib import Path

BOT_DIR = Path(__file__).resolve().parent
PROJECT = BOT_DIR.parent
LOG_PATH = BOT_DIR / "sku_replace.log"
DEFAULT_EXCEL = BOT_DIR / "SkuReplace.xlsx"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("sku_replace")

sys.path.insert(0, str(PROJECT))

from openpyxl import Workbook, load_workbook

import create_listing as cl
import etsy_edit


HEADER_MAP = {
    "store name": "store",
    "store_name": "store",
    "storename": "store",
    "store": "store",
    "sku": "sku",
    "old_sku": "sku",
    "new_sku": "new_sku",
    "new sku": "new_sku",
    "newsku": "new_sku",
    "upc": "upc",
    "gtin": "upc",
    "ean": "upc",
}


@dataclass
class SkuRow:
    row_number: int
    store: str
    sku: str
    new_sku: str
    upc: str
    status: str = ""
    detail: str = ""


def create_sample(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "input"
    ws.append(["store_name", "sku", "new_sku", "upc"])
    ws.append(["AMH", "OLD-SKU-001", "NEW-SKU-001", ""])
    for cell in ws[2]:
        cell.number_format = "@"
    wb.save(path)
    log.info("Wrote sample Excel: %s", path)


def read_excel(path: Path) -> list[SkuRow]:
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
    missing = [name for name in ("store", "sku", "new_sku") if name not in index]
    if missing:
        raise SystemExit(f"Excel is missing columns: {', '.join(missing)}")

    def cell(values: tuple, key: str):
        pos = index.get(key)
        if pos is None or pos >= len(values):
            return None
        return values[pos]

    rows: list[SkuRow] = []
    for row_number, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not values:
            continue
        store = cl.coerce_code(cell(values, "store")).upper()
        sku = cl.coerce_code(cell(values, "sku"))
        new_sku = cl.coerce_code(cell(values, "new_sku"))
        upc = cl.coerce_code(cell(values, "upc")) if "upc" in index else ""
        if not store and not sku and not new_sku:
            continue
        row = SkuRow(row_number=row_number, store=store, sku=sku, new_sku=new_sku, upc=upc)
        if not sku:
            row.status = "FAILED"
            row.detail = "SKU is blank"
        elif not new_sku:
            row.status = "FAILED"
            row.detail = "NEW SKU is blank"
        elif new_sku.lower() == sku.lower() and not upc:
            row.status = "FAILED"
            row.detail = "NEW SKU is the same as SKU"
        rows.append(row)
    wb.close()
    return rows


def apply_row(driver, row: SkuRow) -> str:
    return etsy_edit.update_sku(driver, row.sku, row.new_sku, row.upc)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Replace SKU on existing Etsy listings from Excel")
    parser.add_argument("excel", nargs="?", default=str(DEFAULT_EXCEL), help="Excel file path")
    parser.add_argument("--store", help="Process only this store name")
    parser.add_argument("--limit", type=int, default=0, help="Max rows per store (0 = all)")
    parser.add_argument("--create-sample", action="store_true", help="Write SkuReplace.xlsx and exit")
    parser.add_argument("--no-vpn", action="store_true")
    parser.add_argument("--skip-ip-check", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.create_sample:
        create_sample(DEFAULT_EXCEL)
        return 0
    excel_path = Path(args.excel)
    if not excel_path.is_file():
        log.error("Excel not found: %s", excel_path)
        log.error("Create one with: python sku_replace/sku_replace.py --create-sample")
        return 1
    rows = read_excel(excel_path)
    cl.acquire_run_lock()
    try:
        return etsy_edit.run_job(args, excel_path, rows, apply_row, cl.OUTPUT_DIR / "sku_replace")
    finally:
        cl.release_run_lock()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        log.error("Stopped by user")
        raise SystemExit(130)
