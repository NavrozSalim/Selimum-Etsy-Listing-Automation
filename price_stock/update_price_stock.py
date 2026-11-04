#!/usr/bin/env python3
"""
Update price and quantity on existing Etsy listings.

Excel columns:
  store_name, sku, price, quantity

sku is the SKU already on the listing. A blank price or quantity cell is left unchanged.
At least one of price or quantity is required. The listing is opened in the standard
editor (never Quick listing) and saved, so the live listing changes. No listing is created.

Output, written on every stop (finish, captcha, crash, user stop):
  output/price_stock/Updated(YYYY-MM-DD).xlsx
  output/price_stock/remain-to-update(YYYY-MM-DD).xlsx

Usage:
  python price_stock/update_price_stock.py price_stock/PriceStock.xlsx
  python price_stock/update_price_stock.py --create-sample
  python price_stock/update_price_stock.py PriceStock.xlsx --store AMH --limit 1 --no-vpn --skip-ip-check
"""
from __future__ import annotations

import argparse
import logging
import sys
from dataclasses import dataclass
from pathlib import Path

BOT_DIR = Path(__file__).resolve().parent
PROJECT = BOT_DIR.parent
LOG_PATH = BOT_DIR / "update_price_stock.log"
DEFAULT_EXCEL = BOT_DIR / "PriceStock.xlsx"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("update_price_stock")

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
    "price": "price",
    "quantity": "quantity",
    "qty": "quantity",
    "stock": "quantity",
    "inventory": "quantity",
}


@dataclass
class PriceRow:
    row_number: int
    store: str
    sku: str
    price: str
    quantity: str
    status: str = ""
    detail: str = ""


def create_sample(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "input"
    ws.append(["store_name", "sku", "price", "quantity"])
    ws.append(["AMH", "SKU-001", "19.99", "5"])
    for cell in ws[2]:
        cell.number_format = "@"
    wb.save(path)
    log.info("Wrote sample Excel: %s", path)


def read_excel(path: Path) -> list[PriceRow]:
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
    missing = [name for name in ("store", "sku") if name not in index]
    if missing:
        raise SystemExit(f"Excel is missing columns: {', '.join(missing)}")
    if "price" not in index and "quantity" not in index:
        raise SystemExit("Excel needs a price column or a quantity column")

    def cell(values: tuple, key: str):
        pos = index.get(key)
        if pos is None or pos >= len(values):
            return None
        return values[pos]

    rows: list[PriceRow] = []
    for row_number, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not values:
            continue
        store = cl.coerce_code(cell(values, "store")).upper()
        sku = cl.coerce_code(cell(values, "sku"))
        price = cl.coerce_money(cell(values, "price")) if "price" in index else ""
        quantity = cl.coerce_code(cell(values, "quantity")) if "quantity" in index else ""
        if not store and not sku and not price and not quantity:
            continue
        row = PriceRow(row_number=row_number, store=store, sku=sku, price=price, quantity=quantity)
        if not sku:
            row.status = "FAILED"
            row.detail = "SKU is blank"
        elif not price and not quantity:
            row.status = "FAILED"
            row.detail = "Price or quantity is required"
        elif price and not _is_number(price):
            row.status = "FAILED"
            row.detail = f"Price is not a number: {price}"
        elif quantity and not quantity.isdigit():
            row.status = "FAILED"
            row.detail = f"Quantity is not a whole number: {quantity}"
        rows.append(row)
    wb.close()
    return rows


def _is_number(text: str) -> bool:
    try:
        return float(text) > 0
    except ValueError:
        return False


def apply_row(driver, row: PriceRow) -> str:
    return etsy_edit.update_price_and_stock(driver, row.sku, row.price, row.quantity)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Update Etsy listing price and quantity from Excel")
    parser.add_argument("excel", nargs="?", default=str(DEFAULT_EXCEL), help="Excel file path")
    parser.add_argument("--store", help="Process only this store name")
    parser.add_argument("--limit", type=int, default=0, help="Max rows per store (0 = all)")
    parser.add_argument("--create-sample", action="store_true", help="Write PriceStock.xlsx and exit")
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
        log.error("Create one with: python price_stock/update_price_stock.py --create-sample")
        return 1
    rows = read_excel(excel_path)
    cl.acquire_run_lock()
    try:
        return etsy_edit.run_job(args, excel_path, rows, apply_row, cl.OUTPUT_DIR / "price_stock")
    finally:
        cl.release_run_lock()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        log.error("Stopped by user")
        raise SystemExit(130)
