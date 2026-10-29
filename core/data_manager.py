"""
Data Manager for Etsy Listing Queue.
Loads listings from Excel, CSV, or JSON.
Automatically fills defaults, parses image paths, and auto-generates
13 Etsy SEO tags if tags are not provided or requested to auto-generate.
"""

import os
import csv
import json
from typing import List, Dict, Any, Optional
import pandas as pd

from core.tag_engine import generate_etsy_tags, validate_tags, format_tags_output

DEFAULT_COLUMNS = [
    "title", "description", "images", "price", "quantity", "sku",
    "allow_offers", "category", "tags", "shipping_profile",
    "item_weight", "package_length", "package_width", "package_height",
    "return_policy", "who_made_it", "what_is_it", "when_made",
    "shop_section", "featured", "renewal", "status"
]

DEFAULT_VALUES = {
    "price": "19.99",
    "quantity": "100",
    "sku": "",
    "allow_offers": "no",
    "category": "Craft Supplies & Tools",
    "return_policy": "30 days",
    "who_made_it": "Another company",
    "what_is_it": "A supply or tool",
    "when_made": "2020–2026",
    "shop_section": "",
    "featured": "no",
    "renewal": "Automatic",
    "status": "Pending"
}


def load_listings_from_file(file_path: str, auto_gen_tags_if_empty: bool = True) -> List[Dict[str, Any]]:
    """Loads listings from Excel or CSV file."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(file_path, dtype=str)
    elif ext == ".csv":
        df = pd.read_csv(file_path, dtype=str)
    elif ext == ".json":
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return [normalize_listing_dict(item, auto_gen_tags_if_empty) for item in data]
    else:
        raise ValueError(f"Unsupported file format: {ext}")

    df = df.fillna("")
    records = df.to_dict(orient="records")
    normalized = []
    for row in records:
        normalized.append(normalize_listing_dict(row, auto_gen_tags_if_empty))
    return normalized


def normalize_listing_dict(raw: Dict[str, Any], auto_gen_tags: bool = True) -> Dict[str, Any]:
    """Cleans and sets sensible defaults for a single listing row."""
    item = {}
    for col in DEFAULT_COLUMNS:
        # Match case-insensitively
        val = ""
        for k, v in raw.items():
            if k.strip().lower() == col.lower():
                val = str(v).strip()
                break
        if not val and col in DEFAULT_VALUES:
            val = DEFAULT_VALUES[col]
        item[col] = val

    # Resolve image paths
    raw_images = item.get("images", "")
    item["image_list"] = parse_image_paths(raw_images)

    # Resolve / auto-generate tags
    existing_tags = item.get("tags", "").strip()
    if not existing_tags and auto_gen_tags and item.get("title"):
        generated = generate_etsy_tags(item["title"], item.get("description", ""))
        item["tags"] = ", ".join(generated)
        item["tag_list"] = generated
    elif existing_tags:
        parsed_tags = [t.strip() for t in existing_tags.split(",") if t.strip()]
        item["tag_list"] = parsed_tags
    else:
        item["tag_list"] = []

    return item


def parse_image_paths(raw_str: str) -> List[str]:
    """
    Parses images string. Can be:
    - semicolon or comma separated file paths
    - a directory path containing images
    """
    if not raw_str:
        return []

    # If it's a directory
    if os.path.isdir(raw_str):
        valid_exts = {".jpg", ".jpeg", ".png", ".webp"}
        imgs = []
        for fname in sorted(os.listdir(raw_str)):
            if os.path.splitext(fname)[1].lower() in valid_exts:
                imgs.append(os.path.join(raw_str, fname))
        return imgs

    # Delimited list
    delimiters = [";", ","]
    for d in delimiters:
        if d in raw_str:
            parts = [p.strip() for p in raw_str.split(d) if p.strip()]
            return [p for p in parts if os.path.exists(p)]

    if os.path.exists(raw_str):
        return [raw_str]

    return []


def create_sample_excel_template(output_path: str):
    """Generates an Excel template with comments, sample data, and schema."""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    sample_data = [
        {
            "title": "Burgundy Velvet Ribbon 2.5 Inch Wired Ribbon for Christmas Tree Bow Making, 10 Yards",
            "description": "Premium luxury velvet ribbon with wired edges for elegant holiday bows, gift wrapping, and Christmas tree decorations.",
            "images": "data/sample_images/ribbon1.jpg; data/sample_images/ribbon2.jpg",
            "price": "18.50",
            "quantity": "50",
            "sku": "RIB-BUR-01",
            "allow_offers": "no",
            "category": "Ribbons",
            "tags": "",  # Leaves blank to test automatic rule-based tag generation!
            "shipping_profile": "Standard Ground",
            "item_weight": "0.6",
            "package_length": "8",
            "package_width": "6",
            "package_height": "3",
            "return_policy": "30 days",
            "who_made_it": "Another company",
            "what_is_it": "A supply or tool",
            "when_made": "2020–2026",
            "shop_section": "Holiday Ribbons",
            "featured": "yes",
            "renewal": "Automatic",
            "status": "Pending"
        },
        {
            "title": "Blush Pink and Rose Gold Balloon Garland Arch Kit 120 Pcs for Baby Shower Decor and Birthday",
            "description": "Complete DIY balloon garland arch kit with assorted blush pink, rose gold, and white latex balloons for bridal showers and birthdays.",
            "images": "data/sample_images/balloon1.jpg",
            "price": "24.99",
            "quantity": "100",
            "sku": "BAL-BLUSH-02",
            "allow_offers": "no",
            "category": "Balloons",
            "tags": "",
            "shipping_profile": "Standard Ground",
            "item_weight": "0.9",
            "package_length": "10",
            "package_width": "8",
            "package_height": "2",
            "return_policy": "30 days",
            "who_made_it": "Another company",
            "what_is_it": "A supply or tool",
            "when_made": "2020–2026",
            "shop_section": "Balloon Kits",
            "featured": "no",
            "renewal": "Automatic",
            "status": "Pending"
        }
    ]

    df = pd.DataFrame(sample_data)
    df.to_excel(output_path, index=False)
    csv_path = os.path.splitext(output_path)[0] + ".csv"
    df.to_csv(csv_path, index=False)
    return output_path
