"""
Create root Create.xlsx with proper columns and sample items matching the user's stores.
"""
from pathlib import Path
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent
DEST = ROOT / "Create.xlsx"

wb = Workbook()
ws = wb.active
ws.title = "Listings"

headers = [
    "store_name", "title", "description", "price", "quantity", "sku",
    "photos", "category", "tags", "shipping_profile", "item_weight",
    "package_length", "package_width", "package_height", "return_policy",
    "who_made_it", "what_is_it", "when_made", "shop_section", "allow_offers",
    "featured", "renewal", "status", "publish", "remove_branding"
]
ws.append(headers)

sample_rows = [
    [
        "AMH",
        "Burgundy Velvet Ribbon 2.5 Inch Wired Ribbon for Christmas Tree Bow Making, 10 Yards",
        "Luxury dark red velvet wired ribbon perfect for holiday tree decor, wreath crafting, gift wrapping, and floral arrangements.",
        "18.99",
        "100",
        "RIB-BUR-001",
        "data/sample_images/ribbon1.jpg; data/sample_images/ribbon2.jpg",
        "Ribbons",
        "",  # Left blank to trigger the automatic 20-Rule SEO Tag Generator!
        "Standard Shipping",
        "0.6",
        "8",
        "6",
        "3",
        "30 days",
        "Another company",
        "A supply or tool",
        "2020–2026",
        "Holiday Ribbons",
        "no",
        "yes",
        "Automatic",
        "Pending",
        "no",  # Save as Draft for safe seller review
        ""
    ],
    [
        "TSS",
        "Blush Pink and Rose Gold Balloon Garland Arch Kit 120 Pcs for Baby Shower Decor and Birthday",
        "Elegant DIY balloon garland arch kit featuring metallic rose gold, pastel blush pink, and pearl white latex balloons for bridal showers and celebrations.",
        "24.99",
        "150",
        "BAL-ROSE-002",
        "data/sample_images/balloon1.jpg",
        "Balloons",
        "",  # Auto-generated with 20 rules
        "Standard Shipping",
        "0.9",
        "10",
        "8",
        "2",
        "30 days",
        "Another company",
        "A supply or tool",
        "2020–2026",
        "Party Supplies",
        "no",
        "no",
        "Automatic",
        "Pending",
        "no",
        ""
    ]
]

for row in sample_rows:
    ws.append(row)

wb.save(DEST)
print(f"Generated {DEST}")
