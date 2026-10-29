"""
Command Line Interface (CLI) for Etsy Listing Bot & SEO Tag Engine.
"""

import sys
import argparse
from core.tag_engine import generate_etsy_tags, validate_tags, format_tags_output
from core.store_manager import get_all_stores, save_or_update_store, delete_store, launch_store_session
from core.proxy_helper import test_proxy
from core.data_manager import load_listings_from_file, create_sample_excel_template
from core.etsy_bot import EtsyListingBot


def cmd_tags(args):
    """Generate 13 Etsy SEO tags strictly complying with the 20 rules."""
    title = args.title
    description = args.description or ""
    tags = generate_etsy_tags(title, description)
    print(format_tags_output(tags))

    if args.validate:
        is_valid, errors = validate_tags(tags)
        print("\n--- 20-Rule Validation Report ---")
        if is_valid:
            print("Status: PASSED ALL 20 RULES")
            print(f"- Total tags: {len(tags)}")
            print(f"- Characters: All <= 20")
            print(f"- Last 2 tags: '{tags[-2]}', '{tags[-1]}'")
        else:
            print("Status: FAILED")
            for err in errors:
                print(f"- {err}")


def cmd_list_stores(args):
    stores = get_all_stores()
    print(f"\nFound {len(stores)} store profiles:")
    print("-" * 65)
    for s in stores:
        print(f"ID:      {s.get('store_id')}")
        print(f"Name:    {s.get('name')}")
        print(f"Proxy:   {s.get('proxy') or 'Direct (No Proxy)'}")
        print(f"Profile: {s.get('profile_dir')}")
        print("-" * 65)


def cmd_add_store(args):
    store_data = {
        "store_id": args.id,
        "name": args.name or args.id,
        "proxy": args.proxy or "",
        "profile_dir": f"profiles/{args.id}"
    }
    save_or_update_store(store_data)
    print(f"Store '{args.id}' saved successfully.")


def cmd_test_proxy(args):
    print(f"Testing proxy connection: {args.proxy}...")
    res = test_proxy(args.proxy)
    if res.get("success"):
        print(f"SUCCESS: Connected! Public IP: {res.get('ip')} (Latency: {res.get('latency_ms')}ms)")
    else:
        print(f"FAILED: {res.get('error')}")


def cmd_open_store(args):
    print(f"Launching browser for store '{args.id}'...")
    launch_store_session(args.id)


def cmd_template(args):
    out = args.output or "data/etsy_listings_template.xlsx"
    create_sample_excel_template(out)
    print(f"Template created at: {out}")


def cmd_run(args):
    store_id = args.store
    file_path = args.file
    safe_mode = not args.publish

    print(f"Loading listings from {file_path}...")
    listings = load_listings_from_file(file_path, auto_gen_tags_if_empty=True)
    print(f"Loaded {len(listings)} listings.")

    bot = EtsyListingBot(store_id=store_id, headless=args.headless, safe_mode=safe_mode)
    try:
        bot.start_session()
        if not bot.check_login_status():
            print("Failed to verify Etsy login. Please run 'python cli.py open-store --id <store>' to log in first.")
            return

        for idx, item in enumerate(listings, 1):
            print(f"\nProcessing [{idx}/{len(listings)}]: {item.get('title')}")
            res = bot.process_listing(item)
            print(f"Result: {res.get('status')}")

    finally:
        bot.close()


def main():
    parser = argparse.ArgumentParser(description="Etsy Listing Automation & Tag Generator CLI")
    subparsers = parser.add_subparsers(dest="command")

    # Tags command
    p_tags = subparsers.add_parser("tags", help="Generate 13 Etsy SEO tags")
    p_tags.add_argument("title", help="Product title")
    p_tags.add_argument("--description", "-d", default="", help="Product description")
    p_tags.add_argument("--validate", "-v", action="store_true", help="Print validation report")
    p_tags.set_defaults(func=cmd_tags)

    # Store commands
    p_list = subparsers.add_parser("list-stores", help="List stores")
    p_list.set_defaults(func=cmd_list_stores)

    p_add = subparsers.add_parser("add-store", help="Add or update store")
    p_add.add_argument("--id", required=True, help="Store ID (slug)")
    p_add.add_argument("--name", default="", help="Store Display Name")
    p_add.add_argument("--proxy", default="", help="Proxy URL (http://user:pass@host:port)")
    p_add.set_defaults(func=cmd_add_store)

    p_test = subparsers.add_parser("test-proxy", help="Test proxy connection")
    p_test.add_argument("--proxy", required=True, help="Proxy URL")
    p_test.set_defaults(func=cmd_test_proxy)

    p_open = subparsers.add_parser("open-store", help="Open store browser session")
    p_open.add_argument("--id", required=True, help="Store ID")
    p_open.set_defaults(func=cmd_open_store)

    p_temp = subparsers.add_parser("template", help="Generate Excel/CSV template")
    p_temp.add_argument("--output", "-o", default="data/etsy_listings_template.xlsx", help="Output path")
    p_temp.set_defaults(func=cmd_template)

    p_run = subparsers.add_parser("run", help="Run listing automation")
    p_run.add_argument("--store", "-s", required=True, help="Store ID")
    p_run.add_argument("--file", "-f", required=True, help="Path to Excel/CSV")
    p_run.add_argument("--publish", action="store_true", help="Publish directly instead of saving as draft")
    p_run.add_argument("--headless", action="store_true", help="Run headless")
    p_run.set_defaults(func=cmd_run)

    args = parser.parse_args()
    if hasattr(args, "func"):
        args.func(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
