"""
Store Manager for Multi-Account & Dedicated IP/Proxy Isolation.
Manages persistent Chrome profiles, proxy mappings, and login sessions.
"""

import os
import json
import time
from typing import List, Dict, Any, Optional

from core.proxy_helper import test_proxy
from core.browser_factory import BrowserFactory

STORES_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "stores.json")


def ensure_config_dir():
    os.makedirs(os.path.dirname(STORES_FILE), exist_ok=True)
    if not os.path.exists(STORES_FILE):
        with open(STORES_FILE, "w", encoding="utf-8") as f:
            json.dump([
                {
                    "store_id": "default_store",
                    "name": "Etsy Primary Store",
                    "proxy": "",
                    "profile_dir": "profiles/default_store",
                    "notes": "Primary Etsy store profile"
                }
            ], f, indent=4)


def get_all_stores() -> List[Dict[str, Any]]:
    ensure_config_dir()
    try:
        with open(STORES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def get_store(store_id: str) -> Optional[Dict[str, Any]]:
    stores = get_all_stores()
    for s in stores:
        if s.get("store_id") == store_id:
            return s
    return None


def save_or_update_store(store_data: Dict[str, Any]) -> Dict[str, Any]:
    stores = get_all_stores()
    store_id = store_data.get("store_id") or f"store_{int(time.time())}"
    store_data["store_id"] = store_id

    if not store_data.get("profile_dir"):
        store_data["profile_dir"] = f"profiles/{store_id}"

    updated = False
    for i, s in enumerate(stores):
        if s.get("store_id") == store_id:
            stores[i] = store_data
            updated = True
            break

    if not updated:
        stores.append(store_data)

    with open(STORES_FILE, "w", encoding="utf-8") as f:
        json.dump(stores, f, indent=4)

    return store_data


def delete_store(store_id: str) -> bool:
    stores = get_all_stores()
    new_stores = [s for s in stores if s.get("store_id") != store_id]
    if len(new_stores) < len(stores):
        with open(STORES_FILE, "w", encoding="utf-8") as f:
            json.dump(new_stores, f, indent=4)
        return True
    return False


def launch_store_session(store_id: str, start_url: str = "https://www.etsy.com/your/shops/me"):
    """
    Launches Chrome with the store's associated IP/proxy and persistent profile.
    Allows user to log into Etsy manually, complete 2FA, and store cookies.
    """
    store = get_store(store_id)
    if not store:
        raise ValueError(f"Store '{store_id}' not found.")

    proxy = store.get("proxy", "")
    profile_dir = store.get("profile_dir", f"profiles/{store_id}")

    print(f"[StoreSession] Launching browser for {store.get('name')} with profile '{profile_dir}'...")
    if proxy:
        print(f"[StoreSession] Binding store to proxy: {proxy}")

    driver = BrowserFactory.get_driver(profile_dir=profile_dir, proxy=proxy, headless=False)
    driver.get(start_url)
    print(f"[StoreSession] Browser opened to {start_url}. Keep this browser open or log in and close when done.")
    return driver
