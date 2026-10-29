"""
Inspect Etsy Listing Form DOM.
Connects via VPN and Chrome debug port, navigates to listing editor,
and dumps the exact structure of:
- Category selection
- Who made it
- What is it
- When was it made
"""

from pathlib import Path
import json
import time
from selenium.webdriver.common.by import By
from create_listing import (
    _load_config, resolve_paths, store_lookup, connect_store_vpn, wait_for_vpn,
    close_chrome, automation_user_data, start_chrome, connect_selenium,
    read_ip, wait_for_expected_ip, ETSY_CREATE_URL, wait_for_human_or_login
)

def inspect_form(store_name="IQ"):
    cfg = _load_config()
    chrome, user_data, vpncmd = resolve_paths(cfg)
    info = store_lookup(cfg, store_name)
    profile = info["chrome_profile"]
    expected_ip = info["ip"]
    vpn_account = info["vpn"]
    port = int(cfg.get("debug_port") or 9222)

    print(f"Connecting VPN {vpn_account}...")
    connect_store_vpn(vpncmd, vpn_account)
    wait_for_vpn(vpncmd, vpn_account)

    close_chrome()
    profile_dir = automation_user_data(user_data, profile)
    start_chrome(chrome, profile_dir, profile, port)
    driver = connect_selenium(port)

    ok, ip = wait_for_expected_ip(driver, expected_ip)
    print(f"IP Verified: {ip}")

    print("Navigating to Etsy Listing Editor...")
    driver.get(ETSY_CREATE_URL)
    time.sleep(5)
    wait_for_human_or_login(driver)

    # Dump all labels and inputs
    data = driver.execute_script("""
        const result = {
            headings: [],
            labels: [],
            radios: [],
            selects: [],
            inputs: []
        };
        
        document.querySelectorAll('h1, h2, h3, h4, legend').forEach(h => {
            if (h.innerText) result.headings.push(h.innerText.trim());
        });
        
        document.querySelectorAll('label').forEach(l => {
            if (l.innerText) result.labels.push(l.innerText.trim().replace(/\\n/g, ' '));
        });
        
        document.querySelectorAll('input[type="radio"]').forEach(r => {
            result.radios.push({
                name: r.name,
                value: r.value,
                id: r.id,
                label: r.closest('label')?.innerText?.trim() || r.parentElement?.innerText?.trim()
            });
        });
        
        document.querySelectorAll('select').forEach(s => {
            const opts = Array.from(s.options).map(o => o.text.trim());
            result.selects.push({
                name: s.name,
                id: s.id,
                options: opts
            });
        });
        
        document.querySelectorAll('input, textarea').forEach(i => {
            result.inputs.push({
                type: i.type,
                name: i.name,
                id: i.id,
                placeholder: i.placeholder,
                aria: i.getAttribute('aria-label')
            });
        });
        
        return result;
    """)

    with open("etsy_form_dump.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

    print("Dump saved to etsy_form_dump.json")
    print(f"Found {len(data['headings'])} headings, {len(data['labels'])} labels, {len(data['radios'])} radios, {len(data['selects'])} selects, {len(data['inputs'])} inputs.")
    
    # Take screenshot of page
    driver.save_screenshot("etsy_editor_inspect.png")
    print("Screenshot saved to etsy_editor_inspect.png")
    
    driver.quit()
    close_chrome()

if __name__ == "__main__":
    inspect_form("IQ")
