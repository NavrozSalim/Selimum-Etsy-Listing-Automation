"""
Proxy Helper & Chrome Proxy Authentication Extension Generator.
Ensures store IP isolation, tests proxy connectivity, and generates
custom Chrome extensions for authenticated HTTP/SOCKS5 proxies.
"""

import os
import re
import zipfile
import tempfile
import time
import requests
from typing import Dict, Any, Optional, Tuple


def parse_proxy_url(proxy_str: str) -> Dict[str, Any]:
    """
    Parses proxy strings in multiple standard formats:
    - http://user:pass@host:port
    - socks5://user:pass@host:port
    - host:port:user:pass
    - host:port
    - http://host:port
    """
    proxy_str = proxy_str.strip()
    if not proxy_str:
        return {}

    # Format: host:port:user:pass
    colon_parts = proxy_str.split(":")
    if len(colon_parts) == 4 and not proxy_str.startswith("http") and not proxy_str.startswith("socks"):
        return {
            "scheme": "http",
            "host": colon_parts[0],
            "port": int(colon_parts[1]),
            "username": colon_parts[2],
            "password": colon_parts[3]
        }

    # URL format: scheme://[user:pass@]host:port
    pattern = re.compile(r'^(?:(?P<scheme>https?|socks5)://)?(?:(?P<username>[^:@]+):(?P<password>[^@]+)@)?(?P<host>[^:]+):(?P<port>\d+)$')
    match = pattern.match(proxy_str)
    if match:
        data = match.groupdict()
        scheme = data.get("scheme") or "http"
        return {
            "scheme": scheme,
            "host": data["host"],
            "port": int(data["port"]),
            "username": data.get("username") or "",
            "password": data.get("password") or ""
        }

    raise ValueError(f"Invalid proxy format: {proxy_str}")


def test_proxy(proxy_str: str, timeout: int = 10) -> Dict[str, Any]:
    """
    Tests proxy connection and verifies the outgoing public IP address.
    """
    if not proxy_str:
        return {"success": False, "error": "No proxy provided"}

    try:
        parsed = parse_proxy_url(proxy_str)
        scheme = parsed["scheme"]
        host = parsed["host"]
        port = parsed["port"]
        user = parsed["username"]
        pwd = parsed["password"]

        if user and pwd:
            formatted_url = f"{scheme}://{user}:{pwd}@{host}:{port}"
        else:
            formatted_url = f"{scheme}://{host}:{port}"

        proxies = {
            "http": formatted_url,
            "https": formatted_url
        }

        start_time = time.time()
        # Test against ipify public IP service
        resp = requests.get("https://api.ipify.org?format=json", proxies=proxies, timeout=timeout)
        latency = int((time.time() - start_time) * 1000)

        if resp.status_code == 200:
            ip = resp.json().get("ip", "Unknown")
            return {
                "success": True,
                "ip": ip,
                "latency_ms": latency,
                "message": f"Connected! Public IP: {ip} ({latency}ms)"
            }
        else:
            return {
                "success": False,
                "error": f"HTTP {resp.status_code}: {resp.text}"
            }
    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }


def create_chrome_proxy_extension(proxy_str: str, output_dir: Optional[str] = None) -> str:
    """
    Creates an unpacked Chrome Extension (or ZIP) for proxy authentication.
    Returns path to the extension folder or zip file.
    """
    parsed = parse_proxy_url(proxy_str)
    host = parsed["host"]
    port = parsed["port"]
    user = parsed.get("username", "")
    pwd = parsed.get("password", "")
    scheme = parsed.get("scheme", "http")

    if not output_dir:
        output_dir = tempfile.mkdtemp(prefix="etsy_proxy_ext_")

    manifest_json = """{
    "version": "1.0.0",
    "manifest_version": 2,
    "name": "EtsyStoreProxyAuth",
    "permissions": [
        "proxy",
        "tabs",
        "unlimitedStorage",
        "storage",
        "<all_urls>",
        "webRequest",
        "webRequestBlocking"
    ],
    "background": {
        "scripts": ["background.js"]
    },
    "minimum_chrome_version":"22.0.0"
}"""

    background_js = f"""
var config = {{
    mode: "fixed_servers",
    rules: {{
        singleProxy: {{
            scheme: "{scheme}",
            host: "{host}",
            port: parseInt({port})
        }},
        bypassList: ["localhost", "127.0.0.1"]
    }}
}};

chrome.proxy.settings.set({{value: config, scope: "regular"}}, function() {{}});

chrome.webRequest.onAuthRequired.addListener(
    function(details) {{
        return {{
            authCredentials: {{
                username: "{user}",
                password: "{pwd}"
            }}
        }};
    }},
    {{urls: ["<all_urls>"]}},
    ["blocking"]
);
"""

    manifest_path = os.path.join(output_dir, "manifest.json")
    background_path = os.path.join(output_dir, "background.js")

    with open(manifest_path, "w", encoding="utf-8") as f:
        f.write(manifest_json)

    with open(background_path, "w", encoding="utf-8") as f:
        f.write(background_js)

    # Also build a zip file if needed by certain selenium configurations
    zip_path = os.path.join(output_dir, "proxy_auth.zip")
    with zipfile.ZipFile(zip_path, 'w') as zp:
        zp.write(manifest_path, "manifest.json")
        zp.write(background_path, "background.js")

    return output_dir
