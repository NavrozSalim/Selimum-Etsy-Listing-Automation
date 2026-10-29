"""
Browser Factory for Etsy Multi-Store Isolation.
Supports undetected-chromedriver and stealth Selenium with:
- Persistent profile directories per store (keeps cookies, logins, 2FA saved).
- Store-specific Proxy / Associated IP binding.
- Anti-fingerprint masking (Blink features, navigator.webdriver, WebGL, WebRTC).
"""

import os
import sys
import time
from typing import Optional, Dict, Any
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService

from core.proxy_helper import create_chrome_proxy_extension, parse_proxy_url

try:
    import undetected_chromedriver as uc
    UC_AVAILABLE = True
except ImportError:
    UC_AVAILABLE = False

try:
    import selenium_stealth
    STEALTH_AVAILABLE = True
except ImportError:
    STEALTH_AVAILABLE = False


class BrowserFactory:
    @staticmethod
    def get_driver(
        profile_dir: str,
        proxy: Optional[str] = None,
        headless: bool = False,
        use_uc: bool = False  # default to standard stealth Chrome for maximum stability with Chrome 155
    ) -> webdriver.Chrome:
        """
        Creates and returns a Chrome webdriver configured for a specific store.
        - profile_dir: Path to persistent Chrome user profile directory.
        - proxy: Optional proxy string (http://user:pass@host:port or host:port).
        - headless: Whether to run in headless mode.
        """
        os.makedirs(profile_dir, exist_ok=True)
        abs_profile_dir = os.path.abspath(profile_dir)

        options = ChromeOptions()
        options.add_argument(f"--user-data-dir={abs_profile_dir}")
        options.add_argument("--profile-directory=Default")
        options.add_argument("--no-first-run")
        options.add_argument("--no-service-autorun")
        options.add_argument("--password-store=basic")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--start-maximized")
        options.add_argument("--window-size=1400,900")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--lang=en-US,en")

        # Configure Proxy
        proxy_ext_dir = None
        if proxy and proxy.strip():
            parsed = parse_proxy_url(proxy)
            if parsed.get("username") and parsed.get("password"):
                # Authenticated proxy requires an extension
                proxy_ext_dir = create_chrome_proxy_extension(proxy)
                options.add_argument(f"--load-extension={proxy_ext_dir}")
            else:
                scheme = parsed.get("scheme", "http")
                host = parsed.get("host")
                port = parsed.get("port")
                options.add_argument(f"--proxy-server={scheme}://{host}:{port}")

        if headless:
            options.add_argument("--headless=new")

        # Exclude automation switches
        options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        options.add_experimental_option("useAutomationExtension", False)

        prefs = {
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False,
            "profile.default_content_setting_values.notifications": 2
        }
        options.add_experimental_option("prefs", prefs)

        # Standard Selenium with Stealth injection
        driver = webdriver.Chrome(options=options)

        # Apply selenium-stealth masking
        if STEALTH_AVAILABLE:
            try:
                selenium_stealth.stealth(
                    driver,
                    languages=["en-US", "en"],
                    vendor="Google Inc.",
                    platform="Win32",
                    webgl_vendor="Intel Inc.",
                    renderer="Intel Iris OpenGL Engine",
                    fix_hairline=True,
                )
            except Exception as e:
                print(f"[BrowserFactory] Warning: stealth patch failed: {e}")

        # Remove webdriver flag via CDP
        try:
            driver.execute_cdp_cmd(
                "Page.addScriptToEvaluateOnNewDocument",
                {
                    "source": """
                    Object.defineProperty(navigator, 'webdriver', {
                        get: () => undefined
                    });
                    """
                }
            )
        except Exception:
            pass

        return driver
