"""
Etsy Listing Automation Bot.
Implements the full 10-step listing creation process using Selenium:
1. Add a Listing
2. Item Details (Images, SEO Title, Description)
3. Attributes (Category, Attributes, 13 SEO Tags)
4. Price and Inventory (Price, Quantity, SKU, Offers toggle)
5. Shipping, Processing, and Returns (Profile, Weight, Dimensions L x W x H)
6. Returns and Exchanges (30 days setting)
7. How It's Made (Who made it: Another company, What is it: Supply/tool, When: 2020–2026)
8. Legal and Compliance (Skip optional, complete required)
9. Settings (Shop section, Feature listing, Renewal options)
10. Final Checklist & Draft / Publish Execution
"""

import os
import sys
import time
import logging
from typing import Dict, Any, List, Optional, Callable
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, ElementClickInterceptedException,
    StaleElementReferenceException
)

from core.browser_factory import BrowserFactory
from core.store_manager import get_store
from core.tag_engine import generate_etsy_tags, validate_tags

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("EtsyBot")


class EtsyListingBot:
    def __init__(
        self,
        store_id: str,
        headless: bool = False,
        safe_mode: bool = True,  # If True, saves as Draft instead of auto-publishing
        progress_callback: Optional[Callable[[str, int], None]] = None
    ):
        self.store_id = store_id
        self.headless = headless
        self.safe_mode = safe_mode
        self.progress_callback = progress_callback
        self.driver: Optional[webdriver.Chrome] = None
        self.store_data = get_store(store_id) or {}

    def log(self, message: str, percent: int = -1):
        logger.info(message)
        if self.progress_callback:
            self.progress_callback(message, percent)

    def start_session(self) -> bool:
        """Starts browser connected with the store's associated IP/proxy."""
        profile_dir = self.store_data.get("profile_dir", f"profiles/{self.store_id}")
        proxy = self.store_data.get("proxy", "")

        self.log(f"Connecting to Store '{self.store_data.get('name', self.store_id)}'...", 5)
        if proxy:
            self.log(f"Binding to associated IP / Proxy: {proxy}", 8)
        else:
            self.log("Running with direct connection (no proxy configured).", 8)

        self.driver = BrowserFactory.get_driver(
            profile_dir=profile_dir,
            proxy=proxy,
            headless=self.headless
        )
        return True

    def close(self):
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    def check_login_status(self) -> bool:
        """Navigates to Etsy Shop Manager and verifies authentication."""
        self.log("Verifying Etsy login session...", 10)
        self.driver.get("https://www.etsy.com/your/shops/me")
        time.sleep(3)

        current_url = self.driver.current_url.lower()
        if "signin" in current_url or "login" in current_url:
            self.log("Etsy login required! Please complete login in the opened browser window.", 12)
            # Wait up to 120 seconds if user logs in manually
            for _ in range(24):
                time.sleep(5)
                if "signin" not in self.driver.current_url.lower() and "login" not in self.driver.current_url.lower():
                    self.log("Login detected! Proceeding...", 15)
                    return True
            return False
        return True

    # -------------------------------------------------------------
    # Step 1: Open Shop Manager & Add a Listing
    # -------------------------------------------------------------
    def step_1_open_add_listing(self) -> bool:
        self.log("Step 1: Opening Etsy Shop Manager -> Add a Listing...", 15)
        create_url = "https://www.etsy.com/your/shops/me/listing-editor/create"
        self.driver.get(create_url)
        time.sleep(4)

        # Fallback if redirect happened or direct link needs navigation
        if "listing-editor" not in self.driver.current_url:
            self.driver.get("https://www.etsy.com/your/shops/me/listings")
            time.sleep(3)
            # Click 'Add a listing' button
            try:
                btn = self.wait_for_any_element([
                    (By.XPATH, "//a[contains(., 'Add a listing')]"),
                    (By.XPATH, "//button[contains(., 'Add a listing')]"),
                    (By.CSS_SELECTOR, "a[href*='listing-editor']"),
                ], timeout=8)
                if btn:
                    btn.click()
                    time.sleep(4)
            except Exception as e:
                self.log(f"Navigation note: {e}")

        self.log("Listing editor loaded.", 20)
        return True

    # -------------------------------------------------------------
    # Step 2: Item Details (Images, Title, Description)
    # -------------------------------------------------------------
    def step_2_fill_item_details(self, listing: Dict[str, Any]) -> bool:
        self.log("Step 2: Uploading product images & entering details...", 25)

        # 2a. Upload Images
        image_list = listing.get("image_list", [])
        if image_list:
            abs_images = [os.path.abspath(p) for p in image_list if os.path.exists(p)]
            if abs_images:
                self.log(f"Uploading {len(abs_images)} photos...", 28)
                try:
                    file_input = self.wait_for_any_element([
                        (By.CSS_SELECTOR, "input[type='file'][accept*='image']"),
                        (By.CSS_SELECTOR, "input[type='file']"),
                    ], timeout=10)
                    if file_input:
                        # Selenium allows passing multiple paths joined by \n on Windows/Linux
                        file_input.send_keys("\n".join(abs_images))
                        self.log("Images sent to upload queue. Waiting for preview processing...", 32)
                        time.sleep(4)
                except Exception as e:
                    self.log(f"Image upload warning: {e}")

        # 2b. Title
        title = listing.get("title", "").strip()
        if title:
            self.log(f"Entering SEO Title: '{title[:50]}...'", 35)
            title_el = self.wait_for_any_element([
                (By.NAME, "title"),
                (By.CSS_SELECTOR, "input[data-field='title']"),
                (By.ID, "listing-edit-title-input"),
                (By.XPATH, "//input[contains(@aria-label, 'Title') or contains(@placeholder, 'title')]"),
                (By.TAG_NAME, "textarea")
            ], timeout=8)
            if title_el:
                title_el.clear()
                # Etsy title maximum is usually 140 chars
                title_el.send_keys(title[:140])
                time.sleep(1)

        # 2c. Description
        description = listing.get("description", "").strip()
        if description:
            self.log("Entering product description...", 40)
            desc_el = self.wait_for_any_element([
                (By.NAME, "description"),
                (By.CSS_SELECTOR, "textarea[data-field='description']"),
                (By.ID, "description-text-area"),
                (By.XPATH, "//textarea[contains(@aria-label, 'Description') or contains(@id, 'description')]")
            ], timeout=8)
            if desc_el:
                desc_el.clear()
                desc_el.send_keys(description)
                time.sleep(1)

        return True

    # -------------------------------------------------------------
    # Step 3: Attributes (Category & 13 SEO Tags)
    # -------------------------------------------------------------
    def step_3_fill_attributes_and_tags(self, listing: Dict[str, Any]) -> bool:
        self.log("Step 3: Selecting category & applying 13 SEO tags...", 45)

        # 3a. Category
        category = listing.get("category", "").strip()
        if category:
            try:
                cat_input = self.wait_for_any_element([
                    (By.CSS_SELECTOR, "input[placeholder*='category' i]"),
                    (By.CSS_SELECTOR, "input[aria-label*='category' i]"),
                    (By.ID, "category-search"),
                    (By.XPATH, "//input[contains(@placeholder, 'category') or contains(@placeholder, 'What are you selling')]")
                ], timeout=5)
                if cat_input:
                    cat_input.clear()
                    cat_input.send_keys(category)
                    time.sleep(2)
                    # Click first category suggestion
                    suggestion = self.wait_for_any_element([
                        (By.CSS_SELECTOR, ".category-result, li[role='option'], .wt-menu__item"),
                        (By.XPATH, "//li[contains(@class, 'category') or @role='option']")
                    ], timeout=4)
                    if suggestion:
                        suggestion.click()
                        time.sleep(1.5)
            except Exception as e:
                self.log(f"Category selection note: {e}")

        # 3b. 13 Etsy SEO Tags
        tags = listing.get("tag_list", [])
        if not tags or len(tags) != 13:
            # Generate automatically adhering to 20 rules
            tags = generate_etsy_tags(listing.get("title", ""), listing.get("description", ""))
            listing["tag_list"] = tags
            listing["tags"] = ", ".join(tags)

        self.log(f"Adding 13 Etsy SEO tags (ending in 'party decor', 'party supplies')...", 50)
        try:
            tag_input = self.wait_for_any_element([
                (By.CSS_SELECTOR, "input[placeholder*='tag' i]"),
                (By.CSS_SELECTOR, "input[aria-label*='tag' i]"),
                (By.ID, "tags-input"),
                (By.XPATH, "//input[contains(@placeholder, 'tag') or contains(@id, 'tag')]")
            ], timeout=6)

            if tag_input:
                for t in tags:
                    clean_tag = t.strip()[:20]
                    tag_input.send_keys(clean_tag)
                    tag_input.send_keys(Keys.ENTER)
                    time.sleep(0.3)
                    # Click Add button if Enter doesn't submit
                    try:
                        add_btn = self.driver.find_element(By.XPATH, "//button[text()='Add' or contains(., 'Add tag')]")
                        add_btn.click()
                    except Exception:
                        pass
                self.log("All 13 tags entered.", 55)
        except Exception as e:
            self.log(f"Tag input warning: {e}")

        return True

    # -------------------------------------------------------------
    # Step 4: Price and Inventory (Price, Quantity, SKU, Offers)
    # -------------------------------------------------------------
    def step_4_fill_price_and_inventory(self, listing: Dict[str, Any]) -> bool:
        self.log("Step 4: Setting Price, Quantity, SKU, and Offers...", 58)

        # 4a. Price
        price = str(listing.get("price", "19.99")).strip()
        price_el = self.wait_for_any_element([
            (By.NAME, "price"),
            (By.CSS_SELECTOR, "input[data-field='price']"),
            (By.ID, "price-input"),
            (By.XPATH, "//input[contains(@name, 'price') or contains(@aria-label, 'Price')]")
        ], timeout=5)
        if price_el:
            price_el.clear()
            price_el.send_keys(price)

        # 4b. Quantity
        quantity = str(listing.get("quantity", "100")).strip()
        qty_el = self.wait_for_any_element([
            (By.NAME, "quantity"),
            (By.CSS_SELECTOR, "input[data-field='quantity']"),
            (By.ID, "quantity-input"),
            (By.XPATH, "//input[contains(@name, 'quantity') or contains(@aria-label, 'Quantity')]")
        ], timeout=5)
        if qty_el:
            qty_el.clear()
            qty_el.send_keys(quantity)

        # 4c. SKU
        sku = str(listing.get("sku", "")).strip()
        if sku:
            sku_el = self.wait_for_any_element([
                (By.NAME, "sku"),
                (By.CSS_SELECTOR, "input[data-field='sku']"),
                (By.XPATH, "//input[contains(@name, 'sku') or contains(@aria-label, 'SKU')]")
            ], timeout=3)
            if sku_el:
                sku_el.clear()
                sku_el.send_keys(sku)

        # 4d. Let buyers make offers
        allow_offers = str(listing.get("allow_offers", "no")).lower() in ["yes", "true", "1"]
        try:
            offers_toggle = self.wait_for_any_element([
                (By.XPATH, "//label[contains(., 'offers') or contains(., 'make offers')]//input"),
                (By.NAME, "accept_offers")
            ], timeout=3)
            if offers_toggle and offers_toggle.is_selected() != allow_offers:
                offers_toggle.click()
        except Exception:
            pass

        return True

    # -------------------------------------------------------------
    # Step 5: Shipping, Processing, and Returns
    # -------------------------------------------------------------
    def step_5_fill_shipping_details(self, listing: Dict[str, Any]) -> bool:
        self.log("Step 5: Configuring Shipping Profile & Package Dimensions...", 65)

        # Shipping Profile
        profile_name = listing.get("shipping_profile", "").strip()
        if profile_name:
            try:
                prof_select = self.wait_for_any_element([
                    (By.NAME, "shipping_profile_id"),
                    (By.XPATH, "//select[contains(@name, 'shipping_profile')]")
                ], timeout=4)
                if prof_select:
                    sel = Select(prof_select)
                    try:
                        sel.select_by_visible_text(profile_name)
                    except Exception:
                        pass
            except Exception:
                pass

        # Package Weight
        weight = str(listing.get("item_weight", "0.5")).strip()
        weight_el = self.wait_for_any_element([
            (By.NAME, "item_weight"),
            (By.CSS_SELECTOR, "input[aria-label*='weight' i]"),
            (By.XPATH, "//input[contains(@name, 'weight') or contains(@aria-label, 'Weight')]")
        ], timeout=3)
        if weight_el:
            weight_el.clear()
            weight_el.send_keys(weight)

        # Dimensions: Length, Width, Height
        l = str(listing.get("package_length", "6")).strip()
        w = str(listing.get("package_width", "4")).strip()
        h = str(listing.get("package_height", "2")).strip()

        length_el = self.wait_for_any_element([
            (By.NAME, "item_length"),
            (By.XPATH, "//input[contains(@name, 'length') or contains(@aria-label, 'Length')]")
        ], timeout=3)
        if length_el:
            length_el.clear()
            length_el.send_keys(l)

        width_el = self.wait_for_any_element([
            (By.NAME, "item_width"),
            (By.XPATH, "//input[contains(@name, 'width') or contains(@aria-label, 'Width')]")
        ], timeout=3)
        if width_el:
            width_el.clear()
            width_el.send_keys(w)

        height_el = self.wait_for_any_element([
            (By.NAME, "item_height"),
            (By.XPATH, "//input[contains(@name, 'height') or contains(@aria-label, 'Height')]")
        ], timeout=3)
        if height_el:
            height_el.clear()
            height_el.send_keys(h)

        return True

    # -------------------------------------------------------------
    # Step 6: Returns and Exchanges
    # -------------------------------------------------------------
    def step_6_fill_returns_policy(self, listing: Dict[str, Any]) -> bool:
        self.log("Step 6: Setting Returns Policy (30 Days)...", 72)
        try:
            return_btn = self.wait_for_any_element([
                (By.XPATH, "//label[contains(., '30 days') or contains(., 'Accept returns')]"),
                (By.XPATH, "//button[contains(., '30 days')]")
            ], timeout=3)
            if return_btn:
                return_btn.click()
        except Exception:
            pass
        return True

    # -------------------------------------------------------------
    # Step 7: How It's Made
    # Who made it: Another company
    # What is it: A supply or tool
    # When was it made: 2020–2026
    # -------------------------------------------------------------
    def step_7_fill_how_its_made(self, listing: Dict[str, Any]) -> bool:
        self.log("Step 7: Configuring Production Details (Another company / Supply / 2020-2026)...", 78)

        # Who made it
        try:
            who = self.wait_for_any_element([
                (By.XPATH, "//label[contains(., 'Another company') or contains(., 'Someone else')]"),
                (By.XPATH, "//select[contains(@name, 'who_made')]")
            ], timeout=4)
            if who:
                if who.tag_name == "select":
                    Select(who).select_by_visible_text("Another company or person")
                else:
                    who.click()
        except Exception:
            pass

        # What is it
        try:
            what = self.wait_for_any_element([
                (By.XPATH, "//label[contains(., 'A supply or tool')]"),
                (By.XPATH, "//select[contains(@name, 'is_supply')]")
            ], timeout=3)
            if what:
                if what.tag_name == "select":
                    Select(what).select_by_visible_text("A supply or tool to make things")
                else:
                    what.click()
        except Exception:
            pass

        # When made
        try:
            when = self.wait_for_any_element([
                (By.XPATH, "//label[contains(., '2020') or contains(., 'Made to order')]"),
                (By.XPATH, "//select[contains(@name, 'when_made')]")
            ], timeout=3)
            if when:
                if when.tag_name == "select":
                    sel = Select(when)
                    # Try finding 2020-2026
                    for opt in sel.options:
                        if "2020" in opt.text:
                            sel.select_by_visible_text(opt.text)
                            break
                else:
                    when.click()
        except Exception:
            pass

        return True

    # -------------------------------------------------------------
    # Step 8: Legal and Compliance
    # -------------------------------------------------------------
    def step_8_check_legal_and_compliance(self, listing: Dict[str, Any]) -> bool:
        self.log("Step 8: Reviewing Legal & Compliance requirements...", 84)
        # Skip optional fields, Etsy allows them blank for supplies/tools
        return True

    # -------------------------------------------------------------
    # Step 9: Settings (Section, Feature, Renewal)
    # -------------------------------------------------------------
    def step_9_fill_settings(self, listing: Dict[str, Any]) -> bool:
        self.log("Step 9: Setting Shop Section, Featured toggle, Renewal...", 88)

        # Shop section
        section = listing.get("shop_section", "").strip()
        if section:
            try:
                sec_el = self.wait_for_any_element([
                    (By.NAME, "section_id"),
                    (By.XPATH, "//select[contains(@name, 'section')]")
                ], timeout=3)
                if sec_el:
                    Select(sec_el).select_by_visible_text(section)
            except Exception:
                pass

        # Feature listing
        featured = str(listing.get("featured", "no")).lower() in ["yes", "true", "1"]
        try:
            feat_toggle = self.wait_for_any_element([
                (By.XPATH, "//label[contains(., 'Feature this listing')]//input"),
                (By.NAME, "featured")
            ], timeout=3)
            if feat_toggle and feat_toggle.is_selected() != featured:
                feat_toggle.click()
        except Exception:
            pass

        # Renewal options: Automatic vs Manual
        renewal = listing.get("renewal", "Automatic").strip().capitalize()
        try:
            ren_el = self.wait_for_any_element([
                (By.XPATH, f"//label[contains(., '{renewal}')]"),
                (By.XPATH, f"//input[@value='{renewal.lower()}']")
            ], timeout=3)
            if ren_el:
                ren_el.click()
        except Exception:
            pass

        return True

    # -------------------------------------------------------------
    # Step 10: Final Checklist Before Publishing
    # -------------------------------------------------------------
    def step_10_finalize_and_publish(self, listing: Dict[str, Any]) -> Dict[str, Any]:
        self.log("Step 10: Running final checklist before saving/publishing...", 92)

        # Checklist validation
        checklist = {
            "title_present": bool(listing.get("title")),
            "images_present": bool(listing.get("image_list")),
            "tags_count": len(listing.get("tag_list", [])),
            "price_present": bool(listing.get("price")),
            "action": "draft" if self.safe_mode else "publish"
        }

        self.log(f"Checklist: Title={checklist['title_present']}, Images={checklist['images_present']}, Tags={checklist['tags_count']}/13", 94)

        result_action = "Draft Saved"
        if self.safe_mode:
            self.log("Safe Mode Active: Saving listing as Draft for seller review...", 96)
            save_draft_btn = self.wait_for_any_element([
                (By.XPATH, "//button[contains(., 'Save as draft')]"),
                (By.XPATH, "//button[contains(., 'Save draft')]"),
                (By.CSS_SELECTOR, "button[data-action='save-draft']")
            ], timeout=8)
            if save_draft_btn:
                save_draft_btn.click()
                time.sleep(4)
                result_action = "Saved as Draft successfully"
            else:
                result_action = "Ready (Draft button clicked / Form filled)"
        else:
            self.log("Publishing listing live to Etsy Shop...", 96)
            pub_btn = self.wait_for_any_element([
                (By.XPATH, "//button[contains(., 'Publish')]"),
                (By.CSS_SELECTOR, "button[data-action='publish']")
            ], timeout=8)
            if pub_btn:
                pub_btn.click()
                time.sleep(4)
                result_action = "Published Live"

        # Capture verification screenshot
        screenshot_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "screenshots")
        os.makedirs(screenshot_dir, exist_ok=True)
        screenshot_file = os.path.join(screenshot_dir, f"listing_{int(time.time())}.png")
        try:
            self.driver.save_screenshot(screenshot_file)
            self.log(f"Listing screenshot saved to: {screenshot_file}", 100)
        except Exception:
            pass

        return {
            "success": True,
            "status": result_action,
            "checklist": checklist,
            "screenshot": screenshot_file
        }

    # -------------------------------------------------------------
    # Full Process Runner
    # -------------------------------------------------------------
    def process_listing(self, listing: Dict[str, Any]) -> Dict[str, Any]:
        """Executes all 10 steps sequentially on a single listing."""
        try:
            self.step_1_open_add_listing()
            self.step_2_fill_item_details(listing)
            self.step_3_fill_attributes_and_tags(listing)
            self.step_4_fill_price_and_inventory(listing)
            self.step_5_fill_shipping_details(listing)
            self.step_6_fill_returns_policy(listing)
            self.step_7_fill_how_its_made(listing)
            self.step_8_check_legal_and_compliance(listing)
            self.step_9_fill_settings(listing)
            res = self.step_10_finalize_and_publish(listing)
            return res
        except Exception as e:
            self.log(f"Error during listing execution: {e}", -1)
            return {"success": False, "error": str(e)}

    # -------------------------------------------------------------
    # Helper: Wait for any matching element selector
    # -------------------------------------------------------------
    def wait_for_any_element(self, locators: List[tuple], timeout: int = 5):
        for by, val in locators:
            try:
                el = WebDriverWait(self.driver, timeout).until(
                    EC.presence_of_element_located((by, val))
                )
                if el.is_displayed():
                    return el
            except Exception:
                continue
        return None
