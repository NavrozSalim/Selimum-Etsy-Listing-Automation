"""
Modern PyQt6 Desktop GUI for Etsy Listing Bot & 20-Rule SEO Tag Generator.
Features:
- Store & Associated IP/Proxy Manager with Live IP Tester.
- Excel / CSV Listing Queue with Batch Runner & Progress Tracking.
- Standalone 20-Rule Etsy SEO Tag Generator with Live Validation.
- Real-time Automation Console and Status Monitoring.
"""

import os
import sys
import threading
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QLineEdit, QPushButton, QTextEdit, QTableWidget,
    QTableWidgetItem, QHeaderView, QFileDialog, QMessageBox, QCheckBox,
    QProgressBar, QGroupBox, QSplitter, QComboBox, QFormLayout
)
from PyQt6.QtCore import Qt, pyqtSignal, QObject
from PyQt6.QtGui import QFont, QColor, QIcon

from core.tag_engine import generate_etsy_tags, validate_tags, format_tags_output
from core.store_manager import get_all_stores, save_or_update_store, delete_store, launch_store_session
from core.proxy_helper import test_proxy
from core.data_manager import load_listings_from_file, create_sample_excel_template
from core.etsy_bot import EtsyListingBot


class WorkerSignals(QObject):
    log_message = pyqtSignal(str, int)
    finished = pyqtSignal(dict)
    proxy_result = pyqtSignal(dict)


class EtsyBotGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Etsy Listing Automation Bot & SEO Tag Engine")
        self.resize(1150, 750)
        self.signals = WorkerSignals()
        self.signals.log_message.connect(self.on_log_message)
        self.signals.proxy_result.connect(self.on_proxy_test_result)
        self.signals.finished.connect(self.on_automation_finished)

        self.loaded_listings = []
        self.bot_thread = None
        self.init_ui()
        self.apply_dark_theme()
        self.refresh_stores_list()

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # Header Title
        header_layout = QHBoxLayout()
        title_label = QLabel("⚡ Etsy Multi-Store Listing Bot")
        title_label.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        title_label.setStyleSheet("color: #F1641E;")  # Etsy signature orange
        subtitle = QLabel("Dedicated Store IP Isolation • 10-Step Automation • 20-Rule SEO Tag Engine")
        subtitle.setStyleSheet("color: #888888; font-size: 13px;")

        header_layout.addWidget(title_label)
        header_layout.addSpacing(15)
        header_layout.addWidget(subtitle)
        header_layout.addStretch()
        main_layout.addLayout(header_layout)

        # Tabs
        self.tabs = QTabWidget()
        self.tabs.setFont(QFont("Segoe UI", 10))
        self.tab_tag_engine = QWidget()
        self.tab_stores = QWidget()
        self.tab_listings = QWidget()

        self.setup_tag_engine_tab()
        self.setup_stores_tab()
        self.setup_listings_tab()

        self.tabs.addTab(self.tab_tag_engine, "🏷️ Etsy SEO Tag Generator (20 Rules)")
        self.tabs.addTab(self.tab_stores, "🌐 Store & Associated IP Manager")
        self.tabs.addTab(self.tab_listings, "🚀 Listing Automation & Queue")

        main_layout.addWidget(self.tabs, stretch=3)

        # Bottom Console Log
        console_group = QGroupBox("📋 Automation & Execution Console")
        console_layout = QVBoxLayout(console_group)
        self.console_text = QTextEdit()
        self.console_text.setReadOnly(True)
        self.console_text.setFont(QFont("Consolas", 10))
        self.console_text.setStyleSheet("background-color: #121212; color: #00FF66; border: 1px solid #2A2A2A;")
        console_layout.addWidget(self.console_text)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #333;
                border-radius: 4px;
                text-align: center;
                background-color: #1E1E1E;
                color: #FFF;
            }
            QProgressBar::chunk {
                background-color: #F1641E;
                border-radius: 3px;
            }
        """)
        console_layout.addWidget(self.progress_bar)

        main_layout.addWidget(console_group, stretch=1)

    # -------------------------------------------------------------
    # Tab 1: 20-Rule Etsy SEO Tag Generator
    # -------------------------------------------------------------
    def setup_tag_engine_tab(self):
        layout = QVBoxLayout(self.tab_tag_engine)
        layout.setSpacing(10)

        info_box = QLabel(
            "<b>Strict 20-Rule Tag Engine:</b> Generates exactly 13 tags (< 20 chars), "
            "eliminates measurement/quantity words, applies diverse SEO angles (colors, materials, uses, events), "
            "and strictly sets the last 2 tags to <code>party decor</code> and <code>party supplies</code>."
        )
        info_box.setStyleSheet("background: #23272E; padding: 10px; border-radius: 6px; color: #DCDCDC;")
        layout.addWidget(info_box)

        form_layout = QFormLayout()
        self.input_title = QLineEdit()
        self.input_title.setPlaceholderText("Paste product title here (e.g., Burgundy Velvet Ribbon 2.5 Inch Wired...)")
        self.input_title.setFont(QFont("Segoe UI", 11))
        self.input_title.setStyleSheet("padding: 8px; background: #1E1E1E; color: white; border: 1px solid #444;")

        self.input_desc = QLineEdit()
        self.input_desc.setPlaceholderText("Optional: Add product description or keywords...")
        self.input_desc.setFont(QFont("Segoe UI", 10))
        self.input_desc.setStyleSheet("padding: 6px; background: #1E1E1E; color: white; border: 1px solid #444;")

        form_layout.addRow("<b>Product Title:</b>", self.input_title)
        form_layout.addRow("<b>Description / Notes:</b>", self.input_desc)
        layout.addLayout(form_layout)

        btn_layout = QHBoxLayout()
        self.btn_gen_tags = QPushButton("⚡ Generate 13 Etsy Tags")
        self.btn_gen_tags.setStyleSheet("background-color: #F1641E; color: white; font-weight: bold; padding: 10px; font-size: 13px;")
        self.btn_gen_tags.clicked.connect(self.on_generate_tags_clicked)

        self.btn_copy_tags = QPushButton("📋 Copy Comma-Separated Tags")
        self.btn_copy_tags.setStyleSheet("background-color: #2D3748; color: white; padding: 10px;")
        self.btn_copy_tags.clicked.connect(self.on_copy_tags_clicked)

        btn_layout.addWidget(self.btn_gen_tags)
        btn_layout.addWidget(self.btn_copy_tags)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # Output Text Box (Single Code Block)
        self.output_tags_box = QTextEdit()
        self.output_tags_box.setFont(QFont("Consolas", 11))
        self.output_tags_box.setStyleSheet("background-color: #181A1F; color: #ABB2BF; padding: 10px; border: 1px solid #333;")
        self.output_tags_box.setPlaceholderText("Generated 13 tags in ONE code block will appear here...")
        layout.addWidget(self.output_tags_box)

        # Validation status label
        self.validation_label = QLabel("Status: Ready to generate tags.")
        self.validation_label.setFont(QFont("Segoe UI", 10))
        self.validation_label.setStyleSheet("color: #A0AEC0;")
        layout.addWidget(self.validation_label)

    # -------------------------------------------------------------
    # Tab 2: Store & Associated IP Manager
    # -------------------------------------------------------------
    def setup_stores_tab(self):
        layout = QVBoxLayout(self.tab_stores)

        store_info = QLabel(
            "<b>Multi-Store Isolation:</b> Each store has its own dedicated IP/Proxy and persistent Chrome Profile. "
            "Click <b>'Open Store Browser / Login'</b> to open Chrome with that store's IP, log in to Etsy once, "
            "and all sessions/cookies remain permanently saved without repeated 2FA challenges."
        )
        store_info.setStyleSheet("background: #23272E; padding: 10px; border-radius: 6px; color: #DCDCDC;")
        layout.addWidget(store_info)

        # Form to add/edit store
        form_group = QGroupBox("Add or Edit Store Profile")
        form_layout = QFormLayout(form_group)

        self.input_store_name = QLineEdit()
        self.input_store_name.setPlaceholderText("Store Name (e.g. My Ribbon Shop)")
        self.input_store_proxy = QLineEdit()
        self.input_store_proxy.setPlaceholderText("http://user:pass@host:port or host:port:user:pass (leave blank for direct)")

        form_layout.addRow("Store Name:", self.input_store_name)
        form_layout.addRow("Associated IP / Proxy:", self.input_store_proxy)

        btn_box = QHBoxLayout()
        self.btn_save_store = QPushButton("💾 Save Store")
        self.btn_save_store.setStyleSheet("background-color: #3182CE; color: white; padding: 6px 12px;")
        self.btn_save_store.clicked.connect(self.on_save_store_clicked)

        self.btn_test_proxy = QPushButton("🔍 Test IP & Latency")
        self.btn_test_proxy.setStyleSheet("background-color: #4A5568; color: white; padding: 6px 12px;")
        self.btn_test_proxy.clicked.connect(self.on_test_proxy_clicked)

        btn_box.addWidget(self.btn_save_store)
        btn_box.addWidget(self.btn_test_proxy)
        btn_box.addStretch()
        form_layout.addRow("", btn_box)

        layout.addWidget(form_group)

        # Stores Table
        self.stores_table = QTableWidget()
        self.stores_table.setColumnCount(4)
        self.stores_table.setHorizontalHeaderLabels(["Store Name", "Associated IP / Proxy", "Profile Directory", "Actions"])
        self.stores_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.stores_table)

    # -------------------------------------------------------------
    # Tab 3: Listing Automation & Queue
    # -------------------------------------------------------------
    def setup_listings_tab(self):
        layout = QVBoxLayout(self.tab_listings)

        # Action bar
        actions_bar = QHBoxLayout()

        self.store_selector = QComboBox()
        self.store_selector.setMinimumWidth(200)

        self.btn_load_file = QPushButton("📂 Import Excel/CSV")
        self.btn_load_file.setStyleSheet("background-color: #2D3748; color: white; padding: 6px 12px;")
        self.btn_load_file.clicked.connect(self.on_load_file_clicked)

        self.btn_template = QPushButton("📄 Download Sample Excel")
        self.btn_template.setStyleSheet("background-color: #2D3748; color: white; padding: 6px 12px;")
        self.btn_template.clicked.connect(self.on_download_template_clicked)

        self.chk_safe_mode = QCheckBox("Safe Mode: Save as Draft (Recommended)")
        self.chk_safe_mode.setChecked(True)
        self.chk_safe_mode.setToolTip("Saves listings to Drafts so you can review before spending listing fees.")

        self.chk_headless = QCheckBox("Headless Browser")
        self.chk_headless.setChecked(False)

        self.btn_run_automation = QPushButton("🚀 Start 10-Step Automation")
        self.btn_run_automation.setStyleSheet("background-color: #38A169; color: white; font-weight: bold; padding: 8px 16px;")
        self.btn_run_automation.clicked.connect(self.on_start_automation_clicked)

        actions_bar.addWidget(QLabel("Select Store:"))
        actions_bar.addWidget(self.store_selector)
        actions_bar.addWidget(self.btn_load_file)
        actions_bar.addWidget(self.btn_template)
        actions_bar.addSpacing(10)
        actions_bar.addWidget(self.chk_safe_mode)
        actions_bar.addWidget(self.chk_headless)
        actions_bar.addWidget(self.btn_run_automation)
        layout.addLayout(actions_bar)

        # Listings Table
        self.listings_table = QTableWidget()
        self.listings_table.setColumnCount(7)
        self.listings_table.setHorizontalHeaderLabels([
            "Title", "Price", "Qty", "SKU", "Tags Count", "Images", "Status"
        ])
        self.listings_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.listings_table)

    # -------------------------------------------------------------
    # Logic: Tag Generator
    # -------------------------------------------------------------
    def on_generate_tags_clicked(self):
        title = self.input_title.text().strip()
        if not title:
            QMessageBox.warning(self, "Title Required", "Please paste or enter a product title first.")
            return

        desc = self.input_desc.text().strip()
        tags = generate_etsy_tags(title, desc)
        formatted = format_tags_output(tags)
        self.output_tags_box.setText(formatted)

        is_valid, errs = validate_tags(tags)
        if is_valid:
            self.validation_label.setText(
                "✅ <b>All 20 Rules Passed!</b> Exactly 13 tags • ≤20 chars each • "
                "No banned measurements • Ends with 'party decor, party supplies'."
            )
            self.validation_label.setStyleSheet("color: #48BB78; font-weight: bold;")
        else:
            self.validation_label.setText(f"⚠️ Validation notes: {'; '.join(errs)}")
            self.validation_label.setStyleSheet("color: #E53E3E;")

    def on_copy_tags_clicked(self):
        text = self.output_tags_box.toPlainText()
        if not text:
            return
        # Extract content between ``` if present
        clean_text = text.replace("```", "").strip()
        clipboard = QApplication.clipboard()
        clipboard.setText(clean_text)
        QMessageBox.information(self, "Copied", "13 Etsy tags copied to clipboard!")

    # -------------------------------------------------------------
    # Logic: Stores
    # -------------------------------------------------------------
    def refresh_stores_list(self):
        stores = get_all_stores()
        self.stores_table.setRowCount(len(stores))
        self.store_selector.clear()

        for row, s in enumerate(stores):
            name = s.get("name", "")
            proxy = s.get("proxy", "") or "Direct (No Proxy)"
            profile = s.get("profile_dir", "")
            store_id = s.get("store_id", "")

            self.store_selector.addItem(f"{name} ({proxy})", store_id)

            self.stores_table.setItem(row, 0, QTableWidgetItem(name))
            self.stores_table.setItem(row, 1, QTableWidgetItem(proxy))
            self.stores_table.setItem(row, 2, QTableWidgetItem(profile))

            # Action Buttons: Launch Browser / Delete
            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(2, 2, 2, 2)

            btn_launch = QPushButton("🌐 Open Browser / Login")
            btn_launch.setStyleSheet("background-color: #805AD5; color: white; padding: 4px 8px;")
            btn_launch.clicked.connect(lambda ch, sid=store_id: self.on_launch_store_browser(sid))

            btn_del = QPushButton("❌")
            btn_del.setStyleSheet("background-color: #E53E3E; color: white; padding: 4px;")
            btn_del.clicked.connect(lambda ch, sid=store_id: self.on_delete_store(sid))

            action_layout.addWidget(btn_launch)
            action_layout.addWidget(btn_del)
            self.stores_table.setCellWidget(row, 3, action_widget)

    def on_save_store_clicked(self):
        name = self.input_store_name.text().strip()
        proxy = self.input_store_proxy.text().strip()
        if not name:
            QMessageBox.warning(self, "Name Required", "Please enter a store name.")
            return

        store_id = name.lower().replace(" ", "_")
        save_or_update_store({
            "store_id": store_id,
            "name": name,
            "proxy": proxy,
            "profile_dir": f"profiles/{store_id}"
        })
        self.input_store_name.clear()
        self.input_store_proxy.clear()
        self.refresh_stores_list()
        QMessageBox.information(self, "Store Saved", f"Store '{name}' saved successfully!")

    def on_delete_store(self, store_id: str):
        if QMessageBox.question(self, "Confirm Delete", f"Delete store '{store_id}'?") == QMessageBox.StandardButton.Yes:
            delete_store(store_id)
            self.refresh_stores_list()

    def on_test_proxy_clicked(self):
        proxy = self.input_store_proxy.text().strip()
        if not proxy:
            QMessageBox.warning(self, "Proxy Required", "Enter a proxy string to test.")
            return

        self.console_text.append(f"Testing proxy: {proxy}...")

        def _test():
            res = test_proxy(proxy)
            self.signals.proxy_result.emit(res)

        threading.Thread(target=_test, daemon=True).start()

    def on_proxy_test_result(self, res: dict):
        if res.get("success"):
            ip = res.get("ip")
            latency = res.get("latency_ms")
            msg = f"✅ Proxy Verified! Public IP: {ip} | Latency: {latency}ms"
            self.console_text.append(msg)
            QMessageBox.information(self, "Proxy Test Success", msg)
        else:
            err = res.get("error")
            msg = f"❌ Proxy Test Failed: {err}"
            self.console_text.append(msg)
            QMessageBox.critical(self, "Proxy Test Failed", msg)

    def on_launch_store_browser(self, store_id: str):
        self.console_text.append(f"Opening browser for store '{store_id}' with persistent profile and associated IP...")
        threading.Thread(target=lambda: launch_store_session(store_id), daemon=True).start()

    # -------------------------------------------------------------
    # Logic: Listings Queue & Automation
    # -------------------------------------------------------------
    def on_download_template_clicked(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Excel Template", "etsy_listings_template.xlsx", "Excel Files (*.xlsx)")
        if path:
            create_sample_excel_template(path)
            QMessageBox.information(self, "Template Saved", f"Sample template saved to:\n{path}")

    def on_load_file_clicked(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Listings File", "", "Spreadsheets (*.xlsx *.csv *.json)")
        if not path:
            return

        try:
            listings = load_listings_from_file(path, auto_gen_tags_if_empty=True)
            self.loaded_listings = listings
            self.populate_listings_table(listings)
            self.console_text.append(f"Imported {len(listings)} listings from {os.path.basename(path)}")
        except Exception as e:
            QMessageBox.critical(self, "Error Loading File", str(e))

    def populate_listings_table(self, listings):
        self.listings_table.setRowCount(len(listings))
        for row, item in enumerate(listings):
            self.listings_table.setItem(row, 0, QTableWidgetItem(item.get("title", "")))
            self.listings_table.setItem(row, 1, QTableWidgetItem(f"${item.get('price', '')}"))
            self.listings_table.setItem(row, 2, QTableWidgetItem(str(item.get("quantity", ""))))
            self.listings_table.setItem(row, 3, QTableWidgetItem(item.get("sku", "")))
            tags_count = len(item.get("tag_list", []))
            self.listings_table.setItem(row, 4, QTableWidgetItem(f"{tags_count}/13"))
            imgs_count = len(item.get("image_list", []))
            self.listings_table.setItem(row, 5, QTableWidgetItem(f"{imgs_count} photos"))
            self.listings_table.setItem(row, 6, QTableWidgetItem(item.get("status", "Pending")))

    def on_start_automation_clicked(self):
        if not self.loaded_listings:
            QMessageBox.warning(self, "No Listings", "Please import an Excel or CSV file first.")
            return

        store_id = self.store_selector.currentData()
        if not store_id:
            QMessageBox.warning(self, "No Store", "Please create and select a store profile first.")
            return

        safe_mode = self.chk_safe_mode.isChecked()
        headless = self.chk_headless.isChecked()

        self.btn_run_automation.setEnabled(False)
        self.console_text.append(f"Starting Etsy Listing Bot for Store '{store_id}' (Safe Mode={safe_mode})...")

        def _run():
            bot = EtsyListingBot(
                store_id=store_id,
                headless=headless,
                safe_mode=safe_mode,
                progress_callback=lambda msg, p: self.signals.log_message.emit(msg, p)
            )
            try:
                bot.start_session()
                logged_in = bot.check_login_status()
                if not logged_in:
                    self.signals.log_message.emit("Authentication failed or timed out.", -1)
                    bot.close()
                    self.signals.finished.emit({"success": False, "error": "Not logged in to Etsy"})
                    return

                for idx, listing in enumerate(self.loaded_listings):
                    self.signals.log_message.emit(f"Processing listing #{idx+1}/{len(self.loaded_listings)}: {listing.get('title')[:40]}...", 10)
                    res = bot.process_listing(listing)
                    listing["status"] = res.get("status", "Completed")

                bot.close()
                self.signals.finished.emit({"success": True, "count": len(self.loaded_listings)})
            except Exception as e:
                self.signals.log_message.emit(f"Automation aborted: {e}", -1)
                bot.close()
                self.signals.finished.emit({"success": False, "error": str(e)})

        self.bot_thread = threading.Thread(target=_run, daemon=True)
        self.bot_thread.start()

    def on_log_message(self, message: str, percent: int):
        self.console_text.append(message)
        if percent >= 0:
            self.progress_bar.setValue(percent)

    def on_automation_finished(self, result: dict):
        self.btn_run_automation.setEnabled(True)
        self.populate_listings_table(self.loaded_listings)
        if result.get("success"):
            QMessageBox.information(self, "Automation Complete", f"Successfully processed {result.get('count')} listings!")
        else:
            QMessageBox.warning(self, "Automation Warning", f"Finished with note: {result.get('error')}")

    def apply_dark_theme(self):
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #1A1C23;
                color: #E2E8F0;
            }
            QGroupBox {
                border: 1px solid #2D3748;
                border-radius: 6px;
                margin-top: 10px;
                font-weight: bold;
                padding-top: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #CBD5E0;
            }
            QTabWidget::pane {
                border: 1px solid #2D3748;
                background-color: #1E212B;
            }
            QTabBar::tab {
                background-color: #171923;
                color: #A0AEC0;
                padding: 8px 16px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background-color: #1E212B;
                color: #F1641E;
                font-weight: bold;
            }
            QTableWidget {
                background-color: #171923;
                gridline-color: #2D3748;
                color: #E2E8F0;
                border: 1px solid #2D3748;
            }
            QHeaderView::section {
                background-color: #23272E;
                color: #A0AEC0;
                padding: 4px;
                border: 1px solid #2D3748;
                font-weight: bold;
            }
            QLineEdit, QComboBox {
                background-color: #23272E;
                color: white;
                border: 1px solid #4A5568;
                border-radius: 4px;
                padding: 6px;
            }
            QPushButton {
                border-radius: 4px;
                font-size: 11px;
            }
            QPushButton:hover {
                opacity: 0.9;
            }
        """)


def main():
    app = QApplication(sys.argv)
    window = EtsyBotGUI()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
