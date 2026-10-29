# ⚡ Selenium Etsy Multi-Store Listing Automation Bot

A production-grade Etsy listing automation tool modeled directly after the **Reverb Listing Automation** architecture.

Built specifically for high-volume Etsy sellers managing multiple stores with dedicated IP/VPN connections, isolated Chrome profiles, strict anti-detection, and a built-in **20-Rule Etsy SEO Tag Engine**.

---

## 🌟 Key Features

1. **Dedicated Store IP & VPN Connection**:
   - Integrated with **SoftEther VPN Client** (`vpncmd.exe`) matching your existing Reverb workflow.
   - Connects the store's designated VPN before opening Chrome, verifies outgoing public IP via `https://nordvpn.com/what-is-my-ip/` and `https://api.ipify.org`, and disconnects cleanly when done.
   - Also supports dedicated HTTP/SOCKS5 proxy per store.

2. **Persistent Chrome Profile Isolation**:
   - Uses your real Chrome browser profiles (`Profile 15`, `Profile 17`, etc.) via remote debugging (`--remote-debugging-port=9222`).
   - Keeps your Etsy cookies, login sessions, and 2FA permanently saved so you never get repeatedly challenged.
   - Completely circumvents Cloudflare / bot detection.

3. **Complete 10-Step Etsy Listing Automation**:
   - **Step 1: Add a Listing** (Opens Shop Manager -> Listing Editor)
   - **Step 2: Item Details** (Uploads photos from URLs/local files, SEO Title, Description)
   - **Step 3: Attributes** (Category selection, attributes, and 13 SEO tags)
   - **Step 4: Price & Inventory** (Price, Quantity, SKU, Offers toggle)
   - **Step 5: Shipping, Processing, and Returns** (Shipping profile, weight, package dimensions L x W x H)
   - **Step 6: Returns & Exchanges** (30 days setting)
   - **Step 7: How It's Made** (Who made it: Another company / What is it: A supply or tool / When: 2020–2026)
   - **Step 8: Legal and Compliance** (Skips optional, completes required)
   - **Step 9: Settings** (Shop section, Feature listing, Renewal options)
   - **Step 10: Final Checklist Before Publishing** (Safe Draft mode or Publish mode)

4. **Strict 20-Rule Etsy SEO Tag Engine**:
   - Generates exactly **13 tags** in **ONE code block**, comma-separated.
   - Every tag is **20 characters or fewer** strictly enforced.
   - Prohibits all measurements (`inches`, `yard`, `cm`, `feet`, `size`, `dimensions`) and quantity keywords (`pcs`, `pack`, `counts`).
   - First 11 tags cover diverse SEO angles (colors, materials, uses, events, themes).
   - Last 2 tags are ALWAYS strictly:
     ```
     party decor, party supplies
     ```
   - Automatically populates missing tags in `Create.xlsx`, or can be run standalone from CLI / GUI.

5. **Safe Mode (Draft Review)**:
   - By default, listings are saved as **Drafts** so you can review them in Etsy Shop Manager before publishing and spending listing fees. Add `--publish` when you are ready to publish live.

---

## 📁 Project Structure

```
Selimum Etsy Listing Automation/
├── stores.json              # Store configuration (VPN account, assigned IP, Chrome profile)
├── Create.xlsx              # Master Excel listing input file
├── create_listing.py        # Core automation script (mirrors Reverb bot structure)
├── run.bat                  # 1-Click batch runner
├── generate_tags.bat        # 1-Click prompt to generate 13 SEO tags for any title
├── gui.py                   # PyQt6 Desktop GUI with Store Manager & Tag Generator
├── requirements.txt         # Dependencies
├── core/
│   ├── tag_engine.py        # 20-Rule Etsy SEO Tag Engine & Validator
│   ├── browser_factory.py   # Stealth Chrome & Browser manager
│   ├── store_manager.py     # Multi-store profile manager
│   ├── proxy_helper.py      # Proxy authentication helper & tester
│   ├── data_manager.py      # Excel & template manager
│   └── etsy_bot.py          # 10-step listing automation engine
├── output/
│   ├── failures/            # Failure screenshots
│   ├── photos/              # Downloaded listing photos
│   └── results_*.xlsx       # Execution status reports
└── data/
    └── sample_images/       # Placeholder images for test listings
```

---

## ⚙️ Store Configuration (`stores.json`)

`stores.json` is pre-configured to match your existing Reverb stores and profiles:

```json
{
  "chrome_user_data_dir": "C:\\Users\\Navroz\\AppData\\Local\\Google\\Chrome\\User Data",
  "chrome_path": "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
  "vpncmd_path": "C:\\Program Files\\SoftEther VPN Client\\vpncmd.exe",
  "debug_port": 9222,
  "stores": {
    "AMH": {"ip": "192.3.183.107", "chrome_profile": "Profile 15", "vpn": "AMH"},
    "TSS": {"ip": "204.188.208.166", "chrome_profile": "Profile 17", "vpn": "TSS"},
    "MMS": {"ip": "198.23.208.141", "chrome_profile": "Profile 18", "vpn": "MMS"},
    "MZM": {"ip": "", "chrome_profile": "Profile 19", "vpn": "MZM"},
    "GGL": {"ip": "", "chrome_profile": "Profile 20", "vpn": "GGL"},
    "GG": {"ip": "", "chrome_profile": "Profile 21", "vpn": "GG"},
    "IQ": {"ip": "204.44.93.167", "chrome_profile": "Profile 29", "vpn": "IQ"},
    "IQR": {"ip": "204.44.93.167", "chrome_profile": "Profile 29", "vpn": "IQ"},
    "SSL": {"ip": "107.161.123.246", "chrome_profile": "Profile 36", "vpn": "SSL"}
  }
}
```

---

## 📊 Excel Columns (`Create.xlsx`)

| Column Header | Description | Default / Example |
| :--- | :--- | :--- |
| `store_name` | Store identifier matching `stores.json` | `AMH`, `TSS`, `GG` |
| `title` | SEO Product Title (max 140 chars) | `Burgundy Velvet Ribbon 2.5 Inch Wired...` |
| `description` | Product Description | Full product description |
| `price` | Listing price in USD | `18.99` |
| `quantity` | Stock quantity | `100` |
| `sku` | Unique Product SKU | `RIB-BUR-001` |
| `photos` | Local paths or URLs (separated by `;` or `\|`) | `data/sample_images/ribbon1.jpg; data/...` |
| `category` | Etsy Category | `Ribbons`, `Balloons`, `Craft Supplies` |
| `tags` | 13 Etsy SEO Tags (*Leave empty for auto-generation*) | *Auto-generated with 20 rules if blank* |
| `shipping_profile` | Name of your Etsy Shipping Profile | `Standard Shipping` |
| `item_weight` | Package shipping weight (lbs) | `0.6` |
| `package_length` | Package length (inches) | `8` |
| `package_width` | Package width (inches) | `6` |
| `package_height` | Package height (inches) | `3` |
| `return_policy` | Return policy window | `30 days` |
| `who_made_it` | Etsy production question | `Another company` |
| `what_is_it` | Etsy item classification | `A supply or tool` |
| `when_made` | Production timeframe | `2020–2026` |
| `shop_section` | Optional section in your Etsy shop | `Holiday Ribbons` |
| `allow_offers` | Enable buyer offers | `no` |
| `featured` | Feature this listing | `yes` / `no` |
| `renewal` | Renewal options | `Automatic` |
| `publish` | `no` for Draft (recommended) or `yes` to publish | `no` |

---

## 🚀 How to Run

### 1. Generate 13 Etsy SEO Tags for Any Title
To instantly generate the 13 Etsy SEO tags complying strictly with all 20 rules:
```bash
python create_listing.py --tags "Your Product Title Here"
```
Or double click `generate_tags.bat`.

### 2. Run Listings for All Stores
```bash
python create_listing.py
```
Or double click `run.bat`.

### 3. Run a Specific Store
```bash
python create_listing.py --store AMH
```

### 4. Limit Number of Listings (Test Run)
```bash
python create_listing.py --store TSS --limit 1
```

### 5. Run Without VPN or Skip IP Check (Local Testing)
```bash
python create_listing.py --no-vpn --skip-ip-check
```

### 6. Publish Directly Live (Bypassing Draft Mode)
```bash
python create_listing.py --publish
```

### 7. Launch Graphical User Interface (GUI)
```bash
python gui.py
```
Provides visual store configuration, live proxy latency testing, Excel importing, progress bar, and instant SEO tag generator.
