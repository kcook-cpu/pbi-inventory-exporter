import os
import sys
import time
import requests
import pandas as pd
from playwright.sync_api import sync_playwright

# Configurations
REPORT_URL = "https://app.powerbi.com/groups/me/reports/b569d766-6d6a-4b19-88f5-694caf7add19/d95d619fc9aea1637cba?ctid=87b66f08-478c-40ad-bb09-5e93796da295&experience=power-bi&bookmarkGuid=d1eeaa6b-8551-43cb-876c-06f17e654444"
WEB_APP_URL = os.environ.get("WEB_APP_URL") # Received Laptops Web App Endpoint

CURRENT_FILE = "/tmp/Inventory_Current.xlsx"
STATE_FILE = "state.json"

def export_and_upload():
    pw = sync_playwright().start()

    try:
        print("[1/3] Launching Playwright browser...")
        browser = pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-setuid-sandbox"]
        )

        context_options = {
            "accept_downloads": True,
            "permissions": ["clipboard-read", "clipboard-write"]
        }

        if os.path.exists(STATE_FILE):
            context_options["storage_state"] = STATE_FILE

        context = browser.new_context(**context_options)
        page = context.new_page()

        print("[1/3] Navigating to Power BI Report...")
        page.goto(REPORT_URL)

        try:
            page.wait_for_selector('text="Sublocation"', timeout=20000)
        except Exception:
            page.wait_for_selector('text="Sublocation"', timeout=120000)

        # Focus & click export options
        sublocation_header = page.locator('text="Sublocation"').first
        sublocation_header.focus()
        sublocation_header.hover()
        time.sleep(1)

        more_btn = page.locator('button[title="More options"], button[aria-label="More options"]').last
        more_btn.wait_for(state="visible", timeout=5000)
        more_btn.click(force=True)
        time.sleep(1)

        export_item = page.locator('.pbi-menu-item-text-container, [role="menuitem"]').filter(has_text="Export data").first
        export_item.wait_for(state="visible", timeout=5000)
        export_item.click(force=True)
        time.sleep(1)

        layout_radio = page.locator('label:has-text("Data with current layout"), span:has-text("Data with current layout")').first
        try:
            layout_radio.wait_for(state="visible", timeout=5000)
            layout_radio.click(force=True)
        except Exception:
            page.locator('input[aria-label="Data with current layout"]').first.dispatch_event("click")

        time.sleep(1)

        with page.expect_download(timeout=120000) as download_info:
            confirm_btn = page.locator('button.btn-primary, button:has-text("Export")').last
            confirm_btn.wait_for(state="visible", timeout=10000)
            confirm_btn.click(force=True)

        download = download_info.value
        download.save_as(CURRENT_FILE)
        print(f" -> Export downloaded locally: {CURRENT_FILE}")

        context.close()
        browser.close()

    finally:
        pw.stop()

    # Step 2: Read Excel and format date/datetime objects to strings
    print("[2/3] Reading Excel file and preparing JSON payload...")
    df = pd.read_excel(CURRENT_FILE).fillna("")
    
    # Convert any datetime/date columns to string format so JSON can serialize them
    for col in df.select_dtypes(include=['datetime', 'datetime64']).columns:
        df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')

    rows_data = df.values.tolist()

    # Step 3: POST data to Web App
    response = requests.post(WEB_APP_URL, json=rows_data)
    print(f" -> Web App Response: {response.text}")

if __name__ == "__main__":
    export_and_upload()
