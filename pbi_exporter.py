import os
import sys
import time
import requests
import pandas as pd
from playwright.sync_api import sync_playwright

REPORT_URL = "https://app.powerbi.com/groups/me/reports/b569d766-6d6a-4b19-88f5-694caf7add19/d95d619fc9aea1637cba?ctid=87b66f08-478c-40ad-bb09-5e93796da295&experience=power-bi&bookmarkGuid=d1eeaa6b-8551-43cb-876c-06f17e654444"
WEB_APP_URL = os.environ.get("WEB_APP_URL")
PBI_USER = os.environ.get("PBI_USER")
PBI_PASS = os.environ.get("PBI_PASS")

CURRENT_FILE = "/tmp/Inventory_Current.xlsx"

def export_and_upload():
    pw = sync_playwright().start()

    try:
        print("[1/3] Launching Playwright browser...")
        browser = pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-setuid-sandbox"]
        )

        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        print("[1/3] Navigating to Power BI...")
        page.goto(REPORT_URL)

        # Handle automated login if login prompt appears
        try:
            if page.locator('input[type="email"]').is_visible(timeout=5000):
                print(" -> Logging into Microsoft account...")
                page.fill('input[type="email"]', PBI_USER)
                page.click('input[type="submit"]')
                time.sleep(2)
                
                page.fill('input[type="password"]', PBI_PASS)
                page.click('input[type="submit"]')
                time.sleep(2)
                
                # Click "Stay signed in" if prompted
                if page.locator('input[id="idSIButton9"]').is_visible(timeout=3000):
                    page.click('input[id="idSIButton9"]')
        except Exception as e:
            print(f" -> Login step skipped/passed: {e}")

        # Wait for table
        try:
            page.wait_for_selector('text="Sublocation"', timeout=30000)
        except Exception:
            page.wait_for_selector('text="Sublocation"', timeout=120000)

        # Export table
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
        print(f" -> Export downloaded: {CURRENT_FILE}")

        context.close()
        browser.close()

    finally:
        pw.stop()

    print("[2/3] Reading Excel file and sending payload to Google Drive / Sheets...")
    df = pd.read_excel(CURRENT_FILE).fillna("")
    rows_data = df.values.tolist()

    response = requests.post(WEB_APP_URL, json=rows_data)
    print(f" -> Web App Response: {response.text}")

if __name__ == "__main__":
    export_and_upload()
