import os
from seleniumbase import SB
from selenium.webdriver.common.keys import Keys
import pandas as pd

terms = ["24", "36", "39", "42"]
excel_file = "bmw_financial_data.xlsx"
scraped_data = []

# Load previously scraped model code + description pairs
if os.path.exists(excel_file):
    existing_df = pd.read_excel(excel_file)
    processed_entries = set(zip(existing_df["Model Code"], existing_df["Model Description"]))
    scraped_data = existing_df.to_dict("records")
else:
    processed_entries = set()

def safe_get_text(sb, selector):
    try:
        text = sb.get_text(selector).strip()
        return text if text else "0"
    except Exception:
        return "0"

with SB(uc=True) as sb:
    # Login
    sb.open("https://ssiweb.ssing.eu-central-1.aws.bmw.cloud/inquiry-by-vin/#/init")
    sb.sleep(3)
    sb.type("#idToken2", "john.ha.hayes")
    sb.click('input#callback_2_1')
    sb.click('input#idToken4_0')
    sb.type('#idToken1', '4352')
    sb.click('#idToken2_0')

    input("Press Enter to continue after MFA...")

    sb.click(".btn.btn-init-info")
    sb.sleep(2)
    options = sb.find_elements(".ng-option .ng-option-label")
    total_options = len(options)
    print(f"✅ {total_options} model options found.")
    batch_size = 30
    total_batches = (total_options + batch_size - 1) // batch_size

    # Ask for batch number after knowing how many are available
    print(f"ℹ️ Available batches: 1 to {total_batches} (each has {batch_size} models max)")
    batch = int(input(f"Enter batch number (1–{total_batches}): "))

    # Determine start and end based on batch
    start_index = (batch - 1) * batch_size
    end_index = min(start_index + batch_size, total_options)

    print(f"🔢 Scraping batch {batch}: index {start_index}–{end_index - 1}")

    for i in range(start_index, end_index):
        if i > start_index:
            sb.click(".btn.btn-init-info")
            sb.sleep(2)
        options = sb.find_elements(".ng-option .ng-option-label")
        if i >= len(options):
            print(f"⚠️ Index {i} is out of range. Skipping.")
            break

        label = options[i].text.strip()
        model_code, model_desc = label.split(" ", 1) if " " in label else (label, "")
        entry_key = (model_code, model_desc)

        if entry_key in processed_entries:
            print(f"⏭️ Skipping already processed model: {entry_key}")
            sb.refresh()
            sb.sleep(7)
            continue

        print(f"🔄 Processing: {entry_key}")
        options[i].click()
        sb.sleep(5)

        residuals = {}
        for term in terms:
            try:
                sb.click("div.fsDivBox:nth-of-type(1) .iwp-icon-gen_edit")
                sb.sleep(1)
                sb.type('input.termsText', term + Keys.ENTER)
                sb.sleep(7)
                residual = safe_get_text(sb, "div.fsDivBox:nth-of-type(1) .FSRateInfo div:nth-of-type(3) .pull-right")
                residuals[term] = residual
            except Exception as e:
                print(f"⚠️ Failed to set term {term}: {e}")
                residuals[term] = "0"

        lease_fs_rate = safe_get_text(sb, "div.fsDivBox:nth-of-type(1) .FSRateInfo div:nth-of-type(2) .pull-right")
        lease_credit = safe_get_text(sb, "div.fsDivBox:nth-of-type(1) .inclusiveProgram:nth-of-type(1) .boldFontbig")
        lease_loyalty = safe_get_text(sb, "div.fsDivBox:nth-of-type(1) .inclusiveProgram:nth-of-type(2) .boldFontbig")

        retail_fs_rate = safe_get_text(sb, "div.fsDivBox:nth-of-type(2) .FSRateInfo div:nth-of-type(2) .pull-right")
        retail_credit = safe_get_text(sb, "div.fsDivBox:nth-of-type(2) .inclusiveProgram:nth-of-type(1) .boldFontbig")
        retail_loyalty = safe_get_text(sb, "div.fsDivBox:nth-of-type(2) .inclusiveProgram:nth-of-type(2) .boldFontbig")

        row = {
            "Model Code": model_code,
            "Model Description": model_desc,
            "Lease Rate": lease_fs_rate,
            "Finance Rate": retail_fs_rate,
            "24 mo Residual": residuals.get("24", "0"),
            "36 mo Residual": residuals.get("36", "0"),
            "39 mo Residual": residuals.get("39", "0"),
            "42 mo Residual": residuals.get("42", "0"),
            "Lease Credit": lease_credit,
            "Finance Credit": retail_credit,
            "Lease Loyalty": lease_loyalty,
            "Finance Loyalty": retail_loyalty
        }

        scraped_data.append(row)
        processed_entries.add(entry_key)
        pd.DataFrame(scraped_data).to_excel(excel_file, index=False)
        print(f"✅ Saved model: {entry_key}")

        sb.refresh()
        sb.sleep(7)

print("🎉 Selected batch processed and saved to bmw_financial_data.xlsx")
