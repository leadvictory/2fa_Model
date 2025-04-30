from seleniumbase import SB
from selenium.webdriver.common.keys import Keys
import pandas as pd

terms = ["24", "36", "39", "42"]
scraped_data = []

def safe_get_text(sb, selector):
    try:
        text = sb.get_text(selector).strip()
        return text if text else "0"
    except Exception:
        return "0"

with SB(uc=True) as sb:
    # Login
    sb.open("https://ssiweb.ssing.eu-central-1.aws.bmw.cloud/inquiry-by-vin/#/init")
    sb.sleep(6)
    sb.type("#idToken2", "john.ha.hayes")
    sb.click('input#callback_2_1')
    sb.click('input#idToken4_0')
    sb.type('#idToken1', '4352')
    sb.click('#idToken2_0')

    input("Press Enter to continue after MFA...")

    # Open dropdown
    sb.click(".btn.btn-init-info")
    sb.sleep(2)

    options = sb.find_elements(".ng-option .ng-option-label")

    for i in range(len(options)):
        if i > 0:
            sb.click(".btn.btn-init-info")
            sb.sleep(1)

        options = sb.find_elements(".ng-option .ng-option-label")
        label = options[i].text.strip()
        print(f"Selecting: {label}")
        options[i].click()
        sb.sleep(5)

        model_code, model_desc = label.split(" ", 1) if " " in label else (label, "")

        residuals = {}
        for term in terms:
            try:
                sb.click("div.fsDivBox:nth-of-type(1) .iwp-icon-gen_edit")
                sb.sleep(1)
                sb.type('input.termsText', term + Keys.ENTER)
                sb.sleep(3)
                residual = safe_get_text(sb, "div.fsDivBox:nth-of-type(1) .FSRateInfo div:nth-of-type(3) .pull-right")
                residuals[term] = residual
                print(f"Term {term} applied. Residual: {residual}")
            except Exception as e:
                print(f"Failed to set term {term}: {e}")
                residuals[term] = "0"

        lease_fs_rate = safe_get_text(sb, "div.fsDivBox:nth-of-type(1) .FSRateInfo div:nth-of-type(2) .pull-right")
        lease_credit = safe_get_text(sb, "div.fsDivBox:nth-of-type(1) .inclusiveProgram:nth-of-type(1) .boldFontbig")
        lease_loyalty = safe_get_text(sb, "div.fsDivBox:nth-of-type(1) .inclusiveProgram:nth-of-type(2) .boldFontbig")

        retail_fs_rate = safe_get_text(sb, "div.fsDivBox:nth-of-type(2) .FSRateInfo div:nth-of-type(2) .pull-right")
        retail_credit = safe_get_text(sb, "div.fsDivBox:nth-of-type(2) .inclusiveProgram:nth-of-type(1) .boldFontbig")
        retail_loyalty = safe_get_text(sb, "div.fsDivBox:nth-of-type(2) .inclusiveProgram:nth-of-type(2) .boldFontbig")
        print("Lease FS Rate:", lease_fs_rate)
        # print("Lease Residual %:", lease_residual)
        print("Lease Credit:", lease_credit)
        print("Lease Loyalty:", lease_loyalty)
        print("Retail FS Rate:", retail_fs_rate)
        print("Retail Credit:", retail_credit)
        print("Retail Loyalty:", retail_loyalty)
        scraped_data.append({
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
        })

        sb.refresh()
        sb.sleep(5)

    # Save results to Excel
    df = pd.DataFrame(scraped_data)
    df.to_excel("bmw_financial_data.xlsx", index=False)
    print("✅ Data saved to bmw_financial_data.xlsx")
