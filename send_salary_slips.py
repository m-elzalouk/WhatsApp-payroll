import os
import time
import pandas as pd
import win32com.client as win32
from PyQt5.QtWidgets import QApplication
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# === CONFIGURATION ===
EXCEL_FILE = "salary_template.xlsx"
SHEET_DATA = "data"
SHEET_TEMPLATE = "photo"
RANGE_TO_COPY = "B2:O35"
WHATSAPP_WAIT = 45

# === PyQt5 app setup
app = QApplication.instance()
if not app:
    app = QApplication([])

# === Read Excel data
df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_DATA)

# === Launch Excel
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True  # Required for visual copying
wb = excel.Workbooks.Open(os.path.abspath(EXCEL_FILE))
photo = wb.Sheets(SHEET_TEMPLATE)

# === Launch WhatsApp Web
driver = webdriver.Chrome()
driver.get("https://web.whatsapp.com")
print("📱 Scan QR code in WhatsApp Web...")
time.sleep(WHATSAPP_WAIT)
wait = WebDriverWait(driver, 60)

# === For each employee
for index, row in df.iterrows():
    name = str(row['Contact_Name']).strip()

    try:
        # Fill values in template
        photo.Range("F10").Value = row["Military ID"]
        photo.Range("F12").Value = row["Rank"]
        photo.Range("F14").Value = row["Number of Increments"]
        photo.Range("K14").Value = row["Basic Salary"]
        photo.Range("F16").Value = row["Military Allowance"]
        photo.Range("K16").Value = row["Supply Allowance"]
        photo.Range("F18").Value = row["catering allowance"]
        photo.Range("K18").Value = row["Car Allowance"]
        photo.Range("F20").Value = row["Total Salary"]
        photo.Range("F24").Value = row["Social Security"]
        photo.Range("K24").Value = row["Solidarity"]
        photo.Range("F26").Value = row["Jihad"]
        photo.Range("K26").Value = row["Loan Fund"]
        photo.Range("F28").Value = row["Total Deduction"]
        photo.Range("H31").Value = row["Net Salary"]
        photo.Range("K10").Value = row["Military Name"]

        # Copy visual range as image to clipboard
        photo.Activate()
        photo.Range(RANGE_TO_COPY).CopyPicture(Format=win32.constants.xlPicture)
        time.sleep(2)

        # Search contact in WhatsApp
        print(f"📨 Sending to: {name}")
        search_box = wait.until(EC.presence_of_element_located(
            (By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')))
        search_box.clear()
        search_box.send_keys(name)
        time.sleep(2)
        search_box.send_keys(Keys.ENTER)
        time.sleep(3)

        # Focus message input
        message_box = wait.until(EC.presence_of_element_located(
            (By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')))
        ActionChains(driver).move_to_element(message_box).click().perform()
        time.sleep(1)

        # Paste from clipboard (image copied from Excel)
        message_box.send_keys(Keys.CONTROL, 'v')
        time.sleep(4)

        # Click send
        send_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//div[@role="button" and @aria-label="Send"]')))
        driver.execute_script("arguments[0].click();", send_btn)

        wait.until(EC.presence_of_element_located(
            (By.XPATH, '//span[contains(@data-testid, "msg-time")]')))
        print(f"✅ Sent to {name}")
        time.sleep(3)

    except Exception as e:
        print(f"❌ Failed to send to {name}: {e}")

# === Cleanup
print("🕓 Final delay to ensure all messages are delivered...")
time.sleep(10)
wb.Close(False)
excel.Quit()
driver.quit()
print("✅ All slips sent.")
