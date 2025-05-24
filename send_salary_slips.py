import os
import time
import pandas as pd
import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import win32clipboard
from PIL import Image
import io
import numpy as np

# === CONFIGURATION ===
EXCEL_FILE = "salary_template.xlsx"
EXPORT_DIR = "exports"
SHEET_DATA = "data"
SHEET_TEMPLATE = "photo"
RANGE_TO_EXPORT = "B2:O35"
WHATSAPP_WAIT = 25

# === Ensure export folder exists
if not os.path.exists(EXPORT_DIR):
    os.makedirs(EXPORT_DIR)

# === Read Excel data
df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_DATA)

# === Start Excel COM
# excel = win32.gencache.EnsureDispatch('Excel.Application')
# excel.Visible = False
# wb = excel.Workbooks.Open(os.path.abspath(EXCEL_FILE))
# photo = wb.Sheets(SHEET_TEMPLATE)

# # === Generate images from photo template
# for index, row in df.iterrows():
#     name = str(row['Contact_Name']).strip()

#     # Fill template fields
#     photo.Range("F10").Value = row["Military ID"]
#     photo.Range("F12").Value = row["Rank"]
#     photo.Range("F14").Value = row["Number of Increments"]
#     photo.Range("K14").Value = row["Basic Salary"]
#     photo.Range("F16").Value = row["Military Allowance"]
#     photo.Range("K16").Value = row["Supply Allowance"]
#     photo.Range("F18").Value = row["catering allowance"]
#     photo.Range("K18").Value = row["Car Allowance"]
#     photo.Range("F20").Value = row["Total Salary"]
#     photo.Range("F24").Value = row["Social Security"]
#     photo.Range("K24").Value = row["Solidarity"]
#     photo.Range("F26").Value = row["Jihad"]
#     photo.Range("K26").Value = row["Loan Fund"]
#     photo.Range("F28").Value = row["Total Deduction"]
#     photo.Range("H31").Value = row["Net Salary"]
#     photo.Range("K10").Value = row["Military Name"]

#     # Export image
#     photo.Range(RANGE_TO_EXPORT).CopyPicture(Format=win32.constants.xlPicture)
#     chart = photo.ChartObjects().Add(Left=0, Top=0, Width=600, Height=500)
#     chart.Chart.Paste()
#     image_path = os.path.abspath(os.path.join(EXPORT_DIR, f"{name}.png"))
#     chart.Chart.Export(Filename=image_path)
#     chart.Delete()

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True

try:
    wb = excel.Workbooks.Open(os.path.abspath(EXCEL_FILE))
    photo = wb.Sheets(SHEET_TEMPLATE)

    for index, row in df.iterrows():
        name = str(row["Contact_Name"]).strip()

        # Fill values
        photo.Range("F10").Value = row["Military ID"]
        photo.Range("F12").Value = row["Rank"]
        photo.Range("F14").Value = row["Number of Increments"]
        photo.Range("K12").Value = row["Degree"]
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

        # Export image
        photo.Activate()
        photo.Range("B2:O35").Select()
        time.sleep(0.5)
        photo.Range("B2:O35").CopyPicture(Format=win32.constants.xlPicture)

        chart = photo.ChartObjects().Add(Left=0, Top=0, Width=800, Height=600)
        chart.Activate()
        chart.Chart.Paste()
        image_path = os.path.abspath(os.path.join(EXPORT_DIR, f"{name}.png"))
        chart.Chart.Export(Filename=image_path)
        chart.Delete()

    print("✅ Exported all salary images to 'exports/' folder.")

finally:
    try:
        if 'wb' in locals() and wb:
            wb.Close(False)
    except Exception as e:
        print(f"⚠️ Workbook close error: {e}")
    try:
        if 'excel' in locals() and excel:
            excel.Quit()
    except Exception as e:
        print(f"⚠️ Excel quit error: {e}")

        
# === Open WhatsApp Web
options = Options()
options.add_argument(r"user-data-dir=C:\ChromeProfiles\whatsapp_profile")

driver = webdriver.Chrome(options=options)
time.sleep(2)
driver.get("https://web.whatsapp.com")

print("📱 Please scan the QR code in WhatsApp Web...")
time.sleep(WHATSAPP_WAIT)
wait = WebDriverWait(driver, 60)


def is_image_blank(image_path, threshold=3):
    try:
        img = Image.open(image_path).convert("L")
        pixels = np.array(img)
        std = pixels.std()
        return std < threshold  # low std = blank or near blank
    except Exception as e:
        print(f"⚠️ Error checking if image is blank: {e}")
        return False


def copy_image_to_clipboard(image_path):
    image = Image.open(image_path).convert("RGB")
    output = io.BytesIO()
    image.save(output, format='BMP')
    data = output.getvalue()[14:]  # Strip BMP header
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

# === Send images using WhatsApp Web
# for index, row in df.iterrows():
#     name = str(row["Contact_Name"]).strip()
#     image_path = os.path.abspath(os.path.join(EXPORT_DIR, f"{name}.png"))

#     try:
#         print(f"📤 Sending to {name}")
#         search_box = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')))
#         search_box.clear()
#         time.sleep(2)
#         search_box.send_keys(name)
#         time.sleep(2)
#         search_box.send_keys(Keys.ENTER)
#         time.sleep(3)

#         # Copy image to clipboard
#         copy_image_to_clipboard(image_path)
#         time.sleep(1)

#         # Focus on the message box
#         message_box = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')))
#         ActionChains(driver).move_to_element(message_box).click().perform()
#         time.sleep(1)

#         # Paste from clipboard (Ctrl+V)
#         message_box.send_keys(Keys.CONTROL, 'v')
#         time.sleep(4)

#         search_box.clear()
#         time.sleep(1)
#         # send_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]')))
#         # send_btn.click()
#         # time.sleep(2)

#         # Click send
#         send_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@role="button" and @aria-label="Send"]')))
#         driver.execute_script("arguments[0].click();", send_btn)
#         time.sleep(2)

#         print(f"✅ Sent to {name}")

#     except Exception as e:
#         print(f"❌ Failed to send to {name}: {e}")

failed_contacts = []  # To collect any failures

for index, row in df.iterrows():
    name = str(row["Contact_Name"]).strip()
    image_path = os.path.abspath(os.path.join(EXPORT_DIR, f"{name}.png"))

    if not os.path.exists(image_path):
        print(f"⚠️ Image not found for {name}, skipping.")
        continue

    if is_image_blank(image_path):
        print(f"⚠️ Image for {name} appears blank, skipping.")
        continue

    try:
        print(f"📤 Sending to {name}")
        search_box = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')))
        time.sleep(0.5)
        search_box.clear()
        time.sleep(0.5)
        search_box.send_keys(name)
        time.sleep(1)
        search_box.send_keys(Keys.ENTER)
        time.sleep(2)

        copy_image_to_clipboard(image_path)
        time.sleep(1)

        message_box = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')))
        ActionChains(driver).move_to_element(message_box).click().perform()
        time.sleep(1)

        message_box.send_keys(Keys.CONTROL, 'v')
        time.sleep(2)

        send_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@role="button" and @aria-label="Send"]')))
        driver.execute_script("arguments[0].click();", send_btn)
        time.sleep(3)

        print(f"✅ Sent to {name}")

    except Exception as e:
        print(f"❌ Failed to send to {name}: {e}")
        failed_contacts.append(name)
        continue  # Ensure it moves to the next contact

if failed_contacts:
    print("\n❌ The following contacts failed to receive their image:")
    for contact in failed_contacts:
        print(f" - {contact}")
else:
    print("\n✅ All messages sent successfully.")

time.sleep(5)
excel.Quit()
driver.quit()
print("✅ All slips sent.")