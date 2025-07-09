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
from openpyxl import load_workbook
from PIL import Image
import io
import numpy as np
import pyautogui # type: ignore
import keyboard # type: ignore
import gc
import psutil
import re

# === CONFIGURATION ===
EXCEL_FILE = "salary_template.xlsx"
EXPORT_DIR = "exports"
SHEET_DATA = "MAIN"
SHEET_TEMPLATE = "photo"
RANGE_TO_EXPORT = "B2:L47"
WHATSAPP_WAIT = 25

# === Ensure export folder exists
if not os.path.exists(EXPORT_DIR):
    os.makedirs(EXPORT_DIR)

# === Read Excel data
df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_DATA)

df["Image_Status"] = df["Image_Status"].astype(str)
df["Send_Status"] = df["Send_Status"].astype(str)

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True
wb = excel.Workbooks.Open(os.path.abspath(EXCEL_FILE))

try:
    wb = excel.Workbooks.Open(os.path.abspath(EXCEL_FILE))
    photo = wb.Sheets(SHEET_TEMPLATE)

    for index, row in df.iterrows():
        name = str(row["Contact_Name"]).strip()
        image_path = os.path.abspath(os.path.join(EXPORT_DIR, f"{name}.png"))
        salary_status = str(row.get("salary status", "")).strip()
        # Skip if contact name is empty, image was already generated successfully, or image file exists
        if (
            not name
            or name == 'nan'
            or re.search(r'[\u0600-\u06FF]', salary_status)
            or float(row.get("Net Salary", 0)) <= 0
            or row.get("Image_Status", "").lower() == "success"
            or (os.path.exists(image_path) and os.path.getsize(image_path) > 0)
        ):
            continue
        print(f"📸 Processing {index}...")
        try:

            def safe_cell(value):
                return "" if pd.isna(value) or value == 65535 else value
            # Fill values
            photo.Range("F10").Value = safe_cell(row["Military ID"])   #الرقم العسكري
            photo.Range("F12").Value = safe_cell(row["Rank"])          #الرتبة
            photo.Range("J12").Value = safe_cell(row["Number of Increments"])  #عدد العلاوات
            photo.Range("I12").Value = safe_cell(row["Degree"])       #الدرجة
            photo.Range("F16").Value = safe_cell(row["Basic Salary"])  #الراتب الأساسي
            photo.Range("F18").Value = safe_cell(row["Military Allowance"])    #بدل عسكري
            photo.Range("F22").Value = safe_cell(row["Clothing allowance"])  #بدل ملابس
            photo.Range("F20").Value = safe_cell(row["catering allowance"])    #بدل إطعام
            photo.Range("F24").Value = safe_cell(row["Car Allowance"])   #بدل سيارة
            photo.Range("F26").Value = safe_cell(row["Total Salary"])  #الراتب الإجمالي
            photo.Range("G29").Value = safe_cell(row["Social Security"])   #الضمان
            photo.Range("G33").Value = safe_cell(row["Solidarity"])    #التضامن
            photo.Range("G31").Value = safe_cell(row["Jihad"])         #الجهاد
            photo.Range("G35").Value = safe_cell(row["Internal advance"]) + safe_cell(row["Other iIternal discounts"])   #السلفة الداخلية
            photo.Range("G37").Value = safe_cell(row["Inner box"])         #صندوق داخلي
            photo.Range("G39").Value = safe_cell(row["Loan Fund"])   #صندوق السلف
            photo.Range("G41").Value = safe_cell(row["Total Deduction"])   #الخصومات الإجمالية
            photo.Range("G43").Value = safe_cell(row["Net Salary"])    #الراتب الصافي
            photo.Range("F14").Value = safe_cell(row["Military Name"]) #اسم العسكري

            # Export image
            wb.Activate()
            photo.Activate()
            photo.Range(RANGE_TO_EXPORT).Select()
            time.sleep(0.5)
            
            for attempt in range(2):
                try:
                    photo.Range(RANGE_TO_EXPORT).CopyPicture(Format=win32.constants.xlPicture)
                    chart = photo.ChartObjects().Add(Left=0, Top=0, Width=600, Height=800)
                    chart.Activate()
                    chart.Chart.Paste()
                    chart.Chart.Export(Filename=image_path)
                    chart.Delete()
                    df.iloc[index, df.columns.get_loc("Image_Status")] = ""
                    df.at[index, "Image_Status"] = "Success"
                    break                
                except Exception as e:
                    if attempt == 1:
                        df.at[index, "Image_Status"] = f"Failed: {e}"
        except Exception as e:
            print(f"⚠️ Error processing {name}: {e}")
            df.at[index, "Image_Status"] = f"Failed: {e}"
    print("✅ Exported all salary images to 'exports/' folder.")
except Exception as e:
    print(f"❌ Failed to open Excel file: {e}")
    df["Image_Status"] = "Failed to open Excel file"

 ####################################################################################################       
 ####################################################################################################       
 ####################################################################################################

 # === Ensure ChromeDriver is installed and in PATH
 
# === Open WhatsApp Web

options = Options()
options.add_argument(r"user-data-dir=C:\ChromeProfiles\whatsapp_profile")

driver = webdriver.Chrome(options=options)
time.sleep(1)
driver.get("https://web.whatsapp.com")
pyautogui.press('enter')
print("📱 Please scan the QR code in WhatsApp Web...")
time.sleep(WHATSAPP_WAIT)
wait = WebDriverWait(driver, 60)


# def is_image_blank(image_path, threshold=3):
#     try:
#         img = Image.open(image_path).convert("L")
#         pixels = np.array(img)
#         std = pixels.std()
#         return std < threshold  # low std = blank or near blank
#     except Exception as e:
#         print(f"⚠️ Error checking if image is blank: {e}")
#         return False

# === Check for Blank Image ===
def is_image_blank(image_path, threshold=3):
    try:
        img = Image.open(image_path).convert("L")
        return np.array(img).std() < threshold
    except:
        print(f"⚠️ Error checking if image is blank: {e}")
        return True


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

failed_contacts = []  # To collect any failures

for index, row in df.iterrows():

            # Check if 'q' is pressed to quit
    if keyboard.is_pressed('q'):
        print("\n⛔️ Interrupted by user. Saving progress and exiting...")
        break

    name = str(row["Contact_Name"]).strip().replace("+", "")
    image_path = os.path.abspath(os.path.join(EXPORT_DIR, f"{name}.png"))
    salary_status = str(row.get("salary status", "")).strip()
    # df["Send_Status"] = df.get("Send_Status", pd.Series([""] * len(df)))

# Fill only cells that are NaN, empty string, or not exactly 'Sent'
    # df["Send_Status"] = df["Send_Status"].apply(lambda x: x if str(x).strip().lower() == "sent" else "")

    # Skip if contact name is empty or already sent successfully
    # if not name or name == '0' or row.get("Send_Status", "").lower().startswith("sent"):
    #     continue

    if (
        not name
        or name == '0'
        or name == 'nan'
        or re.search(r'[\u0600-\u06FF]', salary_status)
        or float(row.get("Net Salary", 0)) <= 0
        or row.get("Send_Status", "").lower().startswith("sent")
    ):
        continue

    if not os.path.exists(image_path):
        print(f"⚠️ Image not found for {name}, skipping.")
        df.at[index, "Send_Status"] = "Image Not Found"
        continue
    if is_image_blank(image_path):
        print(f"⚠️ Image for {name} appears blank, skipping.")
        df.at[index, "Send_Status"] = "Blank Image"
        continue

    try:
        df.iloc[index, df.columns.get_loc("Send_Status")] = ""
        print(f"📤 Sending to {name}")
        # search_box = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="8"]')))
        # time.sleep(1)
        # search_box.clear()
        # time.sleep(1)
        # search_box.send_keys(name)
        # search_box.send_keys(Keys.ENTER)
        # time.sleep(2)

        try:
            new_chat_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="new-chat-outline"]')))
            driver.execute_script("arguments[0].click();", new_chat_btn)
            time.sleep(1)
        except Exception as e:
            print(f"❌ Failed to click New Chat button: {e}")
            df.at[index, "Send_Status"] = "Failed: New Chat click error"
            failed_contacts.append(name)
            continue
        

        # Focus on search box in the new chat window
        search_box = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')))
        search_box.clear()
        search_box.send_keys(name)  # name is the phone number
        time.sleep(1)
        search_box.send_keys(Keys.ENTER)
        time.sleep(1)
        wait = WebDriverWait(driver, 2)

        # Add this check to ensure the chat actually opened
        try:
            # Wait for the message box to become available (max 10 seconds)
            message_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
            )
        except Exception as e:
            print(f"❌ Could not open chat for {name} (numeric name issue?): {e}")
            df.at[index, "Send_Status"] = "Failed: Chat not opened"
            failed_contacts.append(name)
            continue

        # Paste image
        copy_image_to_clipboard(image_path)
        time.sleep(1)
        ActionChains(driver).move_to_element(message_box).click().perform()
        time.sleep(1)
        message_box.send_keys(Keys.CONTROL, 'v')
        time.sleep(2)
        
        # Press enter to send
        pyautogui.press('enter')
        df.at[index, "Send_Status"] = "Sent"
        time.sleep(2)
        print(f"✅ Sent to {name}")
        pyautogui.press('esc')

    except Exception as e:
        print(f"❌ Failed to send to {name}: {e}")
        failed_contacts.append(name)
        df.at[index, "Send_Status"] = f"Failed: {e}"
        continue  # Ensure it moves to the next contact

if failed_contacts:
    print("\n❌ The following contacts failed to receive their image:")
    for contact in failed_contacts:
        print(f" - {contact}")
else:
    print("\n✅ All messages sent successfully.")

# === Save results to Excel ===

# First, close the Excel COM objects to release the file
try:
    wait = WebDriverWait(driver, 10)
    wb.Close(False)
    excel.Quit()
    del wb
    del excel
    gc.collect()
    time.sleep(2)  # Give the OS a moment to release the file lock
except Exception as e:
    print(f"⚠️ Error closing Excel COM objects: {e}")

# Save to a new file to avoid file lock issues
try:
    updated_file = "salary_template_UPDATED.xlsx"
    with pd.ExcelWriter(
        updated_file,
        engine='openpyxl',
        mode='w'
    ) as writer:
        df.to_excel(writer, sheet_name=SHEET_DATA, index=False)
    print(f"✅ All messages sent and results saved to {updated_file}.")   
except Exception as e:
    print(f"❌ Failed to save Excel file: {e}")

time.sleep(3)
driver.quit()
print("✅ All slips sent.")