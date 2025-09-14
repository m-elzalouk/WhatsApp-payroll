import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pyautogui # type: ignore
import pyperclip # type: ignore
import keyboard # type: ignore
import gc
import psutil
import re

# === CONFIGURATION ===
EXCEL_FILE = "tedx_txt_msg.xlsx"  # Use your generic Excel file
SHEET_DATA = "misurata"           # Sheet name with contacts and messages
WHATSAPP_WAIT = 25

# === Read Excel data
df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_DATA)
df["Send_Status"] = df.get("Send_Status", pd.Series([""] * len(df))).astype(str)

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

failed_contacts = []

for index, row in df.iterrows():
    name = str(row.get("phone", "")).strip().replace("+", "")
    message ="""
*📣 نداء باستلام تذكرتك من TEDx*
*السلام عليكم ورحمة الله*
نشكر لك اهتمامك بحضور الحدث. لقد تم استقبال بياناتكم بنجاح، وقبول حضوركم للحدث حسب المواعيد المحددة.✅
تذكرتك *مجانية تماماً* وتقدر تستلمها في أحد الأيام التالية:
📅 14 - 15 - 16 يوليو
🕐 من الساعة 5:30 م إلى 8:30 م
📍 في مقر مؤسسة الحوار والمناظرة – شارع الجوازات، مصراتة
🔗 رابط الموقع:https://maps.app.goo.gl/zR8iqAdJ1VxvsHMy8?g_st=com.google.maps.preview.copy
     """ #str(row.get("msg", "")).strip()
    send_status = str(row.get("Send_Status", "")).lower()
    # Skip if no name, no message, or already sent
    if not name or not message or send_status.startswith("sent"):
        continue

    try:
        df.iloc[index, df.columns.get_loc("Send_Status")] = ""
        print(f"📤 Sending to {name}")

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
        search_box.send_keys(name)
        time.sleep(1)
        search_box.send_keys(Keys.ENTER)
        time.sleep(1)

        # Wait for the message box to become available
        try:
            message_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
            )
        except Exception as e:
            print(f"❌ Could not open chat for {name}: {e}")
            df.at[index, "Send_Status"] = "Failed: Chat not opened"
            failed_contacts.append(name)
            continue

        # Send the message
        ActionChains(driver).move_to_element(message_box).click().perform()
        time.sleep(1)
        pyperclip.copy(message)
        message_box.send_keys(Keys.CONTROL, 'v')
        time.sleep(1)
        pyautogui.press('enter')
        df.at[index, "Send_Status"] = "Sent"
        time.sleep(2)
        print(f"✅ Sent to {name}")
        pyautogui.press('esc')

    except Exception as e:
        print(f"❌ Failed to send to {name}: {e}")
        failed_contacts.append(name)
        df.at[index, "Send_Status"] = f"Failed: {e}"
        continue

if failed_contacts:
    print("\n❌ The following contacts failed to receive their message:")
    for contact in failed_contacts:
        print(f" - {contact}")
else:
    print("\n✅ All messages sent successfully.")

# Save to a new file to avoid file lock issues
try:
    updated_file = "contacts_txt_msg_UPDATED.xlsx"
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
print("✅ All WhatsApp texts sent.")