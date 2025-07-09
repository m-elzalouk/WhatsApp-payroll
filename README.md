* `Excel COM Automation (win32com)`
* `Chrome + Selenium` with `chromedriver`
* `clipboard image sending` via `pywin32`
* A `.venv` environment

---

## ✅ Final Verified Steps to Run the WhatsApp Salary Slip Project

### ✅ 1. 📁 Folder Structure

```
FinanceProject/
├── send_salary_slips.py          # Your main Python script
├── salary_template.xlsx          # Excel with 'data' and 'photo' sheets
├── chromedriver.exe              # Must match your Chrome version
├── requirements.txt              # Your Python dependencies
├── exports/                      # Auto-created for PNG slips
└── .venv/                        # Virtual environment (optional but ideal)
```

---

### ✅ 2. 🔧 Initial Setup (First Time on New PC)

#### A. Install Python (3.10+ recommended)

Download from: [https://www.python.org/downloads/](https://www.python.org/downloads/)

> ⚠️ Ensure you check "Add Python to PATH" during install

#### B. Create and activate virtual environment

```bash
cd FinanceProject
python -m venv .venv
.venv\Scripts\activate
```

#### C. Install dependencies

If you have `requirements.txt`:

```bash
pip install -r requirements.txt
```

If not, run:

```bash
pip install pandas selenium pywin32 pillow openpyxl
```

---

### ✅ 3. ⚙️ Setup WhatsApp Chrome Profile (Only Once)

This prevents logging in again every time.
Run the following at RUN Window (Windows + R).

```bash
chrome.exe --user-data-dir="C:\ChromeProfiles\whatsapp_profile"
```

1. This opens a new Chrome window
2. Go to [https://web.whatsapp.com](https://web.whatsapp.com)
3. Scan QR code once, close the window

---

### ✅ 4. 🧾 Prepare `salary_template.xlsx`

* Sheet `data` with columns like `Contact_Name`, `Basic Salary`, etc.
* Sheet `photo` styled like your salary slip, with named cell positions (e.g., `F10`, `K14`)

---

### ✅ 5. 🚀 Run the Script

From the project folder:

```bash
.venv\Scripts\activate
python send_salary_slips.py
```

This will:

* Fill Excel
* Export each slip to `exports/`
* Open Chrome with WhatsApp
* Paste each image via clipboard to the matching contact

---

### ✅ 6. 🔁 Re-run Later

No need to scan WhatsApp QR again.

Just:

```bash
cd FinanceProject
.venv\Scripts\activate
python send_salary_slips.py
```

---

## 🧠 Optional Enhancements

| Feature           | Description                         |
| ----------------- | ----------------------------------- |
| `--noconsole` exe | Hide black window (via PyInstaller) |
| `.bat` launcher   | Auto-run with 2x click              |
| `.csv` log        | Track sent/skipped contacts         |

---