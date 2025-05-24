```markdown
# 💬 WhatsApp Salary Slip Sender

This project automates the generation and sending of salary slips to employees via **WhatsApp Web**, using data from an Excel file and styled salary templates.

---

## 📁 Project Structure

```
FinanceProject/
│
├── .venv/                      # Virtual environment (optional but recommended)
├── send_salary_slips.py       # Main automation script
├── salary_template.xlsx       # Excel file with salary data and the 'photo' template sheet
├── exports/                   # Auto-generated folder for exported salary slip images
├── chromedriver.exe           # Required for controlling Chrome
├── requirements.txt           # List of dependencies
```

---

## 🧰 Features

- Reads employee data from `data` sheet in Excel.
- Fills the `photo` sheet to generate a visual salary slip.
- Exports each slip as a PNG file.
- Detects blank images and skips sending them.
- Sends each image to the corresponding WhatsApp contact using clipboard paste.
- Logs any failed sends and continues gracefully.

---

## 🧪 Requirements

- **Python 3.10+**
- Google Chrome installed
- `chromedriver.exe` that matches your Chrome version (place it in project root)

---

## 📦 Setup Instructions

### 1. Clone the project and install dependencies

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

> If `requirements.txt` is missing, run:

```bash
pip install pandas selenium pywin32 pillow openpyxl
```

---

### 2. Set up WhatsApp Profile (once)

To avoid scanning the QR code every time:

1. Run:
```bash
chrome.exe --user-data-dir="C:\ChromeProfiles\whatsapp_profile"
```

2. Log into WhatsApp Web and close the window.

---

### 3. Run the Script

```bash
.venv\Scripts\activate
python send_salary_slips.py
```

> This will:
> - Fill the Excel template
> - Export PNGs to `exports/`
> - Open WhatsApp Web and send each image to the contact listed

---

## 🧠 Tips

- Do **not** open Excel while the script runs
- Do **not** interact with WhatsApp Web until the script finishes
- Images are sent using `Ctrl+V` via clipboard for best reliability

---

## 🛠 Troubleshooting

- ❌ `DevToolsActivePort` error: You’re using a default Chrome profile — follow the WhatsApp Profile Setup above.
- ❌ COM/Excel error: Excel must be visible (`excel.Visible = True`) and the template range must be properly selected.
- ❌ Blank image sent: The script automatically skips images with no visible data.

---

## 🔐 Security Notice

This script automates WhatsApp for internal use only. Make sure it complies with your organization's data handling and messaging policies.

---

## 📜 License

MIT License
```
