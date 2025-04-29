# KTM Jewellery Limited – EXE Automation Using Python

This project automates manual tasks in desktop applications (EXE files) for **KTM Jewellery Limited**, using Python libraries like `pywinauto` and `pyautogui`. The goal is to eliminate repetitive actions and improve efficiency for branch-level tasks like report generation, data entry, or navigation within business software.

---

## ✅ Features

- Automates opening and navigating EXE-based software.
- Selects specific options like branches, report types, and dates.
- Supports looping through multiple dropdown/combo box options.
- Exports data as PDF, CSV, or other formats.
- Mimics human actions: clicks, keyboard entries, and waits.
- Saves output files with dynamic filenames and timestamps.

---

## ⚙️ Technologies Used

- **Python 3.x**
- [`pywinauto`](https://pywinauto.readthedocs.io/) – for interacting with Windows GUI apps.
- [`pyautogui`](https://pyautogui.readthedocs.io/) – for mouse, keyboard automation.
- [`time`, `os`, `datetime`] – for delays, file handling, and naming.

---

## 🏢 Use Case at KTM

- Daily report exports from in-house software (e.g., *Branch Ornament Stock* report).
- Automates selection of branch codes like `1` and `9`.
- Automatically saves reports in both **PDF** and **CSV** formats.
- Frees staff from repetitive clicking and file saving.

---

## 🚀 How to Run

1. Install dependencies:

   ```bash
   pip install pywinauto pyautogui
   
## Developer: Joycilin
Email: maryjoycilin@gmail.com
Tools: Python + Pywinauto + PyAutoGUI
Client: KTM Jewellery Limited
