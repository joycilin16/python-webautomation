from pywinauto import Application, Desktop
import time
import os
import pyautogui
from datetime import datetime
from pywinauto.keyboard import send_keys




# Launch the app using UIA backend (not win32)
app = Application(backend="uia").start(r"c:\DAIVELKTM\TMProj.exe")  # replace with your actual EXE path

# Wait and connect to the main window 
main_win = app.window(title_re=".*DAIVEL.*")  # Use regex if title may vary slightly
main_win.wait('visible', timeout=10)

# main_win.print_control_identifiers()
main_win.child_window(title="User ID", auto_id="TxtUserId", control_type="Edit").type_keys("mis")
main_win.child_window(title="Password:", auto_id="TxtPassword", control_type="Edit").type_keys("8963")

main_win.child_window(title="Login", auto_id="BtnOk", control_type="Button").click()

time.sleep(2)

# Now you can interact with the menu window
daivel_win= app.window(title_re=".*DAIVEL.*")  # Adjust title as needed
daivel_win.wait('visible', timeout=10)

# daivel_win.print_control_identifiers()

daivel_win.child_window(title="Purchase", control_type="Button").click()

rpt_win=app.window(title_re=".*DAIVEL.*") 
# rpt_win.print_control_identifiers()
purchase_menu=rpt_win.child_window(title="Transaction", control_type="TreeItem").wrapper_object()

# Perform a double click
purchase_menu.double_click_input()

report_menu=rpt_win.child_window(title="Report", control_type="TreeItem").wrapper_object()
report_menu.double_click_input()


branch_item = rpt_win.child_window(title="Branch Ornament Stock", control_type="TreeItem")
branch_item_wrapper = branch_item.wrapper_object()
branch_item_wrapper.click_input()

# rpt_win.print_control_identifiers()

rpt_type=rpt_win.child_window(title="Open", control_type="SplitButton")


# Locate the dropdown (ComboBox)
main_win = app.window(title_re=".*Branch Ornament Stock.*")

# Expand the Report Type dropdown
combo = main_win.child_window(auto_id="CboType", control_type="ComboBox").wrapper_object()
combo.click_input()
time.sleep(1.5)  # Allow the dropdown list to render

from pywinauto.keyboard import send_keys
send_keys("{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{ENTER}") 

branch_button = main_win.child_window(auto_id="TxtWrkunit", control_type="Edit").child_window(control_type="Button")
branch_button.click_input()

# rpt_win.print_control_identifiers()

select_all=rpt_win.child_window(title="Select All", auto_id="btnSelectAll", control_type="Button")
select_all.click_input()

ok_btn=rpt_win.child_window(title="Ok", auto_id="BtnOk", control_type="Button")
ok_btn.click_input()
time.sleep(2)

run_btn=rpt_win.child_window(title="Run", auto_id="CmdGenerate", control_type="Button")
run_btn.click_input()
time.sleep(5)

export_btn=rpt_win.child_window(title="Export Report", control_type="Button")
export_btn.click_input()

app = Desktop(backend="uia").window(title_re=".*Export.*")

# Create timestamp for unique filenames
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

# Coordinates for the filename input (adjust for your screen)
filename_field_coords = (250, 400)

# Function to save file in given format
def save_file(extension, down_presses):
    filename = rf"H:\joedev\Myfolder\report_{timestamp}.{extension}"

    time.sleep(2)  # Wait for Save As dialog to appear
    pyautogui.click(*filename_field_coords)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.write(filename, interval=0.05)

    # Tab to Save as type
    pyautogui.press('tab')
    time.sleep(0.5)

    # Navigate and select format
    pyautogui.press('down', presses=down_presses, interval=0.3)
    pyautogui.press('enter')

    # Press enter to save
    time.sleep(0.5)
    pyautogui.press('enter')

# ---- SAVE PDF FIRST ----
save_file("pdf", down_presses=2)

# ---- Wait and re-open Save As dialog manually or via automation ----
time.sleep(3)
# Optionally: automate reopening the dialog here
pyautogui.press('enter')

export_btn=rpt_win.child_window(title="Export Report", control_type="Button")
export_btn.click_input()

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

# Coordinates for the filename input (adjust for your screen)
filename_field_coords = (250, 400)

# Function to save file in given format
def save_file(extension, down_presses):
    filename = rf"H:\joedev\Myfolder\report_{timestamp}.{extension}"

    time.sleep(2)  # Wait for Save As dialog to appear
    pyautogui.click(*filename_field_coords)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.write(filename, interval=0.05)

    # Tab to Save as type
    pyautogui.press('tab')
    time.sleep(0.5)

    # Navigate and select format
    pyautogui.press('down', presses=down_presses, interval=0.3)
    pyautogui.press('enter')

    # Press enter to save
    time.sleep(0.5)
    pyautogui.press('enter')

# ---- SAVE PDF FIRST ----
save_file("csv", down_presses=3)
time.sleep(3)
# Optionally: automate reopening the dialog here
pyautogui.press('enter')

try:
    app = Application(backend="uia").start(r"c:\DAIVELKTM\TMProj.exe")
except:
    app = Application(backend="win32").start(r"c:\DAIVELKTM\TMProj.exe")

# [rest of your automation logic goes here...]

#  After automation, close the app
try:
    app.kill()  # <- correct call
except Exception as e:
    print("Failed to close the application:", e)
