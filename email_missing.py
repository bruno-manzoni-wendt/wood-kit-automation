# %%
import pyautogui as pyg
import pandas as pd
import pyperclip
from datetime import datetime
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import sys
sys.path.append("/path/to/custom/libs")
import EFX_lib as EFX

print("--- MISSING WOODS EMAIL SCRIPT ---")
print(EFX.today().strftime("%d/%m"))

# %%
excel_path = "/path/to/excel/wood_control.xlsm"
index = max(pd.read_excel(excel_path, sheet_name="Register")["I"])
urgent = pd.read_excel(excel_path, sheet_name="Backend", header=None).at[0, 0]
boats = pd.read_excel(excel_path, sheet_name="Selection")["LINE"].unique()
print(f"\nIndex: {index} | Urgent? {urgent} | Boat(s): {boats}")

def input_order(line: str):
    while True:
        try:
            order = int(input(f"\nEnter the order (OP) for boat {line} (between 1 and 999): "))
            if 1 <= order < 1000:
                print(f"Boat {line}-{order}")
                return order
            else:
                print("Invalid OP, try again.\n")
        except ValueError:
            print("Invalid input, try again.\n")

boat_orders = {line: input_order(line) for line in boats}

# %%
message = ["--- Automated Email via Python ---\n\n"]

if datetime.now().hour > 11:
    message.append("Good afternoon, ")
else:
    message.append("Good morning, ")

today = EFX.format_date(EFX.today(), include_time=True)

if urgent:
    message.append("below is the table with the urgent missing woods for ")
    subject = f"{index}: URGENT MISSING WOODS - {today}"
else:
    message.append("below is the table with the missing woods for ")
    subject = f"{index}: MISSING WOODS - {today}"

for line in boats:
    if line == boats[-1]:
        message.append(f"{line}-{boat_orders[line]}.")
    else:
        message.append(f"{line}-{boat_orders[line]}, ")

message = "".join(message)

EFX.email_outlook(["recipient1@example.com", "recipient2@example.com"], subject, copy=False, body_content=message)

pyg.hotkey("ctrl", "v")
pyg.sleep(0.5)
pyg.press("enter")
pyperclip.copy("If you have any questions, feel free to reach out.")
pyg.hotkey("ctrl", "v")
pyg.sleep(0.5)
pyg.press("del", presses=75)

if urgent:
    pyg.sleep(0.2)
    for key in ["alt", "m", "a", "d"]:
        pyg.press(key)
        pyg.sleep(0.1)

print("Done, email drafted!")
pyg.sleep(2)
