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

print("--- EXTRA WOODS EMAIL SCRIPT ---")
print(EFX.today().strftime("%d/%m"))

# %%
excel_path = "/path/to/excel/wood_control.xlsm"
boats = pd.read_excel(excel_path, sheet_name="Selection")["LINE"].unique()

def input_cost_center():
    while True:
        try:
            cc = int(input("\nEnter the Cost Center (between 1000 and 9999): "))
            if 1000 <= cc < 10000:
                print(f"Cost Center: {cc}")
                return cc
            else:
                print("Invalid Cost Center, try again.")
        except ValueError:
            print("Invalid input, try again.")

def input_requester():
    return str(input("\nEnter the requester name: ")).strip()

def input_order(line: str):
    while True:
        try:
            order = int(input(f"\nEnter the order (OP) for boat {line} (between 1 and 999): "))
            if 1 <= order < 1000:
                print(f"Boat {line}-{order}")
                return order
            else:
                print("Invalid OP, try again.")
        except ValueError:
            print("Invalid input, try again.")

boat_orders = {line: input_order(line) for line in boats}
requester = input_requester()
cost_center = input_cost_center()

# %%
message = []

if datetime.now().hour > 11:
    message.append("Good afternoon, ")
else:
    message.append("Good morning, ")

message.append("below is the table with the extra woods we need for ")

for line in boats:
    if line == boats[-1]:
        message.append(f"{line}-{boat_orders[line]}.")
    else:
        message.append(f"{line}-{boat_orders[line]}, ")

message.append(f"\nDeliver these woods to {requester}.")
message.append(f"\nUse cost center {cost_center}.")
message = "".join(message)

today = EFX.format_date(EFX.today(), include_time=True)
subject = f"EXTRA WOODS {', '.join(boats.astype(str))} - {today}"

EFX.email_outlook(["recipient1@example.com", "recipient2@example.com"], subject, copy=False, body_content=message)

pyg.hotkey("ctrl", "v")
pyg.sleep(0.5)
pyg.press("enter")
pyperclip.copy("If you have any questions, feel free to reach out.")
pyg.hotkey("ctrl", "v")
pyg.sleep(0.5)
pyg.press("del", presses=75)

print("Done, email drafted!")
pyg.sleep(2)
