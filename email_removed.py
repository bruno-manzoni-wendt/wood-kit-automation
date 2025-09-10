# %%
import pyautogui as pyg
import pyperclip
from datetime import datetime

import sys
sys.path.append("/path/to/custom/libs")
import EFX_lib as EFX
# %%
now = EFX.format_date(EFX.today(), include_time=True)
subject = f"REMOVED WOODS - {now}"

if datetime.now().hour > 11:
    body = "--- Automated Email via Python ---\n\nGood afternoon, below is the table with the removed woods:"
else:
    body = "--- Automated Email via Python ---\n\nGood morning, below is the table with the removed woods:"

EFX.capslock(False)
EFX.email_outlook(["recipient1@example.com", "recipient2@example.com"], subject, copy=False, body_content=body)

pyg.hotkey("ctrl", "v")
pyg.sleep(0.5)
pyg.press("enter")
pyperclip.copy("If you have any questions, feel free to reach out.")
pyg.hotkey("ctrl", "v")
pyg.sleep(0.5)
pyg.press("del", presses=75)

print("Done, email drafted!")
pyg.sleep(2)
