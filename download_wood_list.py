# %%
import pyautogui as pyg
import shutil
import os

import sys
libs_path = r'/path/to/project/lib'
sys.path.append(libs_path)
import EFX_lib as EFX

# %%
source = r"/path/to/project/data/WOOD_LIST.xlsx"
EFX.download_drive_excel(r'https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/export?format=xlsx&id=SPREADSHEET_ID',
                        source)

today = EFX.today()
filename = f"WOOD_LIST_{round(today.timestamp())}_{today.day}-{today.month}.xlsx"
destination = r"/path/to/project/backup/WOOD_LIST.xlsx"
renamed = f"/path/to/project/backup/{filename}"

shutil.copy(source, destination)
print("--- Copied ---")
os.rename(destination, renamed)
print("--- Renamed ---")
pyg.sleep(2)