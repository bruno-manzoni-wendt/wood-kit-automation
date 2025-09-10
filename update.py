# %%
import pandas as pd
import datetime as datetime
import time
import os
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", category=FutureWarning)

# Path to the backup folder containing Excel files
backup_path = "/path/to/project/backups"

# Compare the two most recent backup files
file_list = os.listdir(backup_path)
before_path = os.path.join(backup_path, file_list[-2])
after_path = os.path.join(backup_path, file_list[-1])

print(f"Comparing changes from {file_list[-2]} to {file_list[-1]}\n")

# %%
print("--- ORGANIZING DATAFRAMES ---")

columns = [
    "LINE", "CODE", "LC", "LETTER", "WOOD", "QTY", "CUT MODE", "DIMENSION (mm)",
    "THICKNESS (mm)", "BEVEL", "HOLE", "FIBER PART", "TEAM", "SECTOR",
    "LOCATION", "MODEL"
]

def adjust(df_dict):
    """Remove 'Excluded' column and reindex sheets with standard columns."""
    if "Excluded" in df_dict:
        del df_dict["Excluded"]
    for sheet in df_dict.keys():
        df_dict[sheet]["LINE"] = sheet
        df_dict[sheet] = df_dict[sheet].reindex(columns=columns)
    return df_dict

before = adjust(pd.read_excel(before_path, sheet_name=None))
after = adjust(pd.read_excel(after_path, sheet_name=None))

def concat(df):
    """Concatenate sheets, clean values, and normalize types."""
    df = df.dropna(axis=0, how="all")
    df["LC"] = df["LINE"].astype(str) + "-" + df["CODE"].astype(str)
    df = df.sort_values(by="LC")
    df.fillna("", inplace=True)
    for col in df.columns:
        if col in ["CODE", "QTY", "THICKNESS (mm)"]:
            continue
        else:
            df[col] = df[col].astype(str)
    return df

before_concat = concat(pd.concat(before.values(), ignore_index=True))
after_concat = concat(pd.concat(after.values(), ignore_index=True))

print("")
# %%
new = pd.DataFrame(columns=columns)
print("--- New Woods ---")

for i in after_concat.index:
    if after_concat.at[i, "LC"] not in list(before_concat["LC"]):
        print(after_concat.at[i, "LC"])
        row = pd.DataFrame([after_concat.loc[i]], columns=columns)
        new = pd.concat([new, row])

new = new.reset_index(drop=True)
new = new.drop("LC", axis="columns")
new["CODE"] = new["CODE"].astype(int)
print("")

# %%
removed = pd.DataFrame(columns=columns)
print("--- Removed Woods ---")

for i in before_concat.index:
    if before_concat.at[i, "LC"] not in after_concat["LC"].values:
        print(before_concat.at[i, "LC"])
        row = pd.DataFrame([before_concat.loc[i]], columns=columns)
        removed = pd.concat([removed, row])

removed = removed.reset_index(drop=True)
removed = removed.drop("LC", axis="columns")
removed["CODE"] = removed["CODE"].astype(int)
print("")

# %%
modified = pd.DataFrame(columns=columns)
index_values = before_concat["LC"].to_dict()
values_index = {value: key for key, value in index_values.items()}
print("--- Modified Woods ---")

for i, row_after in after_concat.iterrows():
    changes = []
    if row_after["LC"] in values_index:
        before_idx = values_index[row_after["LC"]]
        row_before = before_concat.loc[before_idx]

        if not all(row_before == row_after):
            for col in before_concat.columns:

                if row_before[col] != row_after[col]:
                    changes.append(f'{col}: from "{row_before[col]}" to "{row_after[col]}";')

            row = pd.DataFrame([row_after], columns=columns)
            row["LC"] = "\n".join(changes)
            print(row_after["LC"], row["LC"].item())
            modified = pd.concat([modified, row])

modified = modified.reset_index(drop=True)
modified = modified.rename(columns={"LC": "CHANGES"})
modified["CODE"] = modified["CODE"].astype(int)
print("")

# %%
print("Saving results to Excel...")
output_file = "/path/to/project/results/modified_new_removed.xlsx"
with pd.ExcelWriter(output_file) as writer:
    modified.to_excel(writer, sheet_name="Modified", index=False)
    new.to_excel(writer, sheet_name="New", index=False)
    removed.to_excel(writer, sheet_name="Removed", index=False)
