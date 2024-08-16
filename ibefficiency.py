import pandas as pd
import numpy as np
import glob
import os

def load_data(dirs):
    to_return = []
    for i in dirs:
        paths = glob.glob(os.path.join(i, "*xlsx"))
        if not paths:
            raise FileNotFoundError(i, "File not found")
        for path in paths:
            xl = pd.ExcelFile(path)
            df = pd.read_excel(path, sheet_name=xl.sheet_names[0])            

            while np.nan in df.columns or "Unnamed" in df.columns[1]:
                df.columns = df.iloc[0]
                df = df.iloc[1:].reset_index(drop=True)

        to_return.append(df)
    return to_return

entitlement = "Current Entitlement"
def match_items(data):
    ibdetail, installefficiency = data

    print(ibdetail.columns.to_list())

directories = ["ibdetail", "installefficiency"]

dataset = load_data(directories)

match_items(dataset)