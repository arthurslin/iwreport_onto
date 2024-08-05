import pandas as pd
import numpy as np
import glob
import os

iwdir = "iwcost"
matrixdir = "appmatrix"
splitdir = "splitproducts"


def load_data(dir):
    to_return = []
    paths = glob.glob(os.path.join(dir, "*xlsx"))
    if not paths:
        raise FileNotFoundError(dir, "File not found")
    for path in paths:
        xl = pd.ExcelFile(path)
        df = pd.read_excel(path, sheet_name=xl.sheet_names[0])
    to_return.append(df)
    return to_return


plc = "Part Labor Cost"
rep_view = "Report View"
region = "Employee Region"
pclassdesc = "Product Class Desc"


def add_app_cost(data):
    apps_matrix = load_data(matrixdir)[0].set_index("Region")["Rate"].to_dict()
    for df in data:
        all_app = df.drop(
            df[df["Report View"] != 'APPS'].index, inplace=False)
        all_app[plc] = all_app[region].map(apps_matrix)
        all_app["Total Cost"] = all_app[plc] * all_app["Quantity"]
        all_app.to_excel("apps.xlsx", index=False)


def clean_filename(filename):
    illegal_chars = {'\\', '/', ':', '*', '?', '"', '<', '>', '|', '\0'}

    clean_name = ''.join(
        '_' if char in illegal_chars else char for char in filename)

    return clean_name


def split_by_desc(data):
    if not os.path.exists(splitdir):
        os.makedirs(splitdir)
    for df in data:
        unique_items = df[pclassdesc].unique()

        for item in unique_items:
            if type(item) != str:
                continue
            cur_item = df.drop(
                df[df[pclassdesc] != item].index, inplace=False)
            fn = clean_filename(item)
            cur_item.to_excel(splitdir + "/" + fn + ".xlsx")


iwdata = load_data(iwdir)

# add_app_cost(iwdata)
split_by_desc(iwdata)
