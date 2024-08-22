import pandas as pd
import numpy as np
import datetime
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
            common_items_df = pd.read_excel(path, sheet_name=xl.sheet_names[0])

            while np.nan in common_items_df.columns or "Unnamed" in common_items_df.columns[1]:
                common_items_df.columns = common_items_df.iloc[0]
                common_items_df = common_items_df.iloc[1:].reset_index(drop=True)

        to_return.append(common_items_df)
    return to_return

entitlement = "Current Entitlement"
keep_items = ["Acceptance Warranty", "The Nanometrics Service Contract","The Nanometrics Service Contra",np.nan,"nan"]
cols_to_keep = ['Tool Serial Number', 'Customer Name', 'Current Entitlement', 'Shipped Date', 'Install To Onto Spec Duration ( Days ) ', 'Install To Onto Spec FSE Labor hours Posted',
                'Install To Onto Spec FSE Labor hours Posted Cost ( USD )', 'Install To Cust Spec Duration (  Days )', 'Install To Cust Spec FSE Labor hours Posted',
                'Install To Cust Spec FSE Labor hours Posted Cost  ( USD ) ', '% Warranty Schedule Complete', 'Warranty Cost To Date']
ie_sn = "Tool Serial Number"
ib_sn = "Serial Number"

floor_date = datetime.datetime.today()-datetime.timedelta(days=455)
ceil_date = datetime.datetime.today()-datetime.timedelta(days=90)

def match_items(data):
    ibdetail, installefficiency = data

    #drop within dates
    sd = "Ship Date"
    ibdetail[sd] = pd.to_datetime(ibdetail[sd])
    ibdetail = ibdetail.loc[(ibdetail[sd] >= floor_date) & (ibdetail[sd] <= ceil_date)]

    ibdetail = ibdetail[ibdetail[entitlement].isin(keep_items)]
    installefficiency = installefficiency[installefficiency[entitlement].isin(
        keep_items)]

    # Format Items
    common_items =  installefficiency[installefficiency[ie_sn].isin(ibdetail[ib_sn])]
    common_items_df = common_items[cols_to_keep]
    common_items_df["% Warranty Schedule Complete"] = common_items_df["% Warranty Schedule Complete"] * 100
    common_items_df.loc[common_items_df["% Warranty Schedule Complete"] > 1, "% Warranty Schedule Complete"] = 1
    common_items_df["Shipped Date"] = pd.to_datetime(common_items_df["Shipped Date"]).dt.date
    common_items_df.loc["SUM"] = common_items_df[cols_to_keep[4:]].sum()

    common_items_df.to_excel("common_items.xlsx",index=True)

    # Format Items
    different_items_df = pd.DataFrame()
    different_items_df["Install Efficiency"] = installefficiency[~installefficiency[ie_sn].isin(ibdetail[ib_sn])][ie_sn]
    different_items_df["IB Detail"] = ibdetail[~ibdetail[ib_sn].isin(installefficiency[ie_sn])][ib_sn]

    different_items_df.to_excel("removed_items.xlsx",index=False)



directories = ["ibdetail", "installefficiency"]

dataset = load_data(directories)

match_items(dataset)