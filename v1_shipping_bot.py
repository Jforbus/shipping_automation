# imports
# 16 digit App Password generated for gmail access
from config import gmail_key
# data handling
from datetime import date
import pandas as pd
# created functions for importing and emailing data
from functions import pandas_to_sheets, email_df, email_msg
import os


# Import Production data from Google Sheets
# Collect data using gspread_dataframe

# Input and Output Sheet URLs
VI = 'https://docs.google.com/spreadsheets/d/1bKwL32NDnkvZQdiba4pqfvvGDkqh9CBO3akwBV-uJEc/edit#gid=27'
LM = 'https://docs.google.com/spreadsheets/d/1vP8FzyoRfl1bnO4kSQa6t2xg5RdN4w8u0RAbd6rPzEM/edit#gid=0'

vi_df = pandas_to_sheets(VI,"PR") # get data from gsheets using created function
vi_df.fillna("",inplace=True) # Remove Nonetypes

# Select CA Completed Units
ca_start = vi_df[vi_df['Shipping Arranger'].str.match('V_START', na=False)].index[0] # find start of CA Completed units
cdf = vi_df.iloc[ca_start+2:] # remove data before selection
ca_end = cdf[cdf['VIN #'] == ''].index[0] # find end of CA Completed units
ca_df = vi_df.iloc[ca_start+2:ca_end].copy()

# Select TX Completed Units
tx_start = vi_df[vi_df['Shipping Arranger'].str.match('1954_START', na=False)].index[0] # find start of TX Completed units
tdf = vi_df.iloc[tx_start+1:] # remove data before selection
tx_end = tdf[tdf['VIN #'] == ''].index[0] # find end of TX Completed units
tx_df = vi_df.iloc[tx_start+1:tx_end].copy()

# add origin Ids
ca_df['Origin'] = ["CA" for x in range(len(ca_df))]
tx_df['Origin'] = ["TX" for x in range(len(tx_df))]

# join completed units into single df
joined_df = ca_df.append(tx_df)

# filter completed for unscheduled, ready to ship
joined_df['Notes'] = joined_df['Notes'].str.lower() # lowercase for string matching
joined_df = joined_df[(joined_df['Notes'].str.contains('ready to ship')) & 
                      (joined_df['Shipping Arranger'].str.contains('VALEW|1954 MFG')) & 
                      (joined_df['Est. Ship Date'] == "")]

# remove duplicate location info from customer column
joined_df['Customer'] = joined_df["Customer"].str.extract(r'(.*-)')

# reference dictionaries
# dictionary of Dimensions
dims = {
    "4000" : "32'L x 8'W x 10'H x 22,000lbs",
    "2000" : "25'L x 8'W x 9'H x 15,000lbs",
    "10" : "25'L x 8'W x 9'H x 15,000lbs",
    "15" : "30'L x 8'W x 9'H x 15,000lbs",
    "WATER TOWER" : "36'L x 8'W x 13'H x 19,500lbs",
    "1000" : "22'L x 8'W x 10'H x 20,000lbs",
}
# dictionary of Models
models = {
    "MA" : "MACK",
    "M2" : "FRLR",
    "SD" : "FRLR",
    "F7" : "FORD",
    "F5" : "FORD",
    "F6" : "FORD",
    "KW" : "KW",
    "PB" : "PB",
    "INT" : "INT",
    "WA" : "PORTABLE"
}
# dictionary of Axle quantities
axles = {
    "132" : "3-Axle Truck X 1",
    "084" : "2-Axle Truck X 1",
    "096" : "2-Axle Truck X 1",
    "108" : "2-Axle Truck X 1",
    "120" : "2-Axle Truck X 1",
    "168" : "2-Axle Truck X 1",
    "170" : "2-Axle Truck X 1",
    "WER" : "WATER TOWER"
}
# dictionary of origin adresses
origins = {
    "TX" : "1954 MFG - 4688 SH-16, GRAHAM, TX 76450",
    "CA" : "VALEW - 12522 VIOLET RD, ADELANTO, CA 92301"
}


# Create Shipping Table

# list comprehension for creating ready to ship table
# extract last 6 digits of VIN #
vins = [x[-6:] for x in joined_df['VIN #']] 
# lookup dimensions for body style in created dict
dim = [dims[x] for x in joined_df['Body Type'].str.extract("(^\d+|WATER TOWER)")[0]]
# Comine Customer Name and Shipping Address for Destination Info
dest = [x for x in joined_df[['Customer', 'Shipping Address']].apply(' '.join, axis=1)]
# create item descriptions from created dicts
items = [f'{models[x[:2]]} {axles[x[-3:]]}' for x in joined_df['Truck Type']] # create item descriptions from lookups
# delivery cost
rates = [x for x in joined_df['Delivery Cost']]
# todays date
rts = [date.today().strftime("%m/%d/%Y") for x in range(len(joined_df))]
# origin address from dict
origin = [origins[x] for x in joined_df['Origin']]
# unit quantity
qty = ['1' for x in range(len(joined_df))]
# Status placeholder
stat = ['Ready' for x in range(len(joined_df))]
# id placeholder
id = ['TBD' for x in range(len(joined_df))]
# completed shipping dictionary
shipping_dict = {"STATUS": stat, "LOAD_#": id,"UNIT_QTY" : qty, "VIN" : vins, "TOTAL DIMENSIONS" : dim, "ORIGIN": origin, "DESTINATION" : dest,
           "ITEM DESCRIPTION" : items, "RATE" : rates, "READY TO SHIP" : rts}
# create dataframe from shipping dictionary
shipping_df = pd.DataFrame(shipping_dict)


# if there are units ready to ship
# email dataframe using email_df
# append df to Load Manager Sheet
if len(shipping_df) != 0:
    # arguments for email_df
    today = date.today().strftime("%m/%d/%Y")
    subject = f"Ready to Ship {today}"
    password = gmail_key
    sender = '1954jacobf@gmail.com'
    recipients = 'jacob@1954mfg.com'
    # email df
    email_df(sender, recipients, subject, shipping_df, password)

    # append df to lm sheet by origin worksheet
    # ca origin to ca worksheet
    ca_final = shipping_df[shipping_df['ORIGIN'] == origins["CA"]]
    if len(ca_final) != 0:
        pandas_to_sheets(LM,'CA',df=ca_final,mode='a')
    else:
        print("No Unscheduled CA Units")
    
    # tx origin to tx worksheet
    tx_final = shipping_df[shipping_df['ORIGIN'] == origins['TX']]
    if len(tx_final) != 0:
        pandas_to_sheets(LM,'TX',df=tx_final,mode='a')
    else:
        print("No Unscheduled TX Units")


# if no units ready to ship, email notification
else:
    today = date.today().strftime("%m/%d/%Y")
    subject = f"Ready to Ship {today}"
    password = gmail_key
    sender = '1954jacobf@gmail.com'
    recipients = 'jacob@1954mfg.com'
    message = "No Unscheduled Units Ready to Ship"
    
    email_msg(sender, recipients, subject, message, password)

