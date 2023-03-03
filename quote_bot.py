# -*- coding: utf-8 -*-

# imports
# 16 digit App Password generated for gmail access
from config import gmail_key

# data handling
from datetime import date
import pandas as pd

# created functions for importing and emailing data
from functions import pandas_to_sheets, email_df
import os

# Import Production data from Google Sheets
# Collect data using gspread_dataframe
# Production Data Sheet URL
VI = 'https://docs.google.com/spreadsheets/d/1bKwL32NDnkvZQdiba4pqfvvGDkqh9CBO3akwBV-uJEc/edit#gid=27' 

q_df = pandas_to_sheets(VI,"PR") # get data from gsheets using pandas_to_sheets
q_df.fillna("",inplace=True) # Remove Nonetypes

# Filter Production Data for Units to Quote
# select rows where Delivery Cost is empty, 
# Shipping arranger is Valew or 1954, 
# and Vin number is present
q_df = q_df[(q_df['Shipping Arranger'].str.match('VALEW|1954 MFG') ) & 
            (q_df['Delivery Cost'].str.match("|TBD")) & 
            (q_df['VIN #'].str.match(r".+\d{4}")) &
            (~q_df['Customer'].str.contains('Slot'))].copy()

# Group by delivery address and body type
to_quote = q_df.groupby(['Customer','Body Type', 'Shipping Address']).count()['VIN #']
to_quote = pd.DataFrame(to_quote.reset_index()).rename(columns={'VIN #':'Count'})

# Email the to_quote df
# arguments for email_df
today = date.today().strftime("%m/%d/%Y")
subject = f"To Quote {today}"
password = gmail_key
sender = '1954jacobf@gmail.com'
recipients = 'jacob@1954mfg.com'

email_df(sender, recipients, subject, to_quote, password)
