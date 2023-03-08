# shipping_automation

## Automating Data collection and dissemination for Outbound Freight handling processes.

### Tools: Python, Google Sheets

## Process
1. Data is loaded directly from Google Sheets using Python Libraries gspread and gspread_dataframe.
    - Google Sheets API is turned on for the project in the Google Developer Console.
    - An Oauth2.0 credentials.json file is created from the Console and held in the oauth folder.
    - After authorization the returned authorized_user.json is held in the same oauth folder.
    - A function is created (pandas_to_sheets) that uses gspread and gspread_dataframe to handle the Gsheets connection.
        - Read: Takes URL and Worksheet name as arguments, returns a pandas dataframe with worksheet data.
        - Write: Takes URL, Worksheet name, and a pandas df, clears or creates new worksheet and adds data from the dataframe.
        - Append: Takes URL, Worksheet name, and a pandas df, adds data from the dataframe below existing data.

![pandas_to_sheets](https://github.com/Jforbus/shipping_automation/blob/main/images/pandas_to_sheets2X.png)
2. The completed units for each facility are selected from the dataframe using string matching and index slicing.

3. The completed units are filtered for only those ready to ship.

4. The resulting dataframe of units ready to ship is tranformed into a shipping data table using list comprehension and dictionaries.

5. The resulting dataframe is appended to the output Gsheet for scheduling, and emailed as an xslx file to the necessary parties using smptlib.

## Results
This automation saves the logistics personnel 10+ hours a week of routine data collection and processing. 

    
