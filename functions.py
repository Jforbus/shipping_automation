import os
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
import smtplib
from email import encoders
import gspread_dataframe as gd
import gspread as gs
from datetime import date


# Function for reading/writing Gsheets data with pandas
def pandas_to_sheets(url,ws_name,df=None,mode='r'):
    # helper function for handling gsheets using gspread and gspread_dataframe
    # has 3 modes, w - write: replace existing sheet data or create new, 
    # a - append: keep existing data and add new
    # r - read: read data from gsheets

    gc = gs.oauth(credentials_filename=r"C:\Users\jacob\Desktop\Shipping_Robot\shipping_automation\oauth\credentials.json", 
              authorized_user_filename=r"C:\Users\jacob\Desktop\Shipping_Robot\shipping_automation\oauth\authorized_user.json")
    try:
      ws = gc.open_by_url(url).worksheet(ws_name) # open worksheet by GSheet URL and worksheet name
    except:
       print("Authentication Expired, Refreshing...")
       os.remove(r"C:\Users\jacob\Desktop\Shipping_Robot\shipping_automation\oauth\authorized_user.json")
       ws = gc.open_by_url(url).worksheet(ws_name) # open worksheet by GSheet URL and worksheet name
    finally:
       print("Authentication Successful")
    # clear and write new data to worksheet
    if(mode=='w'):
        ws.clear()
        gd.set_with_dataframe(worksheet=ws,dataframe=df,include_index=False,include_column_header=True,resize=True)
        return True
    # append new data to existing data in worksheet
    elif(mode=='a'):
        ws.add_rows(len(df))
        gd.set_with_dataframe(worksheet=ws,dataframe=df,include_index=False,include_column_header=False,row=len(ws.get_all_values())+1,resize=False)
        return True
    # get data from worksheet as df, here including only necessary columns for shipping robot.
    else:
        return gd.get_as_dataframe(worksheet=ws,usecols=[2,3,4,5,9,15,16,20,22], header=1)
    

# Function for emailing Pandas Dataframes
def email_df(sender, recipients, subject, df, password):
  
  # parse recipients if list
  if type(recipients) == list:
    rec_list = [elem.strip().split(',') for elem in recipients]
  else:
    rec_list = recipients
  today = date.today().strftime("%m/%d/%Y")
  # create message, add subject and sender
  msg = MIMEMultipart()
  msg['Subject'] = subject
  msg['From'] = sender

  # format df as html table
  html = """\
  <html>
    <head></head>
    <body>
      {0}
    </body>
  </html>
  """.format(df.to_html(index=False))
  # add table to message body
  msg_txt = MIMEText(html, 'html')
  msg.attach(msg_txt)
  today = date.today().strftime("%m-%d-%Y")
  # write df to excel file
  df.to_excel(f"attch.xlsx")
  # open excel file and attach to email
  with open(f"attch.xlsx", "rb") as attachment:
        msg_attach = MIMEBase("application", "octet-stream")
        msg_attach.set_payload((attachment).read())
  # encode xlsx file
  encoders.encode_base64(msg_attach)
  # attach file to email
  msg_attach.add_header(
  "Content-Disposition",
  f"attachment; filename= {today}.xlsx")
  msg.attach(msg_attach)
  # remove created file
  os.remove(f"attch.xlsx")

  # connect to email server and send email
  server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
  server.login(sender, password)
  server.sendmail(msg['From'], rec_list , msg.as_string())
  server.quit()

  return print("DF EMAILED SUCCESSFULLY")

# Function for emailing Pandas Dataframes
def email_msg(sender, recipients, subject, message, password):
  
  # parse recipients if list
  if type(recipients) == list:
    rec_list = [elem.strip().split(',') for elem in recipients]
  else:
    rec_list = recipients

  # create message, add subject and sender
  msg = MIMEMultipart()
  msg['Subject'] = subject
  msg['From'] = sender

  # add message to email body
  msg_txt = MIMEText(message)
  msg.attach(msg_txt)

  # connect to email server and send email
  server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
  server.login(sender, password)
  server.sendmail(msg['From'], rec_list , msg.as_string())
  server.quit()

  return print("NO UNSCHEDULED UNITS")
