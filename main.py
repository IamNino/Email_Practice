"""
Created on 
@author: 224939

Send the daily report email to the recipient list.
If something fails, it will send an email saying that there was a failure.

Calls the test_script to build the email body.

"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from socket import gethostname
import time


# Define email sending function
def send_message(recipients, subject, body):
  sender = "samir.nino@aa.com"
  msg = MIMEMultipart('alternative')
  msg['Subject'] = subject
  msg['From'] = sender  
  msg['To'] = ";".join(recipients)
  text = MIMEText(body, 'html')
  msg.attach(text)
  
  hostname = gethostname() 
  if hostname in ("rmappp58", "bdhencdcp07.cdc.aa.com"):
    server = 'smtpdc.aa.com'
  else:
    server = 'relay.lcc.usairways.com'
  smtp = smtplib.SMTP(server) 
  smtp.sendmail(sender, recipients, msg.as_string())
  smtp.close()
  return
 # server

def send_email(html_table_1 = '', html_table_2 = '', err_msg = None):
  # Send email if process completes
  recipients = ['samir.nino@aa.com', 'irakli.gudavadze@aa.com']
  subject = "Testing DF Email"
  
  if err_msg != None:
    header = "This is an automated email<br><br><br><h2>DF Code Failed </h2> <br><br>"
    msg = f"There was an issue with the Math: {err_msg}, please contact Samir Nino at samir.nino@aa.com <br><br><b>to be changed</b>" 
  else:
    header = "This is an automated email<br><br><br> actual DF: "
    # prefix to tab: "<br><br>Actual Show Rates:<br>"
    msg = "<br> Testing a DataFrame <br><br>" + html_table_1 + "<br> The second DataFrame: <br><br>" + html_table_2    
  body = header + msg
  send_message(recipients, subject, body)

if __name__ == "__main__":
  
  attempt = 0
  while attempt < 5 :
    try:
      df = pd.read_excel('Book1.xlsx')

      df_email = df.to_html()

      df2 = pd.read_excel('Book1.xlsx')
      df2.drop(columns=df.columns[0], axis=1, inplace=True)

      df2.to_html()

      send_email(df, df2)
      attempt = 99
      
    except ValueError as err:
      send_email(err_msg = err.args)
      time.sleep(7200)
      attempt += 1

# comparable actual showrates from 'showrates' and 'showrates_comparison' table may differ slightly
# because 'showrates_comparison' table discards flights for which the pure forcast cannot be found.
# Such flights are (some) of the flights that don't fly every day. #In principle, I could find pure forcast for those flights as well but not from mars nsf_fva_post_dptr table


