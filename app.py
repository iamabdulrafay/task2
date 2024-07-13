import pandas as pd 
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime,timedelta
import os
from dotenv import load_dotenv

load_dotenv()

# read banoqabil excel sheet to insert data into csv file rather than instering manuallay ========

data = pd.read_excel("banoqabil.xlsx")

# (data cleaning)
# names geeting
name_bq=data["Name"].dropna()

#remove average,passing,Tootal that is come in it 
remove_names=name_bq.drop(name_bq.tail(3).index)

# getting email 
emails_bq=data["Email"].dropna()



# making csv from pandas from bq sheet ============================

# data frame make for uplaod this data into excel file 

df_obj={
    "Name":remove_names,
    "Email":emails_bq,




}

info=pd.DataFrame(df_obj)
# Save into infromation.Csv 
info.to_csv("information.csv",index=False)
# read info file to get data 
info_file=pd.read_csv("information.csv")


name_info=info_file["Name"]
email_info=info_file["Email"]




# Main email sending fun 
def Email_automate(fromemail, message, subject, toemail):
    try:
        msg = MIMEMultipart()
        msg["From"] = fromemail
        msg["To"] = toemail
        msg["Subject"] = subject
        
  
        message_send = str(message)
        msg.attach(MIMEText(message_send, 'plain'))

        email_server = smtplib.SMTP("smtp.gmail.com", 587)
        email_server.starttls()
        email_server.login(fromemail, os.getenv("EMAIL_PASSWORD"))

        # Send email message only on Mondays (0==Monday)
        today = datetime.today()
        if today.weekday() == 5:
            email_server.send_message(msg)
            print(f"Email sent successfully to {toemail}")

        email_server.quit()

    except Exception as e:
        print(f"Email not sent. Error: {e}")

# reaidng quotes csv ================
quotes_csv = pd.read_csv("quotes.csv")
quotes=quotes_csv["Quotes"]

# Looping to send one quote for one emails 
for i in range(len(email_info)):
    quote = quotes[i]
    toemail = email_info[i]
    Email_automate("abdulrafaydev100@gmail.com", quote, "sheer o shayari", toemail)





