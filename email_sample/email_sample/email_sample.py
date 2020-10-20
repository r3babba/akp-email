import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import datetime
import bs4 as bs
import win32com.client
from win32com.client import Dispatch, constants

now = datetime.datetime.now()
cur_date = now.strftime("%Y%m%d")

send_to = 'Rohan'

#importing html email template 
with open('email_templates/accept_template/index.html', 'r') as f:
  html = f.read()

attachement_file = rf'Z:\autoHR\email_sample\email_sample\sample.jpg'

# Write HTML String to file.html
with open(f"email_{cur_date}.html", "w") as file:
    file.write(html)

const=win32com.client.constants
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = f"Sample Email"
newMail.BodyFormat = 2
newMail.HTMLBody = html
newMail.To = "Rahul.Kalubowila@acuitykp.com"
attachment1 = attachement_file
newMail.Attachments.Add(Source=attachment1, Type=1)
newMail.display()
newMail.Send()
print('done')