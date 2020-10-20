import smtplib
# set up the SMTP server
s = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
s.starttls()
s.login(MY_ADDRESS, PASSWORD)