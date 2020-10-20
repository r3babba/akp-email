# -*- coding: utf-8 -*-
"""
Created on Fri Sep 25 12:00:52 2020

@author: Rahul Kalubowila
"""
# importing required librabraries
import email, smtplib, ssl, getpass
import pandas as pd

from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import xml.etree.ElementTree as ET

# local SMTP server 
# python -m smtpd -c DebuggingServer -n localhost:1025


# use the parse() function to load and parse an XML file
tree = ET.parse('data.xml')
root = tree.getroot()


port = root[0][1].text  
smtp_server = root[0][0].text
sender_email = root[1][0].text 
subject_email = root[1][1].text
reader = pd.read_excel('recruit_info.xlsx', index_col=None, sheet_name='Information')  
sender_info = pd.read_excel('recruit_info.xlsx', index_col=None, header=None,
                            sheet_name='HR Information')  


def setMessageBody(templatePath, message):
    """ 
        The function to set message body - this imports both the html templates and the fail safe of basic text
  
        Parameters: 
            String (templatePath): The path of the email template. 
          
        Returns: 
            None: None
        """
        
    # Create the plain-text and HTML version of your message
    text = """\
    Hi {name},
    How are you?
    Testing mail for HR Automation Project"""
    
    #importing html email template 
    with open(templatePath, 'r') as f:
        html = f.read()

    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")
    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(part1)
    message.attach(part2)

# 
def attachFile(filepath, message):
    """ 
        The function to attach files to the mail body
  
        Parameters: 
            String (filepath): The path of the file to be attached. 
          
        Returns: 
            None: None
        """
    filepath = root[1][2].text
    # Open PDF file in binary mode
    with open(filepath, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filepath= {filepath}",
    )

    # Add attachment to message and convert message to string
    message.attach(part)
    #text = message.as_string()

    



def main():
    try:
        server = smtplib.SMTP(smtp_server,port)
        server.ehlo() # Can be omitted
    
        # creating a list of dataframe columns
        for row in reader.head().itertuples(): 
        
            message = MIMEMultipart("alternative")
        
            message["To"] = row.email 
            message["From"] = sender_email
            message["Subject"] = subject_email
             
            def templatePicker(x):
                """ 
            The function to pick template based on status
  
            Parameters: 
                String (x): State of the candidate. 
          
            Returns: 
                String (): file path of the template
            """
                return {
                    1: 'email_templates/accept_template/index.html',
                    2: 'email_templates/accept_template/index.html'
                }.get(x, 'email_templates/accept_template/index.html') 


            setMessageBody(templatePicker(row.state), message)
        
            attachFile('dummy.pdf', message)
        
            server.sendmail(
                    sender_email,
                    row.email,
                    message.as_string().format(name=row.name),
                )            
    
    except Exception as e:
        # Print any error messages to stdout
        print(e)
    finally:
        server.quit() 

if __name__== "__main__" :
    main()