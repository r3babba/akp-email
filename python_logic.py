# -*- coding: utf-8 -*-
"""
Created on Fri Oct  9 09:41:12 2020

@author: n
"""

#importing required libraries
import smtplib, ssl, getpass

from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import xml.etree.ElementTree as ET

import logging
logging.basicConfig(level=logging.INFO, filename='Log.log', format='%(asctime)s :: %(levelname)s ::%(message)s')


class HrEmailAutomation:
    
    #Class variables
    #use the parse() function to load and parse an XML file
    tree = ET.parse('data.xml')
    root = tree.getroot()
    
    sender_email = root[1][0].text
    subject_email = root[1][1].text
    serverPort = root[0][1].text
    smtpServer = root[0][0].text
    thispassword = root[1][2].text
    
    #Constructor method with instance variables
    def __init__(self, reciever_email, state, name):
        self.reciever_email = reciever_email
        self.state = state
        self.name = name
        logging.info('Constructor Function')
        
    def setMessageBody(self, templatePath):
        """
            The function to set message body - this imports both the html templates and the fail safe 

            Parameters:
                String (templatepath): The path of the email template.
    
            Returns:
                None: None

        """
        
        #Create the plain-text and HTML version of your message
        #importing HTML email template
        with open(templatePath, 'r') as f:
            html = f.read()
        logging.info('Set meassage body for the mail')
        return html
        
        
    def templatePicker(self, x):
        """
            The function to pick template based on status
            
            Parameters:
                String (x): State of the candidate.
                
            Returns:
                String (): file path of the template
        """
        return {
            1: 'email_templates/accept_template/index.html',
            2: 'email_templates/reject_template/index.html',
            }.get(x, 'email_templates/initial_template/index.html',)
    
    def run(self):

        html = self.setMessageBody(self.templatePicker(self.state))

        message = MIMEMultipart()
        message["Subject"] = self.subject_email
        message["From"] = self.sender_email
        message["To"] = self.reciever_email
        body = MIMEText(html.format(name=self.name, jobposition='intern'), "html")
        message.attach(body)
        message = message.as_string()

        try:
            server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            server.ehlo()
            server.login(self.sender_email, self.thispassword)
            server.sendmail(self.sender_email, self.reciever_email, message)
            server.close()

            print('Email sent')
        except:
            print('Something went wrong bro...')
        

        