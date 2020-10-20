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

import datetime
import bs4 as bs
import win32com.client
from win32com.client import Dispatch, constants

import xml.etree.ElementTree as ET

import logging
logging.basicConfig(level=logging.INFO, filename='Log.log', format='%(asctime)s :: %(levelname)s :: %(message)s')



class HrEmailAutomation:

    # Class variables
    # use the parse() function to load and parse an XML file
    tree = ET.parse('data-Copy.xml')
    root = tree.getroot()

    sender_email = root[1][0].text
    subject_email = root[1][1].text
    serverPort = root[0][1].text
    smtpServer = root[0][0].text

    # Constructor method with instance variables 
    def __init__(self, reciever_email, state, name, method):
        self.reciever_email = reciever_email
        self.state = state
        self.name = name
        self.method = method
        logging.info('Constructor Function')

    def setMessageBody(self, templatePath, message):
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
        logging.info('Set message body for the mail')

    def attachFile(self, filepath, message):
        """ 
            The function to attach files to the mail body
    
            Parameters: 
                String (filepath): The path of the file to be attached. 
            
            Returns: 
                None: None
            """
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
        logging.info('Function to attach file')

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
            2: 'email_templates/accept_template/index.html'
        }.get(x, 'email_templates/accept_template/index.html') 

    def run(self):
        if self.method=='outlook':

            #importing html email template 
            with open(self.templatePicker(state), 'r') as f:
                html = f.read()

            attachement_file = rf'Z:\autoHR\dummy.pdf'

            olMailItem = 0x0
            obj = win32com.client.Dispatch("Outlook.Application")
            newMail = obj.CreateItem(olMailItem)
            newMail.Subject = self.subject_email
            newMail.BodyFormat = 2
            newMail.HTMLBody = html
            newMail.To = self.reciever_email
            newMail.Attachments.Add(Source=attachement_file, Type=1)
            newMail.display()
            newMail.Send()
            print('done')


        else:
            # First object, set up instance variables of constructor method
            logging.info('Create class instance')
            try:
                message = MIMEMultipart("alternative")
                logging.info('MIMEMultipart object created')

                message["To"] = self.reciever_email 
                logging.info('Mail is to be sent to '+self.reciever_email)
                message["From"] = self.sender_email
                logging.info('Mail is to be sent from '+self.sender_email)
                message["Subject"] = self.subject_email
                logging.info('Subject of the email is '+self.subject_email)
                    
                


                self.setMessageBody(self.templatePicker(state), message)
            
                self.attachFile('dummy.pdf', message)

                context = ssl.create_default_context()
                with smtplib.SMTP(self.smtpServer, self.serverPort) as server:
                    server.ehlo()  # Can be omitted
                    server.starttls(context=context)
                    server.ehlo()  # Can be omitted
                    password = getpass.getpass(prompt='Password: ', stream=None)
                    server.login(self.sender_email, password)
                    server.sendmail(
                        self.sender_email,
                        self.reciever_email ,
                        message.as_string().format(name=self.name),
                    )  
                logging.info('server smtplib object created')
            
                    
                logging.info('Email is sent to '+str(self.reciever_email)+' from '+str(self.sender_email)+' with the subject '+str(self.subject_email)+' of template '+str(self.state))     
            
            except Exception as e:
                # Print any error messages to stdout
                print(e)
            finally:
                server.quit() 
            """ 
            The function to connect to server and send mail
    
            Parameters: 
                Object (Self): The object created for HrEmailAutomation class
            
            Returns: 
                None: None
            """

        

reciever_email = 'Rahul.Kalubowila@acuitykp.com'
state = 1
name = 'Jeff bezos'
method = 'outlook'
localSession = HrEmailAutomation(reciever_email, state, name, method)
localSession.run()