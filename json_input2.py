# -*- coding: utf-8 -*-
"""
Created on Wed May 19 10:45:18 2021

@author: kalurah
"""
#requests.post('http://127.0.0.1:5000/json-example',data={"toList":"rahulkalubowila@gmail.com","ccList":"rahulkalubowila@gmail.com","message":"This is the PM email message"})

import flask
from flask import request
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import xml.etree.ElementTree as ET
import smtplib, ssl

app = flask.Flask(__name__)
app.config["DEBUG"] = True


@app.route('/', methods=['GET'])
def home():
    return """ <h1>Email Automation API Documentation</h1><p>This site is a API for automating emails</p><p>Homepage: http://127.0.0.1:5000/</p><p>Email API URL: http://127.0.0.1:5000/json-example</p><p>JSON Template: {
    "toList": [
        "rahulkalubowila@gmail.com",
        "abc@gmail.com"
    ],
    "ccList": [
        "rahulkalubowila@gmail.com",
        "rahulkalubowila@gmail.com"
    ],
    "subject": "This is the subject",
    "message": "This is the PM email message",
    "template": 1
}</p> """
@app.errorhandler(404)
def page_not_found(e):
    return "<h1>404</h1><p>The resource could not be found.</p>", 404




def setMessageBody(templatePath):
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
        return html
        
        
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
        2: 'email_templates/reject_template/index.html',
        }.get(x, 'email_templates/initial_template/index.html',)


@app.route('/json-example', methods=['POST'])
def json_example():
    content = request.get_json()
    print(content)
    print(type(content))
    reciever_email = content['toList']
    print(reciever_email)
    print(type(reciever_email))
    cc_email = content['ccList']
    print(cc_email)
    print(type(cc_email))
    message_body = content['message']
    html = setMessageBody(templatePicker(content['template']))
    print(message_body)
    print(type(message_body))
    
    tree = ET.parse('data.xml')
    root = tree.getroot()
    
    sender_email = root[1][0].text
    #subject_email = root[1][1].text
    serverPort = root[0][1].text
    smtpServer = root[0][0].text
    thispassword = root[1][2].text

    message = MIMEMultipart()
    message["Subject"] = content['subject']
    message["From"] = sender_email
    message["To"] = ', '.join(reciever_email)
    message['Cc'] = ', '.join(cc_email)
    message.attach(MIMEText(html.format(name=message_body, jobposition='intern'), "html"))

    try:
        # Create a secure SSL context
        context = ssl.create_default_context()
        server = smtplib.SMTP(smtpServer, serverPort)
        server.ehlo()
        server.starttls(context=context) # Secure the connection
        server.login(sender_email, thispassword)
        server.sendmail(sender_email, reciever_email, message.as_string())
        server.close()
    
        return 'Email sent'
    except Exception as e:
        return "Error"


app.run(debug=True, use_reloader=False)