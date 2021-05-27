# -*- coding: utf-8 -*-
"""
Created on Wed May 19 10:45:18 2021

@author: kalurah
"""
#requests.post('http://127.0.0.1:5000/json-example',data={"toList":"rahulkalubowila@gmail.com","ccList":"rahulkalubowila@gmail.com","message":"This is the PM email message"})

import flask
from flask import request, jsonify
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import xml.etree.ElementTree as ET
import smtplib, ssl

app = flask.Flask(__name__)
app.config["DEBUG"] = True


@app.route('/', methods=['GET'])
def home():
    return "<h1>Email Automation API Documentation</h1><p>This site is a prototype API for distant reading of science fiction novels.</p>"

@app.errorhandler(404)
def page_not_found(e):
    return "<h1>404</h1><p>The resource could not be found.</p>", 404

@app.route('/json-example', methods=['POST'])
def json_example():
    content = request.get_json()
    print(content)
    toList = content['toList']
    ccList = content['ccList']
    message = content['message']
    
    tree = ET.parse('data.xml')
    root = tree.getroot()
    
    sender_email = root[1][0].text
    subject_email = root[1][1].text
    serverPort = root[0][1].text
    smtpServer = root[0][0].text
    thispassword = root[1][2].text


    message = MIMEMultipart()
    message["Subject"] = subject_email
    message["From"] = sender_email
    message["To"] = toList
    message["Cc"] = ccList
    body = message
    message.attach(body)
    message = message.as_string()

    try:
        # Create a secure SSL context
        context = ssl.create_default_context()
        server = smtplib.SMTP(smtpServer, serverPort )
        server.ehlo()
        server.starttls(context=context) # Secure the connection
        server.login(sender_email, thispassword)
        server.sendmail(sender_email, toList, message)
        server.close()

        return 'Email sent'
    except Exception as e:
        return e
    
    
    #return jsonify({"uuid":uuid})

@app.route('/api/sendemail', methods=['GET'])
def api_send_email(): 
    
    query_parameters = request.args

    reciever_email = query_parameters.get('reciever_email')
    cc_email = query_parameters.get('cc_email')
    message_body = query_parameters.get('message_body')
    
    tree = ET.parse('data.xml')
    root = tree.getroot()
    
    sender_email = root[1][0].text
    subject_email = root[1][1].text
    serverPort = root[0][1].text
    smtpServer = root[0][0].text
    thispassword = root[1][2].text


    message = MIMEMultipart()
    message["Subject"] = subject_email
    message["From"] = sender_email
    message["To"] = reciever_email
    body = message_body
    message.attach(MIMEText(body, "plain"))
    message = message.as_string()

    

    try:
        # Create a secure SSL context
        context = ssl.create_default_context()
        server = smtplib.SMTP(smtpServer, serverPort )
        server.ehlo()
        server.starttls(context=context) # Secure the connection
        server.login(sender_email, thispassword)
        server.sendmail(sender_email, reciever_email, message)
        server.close()

        return 'Email sent'
    except Exception as e:
        return e

app.run(debug=True, use_reloader=False)