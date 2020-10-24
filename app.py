# -*- coding: utf-8 -*-
"""
Created on Fri Oct  9 12:11:01 2020

@author: n
"""

from flask import Flask,render_template,request
import pandas as pd
import SMTP_Outlook_Final

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('home.html')
    file = request.form['fname']
    reader = pd.read_excel(file)
    return render_template('home.html',data=reader)

@app.route('/send', methods = ['POST'])
def send():
    #reader = pd.read_excel('recruit_info.xlsx', index_col=None, sheet_name='Sheet1')
    for row in reader.itertuples():
        reciever_email = row.email
        state = row.state
        name = row.name
        method = 'outlook'
        localSession = SMTP_Outlook_Final.HrEmailAutomation(reciever_email, state, name, method)
        localSession.run()
    return 'Done'
        
if __name__=="__main__":
    app.run(debug=True)