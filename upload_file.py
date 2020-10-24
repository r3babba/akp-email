# -*- coding: utf-8 -*-
"""
Created on Fri Oct  9 15:58:33 2020

@author: n
"""

from flask import Flask,render_template,request
import pandas as pd
import SMTP_Outlook_Final

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/display_file', methods = ['POST'])
def display_file():
    file = request.form.get('fname')
    reader = pd.read_excel(file)
    return render_template('home.html',df=reader)
    

@app.route('/send', methods = ['POST'])
def send():
    file = request.form.get('fname')
    reader = pd.read_excel(file)
    e_list=[]
    #reader_dict = reader.to_dict()
    for row in reader.itertuples():
        r_dict = {}
        reciever_email = row.email
        state = row.state
        name = row.name
        method = 'outlook'
        localSession = SMTP_Outlook_Final.HrEmailAutomation(reciever_email, state, name, method)
        localSession.run()
        
        r_dict['name'] = name
        e_list.append(r_dict.copy())
        
    #values = json.loads[e_list]
    return render_template('home.html',data=e_list)

@app.route('/back', methods = ['POST'])
def back():
    return render_template('home.html')

if __name__=="__main__":
    app.run(debug=True)