# -*- coding: utf-8 -*-
"""
Created on Mon Oct 12 11:19:36 2020

@author: n
"""

from flask import Flask,render_template,request
import pandas as pd
import python_logic

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/display_file_send', methods = ['POST'])
def display_file_send():
    file = request.form.get('fname')
    reader = pd.read_excel(file)
    
    if request.form['submit_button'] == 'Preview':
        df=reader.to_html()
        return render_template('index.html',df=df)
    
    elif request.form['submit_button'] == 'Send':
        e_list=[]
        #reader_dict = reader.to_dict()
        for row in reader.itertuples():
            r_dict = {}
            reciever_email = row.email
            state = row.state
            name = row.name
            method = 'SMTP'
            localSession = python_logic.HrEmailAutomation(reciever_email, state, name, method)
            localSession.run()
            
            r_dict['name'] = name
            e_list.append(r_dict.copy())
            
        #values = json.loads[e_list]
        return render_template('index.html',data=e_list)
    
@app.route('/back', methods = ['POST'])
def back():
    return render_template('index.html')

if __name__=="__main__":
    app.run(debug=True)