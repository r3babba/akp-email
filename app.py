# -*- coding: utf-8 -*-
"""
Created on Wed Nov  4 13:35:56 2020

@author: n
"""

from flask import Flask,render_template,request
import pandas as pd
import python_logic

app = Flask(__name__)

@app.route('/', methods = ['GET','POST'])
def home():
    if request.method == 'GET':
        return render_template('index.html')

    if request.method == 'POST':
        
        if request.form['submit_button'] == 'Preview':
            e_list=[]

            print(request.files['fname'])
            f = request.files['fname']
            reader = pd.read_excel(f)
            #file = request.form.get('fname')
            #reader = pd.read_excel(file)
            headings = ("Reciever Email", "State", "Name", "Select")
            for row in reader.itertuples():
                r_dict = {}
                r_dict['reciever_email'] = row.email
                r_dict['state'] = row.state
                r_dict['name'] = row.name
                e_list.append(r_dict.copy())
            return render_template('index.html',headings=headings,data=e_list)
        
        if request.form['submit_button'] == 'Send All Selected':
            values = []
            test = request.form.getlist('check')
            for value in test:
                values.append(value)
            #values = json.loads(json.dumps(values))
            #return jsonify(values)
            result = [dict(
               item.replace("'", '').split(':')
               for item in s[1:-1].split(', ')
               )
          for s in values]
            names = []
            for i in range(len(result)):
                reciever_email = result[i]["reciever_email"]
                state = result[i]["state"]
                name = result[i]["name"]
                names.append(name)
                localSession = python_logic.HrEmailAutomation(reciever_email, state, name)
                localSession.run()
                
            return render_template('index.html', names = names)


if __name__=="__main__":
    app.run(debug=True)