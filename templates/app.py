from flask import Flask, render_template, request, send_file
import pandas as pd
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    # Gets your Excel sheet and your 1-60 tickets
    file = request.files['test_sheet']
    tickets = request.files.getlist('tickets')
    
    # Logic to process the Excel and match tickets goes here
    df = pd.read_excel(file)
    # (Simplified for now: it just adds a column saying it received the files)
    df['Audit_Status'] = f"Processed with {len(tickets)} support documents"
    
    df.to_excel("audit_results.xlsx", index=False)
    return send_file("audit_results.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
