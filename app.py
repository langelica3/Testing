import os
import pandas as pd
import google.generativeai as genai
from flask import Flask, render_template, request, send_file
from io import BytesIO

app = Flask(__name__)

# 1. Setup the AI Brain
# This pulls the key you saved in Render's Environment Variables
genai.configure(api_key=os.environ.get("GEMINI_API_KEY"))
model = genai.GenerativeModel('gemini-1.5-flash')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    # 2. Grab the files from the website
    test_sheet = request.files['test_sheet']
    tickets = request.files.getlist('tickets')
    
    # Read the Excel sheet
    df = pd.read_excel(test_sheet)
    
    # 3. Create a library of your tickets for the AI to search
    ticket_data = {}
    for ticket in tickets:
        # We read the file content (PDF or Image)
        content = ticket.read()
        ticket_data[ticket.filename] = content

    results = []

    # 4. Loop through each row in your Excel (The Audit Test)
    for index, row in df.iterrows():
        sample_id = str(row.get('Sample_ID', ''))
        audit_conclusion = "No matching ticket found"

        # Search the uploaded files for a filename that matches the ID
        for filename, content in ticket_data.items():
            if sample_id in filename:
                # 5. ASK THE AI TO AUDIT THE TICKET
                prompt = f"You are an IT Auditor. Look at this document for Ticket {sample_id}. Is it approved? Who approved it and on what date? Summarize in 10 words."
                
                # Create a 'part' for the AI to read
                doc = {'mime_type': 'application/pdf', 'data': content} if filename.endswith('.pdf') else {'mime_type': 'image/jpeg', 'data': content}
                
                try:
                    response = model.generate_content([prompt, doc])
                    audit_conclusion = response.text
                except Exception as e:
                    audit_conclusion = f"AI Error: {str(e)}"
                break
        
        results.append(audit_conclusion)

    # 6. Save results to a new Excel
    df['AI_Audit_Conclusion'] = results
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    
    return send_file(output, download_name="audit_results.xlsx", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
