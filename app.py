import os
import io
from flask import Flask, render_template, request, send_file
import openpyxl


app = Flask(__name__)

# Load the reference file into memory at startup
REFERENCE_FILE_PATH = os.path.join(os.getcwd(), "public", "COSTING SHEET MASTER NEW.xlsx")

REFERENCE_WORKBOOK = openpyxl.load_workbook(REFERENCE_FILE_PATH, data_only=True)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['excel_file']
        column1 = request.form['column1'].upper()
        column2 = request.form['column2'].upper()
        column3 = request.form['column3'].upper()
        column4 = request.form['column4'].upper()
        # Read the uploaded file into memory
        file_stream = io.BytesIO(file.read())

        # Process the Excel file
        modified_file = process_excel(file_stream, column1, column2, column3,column4)

        # Send the modified file as a response
        return send_file(modified_file, as_attachment=True, download_name="modified_file.xlsx")

    return render_template('index.html')

def process_excel(file_stream, column1, column2, column3,column4):
    # Load workbooks from memory
    db1 = openpyxl.load_workbook(file_stream)
    db2 = REFERENCE_WORKBOOK  # Use preloaded reference file

    # Select worksheets
    ws1 = db1["Table 1"]
    ws2 = db2["CS 0.5 TO 24"]

    col_C = "C"
    col_A =  "A"
    col_D = "D"
    col_Z = "Z"
    
    # Convert column letters to indexes
    c1 = ord(column1.upper()) - 64
    c2 = ord(column2.upper()) - 64
    c3 = ord(column3.upper()) - 64
    c4 = ord(column4.upper()) - 64
    
    c5 = ord(col_C) - 64
    c6 = ord(col_A) - 64
    c7 = ord(col_D) - 64
    c8 = ord(col_Z) - 64

    
    # Save the modified workbook to memory
    modified_file = io.BytesIO()
    db1.save(modified_file)
    modified_file.seek(0)

    return modified_file

if __name__ == '__main__':
    app.run(debug=True)
