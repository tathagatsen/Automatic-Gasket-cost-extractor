import os
from flask import Flask, render_template, request, send_file
import openpyxl
from fractions import Fraction
import io

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Define the path to the second Excel file
REFERENCE_FILE_PATH = 'static/cost/BOQ.xlsx'

def getInch(inch):
    res = ""
    check = 0
    for i in inch:
        if i.isdigit():
            res += i
        if i == '.':
            check = 1
            res += i
        if i == '/':
            res += i
        if i.isalpha():
            break
    if check == 1:
        return Fraction(float(res))
    return Fraction(res)

def getRating(rate):
    res = []
    for char in rate:
        if char.isdigit():
            res.append(char)
    return int("".join(res))

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['excel_file']
        column1 = request.form['column1'].upper()
        column2 = request.form['column2'].upper()
        column3 = request.form['column3'].upper()

        # Ensure the uploads directory exists
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)

        # Save the uploaded file temporarily
        file_path = os.path.join(UPLOAD_FOLDER, 'uploaded_file.xlsx')
        file.save(file_path)

        # Process the Excel file
        modified_file = process_excel(file_path, column1, column2, column3)

        # Send the modified file as a response
        return send_file(modified_file, as_attachment=True, download_name="modified_file.xlsx")

    return render_template('index.html')


def process_excel(file_path, column1, column2, column3):
    # Load workbooks
    db1 = openpyxl.load_workbook(file_path)
    db2 = openpyxl.load_workbook(REFERENCE_FILE_PATH, data_only=True)

    # Select worksheets
    ws1 = db1["Table 1"]
    ws2 = db2["CS 0.5 TO 24"]

    # Read data from ws1
    Desc = [ws1[f"C{i}"].value for i in range(4, 19)]
    Desc_M = [desc.split(',') for desc in Desc]

    Inches = []
    Rating = []
    MOC = []
    for i in range(len(Desc_M)):
        Inches.append(getInch(Desc_M[i][4]))
        Rating.append(getRating(Desc_M[i][5]))
        MOC.append(Desc_M[i][2].replace("L", "-SS316/FG-SS316"))

    # Read data from ws2
    Inches_check = [ws2[f"B{i}"].value for i in range(47, 621)]
    Rating_check = [ws2[f"A{i}"].value for i in range(47, 621)]

    # Perform matching and updating logic
    row1 = ws1.max_row
    row2 = ws2.max_row
    for i in range(len(Inches)):
        for j in range(3, row2 + 1):
            if Inches[i] == ws2[f"C{j}"].value and Rating[i] == ws2[f"A{j}"].value and (MOC[i] in ws2[f"D{j}"].value):
                ws1[f"I{4 + i}"].value = ws2[f"Z{j}"].value
                print(Inches[i], Rating[i], ws2[f"Z{j}"].value)

    # Save the modified workbook to a BytesIO object
    modified_file = io.BytesIO()
    db1.save(modified_file)
    modified_file.seek(0)

    return modified_file

if __name__ == '__main__':
    app.run(debug=True)
