from flask import Flask, request, render_template
from flask.helpers import send_file
import tabula
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows
import os

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Check if a file was uploaded
        if 'file' not in request.files:
            return render_template('index.html', error='No file part')

        file = request.files['file']

        # Check if the file has a name
        if file.filename == '':
            return render_template('index.html', error='No selected file')

        # Check if the file is a PDF
        if file.filename.endswith('.pdf'):
            pdf_file_path = os.path.join('uploads', file.filename)
            file.save(pdf_file_path)

            # Extract tables from the PDF
            tables = tabula.read_pdf(pdf_file_path, pages='all', multiple_tables=True)

            # Create an Excel workbook
            workbook = Workbook()
            default_sheet = workbook.active
            workbook.remove(default_sheet)

            # Create a Font object for bold headers
            bold_font = Font(bold=True)

            # Add each DataFrame to a separate sheet
            for i, df in enumerate(tables):
                sheet_name = f'Table{i + 1}'
                sheet = workbook.create_sheet(title=sheet_name)

                for row in dataframe_to_rows(df, index=False, header=True):
                    sheet.append(row)

                # Set the header row to bold
                for cell in sheet[1]:
                    cell.font = bold_font

            # Save the workbook
            excel_file_path = os.path.join('uploads', 'extracted_tables.xlsx')
            workbook.save(excel_file_path)

            # Download the Excel file
            response = send_file(excel_file_path, as_attachment=True)

            return response

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)

