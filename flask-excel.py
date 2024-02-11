#!/usr/bin/env python3

from flask import Flask, request, send_file
import xlsxwriter
import os
from datetime import datetime

now = datetime.now()
timestamp = now.strftime("%Y%m%d_%H%M%S")


app = Flask(__name__)

@app.route('/')
def home():
    return 'Welcome to the Spreadsheet Creator API! Use /createSpreadsheet to create a spreadsheet.'

@app.route('/createSpreadsheet', methods=['GET'])
def create_spreadsheet():
    # Read the 'columns' query parameter
    columns = int(request.args.get('columns', 0))
    
    # Define the spreadsheet filename
    filename = "spreadsheet.xlsx"
    
    # Create a new Excel file and add a worksheet
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    
    column_titles = ['Account', 'Product', 'Area']
    
    # Write column titles to the first row of the spreadsheet
    for col, title in enumerate(column_titles):
        worksheet.write(0, col, title)
        
    # Close the workbook to finalize the Excel file
    workbook.close()
    
    # Serve the created Excel file
    return send_file(filename, as_attachment=True, download_name=filename)

if __name__ == '__main__':
    app.run(debug=True)
    