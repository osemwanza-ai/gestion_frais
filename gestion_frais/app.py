from flask import Flask, render_template, request, send_file
import openpyxl
from datetime import datetime
import os

app = Flask(__name__)

EXCEL_FILE = "paiements.xlsx"

@app.route('/')
def index():
    return render_template('index.html')

if __name__ == '__main__':
    app.run()
