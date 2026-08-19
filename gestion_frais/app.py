from flask import Flask, render_template, request, send_file
import openpyxl
from datetime import datetime
import os

app = Flask(__name__)

EXCEL_FILE = "paiements.xlsx"

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Récupération des données du formulaire
        nom = request.form.get('nom_eleve') or request.form.get('nom')
        classe = request.form.get('classe')
        frais_scolaires = request.form.get('frais_scolaires')
        frais_inscription = request.form.get('frais_inscription')
        
        # Traitement/Enregistrement des données ici...
        
    return render_template('index.html')

if __name__ == '__main__':
    app.run()
