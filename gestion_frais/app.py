from flask import Flask, render_template, request, send_file
import openpyxl
from datetime import datetime
import os

app = Flask(__name__)

EXCEL_FILE = "paiements.xlsx"

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Paiements"
        ws.append(["Date & Heure", "Nom de l'élève", "Classe", "Frais Scolaires (FC)", "Frais Inscription (FC)"])
        wb.save(EXCEL_FILE)

@app.route('/', methods=['GET', 'POST'])
def index():
    init_excel()
    donnees_recu = None
    
    if request.method == 'POST':
        nom = request.form.get('nom_eleve') or request.form.get('nom')
        classe = request.form.get('classe')
        frais_scolaires = request.form.get('frais_scolaires', 0)
        frais_inscription = request.form.get('frais_inscription', 0)
        date_heure = datetime.now().strftime("%d/%m/%Y %H:%M")

        # Enregistrement dans le fichier Excel
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Paiements"]
        ws.append([date_heure, nom, classe, frais_scolaires, frais_inscription])
        wb.save(EXCEL_FILE)

        # Transmettre les données à la vue pour affichage/impression
        donnees_recu = {
            'nom': nom,
            'classe': classe,
            'frais_scolaires': frais_scolaires,
            'frais_inscription': frais_inscription,
            'date_heure': date_heure
        }

    return render_template('index.html', recu=donnees_recu)

@app.route('/download')
def download():
    init_excel()
    return send_file(EXCEL_FILE, as_attachment=True)

if __name__ == '__main__':
    app.run()
