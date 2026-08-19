from flask import Flask, render_template, request, send_file
import openpyxl
from datetime import datetime
import os

app = Flask(__name__)

EXCEL_FILE = "paiements.xlsx"

def init_excel():
    """ Crée le fichier Excel avec les entêtes s'il n'existe pas encore """
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Paiements"
        ws.append(["Date & Heure", "Nom de l'élève", "Classe", "Type de Frais", "Montant (FC)"])
        wb.save(EXCEL_FILE)

@app.route('/', methods=['GET', 'POST'])
def index():
    init_excel()
    donnees_recu = None
    
    if request.method == 'POST':
        nom = request.form.get('nom_eleve') or request.form.get('nom')
        classe = request.form.get('classe')
        type_frais = request.form.get('type_frais')
        montant = request.form.get('montant', 0)
        date_heure = datetime.now().strftime("%d/%m/%Y %H:%M")

        # 1. Ajout automatique d'une ligne dans le fichier Excel
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Paiements"]
        ws.append([date_heure, nom, classe, type_frais, montant])
        wb.save(EXCEL_FILE)

        # 2. Préparation des données pour le ticket d'impression
        donnees_recu = {
            'nom': nom,
            'classe': classe,
            'type_frais': type_frais,
            'montant': montant,
            'date_heure': date_heure
        }

    return render_template('index.html', recu=donnees_recu)

@app.route('/download')
def download():
    init_excel()
    return send_file(
        EXCEL_FILE,
        as_attachment=True,
        download_name="historique_paiements.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == '__main__':
    app.run()
