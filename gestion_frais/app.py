from flask import Flask, render_template, request, send_file
import openpyxl
import os
from datetime import datetime

app = Flask(__name__)

EXCEL_FILE = "paiements.xlsx"

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Historique Paiements"
        ws.append([
            "N° Reçu", 
            "Date / Heure", 
            "Élève", 
            "Classe", 
            "Frais Scolaires ($)", 
            "Frais Inscription ($)", 
            "Total ($)"
        ])
        wb.save(EXCEL_FILE)

def enregistrer_paiement(eleve, classe, frais_scolaires, frais_inscription):
    init_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active

    num_recu = f"CS-2026-{ws.max_row:03d}"
    date_jour = datetime.now().strftime("%d/%m/%Y %H:%M")
    
    fs = float(frais_scolaires) if frais_scolaires else 0.0
    fi = float(frais_inscription) if frais_inscription else 0.0
    total = fs + fi

    ws.append([num_recu, date_jour, eleve, classe, fs, fi, total])
    wb.save(EXCEL_FILE)

    return {
        "num_recu": num_recu,
        "date": date_jour,
        "eleve": eleve,
        "classe": classe,
        "frais_scolaires": f"{fs:.2f}",
        "frais_inscription": f"{fi:.2f}",
        "total": f"{total:.2f}"
    }

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        eleve = request.form.get('eleve')
        classe = request.form.get('classe')
        frais_scolaires = request.form.get('frais_scolaires', 0)
        frais_inscription = request.form.get('frais_inscription', 0)

        data = enregistrer_paiement(eleve, classe, frais_scolaires, frais_inscription)
        return render_template('recu.html', data=data)

    return render_template('index.html')

@app.route('/download-excel')
def download_excel():
    init_excel()
    return send_file(EXCEL_FILE, as_attachment=True)

if _name_ == '_main_':
    init_excel()
    app.run(host='0.0.0.0', port=5000, debug=True)
