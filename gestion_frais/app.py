from flask import Flask, render_template, request, send_file, redirect, url_for
from flask_sqlalchemy import SQLAlchemy
import openpyxl
from io import BytesIO
from datetime import datetime
import os

app = Flask(__name__)

# Configuration de PostgreSQL
db_url = os.environ.get('DATABASE_URL', 'sqlite:///local.db')
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)

app.config['SQLALCHEMY_DATABASE_URI'] = db_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# Modèle de la table Paiement
class Paiement(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    date_heure = db.Column(db.String(50), nullable=False)
    nom_eleve = db.Column(db.String(100), nullable=False)
    classe = db.Column(db.String(50), nullable=False)
    type_frais = db.Column(db.String(50), nullable=False)
    montant = db.Column(db.Float, nullable=False)

with app.app_context():
    db.create_all()

@app.route('/', methods=['GET', 'POST'])
def index():
    donnees_recu = None
    if request.method == 'POST':
        nom = request.form.get('nom_eleve') or request.form.get('nom')
        classe = request.form.get('classe')
        type_frais = request.form.get('type_frais')
        montant = float(request.form.get('montant', 0))
        date_heure = datetime.now().strftime("%d/%m/%Y %H:%M")

        # Sauvegarde en base de données PostgreSQL
        nouveau_paiement = Paiement(
            date_heure=date_heure,
            nom_eleve=nom,
            classe=classe,
            type_frais=type_frais,
            montant=montant
        )
        db.session.add(nouveau_paiement)
        db.session.commit()

        donnees_recu = {
            'nom': nom,
            'classe': classe,
            'type_frais': type_frais,
            'montant': montant,
            'date_heure': date_heure
        }

    return render_template('index.html', recu=donnees_recu)

# Route pour consulter l'historique directement en ligne
@app.route('/historique')
def historique():
    liste_paiements = Paiement.query.order_by(Paiement.id.desc()).all()
    return render_template('historique.html', paiements=liste_paiements)

# Route pour télécharger le fichier Excel généré dynamiquement depuis PostgreSQL
@app.route('/download')
def download():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Historique Paiements"
    ws.append(["ID", "Date & Heure", "Nom de l'élève", "Classe", "Type de Frais", "Montant (FC)"])

    paiements = Paiement.query.all()
    for p in paiements:
        ws.append([p.id, p.date_heure, p.nom_eleve, p.classe, p.type_frais, p.montant])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="historique_paiements.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == '__main__':
    app.run()
