from flask import Flask, render_template, request, send_file
from flask_sqlalchemy import SQLAlchemy
import openpyxl
from io import BytesIO
from datetime import datetime
import os

app = Flask(__name__)

# Gestion sécurisée de la connexion à PostgreSQL / SQLite
db_url = os.environ.get('DATABASE_URL')
if db_url:
    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql://", 1)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    db_url = f"sqlite:///{os.path.join(BASE_DIR, 'paiements.db')}"

app.config['SQLALCHEMY_DATABASE_URI'] = db_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# Modèle de la table des Paiements
class Paiement(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    date_heure = db.Column(db.String(50), nullable=False)
    nom_eleve = db.Column(db.String(100), nullable=False)
    classe = db.Column(db.String(50), nullable=False)
    type_frais = db.Column(db.String(50), nullable=False)
    montant = db.Column(db.Float, nullable=False)

# Création des tables automatiques au lancement
with app.app_context():
    try:
        db.create_all()
    except Exception as e:
        print("Erreur initialisation DB:", e)

@app.route('/', methods=['GET', 'POST'])
def index():
    donnees_recu = None
    if request.method == 'POST':
        nom = request.form.get('nom_eleve') or request.form.get('nom')
        classe = request.form.get('classe')
        type_frais = request.form.get('type_frais')
        montant = float(request.form.get('montant', 0))
        date_heure = datetime.now().strftime("%d/%m/%Y %H:%M")

        # Enregistrement dans la base de données
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

@app.route('/historique')
def historique():
    try:
        liste_paiements = Paiement.query.order_by(Paiement.id.desc()).all()
    except Exception:
        liste_paiements = []
    return render_template('historique.html', paiements=liste_paiements)

@app.route('/download')
def download():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Historique Paiements"
    ws.append(["N°", "Date & Heure", "Nom de l'élève", "Classe", "Type de Frais", "Montant (FC)"])

    try:
        paiements = Paiement.query.all()
        for p in paiements:
            ws.append([p.id, p.date_heure, p.nom_eleve, p.classe, p.type_frais, p.montant])
    except Exception:
        pass

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
