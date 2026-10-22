from flask import Flask, request, jsonify, send_file, render_template
import json
import os
from flask_cors import CORS
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import io
from reportlab.lib.pagesizes import A4, A5
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.utils import simpleXML
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from io import BytesIO
import secrets

app = Flask(__name__)
CORS(app)

# Configurazione database
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///iscritti.db'
app.config['SQLALCHEMY_TRACK_NOTIFICATIONS'] = False
app.secret_key = secrets.token_hex(32)
db = SQLAlchemy(app)

# Gestione sessioni utente
class UserSession:
    _sessions = {}
    
    @classmethod
    def create(cls, username):
        session_id = secrets.token_hex(16)
        cls._sessions[session_id] = {'username': username, 'admin': True}
        return session_id
    
    @classmethod
    def check(cls, session_id):
        return cls._sessions.get(session_id)
    
    @classmethod
    def invalidate(cls, session_id):
        cls._sessions.pop(session_id, None)

# Modell di database
class Iscritto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(100))
    cognome = db.Column(db.String(100))
    data_nascita = db.Column(db.String(10))
    luogo_nascita = db.Column(db.String(50))
    residenza = db.Column(db.String(200))
    codice_fiscale = db.Column(db.String(16))
    numero_tessera = db.Column(db.String(50))
    data_iscrizione = db.Column(db.DateTime, default=datetime.utcnow)
    firmato = db.Column(db.Boolean, default=False)
    data_firma = db.Column(db.DateTime, nullable=True)

# Crea database all'avvio
with app.app_context():
    db.create_all()

# API - Carica configurazione
@app.route('/api/config', methods=['GET'])
def get_config_api():
    try:
        config_path = os.path.join(os.path.dirname(__file__), 'config.json')
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        return jsonify(config)
    except Exception as e:
        return jsonify({'error': 'Errore nel caricamento configurazione'}), 500

# API - Salva configurazione
@app.route('/api/save-config', methods=['POST'])
def save_config_api():
    try:
        data = request.get_json()
        config_path = os.path.join(os.path.dirname(__file__), 'config.json')
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return jsonify({'success': True, 'message': 'Configurazione salvata con successo!'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# API - Login admin
@app.route('/api/auth', methods=['POST'])
def login_admin():
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')
    
    if username == "admin" and password == "admin123":
        session_id = UserSession.create(username)
        return jsonify({'success': True, 'session_id': session_id, 'user': {'username': username}})
    return jsonify({'success': False, 'error': 'Credenziali non valide'}), 401

# API - Verifica sessione
@app.route('/api/auth/check', methods=['POST'])
def check_session():
    data = request.get_json()
    session_id = data.get('session_id')
    
    session = UserSession.check(session_id)
    if session:
        return jsonify({'valid': True, 'user': session})
    return jsonify({'valid': False}), 401

# API - Logout
@app.route('/api/auth/logout', methods=['POST'])
def logout():
    data = request.get_json()
    session_id = data.get('session_id')
    UserSession.invalidate(session_id)
    return jsonify({'success': True})

# API - Crea nuova iscrizione
@app.route('/api/iscrizioni', methods=['POST'])
def crea_iscrizione():
    data = request.get_json()
    
    try:
        nuova_iscrizione = Iscritto(
            nome=data['nome'],
            cognome=data['cognome'],
            data_nascita=data['data_nascita'],
            luogo_nascita=data['luogo_nascita'],
            residenza=data['residenza'],
            codice_fiscale=data['codice_fiscale'],
            numero_tessera=data['numero_tessera']
        )
        
        db.session.add(nuova_iscrizione)
        db.session.commit()
        
        return jsonify({'success': True, 'id': nuova_iscrizione.id, 'message': 'Iscrizione creata con successo!'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'error': str(e), 'message': 'Errore nella creazione iscrizione'}), 400

# API - Ottieni tutte le iscrizioni
@app.route('/api/iscrizioni', methods=['GET'])
def get_iscrizioni():
    try:
        iscrizioni = db.session.query(Iscritto).order_by(Iscritto.data_iscrizione.desc()).all()
        data = []
        for i in iscrizioni:
            data.append({
                'id': i.id,
                'nome': i.nome,
                'cognome': i.cognome,
                'data': i.data_nascita,
                'luo': i.luogo_nascita,
                'res': i.residenza,
                'codice': i.codice_fiscale,
                'tessera': i.numero_tessera,
                'firma': i.firma,
                'data_iscrizione': i.data_iscrizione.strftime('%Y-%m-%d'),
                'data_firma': i.data_firma.strftime('%Y-%m-%d') if i.data_firma else None
            })
        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# API - Conferma firma iscrizione
@app.route('/api/iscrizioni/<int:id>/conferma', methods=['POST'])
def conferma_firma(id):
    session_data = request.get_json()
    session_id = session_data.get('session_id')
    
    session = UserSession.check(session_id)
    if not session or not session.get('admin'):
        return jsonify({'success': False, 'error': 'Accesso non autorizzato'}), 403
    
    try:
        iscrizione = Iscritto.query.get(id)
        if iscrizione:
            iscrizione.firma = True
            iscrizione.data_firma = datetime.utcnow()
            db.session.commit()
            return jsonify({'success': True, 'message': 'Firma confermata con successo!'})
        else:
            return jsonify({'success': False, 'error': 'Iscrizione non trovata'}), 404
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'error': str(e)}), 500

# API - Elimina iscrizione
@app.route('/api/iscrizioni/<int:id>', methods=['DELETE'])
def elimina_iscrizione(id):
    session_data = request.get_json()
    session_id = session_data.get('session_id')
    
    session = UserSession.check(session_id)
    if not session or not session.get('admin'):
        return jsonify({'success': False, 'error': 'Accesso non autorizzato'}), 403
    
    try:
        iscrizione = Iscritto.query.get(id)
        if iscrizione:
            db.session.delete(iscrizione)
            db.session.commit()
            return jsonify({'success': True, 'message': 'Iscrizione eliminata con successo!'})
        else:
            return jsonify({'success': False, 'error': 'Iscrizione non trovata'}), 404
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'error': str(e)}), 500

# API - Genera PDF (ricevuta A5)
@app.route('/api/pdf/<int:id>', methods=['GET'])
def genera_pdf(id):
    session_data = request.args.get('session_data', '{}')
    session_data = session_data if session_data else '{}'
    session_id = None
    try:
        session_data = eval(session_data)
        session_id = session_data.get('session_id')
    except:
        pass
    
    session = UserSession.check(session_id)
    if session and not session.get('admin'):
        return jsonify({'error': 'Accesso non autorizzato'}), 403
    
    try:
        iscrizione = Iscritto.query.get(id)
        if not iscrizione:
            return jsonify({'error': 'Iscrizione non trovata'}), 404
        
        buf = BytesIO()
        c = canvas.Canvas(buf, pagesize=A5)
        
        # Intestazione
        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, 700, "COMUNICAZIONE DI Ricezione")
        c.setFont("Helvetica", 12)
        c.drawString(50, 670, "Associazione")
        c.setFont("Helvetica", 11)
        c.drawString(50, 650, "Data: %s" % iscrizione.data_nascita)
        c.line(50, 620, 520, 620)
        c.setFont("Helvetica-Oblique", 11)
        c.drawString(50, 605, "Data emissione documento: %s" % datetime.utcnow().strftime('%d/%m/%Y %H:%M'))
        
        # Dati utente
        y = 570
        c.setFont("Helvetica-Bold", 12)
        c.drawString(50, y, "Spett.le")
        y -= 25
        c.setFont("Helvetica-Bold", 14)
        c.drawString(50, y, "Gentile %s %s," % (iscrizione.nome, iscrizione.cognome))
        y -= 40
        
        c.setFont("Helvetica", 11)
        c.setFont("Helvetica", 11)
        y -= 35
        c.drawString(50, y, "Con la presente si attesta la ricezione della")
        y -= 22
        c.drawString(50, y, "richiesta di iscrizione all\'Associazione")
        y -= 30
        c.setFont("Helvetica-Bold", 12)
        c.drawString(50, y, "N° Tessera Irreducibili: %s" % iscrizione.numero_tessera)
        y -= 50
        
        # Spazio per la tua firma (admin)
        y -= 55
        c.setFont("Helvetica", 10)
        c.setFont("Helvetica", 10)
        c.drawString(50, y, "Ricevuta di protocollo")
        y -= 20
        c.setFont("Helvetica", 10)
        c.drawString(150, y, "Data ricezione: ______")
        y -= 25
        c.setFont("Helvetica", 10)
        c.drawString(280, y, "Luogo: ______")
        y -= 25
        c.setFont("Helvetica", 10)
        c.drawString(360, y, "Firma del referente: ______")
        
        c.save()
        
        buf.seek(0)
        return send_file(buf, mimetype='application/pdf')
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

# API - Genera PDF A4 con modulo completo e spazio per firma
@app.route('/api/pdf/<int:id>/completo', methods=['GET'])
def genera_pdf_completo(id):
    session_data = request.args.get('session_data', '{}')
    session_id = None
    try:
        session_data = eval(session_data)
        session_id = session_data.get('session_id')
    except:
        pass
    
    session = UserSession.check(session_id)
    if session and not session.get('admin'):
        return jsonify({'error': 'Accesso non autorizzato'}), 403
    
    try:
        iscrizione = Iscritto.query.get(id)
        if not iscrizione:
            return jsonify({'error': 'Iscrizione non trovata'}), 404
        
        buf = BytesIO()
        c = canvas.Canvas(buf, pagesize=A4)
        
        # Intestazione
        c.setFont("Helvetica-Bold", 22)
        c.drawString(70, 750, "MODULO DI ISCRIZIONE ASSOCIAZIONE")
        c.setFont("Helvetica", 14)
        c.drawString(70, 720, "Associazione Italiana Irreducibili")
        c.setFont("Helvetica", 11)
        c.drawString(70, 695, "Via Roma 123, 00100 Città (RO)")
        c.setFont("Helvetica", 11)
        c.drawString(70, 675, "Tel: +39 010 1234567 - Email: segreteria@associazione.it")
        c.line(70, 660, 510, 660)
        c.setFont("Helvetica", 11)
        c.drawString(70, 640, "Data rilascio modulo: %s" % datetime.utcnow().strftime('%d/%m/%Y'))
        
        # Sezione 1: Dati dell'iscritto
        y = 600
        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, y, "1. DATI DELL'ISCRITTO")
        y -= 25
        c.line(50, y, 550, y)
        y -= 35
        
        c.setFont("Helvetica", 12)
        fields = [
            ("A) Cognome:", iscrizione.cognome),
            ("B) Nome:", iscrizione.nome),
            ("C) Data di nascita:", iscrizione.data_nascita),
            ("D) Luogo di nascita:", iscrizione.luogo_nascita),
            ("E) Residenza attuale:", iscrizione.residenza),
            ("F) Codice Fiscale:", iscrizione.codice_fiscale),
            ("G) N° Tessera Irreducibili:", iscrizione.numero_tessera)
        ]
        
        for label, value in fields:
            c.drawString(50, y, label)
            y -= 20
            c.drawString(340, y, value)
            y -= 35
        
        # Sezione 2: Dichiarazione
        y -= 15
        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, y, "2. DICHIARAZIONE")
        y -= 25
        c.line(50, y, 550, y)
        y -= 35
        
        c.setFont("Helvetica", 11)
        c.setFont("Helvetica-Oblique", 11)
        y -= 25
        c.drawString(50, y, "Io sottoscritto/a %s %s, nato/a a %s il %s," % (
            iscrizione.nome, iscrizione.cognome, iscrizione.luogo_nascita, iscrizione.data_nascita
        ))
        y -= 25
        c.drawString(50, y, "residente a %s," % iscrizione.residenza)
        y -= 25
        c.drawString(50, y, "dichiaro di essere a conoscenza e di accettare le seguenti condizioni:")
        y -= 25
        c.drawString(50, y, "1) Partecipazione attiva alle attività dell'Associazione")
        y -= 22
        c.drawString(50, y, "2) Rispetto del regolamento associativo")
        y -= 22
        c.drawString(50, y, "3) Pagamento del tesseramento annuale")
        y -= 22
        c.drawString(50, y, "4) Riconoscimento della tessera irreducibili sopra riportata")
        y -= 50
        
        # Sezione 3: Spazio per firme
        y -= 55
        c.setFont("Helvetica-Bold", 14)
        c.drawString(50, y, "3. FERMATA PER ACCETTAZIONE")
        y -= 30
        c.line(50, y, 500, y)
        y -= 40
        
        c.setFont("Helvetica-Oblique", 12)
        c.setFont("Helvetica-Oblique", 12)
        c.setFont("Helvetica-Oblique", 12)
        y -= 25
        c.drawString(50, y, "Data: ____/____/____")
        y -= 25
        c.drawString(50, y, "Firma dell'iscritto:")
        y -= 30
        c.drawString(130, y, "__________________________")
        y -= 55
        
        c.setFont("Helvetica", 12)
        c.setFont("Helvetica", 12)
        c.setFont("Helvetica", 12)
        y -= 25
        c.drawString(350, y, "Data: ____/____/____")
        y -= 25
        c.drawString(350, y, "Firma dell'Associazione:")
        y -= 30
        c.drawString(480, y, "__________________________")
        y -= 40
        c.setFont("Helvetica", 9)
        c.setFont("Helvetica", 9)
        c.setFont("Helvetica", 9)
        c.drawString(430, y, "Il presente modulo dovrà essere restituito firmato.")
        
        c.save()
        
        buf.seek(0)
        return send_file(buf, mimetype='application/pdf')
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

# API - Genera Excel con riepilogo
@app.route('/api/excel', methods=['GET'])
def genera_excel():
    session_data = request.args.get('session_data', '{}')
    session_id = None
    try:
        session_data = eval(session_data)
        session_id = session_data.get('session_id')
    except:
        pass
    
    session = UserSession.check(session_id)
    if session and not session.get('admin'):
        return jsonify({'error': 'Accesso non autorizzato'}), 403
    
    try:
        iscrizioni = db.session.query(Iscritto).order_by(Iscritto.data_iscrizione.desc()).all()
        
        wb = openpyxl.Workbook()
        ws_main = wb.active
        ws_main.title = "Elenco Iscrizioni"
        
        # Header
        headers = ['ID', 'Cognome', 'Nome', 'Data Nascita', 'Luogo Nascita', 
                  'Residenza', 'Codice Fiscale', 'Numero Tessera', 
                  'Data Iscrizione', 'Data Firma', 'Firmato']
        for col, header in enumerate(headers, 1):
            ws_main.cell(row=1, column=col, value=header)
            fill = PatternFill(start_color='673AB7', end_color='673AB7', fill_type='solid')
            ws_main.cell(row=1, column=col).fill = fill
            ws_main.cell(row=1, column=col).font = Font(bold=True, size=12, name='Arial')
            ws_main.cell(row=1, column=col).alignment = Alignment(horizontal='center')
        
        # Dati
        for row, iscrizione in enumerate(iscrizioni, 2):
            ws_main.cell(row=row, column=1, value=iscrizione.id)
            ws_main.cell(row=row, column=2, value=iscrizione.cognome)
            ws_main.cell(row=row, column=3, value=iscrizione.nome)
            ws_main.cell(row=row, column=4, value=iscrizione.data_nascita)
            ws_main.cell(row=row, column=5, value=iscrizione.luogo_nascita)
            ws_main.cell(row=row, column=6, value=iscrizione.residenza)
            ws_main.cell(row=row, column=7, value=iscrizione.codice_fiscale)
            ws_main.cell(row=row, column=8, value=iscrizione.numero_tessera)
            ws_main.cell(row=row, column=9, value=iscrizione.data_iscrizione.strftime('%d/%m/%Y'))
            ws_main.cell(row=row, column=11, value="Sì" if iscrizione.firma else "No")
            ws_main.cell(row=row, column=10, value=iscrizione.data_firma.strftime('%d/%m/%Y') if iscrizione.data_firma else "")
        
        # Stile per il nome
        for col in range(2, 9):
            ws_main.cell(row=1, column=col).font = Font(name='Arial', size=11, bold=False)
        for col in range(1, 11):
            ws_main.cell(row=1, column=col).alignment = Alignment(horizontal='center', vertical='center')
        
        ws_main.font = Font(name='Arial', size=11, bold=False)
        
        # Crea foglio statistiche
        ws_stat = wb.create_sheet("Statistiche")
        ws_stat['A1'] = "Riepilogo Iscrizioni"
        ws_stat['A2'] = datetime.utcnow().strftime('%d/%m/%Y - %H:%M')
        
        ws_stat['A4'] = "TOTALE ISCRITTI: %d" % len(iscrizioni)
        ws_stat['B4'] = "TOTALI: %d" % db.session.query(Iscritto).count()
        ws_stat['A5'] = "SCRITTI E VISTOSI: %d" % db.session.query(Iscritto).filter_by(firma=True).count()
        ws_stat['A6'] = "SCRITTI E NON VISTOSI: %d" % (len(iscrizioni) - db.session.query(Iscritto).filter_by(firma=True).count())
        
        # Colori e formattazione statistiche
        fill = PatternFill(start_color='4CAF50', end_color='4CAF50', fill_type='solid')
        fill_stat = PatternFill(start_color='2196F3', end_color='2196F3', fill_type='solid')
        fill_orange = PatternFill(start_color='ff9800', end_color='ff9800', fill_type='solid')
        
        fill_stat.apply(ws_stat.cell(row=4, column=1))
        fill_stat.apply(ws_stat.cell(row=4, column=2))
        fill_stat.apply(ws_stat.cell(row=5, column=2))
        fill_orange.apply(ws_stat.cell(row=6, column=2))
        
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        
        return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)