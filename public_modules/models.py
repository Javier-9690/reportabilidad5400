"""Tablas nuevas y aisladas: nunca eliminan ni modifican tablas previas."""
from public_modules.shared_db import db

class SurveyResponse(db.Model):
    __tablename__ = 'public_survey_5400'
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.DateTime, nullable=False, index=True)
    q1_puntaje = db.Column(db.Integer, nullable=False)
    q2_puntaje = db.Column(db.Integer, nullable=False)
    q3_puntaje = db.Column(db.Integer, nullable=False)
    q4_puntaje = db.Column(db.Integer, nullable=False)
    q5_puntaje = db.Column(db.Integer, nullable=False)
    comentarios = db.Column(db.Text, nullable=False, default='')
    source = db.Column(db.String(32), nullable=False, default='web')

class Feedback(db.Model):
    __tablename__ = 'public_feedback_5400'
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.DateTime, nullable=False, index=True)
    nombre = db.Column(db.String(180), nullable=False)
    empresa = db.Column(db.String(180), nullable=False, default='')
    user_id = db.Column(db.String(100), nullable=False, default='')
    telefono = db.Column(db.String(30), nullable=False, default='')
    modulo = db.Column(db.String(100), nullable=False, default='')
    habitacion = db.Column(db.String(100), nullable=False, default='')
    detalle = db.Column(db.Text, nullable=False)

class Onboarding(db.Model):
    __tablename__ = 'public_onboarding_5400'
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.DateTime, nullable=False, index=True)
    nombre = db.Column(db.String(180), nullable=False)
    rut = db.Column(db.String(20), nullable=False)
    empresa = db.Column(db.String(180), nullable=False)
    user_id = db.Column(db.String(100), nullable=False)
    aceptado = db.Column(db.Boolean, nullable=False, default=True)
    version_reglamento = db.Column(db.String(35), nullable=False, default='2026')
    comprobante_pdf = db.deferred(db.Column(db.LargeBinary, nullable=False))
