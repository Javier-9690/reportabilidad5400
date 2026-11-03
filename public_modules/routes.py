"""Formularios públicos y módulo privado de encuestas y onboarding."""
import base64
import csv
import io
import os
import secrets
import hmac
from datetime import datetime
from zoneinfo import ZoneInfo
from flask import Blueprint, abort, flash, jsonify, redirect, render_template, request, session, url_for, Response, send_file
from sqlalchemy import and_
from werkzeug.exceptions import RequestEntityTooLarge
from public_modules.models import SurveyResponse, Feedback, Onboarding
from public_modules.shared_db import db
from public_modules.onboarding_pdf import build_pdf
from public_modules import survey_engine as eng

public = Blueprint('public5400', __name__)
admin = Blueprint('admin5400', __name__)
SANTIAGO = ZoneInfo('America/Santiago')
QUESTIONS = [
  'Amabilidad y cordialidad',
  'Tiempo de espera para ser atendido',
  'La información entregada por el personal fue clara y comprensible',
  'Presentación personal del equipo de Custodia',
  'Indíquenos su satisfacción general con el servicio de Custodia 5400.'
]
LABELS = {1:'Muy insatisfecho',2:'Insatisfecho',3:'Neutral',4:'Satisfecho',5:'Muy satisfecho'}

def now_local():
    return datetime.now(SANTIAGO).replace(tzinfo=None)

def new_token(name):
    token=secrets.token_urlsafe(32)
    session[name] = token
    return token

def verify_token(name):
    token=session.get(name,'')
    sent=request.form.get('csrf_token','')
    if not token or not sent or not hmac.compare_digest(token,sent):
        abort(400,'La sesión del formulario expiró; recarga la página.')
    session.pop(name,None)

def fields_valid():
    if request.form.get('website','').strip():
        abort(400)

def checked_text(key, max_size=200, required=False):
    value=request.form.get(key,'').strip()
    if len(value)>max_size or (required and not value):
        abort(400,f'Campo {key} inválido.')
    return value

def admin_only():
    if not session.get('r5400_admin_authenticated'):
        abort(401,'Debes iniciar sesión como administrador.')

def csv_response(filename, headers, rows):
    output=io.StringIO(); writer=csv.writer(output);writer.writerow(headers);writer.writerows(rows)
    return Response('\ufeff'+output.getvalue(),mimetype='text/csv',headers={'Content-Disposition':f'attachment; filename="{filename}"','Cache-Control':'no-store'})

@public.get('/encuesta')
def encuesta_form():
    return render_template('public5400/encuesta.html',csrf_token=new_token('survey_csrf'), questions=QUESTIONS,labels=LABELS)

@public.post('/encuesta')
def encuesta_save():
    verify_token('survey_csrf');fields_valid()
    try:
        scores=[int(request.form.get(f'q{i}','')) for i in range(1,6)]
    except (ValueError,TypeError):
        abort(400,'Debes responder las cinco preguntas.')
    if not all(1<=x<=5 for x in scores):abort(400,'Respuesta fuera de rango.')
    comments=checked_text('comentarios',2000)
    rec=SurveyResponse(fecha=now_local(),comentarios=comments,**{f'q{i}_puntaje':scores[i-1] for i in range(1,6)})
    db.session.add(rec);db.session.commit()
    return render_template('public5400/gracias.html',tipo='encuesta')

@public.get('/sugerencias')
def feedback_form():
    return render_template('public5400/feedback.html',csrf_token=new_token('feedback_csrf'))

@public.post('/sugerencias')
def feedback_save():
    verify_token('feedback_csrf');fields_valid()
    rec=Feedback(fecha=now_local(),nombre=checked_text('nombre',180,True), empresa=checked_text('empresa',180),user_id=checked_text('user_id',100),telefono=checked_text('telefono',30),modulo=checked_text('modulo',100),habitacion=checked_text('habitacion',100),detalle=checked_text('detalle',5000,True))
    db.session.add(rec);db.session.commit()
    return render_template('public5400/gracias.html',tipo='sugerencia o reclamo')

@public.get('/onboarding')
def onboarding_form():
    return render_template('public5400/onboarding.html',csrf_token=new_token('onboarding_csrf'))

@public.post('/onboarding')
def onboarding_save():
    verify_token('onboarding_csrf'); fields_valid()
    if request.form.get('aceptado')!='1':abort(400,'Debes aceptar el reglamento.')
    nombre=checked_text('nombre',180,True);rut=checked_text('rut',20,True)
    empresa=checked_text('empresa',180,True);user_id=checked_text('user_id',100,True)
    raw=request.form.get('firma','')
    if not raw.startswith('data:image/png;base64,') or len(raw)>550000:
        abort(400,'Firma obligatoria o inválida.')
    try:signature=base64.b64decode(raw.split(',',1)[1],validate=True)
    except Exception:abort(400,'Firma inválida.')
    if not signature.startswith(b'\x89PNG\r\n\x1a\n') or len(signature)<100 or len(signature)>400000:
        abort(400,'Firma inválida.')
    fecha=now_local()
    pdf=build_pdf(nombre,rut,empresa,user_id,fecha,signature)
    rec=Onboarding(fecha=fecha,nombre=nombre,rut=rut,empresa=empresa,user_id=user_id,comprobante_pdf=pdf)
    db.session.add(rec);db.session.commit()
    return render_template('public5400/gracias.html',tipo='onboarding')

# Las funciones de reportabilidad se conservan del proyecto original Encuesta 5400.
def report_rows():
    rows=SurveyResponse.query.order_by(SurveyResponse.fecha.asc(),SurveyResponse.id.asc()).all()
    result=[]
    for rec in rows:
        scores=[getattr(rec,f'q{i}_puntaje') for i in range(1,6)]
        obj={'id':rec.id,'fecha':rec.fecha.strftime('%Y-%m-%d %H:%M:%S'), 'fecha_dt':rec.fecha,
             'total':sum(scores),'promedio':sum(scores)/5,'comentarios':rec.comentarios or '',
             'imported_at':rec.fecha.strftime('%Y-%m-%d %H:%M:%S')}
        for i,n in enumerate(scores,1):obj[f'q{i}_puntaje']=n;obj[f'q{i}_respuesta']=LABELS[n]
        result.append(obj)
    return result

def metrics_and_filter():
    data=report_rows(); closed=eng.filter_data_through_closed_week(data)
    start=(request.args.get('fecha_inicio') or '').strip();end=(request.args.get('fecha_fin') or '').strip()
    active=bool(start or end)
    filtered=eng.filter_data_by_dates(closed,start,end) if active else []
    return data,closed,eng.build_metrics(closed),start,end,active,filtered,eng.build_metrics(filtered,include_current_closed_week=False) if active else eng.build_metrics([])

@admin.get('/encuestas')
def dashboard():
    admin_only()
    data,closed,metrics,start,end,active,filtered,fmetrics=metrics_and_filter()
    return render_template('encuesta_admin/dashboard.html',title=eng.APP_TITLE,metrics=metrics,filtered_metrics=fmetrics,filter_start=start,filter_end=end,filter_active=active,filter_label=eng.build_filter_label(start,end),status_class=eng.status_class,json_dumps=eng.json.dumps,import_csrf=new_token('survey_import_csrf'))

@admin.post('/encuestas/importar')
def importar():
    admin_only(); verify_token('survey_import_csrf')
    f=request.files.get('archivo')
    if not f or not f.filename:abort(400,'Selecciona un archivo .xlsx.')
    rows=eng.read_uploaded_workbook(f)
    for row in rows:
        dt=eng.parse_date(row['fecha'])
        if dt is None:continue
        scores=[row.get(f'q{i}_puntaje') for i in range(1,6)]
        if any(x is None or not 1<=float(x)<=5 for x in scores):continue
        db.session.add(SurveyResponse(fecha=dt,comentarios=row.get('comentarios',''),source='excel',**{f'q{i}_puntaje':int(round(scores[i-1])) for i in range(1,6)}))
    db.session.commit()
    flash('Encuestas importadas sin borrar registros anteriores.','success')
    return redirect('/encuestas')

@admin.get('/encuestas/plantilla')
def plantilla():
    admin_only()
    return send_file(eng.build_template_buffer(),as_attachment=True,download_name='plantilla_encuesta_5400.xlsx')

@admin.get('/encuestas/exportar/csv')
def exportar_csv():
    admin_only()
    return Response(eng.export_csv_buffer(report_rows()),mimetype='text/csv',headers={'Content-Disposition':'attachment; filename="encuestas_5400.csv"'})

@admin.get('/encuestas/exportar/excel')
def exportar_excel():
    admin_only()
    data,closed,metrics,start,end,active,filtered,fmetrics=metrics_and_filter()
    output=eng.build_excel_buffer(closed,metrics,filtered if active else None,fmetrics if active else None,eng.build_filter_label(start,end) if active else None)
    return send_file(output,as_attachment=True,download_name='reporte_encuestas_5400.xlsx')

@admin.get('/encuestas/exportar/pdf')
def exportar_pdf():
    admin_only()
    data,closed,metrics,start,end,active,filtered,fmetrics=metrics_and_filter()
    if not metrics['has_data']:abort(404,'No hay semanas cerradas para reportar.')
    pdf=eng.generate_pdf_report(metrics,fmetrics if active else None,eng.build_filter_label(start,end) if active else None)
    return send_file(pdf if hasattr(pdf,'read') else io.BytesIO(pdf),mimetype='application/pdf',as_attachment=True,download_name='reporte_encuestas_5400.pdf')

@admin.get('/encuestas/registros')
def encuesta_records():
    admin_only()
    start=eng.parse_filter_date(request.args.get('fecha_inicio'));end=eng.parse_filter_date(request.args.get('fecha_fin'))
    records=SurveyResponse.query.order_by(SurveyResponse.fecha.desc()).limit(1000).all()
    records=[r for r in records if (not start or r.fecha.date()>=start) and (not end or r.fecha.date()<=end)]
    return render_template('public5400/admin_records.html',title='Registros de encuestas',kind='encuesta',records=records)

@admin.get('/sugerencias/registros')
def feedback_records():
    admin_only()
    return render_template('public5400/admin_records.html',title='Sugerencias y reclamos',kind='feedback',records=Feedback.query.order_by(Feedback.fecha.desc()).limit(1000).all())

@admin.get('/sugerencias/exportar/csv')
def feedback_csv():
    admin_only()
    return csv_response('sugerencias_reclamos_5400.csv',['FECHA_HORA','NOMBRE','EMPRESA','ID','TELEFONO','MODULO','HABITACION','DETALLE'],[(x.fecha,x.nombre,x.empresa,x.user_id,x.telefono,x.modulo,x.habitacion,x.detalle) for x in Feedback.query.order_by(Feedback.fecha).all()])

@admin.get('/onboarding/registros')
def onboarding_records():
    admin_only()
    return render_template('public5400/admin_records.html',title='Onboarding - registros',kind='onboarding',records=Onboarding.query.order_by(Onboarding.fecha.desc()).limit(1000).all())

@admin.get('/onboarding/exportar/csv')
def onboarding_csv():
    admin_only()
    return csv_response('onboarding_registros_5400.csv',['FECHA_HORA','NOMBRE','RUT','EMPRESA','ID','ACEPTADO'],[(x.fecha,x.nombre,x.rut,x.empresa,x.user_id,'SI' if x.aceptado else 'NO') for x in Onboarding.query.order_by(Onboarding.fecha).all()])

@admin.get('/onboarding/pdf/<int:record_id>')
def onboarding_pdf(record_id):
    admin_only();record=Onboarding.query.get_or_404(record_id)
    return send_file(io.BytesIO(record.comprobante_pdf),mimetype='application/pdf',as_attachment=True,download_name=f'onboarding_{record.id}.pdf')

# Eliminación protegida con la misma validación del sistema (security.destructive_path).
@admin.post('/encuestas/registros/<int:record_id>/eliminar')
def encuesta_delete(record_id):
    admin_only();item=SurveyResponse.query.get_or_404(record_id);db.session.delete(item);db.session.commit();return redirect('/encuestas/registros')

@admin.post('/sugerencias/registros/<int:record_id>/eliminar')
def feedback_delete(record_id):
    admin_only();item=Feedback.query.get_or_404(record_id);db.session.delete(item);db.session.commit();return redirect('/sugerencias/registros')

@admin.post('/onboarding/registros/<int:record_id>/eliminar')
def onboarding_delete(record_id):
    admin_only();item=Onboarding.query.get_or_404(record_id);db.session.delete(item);db.session.commit();return redirect('/onboarding/registros')

