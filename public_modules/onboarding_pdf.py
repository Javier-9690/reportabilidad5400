import base64
import io
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import cm
from xml.sax.saxutils import escape

DECLARACION = [
    'Declaro que conozco y acepto los siguientes puntos del Reglamento del Campamento 5400:',
    '1. Las reservas y uso de alojamiento deben realizarse con anticipación y es obligatorio portar la credencial.',
    '2. Las habitaciones son de uso exclusivo, deben mantenerse limpias y está prohibido instalar artefactos eléctricos no autorizados.',
    '3. Está prohibido fumar en áreas no habilitadas, portar armas, encender fuego o manipular sistemas de emergencia.',
    '4. En el casino debo respetar los horarios, registrar mi consumo y no retirar alimentos.',
    '5. Las instalaciones recreativas (minimarket, PUB, gimnasio, multicanchas) tienen horarios definidos y deben usarse responsablemente.',
    '6. En caso de emergencia debo evacuar hacia los Puntos de Encuentro de Emergencia y seguir instrucciones.',
    '7. Existe tolerancia cero al acoso sexual; toda conducta inapropiada será investigada y sancionada.',
    '8. Debo presentarme con vestimenta y calzado deportivo adecuados, mantener conducta respetuosa (sin ofensas ni gritos obscenos), usar y devolver con cuidado los implementos, cambiarme en los baños y no en la cancha, y entiendo que la actividad puede ser suspendida inmediatamente si no cumplo estas normas.',
    'Con mi firma confirmo que he leído, entiendo y acepto estas normas básicas de convivencia y seguridad en el Campamento 5400.'
]

def build_pdf(nombre, rut, empresa, user_id, fecha, signature_bytes):
    if len(signature_bytes) > 400000:
        raise ValueError('La firma excede el tamaño permitido.')
    out = io.BytesIO()
    doc = SimpleDocTemplate(out, pagesize=A4, leftMargin=1.8*cm, rightMargin=1.8*cm, topMargin=1.7*cm, bottomMargin=1.7*cm)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Title5400', parent=styles['Heading1'], alignment=TA_CENTER, fontSize=15, leading=19, spaceAfter=16))
    styles.add(ParagraphStyle(name='Text5400', parent=styles['BodyText'], fontSize=9.5, leading=13, spaceAfter=8))
    story = [Paragraph('Registro de firma de Onboarding — Campamento 5400', styles['Title5400']),
             Paragraph('<b>DECLARACIÓN DE CONOCIMIENTO Y ACEPTACIÓN DEL REGLAMENTO</b>', styles['Text5400'])]
    for line in DECLARACION:
        story.append(Paragraph(escape(line), styles['Text5400']))
    data = [['Nombre', escape(nombre)], ['RUT', escape(rut)], ['Empresa', escape(empresa)], ['ID', escape(user_id)], ['Fecha y hora', fecha.strftime('%d/%m/%Y %H:%M:%S')]]
    grid = Table([[Paragraph(str(a), styles['Text5400']), Paragraph(str(b), styles['Text5400'])] for a,b in data], colWidths=[3.2*cm, 13*cm])
    grid.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('LINEBELOW',(0,0),(-1,-1),0.3,colors.lightgrey)]))
    story.extend([Spacer(1,10), grid, Spacer(1,14), Paragraph('<b>Firma electrónica manuscrita:</b>',styles['Text5400']), Image(io.BytesIO(signature_bytes), width=11*cm, height=4*cm)])
    doc.build(story)
    return out.getvalue()
