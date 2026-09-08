"""Campos y validación para editar registros existentes de hotelería."""

import hashlib
import json
import math
import re
from datetime import date, datetime, time

from sqlalchemy import Date, DateTime, Float, Integer, Text, Time


# Lista explícita: el formulario nunca puede modificar la clave ni la creación.
EDIT_CONFIG = {
    "censo": ("Censo", "fecha censo_dia censo_noche total"),
    "eventos": ("Eventos de seguridad", "fecha horario que_ocurrio nombre_afectado accion"),
    "duplicidades": ("Duplicidades", "semana fecha id_interno empresa_contratista descripcion_problema tipo_riesgo pabellon habitacion ingresar_contacto nombre_usuario responsable estatus notificacion_usuario plan_accion fecha_cierre"),
    "encuestas": ("Encuesta de satisfacción", "fecha_hora q1_respuesta q1_puntaje q2_respuesta q2_puntaje q3_respuesta q3_puntaje q4_respuesta q4_puntaje q5_respuesta q5_puntaje comentarios"),
    "atencion": ("Atención al público", "fecha tiempo_promedio_sec cantidad"),
    "robos": ("Robos y hurtos", "fecha hora modulo habitacion empresa nombre_cliente rut medio_reclamo especies observaciones recepciona"),
    "miscelaneo": ("Misceláneo", "ot division area lugar ubicacion disciplina especialidad falla empresa fecha_creacion fecha_inicio fecha_termino fecha_aprobacion estado comentario"),
    "desviaciones": ("Desviaciones", "n_solicitud fecha id_interno empresa_contratista descripcion_problema tipo_riesgo tipo_solicitud pabellon habitacion via_solicitud quien_informa riesgo_material correo_destino acciones"),
    "solicitud_ot": ("Solicitud y OT de usuario", "n_solicitud descripcion_problema tipo_solicitud modulo habitacion tipo_turno jornada via_solicitud correo_usuario tipo_tarea ot fecha_inicio estado tiempo_respuesta_sec satisfaccion_reclamo motivo observacion"),
    "reclamos": ("Reclamos de usuarios", "n_solicitud fecha id_interno empresa_contratista descripcion_problema tipo_solicitud pabellon habitacion via_solicitud ingresar_contacto nombre_usuario responsable estatus notificacion_usuario plan_accion"),
    "alarmas": ("Activación de alarma", "modulo n_habitacion nombre_recepcionista fecha empresa id_interno co aviso_mantencion_h llegada_mantencion_h aviso_lider_h llegada_lider_h hora_reporte_salfa tipo_evento tipo_actividad fecha_reporte turno_recepcion_ingresos observaciones"),
    "extensiones": ("Extensión y excepción", "fecha_solicitud id_interno empresa co gerencia proyecto cant_clientes desde hasta aprobador observacion"),
    "onboarding": ("Onboarding", "fecha_hora nombre rut empresa id_interno archivo_pdf"),
    "apertura": ("Apertura de habitación", "fecha habitacion hora responsable estado_chapa"),
    "cumplimiento": ("Cumplimiento EECC", "fecha empresa n_contrato co correo_electronico id_interno turno"),
}

FIELD_LABELS = {
    "fecha": "Fecha", "censo_dia": "Censo día", "censo_noche": "Censo noche",
    "total": "Total", "que_ocurrio": "¿Qué ocurrió?", "accion": "Acción",
    "id_interno": "ID", "descripcion_problema": "Descripción del problema",
    "pabellon": "Pabellón", "habitacion": "Habitación",
    "ingresar_contacto": "Contacto", "notificacion_usuario": "Notificación al usuario",
    "plan_accion": "Plan de acción", "fecha_hora": "Fecha y hora",
    "tiempo_promedio_sec": "Tiempo promedio (mm:ss)", "cantidad": "Cantidad de atenciones",
    "modulo": "Módulo", "rut": "RUT", "ot": "OT", "division": "División",
    "area": "Área", "ubicacion": "Ubicación", "fecha_creacion": "Fecha de creación",
    "fecha_termino": "Fecha de término", "fecha_aprobacion": "Fecha de aprobación",
    "n_solicitud": "N.º de solicitud", "via_solicitud": "Vía de solicitud",
    "quien_informa": "Quién informa", "tiempo_respuesta_sec": "Tiempo de respuesta (mm:ss)",
    "satisfaccion_reclamo": "Satisfacción del reclamo", "observacion": "Observación",
    "n_habitacion": "N.º de habitación", "co": "C.O.",
    "aviso_mantencion_h": "Aviso a mantención (horas)",
    "llegada_mantencion_h": "Llegada de mantención (horas)",
    "aviso_lider_h": "Aviso al líder de emergencias (horas)",
    "llegada_lider_h": "Llegada del líder (horas)",
    "hora_reporte_salfa": "Hora de reporte SALFA",
    "turno_recepcion_ingresos": "Turno de recepción / ingresos",
    "cant_clientes": "Cantidad de clientes", "archivo_pdf": "Archivo PDF (nombre)",
    "n_contrato": "N.º de contrato", "correo_electronico": "Correo electrónico",
    "estado_chapa": "Estado de la chapa",
}
for question in range(1, 6):
    FIELD_LABELS[f"q{question}_respuesta"] = f"Pregunta {question}: respuesta"
    FIELD_LABELS[f"q{question}_puntaje"] = f"Pregunta {question}: puntaje"

DURATION_FIELDS = {"tiempo_promedio_sec", "tiempo_respuesta_sec"}
MAX_INTEGER = 2147483647


def record_version(record):
    values = {column.name: getattr(record, column.name) for column in record.__table__.columns}
    payload = json.dumps(values, sort_keys=True, default=str, ensure_ascii=False)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def edit_fields(entity, record):
    fields = []
    for name in EDIT_CONFIG[entity][1].split():
        column = record.__table__.columns[name]
        value = getattr(record, name)
        kind = "text"
        if name in DURATION_FIELDS:
            kind = "duration"
        elif isinstance(column.type, DateTime):
            kind = "datetime-local"
        elif isinstance(column.type, Date):
            kind = "date"
        elif isinstance(column.type, Time):
            kind = "time"
        elif isinstance(column.type, (Integer, Float)):
            kind = "number"
        elif isinstance(column.type, Text):
            kind = "textarea"

        if value is None:
            display = ""
        elif kind == "duration":
            minutes, seconds = divmod(value, 60)
            display = f"{minutes:02d}:{seconds:02d}"
        elif isinstance(value, (datetime, date, time)):
            display = value.isoformat()
        else:
            display = str(value)

        fields.append({
            "name": name, "label": FIELD_LABELS.get(name, name.replace("_", " ").capitalize()),
            "kind": kind, "value": display, "column": column,
            "required": not column.nullable and column.default is None,
            "maxlength": getattr(column.type, "length", None),
            "step": "1" if isinstance(column.type, Integer) else "any",
            "minimum": 1 if name == "semana" or name.endswith("_puntaje") else 0,
            "maximum": 5 if name.endswith("_puntaje") else (MAX_INTEGER if isinstance(column.type, Integer) else None),
        })
    return fields


def parse_edit_values(entity, fields, form):
    """Valida todo antes de modificar el objeto; devuelve errores por campo."""
    values, errors = {}, {}
    for field in fields:
        name, column = field["name"], field["column"]
        raw = form.get(name, field["value"]).strip()
        if entity == "censo" and name == "total" and form.get("censo_total_auto") == "1":
            continue
        try:
            if not raw:
                if field["required"]:
                    raise ValueError("Completa este campo.")
                value = None if column.nullable or name == "total" else 0
            elif field["kind"] == "duration":
                if not re.fullmatch(r"[0-9]+:[0-5][0-9]", raw):
                    raise ValueError("Usa minutos:segundos, por ejemplo 03:54. Los segundos van de 00 a 59.")
                minutes, seconds = map(int, raw.split(":"))
                value = minutes * 60 + seconds
                if value > MAX_INTEGER:
                    raise ValueError("El tiempo ingresado es demasiado grande.")
            elif isinstance(column.type, DateTime):
                value = datetime.fromisoformat(raw)
                if value.tzinfo is not None:
                    raise ValueError("Ingresa una fecha y hora local.")
            elif isinstance(column.type, Date):
                value = date.fromisoformat(raw)
            elif isinstance(column.type, Time):
                value = time.fromisoformat(raw)
                if value.tzinfo is not None:
                    raise ValueError("Ingresa una hora local.")
            elif isinstance(column.type, Integer):
                if not re.fullmatch(r"[0-9]+", raw):
                    raise ValueError("Ingresa un número entero sin decimales ni valores negativos.")
                value = int(raw)
                if not field["minimum"] <= value <= field["maximum"]:
                    raise ValueError(f"Ingresa un valor entre {field['minimum']} y {field['maximum']}.")
            elif isinstance(column.type, Float):
                value = float(raw.replace(",", "."))
                if not math.isfinite(value) or value < 0:
                    raise ValueError("Ingresa un número válido mayor o igual a cero.")
            else:
                value = raw
                if field["maxlength"] and len(value) > field["maxlength"]:
                    raise ValueError(f"Usa un máximo de {field['maxlength']} caracteres.")
            values[name] = value
        except (ValueError, OverflowError) as exc:
            if isinstance(column.type, (Date, DateTime, Time)):
                errors[name] = "Ingresa una fecha u hora válida."
            elif isinstance(column.type, Float):
                errors[name] = "Ingresa un número válido mayor o igual a cero."
            else:
                errors[name] = str(exc)

    if not errors:
        if entity == "censo" and (form.get("censo_total_auto") == "1" or values.get("total") is None):
            values["total"] = values["censo_dia"] + values["censo_noche"]
            if values["total"] > MAX_INTEGER:
                errors["total"] = "La suma del censo es demasiado grande."
        elif entity == "encuestas":
            scores = [values[f"q{i}_puntaje"] for i in range(1, 6) if values[f"q{i}_puntaje"] is not None]
            values["total"] = sum(scores) if scores else None
            values["promedio"] = round(sum(scores) / len(scores), 2) if scores else None
        elif entity == "extensiones" and values["desde"] and values["hasta"] and values["hasta"] < values["desde"]:
            errors["hasta"] = "La fecha hasta debe ser igual o posterior a la fecha desde."
    return values, errors
