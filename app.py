
import hashlib
import os
import re
import unicodedata
from collections import defaultdict
from difflib import SequenceMatcher
from datetime import date, datetime, timedelta
from io import BytesIO

import pandas as pd
from dotenv import load_dotenv
from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, url_for
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from sqlalchemy import func, or_

load_dotenv()
db = SQLAlchemy()

DATE_RE = re.compile(r"(\d{1,2})[-_/\.](\d{1,2})[-_/\.](\d{2,4})")
DEFAULT_GERENCIAS = [
    "BHP Projects", "Cátodos", "Concentradora", "Development&Strategic", "HSE",
    "Integrated Operations", "Mine Operations", "NPI&CHO", "PT&E", "VPP", "VPP Growth",
    "VP Sustaining", "Infraestructura & Servicios",
]


def now_utc():
    return datetime.utcnow()


class UploadedFile(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    file_type = db.Column(db.String(20), nullable=False)  # curva | censo
    size_bytes = db.Column(db.Integer, nullable=False, default=0)
    sha256 = db.Column(db.String(64), nullable=False, index=True)
    content = db.Column(db.LargeBinary, nullable=False)
    uploaded_at = db.Column(db.DateTime, nullable=False, default=now_utc)


class CurvaVersion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    file_id = db.Column(db.Integer, db.ForeignKey("uploaded_file.id"), nullable=False)
    name = db.Column(db.String(255), nullable=False)
    sheet_name = db.Column(db.String(120), nullable=False, default="Fcst_5400")
    is_active = db.Column(db.Boolean, nullable=False, default=True)
    total_items = db.Column(db.Integer, nullable=False, default=0)
    total_daily_values = db.Column(db.Integer, nullable=False, default=0)
    uploaded_at = db.Column(db.DateTime, nullable=False, default=now_utc)
    file = db.relationship("UploadedFile", backref="curvas")


class CurvaItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    curva_version_id = db.Column(db.Integer, db.ForeignKey("curva_version.id"), nullable=False, index=True)
    solicitud_id = db.Column(db.String(80), nullable=False, index=True)
    gerencia = db.Column(db.String(180), nullable=False, index=True)
    area = db.Column(db.String(180), default="")
    empresa = db.Column(db.String(180), default="")
    turno = db.Column(db.String(80), default="")
    tipo_contrato = db.Column(db.String(120), default="")
    formato = db.Column(db.String(120), default="")
    camp = db.Column(db.String(120), default="")
    version = db.relationship("CurvaVersion", backref="items")
    __table_args__ = (db.UniqueConstraint("curva_version_id", "solicitud_id", name="uq_curva_solicitud"),)


class CurvaDailyValue(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    curva_item_id = db.Column(db.Integer, db.ForeignKey("curva_item.id"), nullable=False, index=True)
    fecha = db.Column(db.Date, nullable=False, index=True)
    dotacion_planificada = db.Column(db.Float, nullable=False, default=0)
    item = db.relationship("CurvaItem", backref="daily_values")


class Censo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    file_id = db.Column(db.Integer, db.ForeignKey("uploaded_file.id"), nullable=False)
    fecha_censo = db.Column(db.Date, nullable=False, index=True)
    sheet_name = db.Column(db.String(120), default="")
    total_records = db.Column(db.Integer, nullable=False, default=0)
    total_occupied = db.Column(db.Float, nullable=False, default=0)
    matched_count = db.Column(db.Integer, nullable=False, default=0)
    unmatched_count = db.Column(db.Integer, nullable=False, default=0)
    imported_at = db.Column(db.DateTime, nullable=False, default=now_utc)
    file = db.relationship("UploadedFile", backref="censos")

    @property
    def fecha_label(self):
        return self.fecha_censo.strftime("%d/%m/%Y")


class CensoRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    censo_id = db.Column(db.Integer, db.ForeignKey("censo.id"), nullable=False, index=True)
    curva_item_id = db.Column(db.Integer, db.ForeignKey("curva_item.id"), nullable=True, index=True)
    solicitud_id = db.Column(db.String(80), index=True)
    modulo = db.Column(db.String(120), default="")
    lugar = db.Column(db.String(120), default="")
    habitacion = db.Column(db.String(120), default="")
    empresa = db.Column(db.String(180), default="")
    cama = db.Column(db.String(80), default="")
    dia = db.Column(db.String(80), default="")
    camas_ocupadas = db.Column(db.Float, nullable=False, default=1)
    turno = db.Column(db.String(80), default="")
    gerencia_censo = db.Column(db.String(180), default="")
    area = db.Column(db.String(180), default="")
    rut = db.Column(db.String(80), default="")
    estado = db.Column(db.String(120), default="")
    censo = db.relationship("Censo", backref="records")
    curva_item = db.relationship("CurvaItem", backref="censo_records")


def get_database_url(app):
    """
    Resolve la URL de base de datos.

    En Render debe venir desde la variable DATABASE_URL de PostgreSQL.
    Para desarrollo local, crea automáticamente la carpeta instance/ y usa SQLite.
    """
    db_url = os.getenv("DATABASE_URL", "").strip()

    if db_url:
        # Render/PostgreSQL antiguos pueden entregar postgres://, pero SQLAlchemy usa postgresql://
        if db_url.startswith("postgres://"):
            db_url = db_url.replace("postgres://", "postgresql://", 1)
        return db_url

    # Fallback solo para desarrollo local.
    # Esto evita el error: sqlite3.OperationalError: unable to open database file
    # cuando la carpeta instance/ no existe.
    os.makedirs(app.instance_path, exist_ok=True)
    return "sqlite:///" + os.path.join(app.instance_path, "dotacion_reportes.db")


def create_app():
    app = Flask(__name__, instance_relative_config=True)
    app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "dev-change-me")
    app.config["SQLALCHEMY_DATABASE_URI"] = get_database_url(app)
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {"pool_pre_ping": True}
    app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_CONTENT_LENGTH_MB", "80")) * 1024 * 1024
    db.init_app(app)
    with app.app_context():
        db.create_all()
    register_routes(app)
    return app


def strip_accents(text):
    text = unicodedata.normalize("NFKD", str(text))
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def norm(value):
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", strip_accents(value).lower().replace("_", " ")).strip()


def clean(value):
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def norm_id(value):
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text.upper()


def as_number(value):
    if value is None or pd.isna(value):
        return 0.0
    if isinstance(value, str):
        text = value.strip().replace(",", ".")
        if norm(text) in {"si", "sí", "x", "ok", "ocupado", "ocupada"}:
            return 1.0
        try:
            return float(text)
        except ValueError:
            return 0.0
    try:
        return float(value)
    except Exception:
        return 0.0


def parse_date(value):
    if value is None or value == "" or pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and 20000 <= float(value) <= 60000:
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
    parsed = pd.to_datetime(str(value), errors="coerce", dayfirst=True)
    return None if pd.isna(parsed) else parsed.date()


def date_from_text(text):
    m = DATE_RE.search(str(text or ""))
    if not m:
        return None
    d, mo, y = m.groups()
    y = int(y) + (2000 if int(y) < 100 else 0)
    try:
        return date(y, int(mo), int(d))
    except ValueError:
        return None


def find_header(raw, required):
    required = [norm(x) for x in required]
    for i in range(min(len(raw), 50)):
        values = {norm(v) for v in raw.iloc[i].tolist() if clean(v)}
        if sum(1 for x in required if x in values) >= min(2, len(required)):
            return i
    return 0


def colmap(columns):
    aliases = {
        "id": ["id de la solicitud", "id solicitud", "id", "solicitud"],
        "gerencia": ["gerencia general", "gerencia"],
        "area": ["area", "área"], "empresa": ["empresa"], "turno": ["turno"],
        "tipo_contrato": ["tipo contrato", "tipo de contrato"], "formato": ["formato"], "camp": ["camp", "campamento"],
        "modulo": ["modulo", "módulo"], "lugar": ["lugar"], "habitacion": ["habitacion", "habitación"],
        "cama": ["cama"], "dia": ["dia", "día"], "ocupadas": ["camas ocupadas", "camas ocupdas", "ocupadas"],
        "rut": ["rut", "run"], "estado": ["estado"],
    }
    available = {norm(c): c for c in columns}
    out = {}
    for key, opts in aliases.items():
        for opt in opts:
            if norm(opt) in available:
                out[key] = available[norm(opt)]
                break
    return out


def save_upload(file, file_type):
    content = file.read()
    f = UploadedFile(filename=file.filename, file_type=file_type, size_bytes=len(content), sha256=hashlib.sha256(content).hexdigest(), content=content)
    db.session.add(f); db.session.flush()
    return f, content



def xlsx_sheet_paths(content):
    """
    Lee los nombres de hojas de un .xlsx directamente desde el ZIP interno.
    No usa openpyxl para evitar cargar estilos pesados en Render.
    """
    import zipfile
    import xml.etree.ElementTree as ET

    ns_main = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    ns_rel = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}

    with zipfile.ZipFile(BytesIO(content)) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

        rel_map = {}
        for rel in rels.findall("r:Relationship", ns_rel):
            target = rel.attrib.get("Target", "")
            if not target.startswith("/"):
                target = "xl/" + target
            else:
                target = target.lstrip("/")
            rel_map[rel.attrib.get("Id")] = target

        sheets = []
        for sheet in workbook.findall("a:sheets/a:sheet", ns_main):
            name = sheet.attrib.get("name")
            rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            path = rel_map.get(rel_id)
            if name and path:
                sheets.append((name, path))
        return sheets


def get_sheet_names(content, filename):
    name = (filename or "").lower()
    if name.endswith(".xlsx") or name.endswith(".xlsm"):
        return [name for name, _ in xlsx_sheet_paths(content)]
    if name.endswith(".xlsb"):
        return pd.ExcelFile(BytesIO(content), engine="pyxlsb").sheet_names
    return pd.ExcelFile(BytesIO(content)).sheet_names


def column_index_from_cell_ref(cell_ref):
    letters = "".join(ch for ch in str(cell_ref or "") if ch.isalpha()).upper()
    if not letters:
        return 0
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def read_xlsx_shared_strings(zf):
    import xml.etree.ElementTree as ET
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []

    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    root = ET.fromstring(data)
    values = []
    for si in root.findall("a:si", ns):
        texts = [node.text or "" for node in si.findall(".//a:t", ns)]
        values.append("".join(texts))
    return values


def parse_xlsx_cell(cell, shared_strings):
    import xml.etree.ElementTree as ET
    cell_type = cell.attrib.get("t")
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

    if cell_type == "inlineStr":
        texts = [node.text or "" for node in cell.findall(f".//{ns}t")]
        return "".join(texts)

    value_node = cell.find(f"{ns}v")
    if value_node is None or value_node.text is None:
        return ""

    raw = value_node.text
    if cell_type == "s":
        try:
            return shared_strings[int(raw)]
        except Exception:
            return ""
    if cell_type == "b":
        return raw == "1"

    try:
        num = float(raw)
        return int(num) if num.is_integer() else num
    except Exception:
        return raw


def read_xlsx_sheet_df(content, sheet_name, header=None, nrows=None):
    """
    Convierte una hoja .xlsx a DataFrame leyendo XML crudo.
    Esto evita openpyxl/pandas para la curva, que en Render puede morir por estilos.
    """
    import zipfile
    import xml.etree.ElementTree as ET

    sheet_map = dict(xlsx_sheet_paths(content))
    if sheet_name not in sheet_map:
        raise ValueError(f"La hoja {sheet_name} no existe en el archivo.")

    rows = []
    max_cols = 0
    with zipfile.ZipFile(BytesIO(content)) as zf:
        shared_strings = read_xlsx_shared_strings(zf)
        sheet_path = sheet_map[sheet_name]
        with zf.open(sheet_path) as fh:
            context = ET.iterparse(fh, events=("end",))
            ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            for _, elem in context:
                if elem.tag != f"{ns}row":
                    continue
                row_values = []
                for cell in elem.findall(f"{ns}c"):
                    col_idx = column_index_from_cell_ref(cell.attrib.get("r"))
                    while len(row_values) <= col_idx:
                        row_values.append("")
                    row_values[col_idx] = parse_xlsx_cell(cell, shared_strings)
                max_cols = max(max_cols, len(row_values))
                rows.append(row_values)
                elem.clear()
                if nrows is not None and len(rows) >= nrows:
                    break

    if not rows:
        return pd.DataFrame()

    for row in rows:
        if len(row) < max_cols:
            row.extend([""] * (max_cols - len(row)))

    if header is None:
        return pd.DataFrame(rows)

    header = int(header)
    columns = rows[header]
    # Asegura nombres únicos para columnas vacías o repetidas.
    cleaned_columns = []
    seen = {}
    for i, col in enumerate(columns):
        # Mantiene los encabezados numéricos de fecha como número Excel.
        # Así parse_date(46083) los convierte correctamente a fecha real.
        if col is None or col == "" or pd.isna(col):
            name = f"col_{i + 1}"
        else:
            name = col

        key = str(name)
        if key in seen:
            seen[key] += 1
            name = f"{key}_{seen[key]}"
        else:
            seen[key] = 0
        cleaned_columns.append(name)

    data_rows = rows[header + 1:]
    return pd.DataFrame(data_rows, columns=cleaned_columns)


def read_excel_df(content, filename, sheet_name, header=None, nrows=None):
    name = (filename or "").lower()
    if name.endswith(".xlsx") or name.endswith(".xlsm"):
        return read_xlsx_sheet_df(content, sheet_name=sheet_name, header=header, nrows=nrows)

    engine = "pyxlsb" if name.endswith(".xlsb") else None
    kwargs = {
        "io": BytesIO(content),
        "sheet_name": sheet_name,
        "header": header,
    }
    if engine:
        kwargs["engine"] = engine
    if nrows is not None:
        kwargs["nrows"] = nrows
    return pd.read_excel(**kwargs)


def import_curva(file, sheet_name="Fcst_5400", version_name=None):
    uploaded, content = save_upload(file, "curva")
    sheet_names = get_sheet_names(content, uploaded.filename)
    sheet = sheet_name if sheet_name in sheet_names else next((s for s in sheet_names if "5400" in s), sheet_names[0])
    raw = read_excel_df(content, uploaded.filename, sheet_name=sheet, header=None, nrows=80)
    header = find_header(raw, ["ID de la solicitud", "Gerencia General"])
    df = read_excel_df(content, uploaded.filename, sheet_name=sheet, header=header).dropna(how="all")
    cm = colmap(df.columns)
    if "id" not in cm or "gerencia" not in cm:
        raise ValueError("No se encontraron columnas ID de la solicitud y Gerencia General en la curva.")
    date_cols = [(c, parse_date(c)) for c in df.columns]
    date_cols = [(c, d) for c, d in date_cols if d]
    CurvaVersion.query.update({CurvaVersion.is_active: False})
    version = CurvaVersion(file_id=uploaded.id, name=version_name or uploaded.filename, sheet_name=sheet, is_active=True)
    db.session.add(version); db.session.flush()
    seen, n_items, n_values = set(), 0, 0
    for _, row in df.iterrows():
        sid = norm_id(row.get(cm["id"]))
        if not sid or sid in seen:
            continue
        seen.add(sid)
        item = CurvaItem(
            curva_version_id=version.id, solicitud_id=sid, gerencia=clean(row.get(cm["gerencia"])) or "SIN GERENCIA",
            area=clean(row.get(cm.get("area"))) if cm.get("area") else "",
            empresa=clean(row.get(cm.get("empresa"))) if cm.get("empresa") else "",
            turno=clean(row.get(cm.get("turno"))) if cm.get("turno") else "",
            tipo_contrato=clean(row.get(cm.get("tipo_contrato"))) if cm.get("tipo_contrato") else "",
            formato=clean(row.get(cm.get("formato"))) if cm.get("formato") else "",
            camp=clean(row.get(cm.get("camp"))) if cm.get("camp") else "",
        )
        db.session.add(item); db.session.flush(); n_items += 1
        vals = [CurvaDailyValue(curva_item_id=item.id, fecha=d, dotacion_planificada=as_number(row.get(c))) for c, d in date_cols]
        if vals:
            db.session.bulk_save_objects(vals); n_values += len(vals)
    version.total_items, version.total_daily_values = n_items, n_values
    db.session.commit()
    return version


def read_censo(content, filename):
    sheet_names = get_sheet_names(content, filename)
    sheet = next((s for s in sheet_names if date_from_text(s)), sheet_names[0])
    detected_date = date_from_text(sheet) or date_from_text(filename)
    raw = read_excel_df(content, filename, sheet_name=sheet, header=None, nrows=80)
    header = find_header(raw, ["Id", "Camas Ocupadas"])
    df = read_excel_df(content, filename, sheet_name=sheet, header=header).dropna(how="all")
    return df, sheet, detected_date


def import_censo(file):
    uploaded, content = save_upload(file, "censo")
    df, sheet, fecha = read_censo(content, uploaded.filename)
    cm = colmap(df.columns)
    if "id" not in cm:
        raise ValueError("No se encontró la columna Id en el censo.")
    if not fecha and "dia" in cm:
        dates = [parse_date(x) for x in df[cm["dia"]].dropna().tolist()]
        dates = [x for x in dates if x]
        fecha = max(set(dates), key=dates.count) if dates else None
    if not fecha:
        raise ValueError("No se pudo detectar la fecha del censo.")
    active = CurvaVersion.query.filter_by(is_active=True).order_by(CurvaVersion.uploaded_at.desc()).first()
    item_map = {i.solicitud_id: i.id for i in CurvaItem.query.filter_by(curva_version_id=active.id).all()} if active else {}
    censo = Censo(file_id=uploaded.id, fecha_censo=fecha, sheet_name=sheet)
    db.session.add(censo); db.session.flush()
    records, matched, unmatched, occupied_total = [], 0, 0, 0
    for _, row in df.iterrows():
        sid = norm_id(row.get(cm["id"]))
        if not sid: continue
        occupied = as_number(row.get(cm.get("ocupadas"))) if cm.get("ocupadas") else 1.0
        if occupied <= 0: continue
        item_id = item_map.get(sid)
        matched += 1 if item_id else 0; unmatched += 0 if item_id else 1
        occupied_total += occupied
        records.append(CensoRecord(
            censo_id=censo.id, curva_item_id=item_id, solicitud_id=sid,
            modulo=clean(row.get(cm.get("modulo"))) if cm.get("modulo") else "",
            lugar=clean(row.get(cm.get("lugar"))) if cm.get("lugar") else "",
            habitacion=clean(row.get(cm.get("habitacion"))) if cm.get("habitacion") else "",
            empresa=clean(row.get(cm.get("empresa"))) if cm.get("empresa") else "",
            cama=clean(row.get(cm.get("cama"))) if cm.get("cama") else "",
            dia=clean(row.get(cm.get("dia"))) if cm.get("dia") else fecha.strftime("%d/%m/%Y"),
            camas_ocupadas=occupied, turno=clean(row.get(cm.get("turno"))) if cm.get("turno") else "",
            gerencia_censo=clean(row.get(cm.get("gerencia"))) if cm.get("gerencia") else "",
            area=clean(row.get(cm.get("area"))) if cm.get("area") else "", rut=clean(row.get(cm.get("rut"))) if cm.get("rut") else "",
            estado=clean(row.get(cm.get("estado"))) if cm.get("estado") else "",
        ))
    if records: db.session.bulk_save_objects(records)
    censo.total_records, censo.total_occupied, censo.matched_count, censo.unmatched_count = len(records), occupied_total, matched, unmatched
    db.session.commit()
    return censo


def date_span(start, end):
    while start <= end:
        yield start
        start += timedelta(days=1)


def report_data(start=None, end=None, curve_id=None):
    if not start or not end:
        minmax = db.session.query(func.min(Censo.fecha_censo), func.max(Censo.fecha_censo)).one()
        start = start or minmax[0]; end = end or minmax[1]
    if not start or not end:
        return {"dates": [], "date_labels": [], "rows": [], "totals_by_date": [], "grand_total": 0, "conclusions": ["No hay censos importados."]}
    dates = list(date_span(start, end))
    curve = CurvaVersion.query.get(curve_id) if curve_id else CurvaVersion.query.filter_by(is_active=True).order_by(CurvaVersion.uploaded_at.desc()).first()
    actual, planned, gerencias = defaultdict(lambda: defaultdict(float)), defaultdict(lambda: defaultdict(float)), set()
    rows_actual = db.session.query(Censo.fecha_censo, func.coalesce(CurvaItem.gerencia, "SIN MATCH EN CURVA"), func.sum(CensoRecord.camas_ocupadas)).join(CensoRecord, CensoRecord.censo_id==Censo.id).outerjoin(CurvaItem, CurvaItem.id==CensoRecord.curva_item_id).filter(Censo.fecha_censo>=start, Censo.fecha_censo<=end).group_by(Censo.fecha_censo, func.coalesce(CurvaItem.gerencia, "SIN MATCH EN CURVA")).all()
    for d, g, val in rows_actual:
        actual[g][d] += float(val or 0); gerencias.add(g)
    if curve:
        rows_plan = db.session.query(CurvaItem.gerencia, CurvaDailyValue.fecha, func.sum(CurvaDailyValue.dotacion_planificada)).join(CurvaDailyValue, CurvaDailyValue.curva_item_id==CurvaItem.id).filter(CurvaItem.curva_version_id==curve.id, CurvaDailyValue.fecha>=start, CurvaDailyValue.fecha<=end).group_by(CurvaItem.gerencia, CurvaDailyValue.fecha).all()
        for g, d, val in rows_plan:
            planned[g][d] += float(val or 0); gerencias.add(g)
    order = [g for g in DEFAULT_GERENCIAS if g in gerencias] + sorted([g for g in gerencias if g not in DEFAULT_GERENCIAS and g != "SIN MATCH EN CURVA"])
    if "SIN MATCH EN CURVA" in gerencias: order.append("SIN MATCH EN CURVA")
    totals = [0.0 for _ in dates]; planned_totals = [0.0 for _ in dates]; out_rows=[]
    for g in order:
        vals=[actual[g].get(d,0.0) for d in dates]; plans=[planned[g].get(d,0.0) for d in dates]
        for i,v in enumerate(vals): totals[i]+=v
        for i,v in enumerate(plans): planned_totals[i]+=v
        out_rows.append({"gerencia":g,"values":vals,"planned_values":plans,"total":sum(vals),"planned_total":sum(plans),"difference":sum(vals)-sum(plans)})
    grand, plan_grand = sum(totals), sum(planned_totals)
    conclusions = [f"Dotación real acumulada: {grand:.0f}. Planificación acumulada: {plan_grand:.0f}."]
    if plan_grand: conclusions.append(f"Cumplimiento global contra curva: {(grand/plan_grand*100):.1f}%.")
    if dates: conclusions.append(f"Día con mayor dotación real: {dates[max(range(len(totals)), key=lambda i: totals[i])].strftime('%d/%m/%Y')} ({max(totals):.0f}).")
    if out_rows:
        top=max(out_rows, key=lambda r:r['total']); conclusions.append(f"Gerencia con mayor dotación real: {top['gerencia']} ({top['total']:.0f}).")
    no_match=next((r for r in out_rows if r['gerencia']=='SIN MATCH EN CURVA'), None)
    if no_match and no_match['total']>0: conclusions.append(f"Existen {no_match['total']:.0f} registros sin match de ID en la curva activa.")
    return {"start_date":start.isoformat(),"end_date":end.isoformat(),"dates":[d.isoformat() for d in dates],"date_labels":[d.strftime('%d/%m/%y') for d in dates],"rows":out_rows,"totals_by_date":totals,"planned_totals_by_date":planned_totals,"grand_total":grand,"planned_grand_total":plan_grand,"conclusions":conclusions,"curve":{"id":curve.id,"name":curve.name} if curve else None}


def report_xlsx(data):
    wb=Workbook(); ws=wb.active; ws.title='Dotación Gerencia'
    dates=data.get('date_labels',[]); rows=data.get('rows',[]); total_cols=len(dates)+2
    red='FF0000'; pink='F8C9CD'; yellow='FFFF00'; white='FFFFFF'; black='000000'
    border=Border(left=Side(style='thin',color=black),right=Side(style='thin',color=black),top=Side(style='thin',color=black),bottom=Side(style='thin',color=black))
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=total_cols)
    ws.cell(1,1,'RESUMEN DE DOTACIÓN POR GERENCIA')
    ws.cell(1,1).fill=PatternFill('solid',fgColor=red); ws.cell(1,1).font=Font(color=white,bold=True,size=14); ws.cell(1,1).alignment=Alignment(horizontal='center')
    ws.row_dimensions[1].height=30
    ws.cell(2,1,'GERENCIAS')
    for i,d in enumerate(dates,2): ws.cell(2,i,d)
    ws.cell(2,total_cols,'TOTAL GENERAL')
    for c in range(1,total_cols+1):
        cell=ws.cell(2,c); cell.fill=PatternFill('solid',fgColor=red); cell.font=Font(color=white,bold=True,size=8); cell.border=border; cell.alignment=Alignment(horizontal='center',vertical='center',text_rotation=90 if 1<c<total_cols else 0)
    ws.row_dimensions[2].height=75
    for r,item in enumerate(rows,3):
        ws.cell(r,1,item['gerencia'])
        for c,val in enumerate(item['values'],2): ws.cell(r,c,int(round(val)))
        ws.cell(r,total_cols,int(round(item['total'])))
        for c in range(1,total_cols+1):
            cell=ws.cell(r,c); cell.border=border; cell.font=Font(size=8,bold=(c in [1,total_cols])); cell.alignment=Alignment(horizontal='left' if c==1 else 'center')
            if 1<c<total_cols: cell.fill=PatternFill('solid',fgColor=pink)
            if c==total_cols: cell.fill=PatternFill('solid',fgColor=yellow if item['total'] else pink)
    tr=len(rows)+3; ws.cell(tr,1,'TOTAL')
    for c,val in enumerate(data.get('totals_by_date',[]),2): ws.cell(tr,c,int(round(val)))
    ws.cell(tr,total_cols,int(round(data.get('grand_total',0))))
    for c in range(1,total_cols+1):
        cell=ws.cell(tr,c); cell.border=border; cell.fill=PatternFill('solid',fgColor=yellow if c==total_cols else red); cell.font=Font(color=black if c==total_cols else white,bold=True,size=8); cell.alignment=Alignment(horizontal='left' if c==1 else 'center')
    ws.column_dimensions['A'].width=28
    for c in range(2,total_cols): ws.column_dimensions[get_column_letter(c)].width=5
    ws.column_dimensions[get_column_letter(total_cols)].width=14
    ws.freeze_panes='B3'
    rr=tr+3; ws.cell(rr,1,'CONCLUSIONES'); ws.cell(rr,1).font=Font(bold=True,color='B00000')
    for i,text in enumerate(data.get('conclusions',[]),rr+1):
        ws.cell(i,1,'• '+text); ws.merge_cells(start_row=i,start_column=1,end_row=i,end_column=min(total_cols,8)); ws.cell(i,1).alignment=Alignment(wrap_text=True)
    bio=BytesIO(); wb.save(bio); bio.seek(0); return bio



def compact_id(value):
    """Normaliza un ID para comparación flexible: quita espacios, guiones y símbolos."""
    return re.sub(r"[^A-Z0-9]", "", norm_id(value))


def digits_id(value):
    """Retorna solo los dígitos del ID para detectar diferencias por ceros, puntos o guiones."""
    return re.sub(r"\D", "", str(value or ""))


def recalc_censo_stats(censo_id):
    censo = Censo.query.get(censo_id)
    if not censo:
        return None
    records = CensoRecord.query.filter_by(censo_id=censo_id).all()
    censo.total_records = len(records)
    censo.total_occupied = sum(float(r.camas_ocupadas or 0) for r in records)
    censo.matched_count = sum(1 for r in records if r.curva_item_id)
    censo.unmatched_count = sum(1 for r in records if not r.curva_item_id)
    return censo


def set_latest_curve_active_if_needed():
    active = CurvaVersion.query.filter_by(is_active=True).first()
    if active:
        return active
    latest = CurvaVersion.query.order_by(CurvaVersion.uploaded_at.desc()).first()
    if latest:
        latest.is_active = True
    return latest


def delete_uploaded_file_data(file_id):
    """
    Elimina un archivo subido y toda la información derivada.

    - Si es censo: elimina registros del censo y el censo.
    - Si es curva: elimina versiones, items y valores diarios; los registros de censo que
      apuntaban a esa curva quedan como sin match para poder corregirlos con otra curva.
    """
    uploaded = UploadedFile.query.get_or_404(file_id)
    impacted_censo_ids = set()

    if uploaded.file_type == "censo":
        censos = Censo.query.filter_by(file_id=uploaded.id).all()
        censo_ids = [c.id for c in censos]
        if censo_ids:
            CensoRecord.query.filter(CensoRecord.censo_id.in_(censo_ids)).delete(synchronize_session=False)
            Censo.query.filter(Censo.id.in_(censo_ids)).delete(synchronize_session=False)

    elif uploaded.file_type == "curva":
        curves = CurvaVersion.query.filter_by(file_id=uploaded.id).all()
        for curve in curves:
            item_ids = [row[0] for row in db.session.query(CurvaItem.id).filter_by(curva_version_id=curve.id).all()]
            if item_ids:
                impacted_censo_ids.update(
                    row[0] for row in db.session.query(CensoRecord.censo_id)
                    .filter(CensoRecord.curva_item_id.in_(item_ids))
                    .distinct()
                    .all()
                )
                CensoRecord.query.filter(CensoRecord.curva_item_id.in_(item_ids)).update(
                    {CensoRecord.curva_item_id: None}, synchronize_session=False
                )
                CurvaDailyValue.query.filter(CurvaDailyValue.curva_item_id.in_(item_ids)).delete(synchronize_session=False)
                CurvaItem.query.filter(CurvaItem.id.in_(item_ids)).delete(synchronize_session=False)
            db.session.delete(curve)
    else:
        raise ValueError("Tipo de archivo no reconocido.")

    db.session.delete(uploaded)
    db.session.flush()

    for censo_id in impacted_censo_ids:
        recalc_censo_stats(censo_id)

    set_latest_curve_active_if_needed()
    db.session.commit()
    return uploaded


def get_curve_for_matching(curve_id=None):
    if curve_id:
        return CurvaVersion.query.get(curve_id)
    return CurvaVersion.query.filter_by(is_active=True).order_by(CurvaVersion.uploaded_at.desc()).first()


def build_curve_candidates(curve):
    if not curve:
        return []
    return CurvaItem.query.filter_by(curva_version_id=curve.id).all()


def build_curve_match_index(curve):
    """
    Construye un índice en memoria para proponer correcciones sin comparar cada
    registro contra toda la curva. Esto reemplaza el cálculo O(registros * curva)
    por búsquedas directas por ID normalizado, dígitos y sufijos.
    """
    items = build_curve_candidates(curve)
    index = {
        "items": items,
        "by_compact": defaultdict(list),
        "by_digits": defaultdict(list),
        "by_suffix": defaultdict(list),
    }

    for item in items:
        compact = compact_id(item.solicitud_id)
        digits = digits_id(item.solicitud_id).lstrip("0")

        if compact:
            index["by_compact"][compact].append(item)
        if digits:
            index["by_digits"][digits].append(item)
            # Sufijos largos primero; evita recorrer toda la curva para IDs parecidos.
            for n in range(4, min(len(digits), 12) + 1):
                index["by_suffix"][(n, digits[-n:])].append(item)

    return index


def candidate_score(record, item):
    target = compact_id(record.solicitud_id)
    candidate = compact_id(item.solicitud_id)
    target_digits = digits_id(record.solicitud_id).lstrip("0")
    candidate_digits = digits_id(item.solicitud_id).lstrip("0")

    if not target or not candidate:
        return 0, ""

    score = 0
    reason = "Similitud de texto"

    if target == candidate:
        score, reason = 100, "Coincidencia exacta ignorando formato"
    elif target_digits and candidate_digits and target_digits == candidate_digits:
        score, reason = 98, "Coinciden los dígitos ignorando ceros o símbolos"
    elif target in candidate or candidate in target:
        score, reason = 90, "Un ID contiene al otro"
    elif target_digits and candidate_digits:
        max_suffix = 0
        for n in range(min(len(target_digits), len(candidate_digits)), 3, -1):
            if target_digits[-n:] == candidate_digits[-n:]:
                max_suffix = n
                break
        if max_suffix:
            score = min(88, 70 + max_suffix * 2)
            reason = f"Coinciden los últimos {max_suffix} dígitos"

    # Solo se ejecuta sobre candidatos prefiltrados, no sobre toda la curva.
    ratio = SequenceMatcher(None, target, candidate).ratio()
    if ratio >= 0.70 and int(ratio * 85) > score:
        score = int(ratio * 85)
        reason = f"Similitud de ID {ratio * 100:.0f}%"

    if record.empresa and item.empresa and norm(record.empresa) == norm(item.empresa):
        score += 4
        reason += " + misma empresa"
    if record.gerencia_censo and item.gerencia and norm(record.gerencia_censo) == norm(item.gerencia):
        score += 4
        reason += " + misma gerencia"

    return min(score, 100), reason


def candidate_pool_for_record(record, match_index, max_candidates=250):
    """Obtiene un subconjunto pequeño de curva para calcular sugerencias."""
    if not match_index:
        return []

    target = compact_id(record.solicitud_id)
    target_digits = digits_id(record.solicitud_id).lstrip("0")
    pool = []
    seen = set()

    def add(items):
        for item in items or []:
            if item.id not in seen:
                seen.add(item.id)
                pool.append(item)
                if len(pool) >= max_candidates:
                    return

    if target:
        add(match_index["by_compact"].get(target))
    if target_digits:
        add(match_index["by_digits"].get(target_digits))
        for n in range(min(len(target_digits), 12), 3, -1):
            add(match_index["by_suffix"].get((n, target_digits[-n:])))
            if len(pool) >= max_candidates:
                break

    # Respaldo liviano: si no hay candidatos por dígitos, compara con una muestra
    # acotada de IDs que compartan inicio o término normalizado.
    if not pool and target:
        for item in match_index["items"][:max_candidates]:
            candidate = compact_id(item.solicitud_id)
            if candidate and (candidate[:4] == target[:4] or candidate[-4:] == target[-4:]):
                add([item])

    return pool[:max_candidates]


def suggestions_for_record(record, match_index, limit=3):
    suggestions = []
    for item in candidate_pool_for_record(record, match_index):
        score, reason = candidate_score(record, item)
        if score >= 60:
            suggestions.append({
                "item": item,
                "score": score,
                "reason": reason,
            })
    suggestions.sort(key=lambda x: x["score"], reverse=True)
    return suggestions[:limit]


def find_curve_item_by_manual_id(curve, manual_id):
    """Busca un ID manual en la curva, aceptando diferencias simples de formato."""
    if not curve or not manual_id:
        return None

    manual_id = norm_id(manual_id)
    item = CurvaItem.query.filter_by(curva_version_id=curve.id, solicitud_id=manual_id).first()
    if item:
        return item

    target = compact_id(manual_id)
    target_digits = digits_id(manual_id).lstrip("0")
    for item in CurvaItem.query.filter_by(curva_version_id=curve.id).all():
        if target and compact_id(item.solicitud_id) == target:
            return item
        if target_digits and digits_id(item.solicitud_id).lstrip("0") == target_digits:
            return item
    return None


def apply_correction_to_same_id(record, item):
    """
    Corrige todos los registros sin match que tengan el mismo ID de censo.
    Se aplica a todos los censos almacenados para evitar corregir uno por uno.
    """
    same_id = norm_id(record.solicitud_id)
    if not same_id:
        same_id = record.solicitud_id

    records = CensoRecord.query.filter(
        CensoRecord.curva_item_id.is_(None),
        CensoRecord.solicitud_id == same_id,
    ).all()

    # Respaldo por si algún registro quedó con espacios/formato distinto antes de normalizar.
    if not records:
        target = compact_id(record.solicitud_id)
        all_unmatched = CensoRecord.query.filter(CensoRecord.curva_item_id.is_(None)).all()
        records = [r for r in all_unmatched if compact_id(r.solicitud_id) == target]

    censo_ids = set()
    for row in records:
        row.curva_item_id = item.id
        censo_ids.add(row.censo_id)

    for censo_id in censo_ids:
        recalc_censo_stats(censo_id)

    return len(records), censo_ids


def no_match_payload(censo_id=None, curve_id=None, search_text="", limit=200):
    curve = get_curve_for_matching(curve_id)
    match_index = build_curve_match_index(curve)

    base_query = CensoRecord.query.join(Censo, Censo.id == CensoRecord.censo_id).filter(CensoRecord.curva_item_id.is_(None))
    if censo_id:
        base_query = base_query.filter(CensoRecord.censo_id == censo_id)

    search_text = (search_text or "").strip()
    if search_text:
        like = f"%{search_text}%"
        base_query = base_query.filter(
            or_(
                CensoRecord.solicitud_id.ilike(like),
                CensoRecord.empresa.ilike(like),
                CensoRecord.gerencia_censo.ilike(like),
                CensoRecord.rut.ilike(like),
                CensoRecord.habitacion.ilike(like),
            )
        )

    total_records = base_query.count()
    total_ids = base_query.with_entities(CensoRecord.solicitud_id).distinct().count()

    grouped = base_query.with_entities(
        CensoRecord.solicitud_id.label("solicitud_id"),
        func.count(CensoRecord.id).label("same_id_count"),
        func.min(Censo.fecha_censo).label("first_date"),
        func.max(Censo.fecha_censo).label("last_date"),
        func.min(CensoRecord.id).label("sample_id"),
    ).group_by(
        CensoRecord.solicitud_id
    ).order_by(
        func.max(Censo.fecha_censo).desc(),
        CensoRecord.solicitud_id.asc()
    ).limit(limit).all()

    sample_ids = [row.sample_id for row in grouped]
    sample_records = {
        record.id: record
        for record in CensoRecord.query.filter(CensoRecord.id.in_(sample_ids)).all()
    } if sample_ids else {}

    rows = []
    for group in grouped:
        record = sample_records.get(group.sample_id)
        if not record:
            continue
        rows.append({
            "record": record,
            "censo": record.censo,
            "same_id_count": int(group.same_id_count or 0),
            "first_date": group.first_date,
            "last_date": group.last_date,
            "suggestions": suggestions_for_record(record, match_index),
        })

    return {
        "curve": curve,
        "total": total_ids,
        "total_records": total_records,
        "limit": limit,
        "rows": rows,
    }

def parse_arg(name):
    v=(request.args.get(name) or '').strip()
    return datetime.strptime(v,'%Y-%m-%d').date() if v else None


def register_routes(app):
    @app.route('/')
    def dashboard():
        active=CurvaVersion.query.filter_by(is_active=True).order_by(CurvaVersion.uploaded_at.desc()).first()
        latest=Censo.query.order_by(Censo.fecha_censo.desc()).first()
        return render_template('dashboard.html', active=active, latest=latest, total_censos=Censo.query.count(), total_files=UploadedFile.query.count(), total_unmatched=db.session.query(func.sum(Censo.unmatched_count)).scalar() or 0)

    @app.route('/imports')
    def imports_page():
        return render_template('imports.html', files=UploadedFile.query.order_by(UploadedFile.uploaded_at.desc()).limit(20).all(), curves=CurvaVersion.query.order_by(CurvaVersion.uploaded_at.desc()).all())

    @app.post('/api/import/curva')
    def upload_curva():
        f=request.files.get('file')
        if not f or not f.filename: flash('Selecciona un archivo de curva.', 'danger'); return redirect(url_for('imports_page'))
        try:
            v=import_curva(f, request.form.get('sheet_name') or 'Fcst_5400', request.form.get('version_name') or None)
            flash(f'Curva importada: {v.total_items} IDs y {v.total_daily_values} valores diarios.', 'success')
        except Exception as e:
            db.session.rollback(); flash(f'Error al importar curva: {e}', 'danger')
        return redirect(url_for('imports_page'))

    @app.post('/api/import/censo')
    def upload_censo():
        f=request.files.get('file')
        if not f or not f.filename: flash('Selecciona un archivo de censo.', 'danger'); return redirect(url_for('imports_page'))
        try:
            c=import_censo(f); flash(f'Censo {c.fecha_label} importado: {c.total_records} registros, {c.matched_count} cruzados, {c.unmatched_count} sin match.', 'success')
        except Exception as e:
            db.session.rollback(); flash(f'Error al importar censo: {e}', 'danger')
        return redirect(url_for('imports_page'))

    @app.post('/api/curves/<int:curve_id>/activate')
    def activate_curve(curve_id):
        CurvaVersion.query.update({CurvaVersion.is_active: False}); curve=CurvaVersion.query.get_or_404(curve_id); curve.is_active=True; db.session.commit(); flash(f'Curva activa: {curve.name}', 'success'); return redirect(url_for('imports_page'))


    @app.post('/api/files/<int:file_id>/delete')
    def delete_file(file_id):
        try:
            uploaded = delete_uploaded_file_data(file_id)
            flash(f'Archivo eliminado: {uploaded.filename}', 'success')
        except Exception as e:
            db.session.rollback()
            flash(f'Error al eliminar archivo: {e}', 'danger')
        return redirect(request.referrer or url_for('imports_page'))

    @app.route('/no-match')
    def no_match_page():
        censo_id = request.args.get('censo_id', type=int)
        curve_id = request.args.get('curve_id', type=int)
        search_text = request.args.get('q', '')
        limit = min(request.args.get('limit', 100, type=int) or 100, 1000)
        payload = no_match_payload(censo_id=censo_id, curve_id=curve_id, search_text=search_text, limit=limit)
        return render_template(
            'no_match.html',
            rows=payload['rows'],
            curve=payload['curve'],
            total=payload['total'],
            total_records=payload['total_records'],
            limit=payload['limit'],
            censos=Censo.query.order_by(Censo.fecha_censo.desc()).all(),
            curves=CurvaVersion.query.order_by(CurvaVersion.uploaded_at.desc()).all(),
            selected_censo_id=censo_id,
            selected_curve_id=curve_id,
            q=search_text,
        )

    @app.post('/api/no-match/<int:record_id>/apply')
    def apply_no_match(record_id):
        record = CensoRecord.query.get_or_404(record_id)
        item_id = request.form.get('curva_item_id', type=int)
        item = CurvaItem.query.get_or_404(item_id)
        updated_count, censo_ids = apply_correction_to_same_id(record, item)
        db.session.commit()
        flash(
            f'Corrección aplicada a {updated_count} registro(s) con ID censo {record.solicitud_id} → curva {item.solicitud_id} / {item.gerencia}.',
            'success'
        )
        return redirect(request.referrer or url_for('no_match_page'))

    @app.post('/api/no-match/<int:record_id>/manual')
    def manual_no_match(record_id):
        record = CensoRecord.query.get_or_404(record_id)
        curve_id = request.form.get('curve_id', type=int)
        manual_id = norm_id(request.form.get('manual_id'))
        curve = get_curve_for_matching(curve_id)
        if not curve:
            flash('No hay curva disponible para buscar el ID manual.', 'danger')
            return redirect(request.referrer or url_for('no_match_page'))
        item = find_curve_item_by_manual_id(curve, manual_id)
        if not item:
            flash(f'No se encontró el ID {manual_id} en la curva {curve.name}.', 'warning')
            return redirect(request.referrer or url_for('no_match_page'))
        updated_count, censo_ids = apply_correction_to_same_id(record, item)
        db.session.commit()
        flash(
            f'Corrección manual aplicada a {updated_count} registro(s): {record.solicitud_id} → {item.solicitud_id} / {item.gerencia}.',
            'success'
        )
        return redirect(request.referrer or url_for('no_match_page'))

    @app.route('/censos')
    def censos_page():
        return render_template('censos.html', censos=Censo.query.order_by(Censo.fecha_censo.desc()).all())

    @app.route('/reports/dotacion-gerencia')
    def report_page():
        return render_template('report.html', curves=CurvaVersion.query.order_by(CurvaVersion.uploaded_at.desc()).all())

    @app.route('/api/reports/dotacion-gerencia')
    def report_api():
        return jsonify(report_data(parse_arg('start_date'), parse_arg('end_date'), request.args.get('curve_id', type=int)))

    @app.route('/api/reports/dotacion-gerencia/export')
    def report_export():
        data=report_data(parse_arg('start_date'), parse_arg('end_date'), request.args.get('curve_id', type=int))
        return send_file(report_xlsx(data), as_attachment=True, download_name=f'dotacion_gerencia_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

app=create_app()
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT',5000)), debug=os.getenv('FLASK_DEBUG')=='1')
