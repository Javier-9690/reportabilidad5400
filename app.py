
import hashlib
import os
import re
import unicodedata
from collections import defaultdict
from datetime import date, datetime, timedelta
from io import BytesIO

import pandas as pd
from dotenv import load_dotenv
from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, url_for
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from sqlalchemy import func

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


def parse_arg(name):
    v=(request.args.get(name) or '').strip()
    return datetime.strptime(v,'%Y-%m-%d').date() if v else None


def register_routes(app):
    @app.route('/')
    def dashboard():
        active=CurvaVersion.query.filter_by(is_active=True).order_by(CurvaVersion.uploaded_at.desc()).first()
        latest=Censo.query.order_by(Censo.fecha_censo.desc()).first()
        return render_template('dashboard.html', active=active, latest=latest, total_censos=Censo.query.count(), total_files=UploadedFile.query.count())

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
