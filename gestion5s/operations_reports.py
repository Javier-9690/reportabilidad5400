"""Reporte de las siete categorías del libro de desviaciones y solicitudes.

Consulta las bases vigentes. El archivo de referencia no se importa ni modifica
registros. Los indicadores y el detalle comparten filtros y una misma transacción.
"""

from collections import Counter
from datetime import date, datetime, timedelta, timezone

from sqlalchemy import func, or_

from gestion5s.editing import list_fields
from gestion5s.orders import normalize_order_text, order_reference_column, order_reference_date
from gestion5s.searching import clean_search_term, record_search_condition
from gestion5s.user_reports import OPEN_STATUSES, classify_status, parse_report_range


REPORT_TITLE = "Reporte de desviaciones, solicitudes y reclamos"
PAGE_SIZE = 50
SOURCES = (
    dict(entity="miscelaneo", label="Misceláneos", sheet="Misceláneos", date="fecha_creacion",
         fallback=True, date_label="Fecha creación; si falta, Fecha inicio", status="estado",
         company="empresa", location=("lugar",), color="#916335"),
    dict(entity="solicitud_ot", label="Solicitudes y OT de usuario", sheet="Solicitudes y OT de usuario",
         date="fecha_creacion", fallback=True, date_label="Fecha creación; si falta, Fecha inicio",
         status="estado", company="empresa", location=("lugar", "modulo"), color="#28699C"),
    dict(entity="reclamos", label="Reclamos de usuarios", sheet="Reclamos de usuarios", date="fecha",
         date_label="Fecha", status="estatus", company="empresa_contratista", location=("pabellon",), color="#B42318"),
    dict(entity="robos", label="Robos y hurtos", sheet="Robos y hurtos", date="fecha", date_label="Fecha",
         status=None, company="empresa", location=("modulo",), color="#8A507C",
         tracking=("recepciona",), tracking_label="Con receptor de denuncia", missing_label="Sin receptor de denuncia"),
    dict(entity="duplicidades", label="Doble asignación", sheet="Doble asignación", date="fecha",
         date_label="Fecha", status="estatus", company="empresa_contratista", location=("pabellon",), color="#6560A4"),
    dict(entity="habitaciones_bloqueadas", label="Habitaciones bloqueadas", sheet="Habitaciones bloqueadas",
         date="fecha_bloqueo", date_label="Fecha de bloqueo", status=None,
         company="empresa", location=("habitacion",), color="#247B79",
         tracking=("fecha_liberada_ingeclean", "fecha_liberada_facility", "fecha_liberada_mantencion", "fecha_liberada_investigacion"),
         tracking_label="Con alguna fecha de liberación", missing_label="Sin fechas de liberación"),
    dict(entity="desviaciones", label="Desviaciones", sheet="Desviaciones", date="fecha", date_label="Fecha",
         status=None, company="empresa_contratista", location=("pabellon",), color="#BD771C",
         tracking=("acciones",), tracking_label="Con acciones registradas", missing_label="Sin acciones registradas"),
)
SOURCE_MAP = {source["entity"]: source for source in SOURCES}
STATE_LABELS = {"open": "Abierto", "closed": "Cerrado", "unknown": "Sin clasificar", "not_applicable": "No aplica"}
METHOD_NOTE = (
    "Se cuentan filas de los registros, no personas ni tickets únicos. Los estados son los actuales "
    "de los registros originados en el período; no representan el historial de cierres."
)
RATE_NOTE = (
    "Cierre = cerrados / (abiertos + cerrados). Sin un estado clasificable se muestra —. "
    "Desviaciones, Robos y hurtos y Habitaciones bloqueadas no tienen un campo de estado "
    "en el sistema y no participan en esta tasa."
)
SOURCE_NOTE = (
    "Misceláneos y Solicitudes OT se alimentan de la importación general Samtech: "
    "CARPINTERIA MENOR va a Misceláneos y el resto a Solicitudes OT. "
    "El reporte Gestión de usuarios mantiene el seguimiento separado de solicitudes QR."
)


def parse_filters(args):
    today = date.today()
    start, end = parse_report_range({
        "start_date": args.get("start_date", today.replace(day=1).isoformat()),
        "end_date": args.get("end_date", today.isoformat()),
    })
    tab = args.get("tab", "general")
    if tab not in ("general", *SOURCE_MAP):
        raise ValueError("La pestaña solicitada no existe.")
    raw_page = args.get("page", "1")
    if not raw_page.isascii() or not raw_page.isdigit() or len(raw_page) > 8 or int(raw_page) < 1:
        raise ValueError("La página solicitada no es válida.")
    undated = args.get("include_undated", "0")
    if undated not in ("0", "1"):
        raise ValueError("La opción de registros sin fecha no es válida.")
    return dict(start=start, end=end, search=clean_search_term(args.get("q", "")),
                include_undated=undated == "1", tab=tab, page=int(raw_page))


def report_state(source, raw):
    if not source["status"]:
        return "not_applicable"
    code = classify_status(raw, entity=source["entity"])
    return "open" if code in OPEN_STATUSES else "closed" if code == "cerrado" else "unknown"


def _ranks(groups):
    counts, labels = Counter(), {}
    for raw, count in groups:
        label = " ".join((raw or "").split()) or "Sin informar"
        key = normalize_order_text(label)
        counts[key] += count
        labels.setdefault(key, label)
    return [{"label": labels[key], "count": count} for key, count in sorted(counts.items(), key=lambda x: (-x[1], x[0]))]


def _rank_query(db, model, scope, names):
    columns = [func.nullif(func.trim(getattr(model, name)), "") for name in names]
    column = columns[0] if len(columns) == 1 else func.coalesce(*columns)
    return _ranks(db.query(column, func.count(model.id)).filter(*scope).group_by(column).all())


def _rate(totals):
    denominator = totals["open"] + totals["closed"]
    return totals["closed"] / denominator if denominator else None


def build_operations_report(db, models, filters, export=None):
    """export='general' incluye todos los detalles; un entity incluye solo ese detalle."""
    if export is not None and export not in ("general", *SOURCE_MAP):
        raise ValueError("La pestaña solicitada no existe.")
    sources = SOURCES if export in (None, "general") else (SOURCE_MAP[export],)
    start, end = filters["start"], filters["end"]
    days = [start + timedelta(days=i) for i in range((end - start).days + 1)]
    categories = []
    for source in sources:
        entity, model = source["entity"], models[source["entity"]]
        fields = [{key: f[key] for key in ("name", "label", "kind")} for f in list_fields(entity, model())]
        date_column = order_reference_column(model) if source.get("fallback") else getattr(model, source["date"])
        base = []
        if filters["search"]:
            base.append(record_search_condition(entity, model, filters["search"], db.get_bind().dialect.name))
        period = (date_column >= start) & (date_column <= end)
        if filters["include_undated"]:
            period = or_(period, date_column.is_(None))
        scope = [*base, period]
        undated_available = db.query(func.count(model.id)).filter(*base, date_column.is_(None)).scalar()
        if source["status"]:
            state_column = getattr(model, source["status"])
            groups = db.query(date_column, state_column, func.count(model.id)).filter(*scope).group_by(date_column, state_column).all()
        else:
            groups = [(day, None, count) for day, count in db.query(date_column, func.count(model.id))
                      .filter(*scope).group_by(date_column).all()]
        totals = dict(records=0, open=0, closed=0, unknown=0, not_applicable=0, undated=0)
        dated = Counter()
        for day, raw, count in groups:
            totals["records"] += count
            totals[report_state(source, raw)] += count
            if day is None:
                totals["undated"] += count
            else:
                dated[day] += count
        totals["rate"] = _rate(totals)
        companies = _rank_query(db, model, scope, (source["company"],))
        locations = _rank_query(db, model, scope, source["location"])
        totals["companies"] = sum(r["label"] != "Sin informar" for r in companies)
        if source.get("tracking"):
            populated = []
            for name in source["tracking"]:
                column = getattr(model, name)
                populated.append(column.is_not(None) if name.startswith("fecha_") else func.length(func.trim(column)) > 0)
            totals["tracked"] = db.query(func.count(model.id)).filter(*scope, or_(*populated)).scalar()
            totals["untracked"] = totals["records"] - totals["tracked"]
        page_count = max(1, (totals["records"] + PAGE_SIZE - 1) // PAGE_SIZE)
        page = min(filters["page"], page_count)
        records = []
        if export is not None or filters["tab"] == entity:
            query = db.query(model).filter(*scope).order_by(date_column.is_(None), date_column.desc(), model.id.desc())
            if export is None:
                query = query.limit(PAGE_SIZE).offset((page - 1) * PAGE_SIZE)
            for record in query.yield_per(1000):
                detail = {field["name"]: getattr(record, field["name"]) for field in fields}
                detail["_report_date"] = order_reference_date(record) if source.get("fallback") else getattr(record, source["date"])
                detail["_report_state"] = STATE_LABELS[report_state(source, getattr(record, source["status"]) if source["status"] else None)]
                records.append(detail)
        categories.append({**source, "totals": totals, "daily": [dated[day] for day in days],
                           "fields": fields, "details": records, "companies": companies, "locations": locations,
                           "undated_available": undated_available, "page": page, "pages": page_count,
                           "first_row": (page - 1) * PAGE_SIZE + 1 if totals["records"] else 0,
                           "last_row": min(page * PAGE_SIZE, totals["records"])})
    totals = {key: sum(c["totals"][key] for c in categories)
              for key in ("records", "open", "closed", "unknown", "not_applicable", "undated")}
    totals["rate"] = _rate(totals)
    totals["undated_available"] = sum(c["undated_available"] for c in categories)
    daily = [sum(c["daily"][i] for c in categories) for i in range(len(days))]
    report = dict(title=REPORT_TITLE, filters=filters, start=start, end=end, days=days, daily=daily,
                  categories=categories, totals=totals, generated_at=datetime.now(timezone.utc),
                  method_note=METHOD_NOTE, rate_note=RATE_NOTE, source_note=SOURCE_NOTE)
    report["conclusions"] = conclusions(report)
    return report


def conclusions(report):
    totals, categories = report["totals"], report["categories"]
    if not totals["records"]:
        return ["No hay registros que coincidan con los filtros seleccionados."]
    notes = []
    largest = max(categories, key=lambda c: c["totals"]["records"])
    if len(categories) > 1:
        share = f"{largest['totals']['records'] / totals['records'] * 100:.1f}".replace(".", ",")
        notes.append(f"{largest['label']} concentra {largest['totals']['records']} de los {totals['records']} registros "
                     f"({share}%).")
    if totals["rate"] is not None:
        rate = f"{totals['rate'] * 100:.1f}".replace(".", ",")
        notes.append(f"El cierre actual es {rate}%: {totals['closed']} cerrados y {totals['open']} abiertos con estado reconocido.")
    else:
        notes.append("No hay estados clasificables para calcular un porcentaje de cierre.")
    if totals["open"]:
        pending = max(categories, key=lambda c: c["totals"]["open"])
        notes.append(f"{pending['label']} tiene la mayor cantidad de registros abiertos: {pending['totals']['open']}.")
    if totals["unknown"]:
        notes.append(f"Hay {totals['unknown']} registros con estado vacío o no reconocido para revisar.")
    deviations = next((c for c in categories if c["entity"] == "desviaciones"), None)
    if deviations and deviations["totals"].get("untracked"):
        notes.append(f"{deviations['totals']['untracked']} desviaciones no tienen acciones registradas.")
    if max(report["daily"], default=0):
        peak = max(report["daily"])
        index = report["daily"].index(peak)
        notes.append(f"Mayor volumen diario: {peak} registros el {report['days'][index]:%d/%m/%Y}. "
                     "Si hay empate, se muestra la primera fecha.")
    return notes


def chart_data(report, selected=None):
    categories = [selected] if selected else report["categories"]
    monthly = len(report["days"]) > 93
    labels, keys = [], []
    for day in report["days"]:
        key = day.strftime("%m/%Y" if monthly else "%d/%m/%Y")
        if not keys or keys[-1] != key:
            labels.append(key)
        keys.append(key)
    series = []
    for category in categories:
        counts = Counter()
        for key, value in zip(keys, category["daily"]):
            counts[key] += value
        series.append(dict(label=category["label"], color=category["color"], values=[counts[label] for label in labels]))
    return dict(labels=labels, series=series, monthly=monthly,
                categories=[c["label"] for c in categories],
                states={state: [c["totals"][state] for c in categories] for state in STATE_LABELS},
                companies=selected["companies"][:7] if selected else [])
