"""Informe de casos de usuarios, agrupados por su fecha y estado actual."""

from collections import Counter
from datetime import date, datetime, timedelta, timezone
import re
import unicodedata

from sqlalchemy import func

from gestion5s.editing import list_fields


SOURCES = (
    {"entity": "duplicidades", "label": "Doble asignación", "date": "fecha", "date_label": "Fecha",
     "status": "estatus", "sheet": "Doble asignación", "expected_open": ("abierto",), "color": "#7950A3"},
    {"entity": "reclamos", "label": "Reclamos usuarios", "date": "fecha", "date_label": "Fecha",
     "status": "estatus", "sheet": "Reclamos usuarios", "expected_open": ("abierto",), "color": "#B42318"},
    {"entity": "solicitud_ot", "label": "Solicitudes de usuarios", "date": "fecha_inicio",
     "date_label": "Fecha inicio", "status": "estado", "sheet": "Solicitudes OT",
     "expected_open": ("no_iniciada",), "color": "#1667A5"},
    {"entity": "samtech_usuarios", "label": "Samtech usuarios", "date": "fecha_creacion",
     "date_label": "Fecha creación", "status": "estado", "sheet": "Samtech usuarios",
     "expected_open": ("no_iniciada", "en_progreso"), "color": "#198754"},
    {"entity": "desviaciones", "label": "Desviaciones clientes", "date": "fecha", "date_label": "Fecha",
     "status": None, "sheet": "Desviaciones clientes", "expected_open": (), "color": "#B86A00"},
)
STATUS_LABELS = {
    "abierto": "Abierto", "no_iniciada": "No iniciada", "en_progreso": "En progreso",
    "pendiente": "Pendiente", "cerrado": "Cerrado", "sin_clasificar": "Sin clasificar",
}
STATUS_ALIASES = {
    "abierto": ("abierto", "abierta"),
    "no_iniciada": ("no iniciada", "no iniciado"),
    "en_progreso": ("en progreso", "en proceso", "en curso", "en ejecucion", "iniciado", "iniciada"),
    "pendiente": ("pendiente", "en espera"),
    "cerrado": ("cerrado", "cerrada", "resuelto", "resuelta", "finalizado", "finalizada",
                "completado", "completada", "terminado", "terminada"),
}
OPEN_STATUSES = ("abierto", "no_iniciada", "en_progreso", "pendiente")
STATUS_LOOKUP = {alias: key for key, aliases in STATUS_ALIASES.items() for alias in aliases}
MONTHS = ("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic")
METHOD_NOTE = (
    "Los casos se agrupan por la fecha del registro y, cuando existe ese campo, por su estado actual. "
    "Un caso cerrado permanece en la fecha de origen; no indica el día en que se cerró."
)
RATE_NOTE = (
    "% cerrado = cerrados / (abiertos + cerrados). Los estados sin clasificar se muestran aparte. "
    "Cuando no hay casos clasificables se muestra 0%. El total del período usa los conteos acumulados."
)
WITHOUT_STATUS_NOTE = (
    "Desviaciones clientes no tiene un campo de estado. Sus registros se incluyen en el volumen total "
    "y en Sin clasificar, pero no en abiertos, cerrados ni en el porcentaje de cierre."
)


def classify_status(value):
    normalized = unicodedata.normalize("NFKD", value or "")
    normalized = "".join(char for char in normalized if not unicodedata.combining(char)).casefold()
    normalized = re.sub(r"[\s_-]+", " ", normalized).strip()
    return STATUS_LOOKUP.get(normalized, "sin_clasificar")


def parse_report_range(args):
    dates = []
    for key, label in (("start_date", "Desde"), ("end_date", "Hasta")):
        raw = args.get(key, "").strip()
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw):
            raise ValueError(f"Completa {label} con una fecha válida.")
        try:
            parsed = date.fromisoformat(raw)
        except ValueError:
            raise ValueError(f"Completa {label} con una fecha válida.") from None
        if parsed < date(1900, 1, 1):
            raise ValueError("El informe Excel admite fechas desde el 01/01/1900.")
        dates.append(parsed)
    start, end = dates
    if end < start:
        raise ValueError("Hasta debe ser igual o posterior a Desde.")
    if (end - start).days + 1 > 16382:
        raise ValueError("El rango excede las columnas disponibles en Excel. Selecciona un período más corto.")
    return start, end


def percent_text(value):
    return f"{value * 100:.1f}%".replace(".", ",")


def build_user_report(db, models, start, end, include_details=False):
    days = [start + timedelta(days=offset) for offset in range((end - start).days + 1)]
    day_index = {day: index for index, day in enumerate(days)}
    categories = []
    for source in SOURCES:
        model = models[source["entity"]]
        date_column = getattr(model, source["date"])
        status_column = getattr(model, source["status"]) if source["status"] else None
        scope = (date_column >= start, date_column <= end)
        fields = list_fields(source["entity"], model())
        details = []
        if include_details:
            grouped = Counter()
            for record in db.query(model).filter(*scope).order_by(date_column, model.id).yield_per(1000):
                state = getattr(record, source["status"]) if source["status"] else None
                grouped[(getattr(record, source["date"]), state)] += 1
                details.append({field["name"]: getattr(record, field["name"]) for field in fields})
            groups = [(day, state, count) for (day, state), count in grouped.items()]
        elif status_column is not None:
            groups = db.query(date_column, status_column, func.count(model.id)).filter(*scope).group_by(
                date_column, status_column,
            ).all()
        else:
            groups = [(day, None, count) for day, count in db.query(date_column, func.count(model.id))
                      .filter(*scope).group_by(date_column).all()]
        series = {key: [0] * len(days) for key in STATUS_LABELS}
        unknown = Counter()
        for day, state, count in groups:
            code = classify_status(state)
            series[code][day_index[day]] += count
            if code == "sin_clasificar" and source["status"]:
                unknown[(state or "").strip() or "(vacío)"] += count
        opened = [sum(series[key][i] for key in OPEN_STATUSES) for i in range(len(days))]
        closed, unclassified = series["cerrado"], series["sin_clasificar"]
        counts = {
            "records": [a + c + u for a, c, u in zip(opened, closed, unclassified)],
            "open": opened, "closed": closed, "unclassified": unclassified,
        }
        totals = {key: sum(values) for key, values in counts.items()}
        denominator = totals["open"] + totals["closed"]
        totals["rate"] = (totals["closed"] / denominator if denominator else 0) if source["status"] else None
        totals["without_status"] = totals["records"] if not source["status"] else 0
        categories.append({
            **source, "series": series, "counts": counts, "totals": totals,
            "unknown_states": [{"state": state, "count": count} for state, count in sorted(unknown.items())],
            "undated": db.query(func.count(model.id)).filter(date_column.is_(None)).scalar(),
            "fields": [{key: field[key] for key in ("name", "label", "kind")} for field in fields],
            "details": details,
        })

    counts = {key: [sum(category["counts"][key][i] for category in categories) for i in range(len(days))]
              for key in ("records", "open", "closed", "unclassified")}
    totals = {key: sum(values) for key, values in counts.items()}
    totals["classified"] = totals["open"] + totals["closed"]
    totals["rate"] = totals["closed"] / totals["classified"] if totals["classified"] else 0
    totals["undated"] = sum(category["undated"] for category in categories)
    totals["without_status"] = sum(category["totals"]["without_status"] for category in categories)
    totals["unknown_status"] = totals["unclassified"] - totals["without_status"]
    rates = [closed / (opened + closed) if opened + closed else 0
             for opened, closed in zip(counts["open"], counts["closed"])]

    def row(key, label, values, total=None, kind="count", refs=()):
        return {"key": key, "label": label, "values": values, "total": sum(values) if total is None else total,
                "kind": kind, "refs": refs}

    record_rows = [row(c["entity"] + "_records", c["label"], c["counts"]["records"]) for c in categories]
    record_rows.append(row("total_records", "Total registros", counts["records"], kind="total",
                           refs=[r["key"] for r in record_rows]))
    open_rows = []
    for category in categories:
        if not category["status"]:
            continue
        statuses = list(category["expected_open"]) + [key for key in OPEN_STATUSES
            if key not in category["expected_open"] and sum(category["series"][key])]
        for key in statuses:
            open_rows.append(row(category["entity"] + "_" + key,
                                 f"{category['label']} ({STATUS_LABELS[key]})", category["series"][key]))
    open_rows.append(row("total_open", "Total abiertos", counts["open"], kind="total", refs=[r["key"] for r in open_rows]))
    closed_rows = [row(c["entity"] + "_closed", f"{c['label']} (Cerrado)", c["counts"]["closed"])
                   for c in categories if c["status"]]
    closed_rows.append(row("total_closed", "Total cerrados", counts["closed"], kind="total", refs=[r["key"] for r in closed_rows]))
    sections = [
        {"key": "records", "title": "Registros", "rows": record_rows},
        {"key": "open", "title": "Casos abiertos (Abierto / No iniciada / En progreso / Pendiente)", "rows": open_rows},
        {"key": "closed", "title": "Casos cerrados", "rows": closed_rows},
    ]
    if totals["unclassified"]:
        unknown_rows = [row(c["entity"] + "_unclassified", c["label"], c["counts"]["unclassified"]) for c in categories]
        unknown_rows.append(row("total_unclassified", "Total sin clasificar", counts["unclassified"],
                                kind="total", refs=[r["key"] for r in unknown_rows]))
        sections.append({"key": "unclassified", "title": "Registros sin clasificación de estado", "rows": unknown_rows})
    comparison = [
        row("comparison_open", "Total abiertos", counts["open"], refs=["total_open"]),
        row("comparison_closed", "Total cerrados", counts["closed"], refs=["total_closed"]),
    ]
    if totals["unclassified"]:
        comparison.append(row("comparison_unclassified", "Total sin clasificar", counts["unclassified"], refs=["total_unclassified"]))
    comparison.append(row("closure_rate", "% cerrado", rates, total=totals["rate"], kind="percent"))
    sections.append({"key": "comparison", "title": "Comparación: abiertos vs. cerrados", "rows": comparison})
    report = {
        "start": start, "end": end, "days": days,
        "day_labels": [f"{day.day}-{MONTHS[day.month - 1]}" for day in days],
        "categories": categories, "counts": counts, "totals": totals, "sections": sections,
        "generated_at": datetime.now(timezone.utc), "method_note": METHOD_NOTE, "rate_note": RATE_NOTE,
        "without_status_note": WITHOUT_STATUS_NOTE,
    }
    report["conclusions"] = make_conclusions(report)
    return report


def make_conclusions(report):
    totals, categories = report["totals"], report["categories"]
    conclusions = []

    def cases(count):
        return f"{count} {'caso' if count == 1 else 'casos'}"

    if not totals["records"]:
        conclusions.append("No hay registros con fecha dentro del rango seleccionado.")
    else:
        conclusions.append(f"Se {'registró' if totals['records'] == 1 else 'registraron'} {cases(totals['records'])}: "
                           f"{totals['open']} {'abierto' if totals['open'] == 1 else 'abiertos'}, "
                           f"{totals['closed']} {'cerrado' if totals['closed'] == 1 else 'cerrados'} y {totals['unclassified']} sin clasificar.")
        if totals["classified"]:
            conclusions.append(f"El cierre es {percent_text(totals['rate'])} entre {cases(totals['classified'])} "
                               "con estado clasificable, considerando su estado actual.")
        else:
            conclusions.append("No hay estados clasificables para evaluar el porcentaje de cierre.")
        maximum = max(c["totals"]["records"] for c in categories)
        leaders = [c["label"] for c in categories if c["totals"]["records"] == maximum]
        if len(leaders) == 1:
            conclusions.append(f"La categoría con más registros es {leaders[0]}: {cases(maximum)} "
                               f"({percent_text(maximum / totals['records'])} del total).")
        else:
            conclusions.append(f"{', '.join(leaders)} comparten el mayor volumen, con {cases(maximum)} "
                               f"por categoría ({percent_text(maximum / totals['records'])} del total cada una).")
        if totals["open"]:
            maximum = max(c["totals"]["open"] for c in categories)
            leaders = [c["label"] for c in categories if c["totals"]["open"] == maximum]
            conclusions.append(f"Prioridad de seguimiento por cantidad de abiertos: {', '.join(leaders)} "
                               f"({maximum} casos abiertos por categoría).")
        peak = max(report["counts"]["records"])
        peak_days = [day.strftime("%d/%m/%Y") for day, count in zip(report["days"], report["counts"]["records"]) if count == peak]
        dates = ", ".join(peak_days[:4]) + (f" y {len(peak_days) - 4} días más" if len(peak_days) > 4 else "")
        conclusions.append(f"El mayor volumen diario fue de {cases(peak)}, registrado en {dates}.")
    if totals["unknown_status"]:
        conclusions.append(f"Revisar estados: {cases(totals['unknown_status'])} del período "
                           f"{'requiere' if totals['unknown_status'] == 1 else 'requieren'} clasificación. "
                           "Los estados vacíos o no reconocidos quedan fuera del porcentaje de cierre.")
    if totals["without_status"]:
        conclusions.append(f"Desviaciones clientes aporta {cases(totals['without_status'])} al total. "
                           "Este registro no tiene campo de estado; aparece en Sin clasificar y queda fuera del porcentaje de cierre.")
    if totals["undated"]:
        conclusions.append(f"Hay {totals['undated']} {'registro' if totals['undated'] == 1 else 'registros'} "
                           "sin fecha de referencia en el histórico de estas categorías, fuera del rango. "
                           "Completa las fechas faltantes para incluirlos en el período correspondiente.")
    return conclusions
