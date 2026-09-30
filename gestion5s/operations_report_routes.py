"""Pantalla y descargas del reporte de desviaciones, solicitudes y reclamos."""

from datetime import date

from flask import Blueprint, current_app, make_response, render_template, request, send_file
from sqlalchemy import text
from sqlalchemy.exc import SQLAlchemyError

from gestion5s.editing import display_record_value
from gestion5s.operations_reports import (REPORT_TITLE, SOURCES, SOURCE_MAP, build_operations_report,
                                          chart_data, parse_filters)

bp = Blueprint("operations_reports", __name__)
BASE_PATH = "/reports/desviaciones-solicitudes-reclamos"


def load_report(filters, export=None):
    from gestion5s import web
    with web.SessionLocal() as db:
        if db.get_bind().dialect.name == "postgresql":
            db.connection(execution_options={"isolation_level": "REPEATABLE READ"})
        elif db.get_bind().dialect.name == "sqlite":
            db.execute(text("BEGIN"))
        return build_operations_report(db, web.ENTITY_MODEL, filters, export=export)


def page_context():
    today = date.today()
    return dict(title=REPORT_TITLE, sources=SOURCES, report=None, selected=None, error=None,
                start_value=request.args.get("start_date", today.replace(day=1).isoformat()),
                end_value=request.args.get("end_date", today.isoformat()),
                search_value=request.args.get("q", ""),
                include_undated=request.args.get("include_undated") == "1",
                active_tab=request.args.get("tab", "general"), display_value=display_record_value)


def html_response(context, status=200):
    response = make_response(render_template("operations_report.html", **context), status)
    response.headers["Cache-Control"] = "no-store"
    return response


@bp.get(BASE_PATH)
def report_page():
    context = page_context()
    try:
        filters = parse_filters(request.args)
        report = load_report(filters)
        selected = next((c for c in report["categories"] if c["entity"] == filters["tab"]), None)
        context.update(report=report, selected=selected, charts=chart_data(report, selected),
                       params=dict(start_date=filters["start"].isoformat(), end_date=filters["end"].isoformat(),
                                   q=filters["search"], include_undated="1" if filters["include_undated"] else "0"))
    except ValueError as exc:
        context["error"] = str(exc)
        return html_response(context, 400)
    except SQLAlchemyError:
        current_app.logger.exception("Error al consultar el reporte de desviaciones, solicitudes y reclamos")
        context["error"] = "No se pudo consultar el reporte. Inténtalo nuevamente."
        return html_response(context, 503)
    return html_response(context)


@bp.get(BASE_PATH + ".xlsx", defaults={"category": "general"})
@bp.get(BASE_PATH + "/<category>.xlsx")
def export_excel(category):
    from gestion5s.operations_report_excel import export_operations_report
    context = page_context()
    if category not in ("general", *SOURCE_MAP):
        context["error"] = "La pestaña solicitada no existe."
        return html_response(context, 404)
    try:
        filters = parse_filters(request.args)
        report = load_report(filters, export=category)
        content = export_operations_report(report)
    except (ValueError, SQLAlchemyError) as exc:
        if isinstance(exc, ValueError):
            context["error"], status = str(exc), 400
        else:
            current_app.logger.exception("Error al exportar el reporte de desviaciones, solicitudes y reclamos")
            context["error"], status = "No se pudo exportar el reporte. Inténtalo nuevamente.", 503
        return html_response(context, status)
    response = send_file(content, as_attachment=True,
                         download_name=f"reporte_desviaciones_solicitudes_reclamos_{category}_{report['start']}_{report['end']}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response.headers["Cache-Control"] = "no-store"
    return response
