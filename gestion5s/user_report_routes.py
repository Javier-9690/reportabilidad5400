"""Rutas del reporte en el menú principal, con acceso a las tablas de hotelería."""

from flask import Blueprint, current_app, make_response, render_template, request, send_file
from sqlalchemy import text
from sqlalchemy.exc import SQLAlchemyError

from gestion5s.user_reports import SOURCES, STATUS_ALIASES, STATUS_LABELS, build_user_report, parse_report_range


bp = Blueprint("user_reports", __name__)


def report_context():
    return {
        "start_value": request.args.get("start_date", ""), "end_value": request.args.get("end_date", ""),
        "report": None, "error": None, "sources": SOURCES, "status_aliases": STATUS_ALIASES,
        "status_labels": STATUS_LABELS,
    }


def load_report(include_details=False):
    # Importación diferida: conserva la instancia y conexión del módulo montado.
    from gestion5s import web
    start, end = parse_report_range(request.args)
    with web.SessionLocal() as db:
        # Todas las categorías deben ver la misma carga aunque otro worker
        # confirme una importación mientras se está generando este reporte.
        if db.get_bind().dialect.name == "postgresql":
            db.connection(execution_options={"isolation_level": "REPEATABLE READ"})
        elif db.get_bind().dialect.name == "sqlite":
            db.execute(text("BEGIN"))
        return build_user_report(db, web.ENTITY_MODEL, start, end, include_details=include_details)


@bp.get("/reports/usuarios")
def report_page():
    context, status = report_context(), 200
    if "start_date" in request.args or "end_date" in request.args:
        try:
            context["report"] = load_report()
        except ValueError as exc:
            context["error"], status = str(exc), 400
        except SQLAlchemyError:
            current_app.logger.exception("Error al consultar el reporte de usuarios")
            context["error"], status = "No se pudo consultar el reporte. Inténtalo nuevamente.", 503
    response = make_response(render_template("user_report.html", **context), status)
    response.headers["Cache-Control"] = "no-store"
    return response


@bp.get("/reports/usuarios.xlsx")
def export_excel():
    from gestion5s.user_report_excel import export_user_report
    try:
        report = load_report(include_details=True)
        content = export_user_report(report)
    except (ValueError, SQLAlchemyError) as exc:
        context = report_context()
        if isinstance(exc, ValueError):
            context["error"], status = str(exc), 400
        else:
            current_app.logger.exception("Error al exportar el reporte de usuarios")
            context["error"], status = "No se pudo exportar el reporte. Inténtalo nuevamente.", 503
        return render_template("user_report.html", **context), status
    response = send_file(content, as_attachment=True,
                         download_name=f"reporte_usuarios_{report['start']}_{report['end']}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response.headers["Cache-Control"] = "no-store"
    return response
