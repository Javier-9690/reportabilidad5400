"""Reemplazo confirmado de la carga general o, por separado, de solicitudes QR."""

from datetime import timedelta
import hashlib
import json
from pathlib import PurePosixPath
import secrets

from flask import Blueprint, current_app, flash, make_response, redirect, render_template, request, session, url_for
from itsdangerous import BadData, URLSafeTimedSerializer
from sqlalchemy import text
from sqlalchemy.exc import SQLAlchemyError

from gestion5s.editing import ORDER_IMPORT_FIELDS, SAMTECH_USER_FIELDS, record_version
from gestion5s.orders import (MAX_UPLOAD_BYTES, ORDER_ENTITIES, ORDER_TITLES,
                             ORDER_REPORT_NOTE, order_import_targets, parse_orders)
from gestion5s.user_reports import classify_status, STATUS_LABELS

order_imports = Blueprint("order_imports", __name__)
PREVIEW_SECONDS = 1800


def _web():
    from gestion5s import web
    return web


def _error(message, tab="solicitud_ot", status=400):
    response = make_response(render_template("order_import_error.html", message=message, tab=tab), status)
    response.headers["Cache-Control"] = "no-store"
    return response


def _csrf_valid():
    expected = session.get("hotel_orders_csrf", "")
    received = request.form.get("csrf_token", "")
    return bool(expected) and secrets.compare_digest(expected.encode(), received.encode())


def _owner():
    return hashlib.sha256(session["hotel_orders_csrf"].encode()).hexdigest()


def _serializer():
    return URLSafeTimedSerializer(current_app.secret_key, salt="hotel-orders-replacement")


def orders_snapshot(db, targets=ORDER_ENTITIES):
    """Detecta altas, bajas y ediciones desde la vista previa, incluso sin cambiar el total."""
    web = _web()
    result = {}
    for entity in targets:
        model = web.ENTITY_MODEL[entity]
        digest, count = hashlib.sha256(), 0
        for record in db.query(model).order_by(model.id).yield_per(1000):
            digest.update(record_version(record).encode())
            count += 1
        result[entity] = {"count": count, "digest": digest.hexdigest()}
    return result


def preview_order_import(entity):
    targets = order_import_targets(entity)
    if not _csrf_valid():
        return _error("La sesión del formulario expiró. Vuelve al formulario y selecciona el archivo nuevamente.", entity)
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename or not uploaded.filename.lower().endswith(".xlsx"):
        return _error("Selecciona el archivo de órdenes en formato .xlsx.", entity)
    content = uploaded.read(MAX_UPLOAD_BYTES + 1)
    try:
        parsed = parse_orders(content, entity=entity)
    except ValueError as exc:
        return _error(str(exc), entity)
    except Exception:
        current_app.logger.exception("No se pudo leer el archivo de órdenes")
        return _error("No se pudo leer el Excel. Revisa el archivo y vuelve a cargarlo.", entity)
    web = _web()
    with web.SessionLocal() as db:
        try:
            snapshot = orders_snapshot(db, targets)
            owner = _owner()
            db.query(web.OrderImportBatch).filter(
                (web.OrderImportBatch.expires_at < web.now_utc()) |
                ((web.OrderImportBatch.owner == owner) & web.OrderImportBatch.return_tab.in_(targets))
            ).delete(synchronize_session=False)
            batch = web.OrderImportBatch(
                id=secrets.token_hex(32), owner=owner,
                filename=PurePosixPath(uploaded.filename.replace("\\", "/")).name[:200],
                content=content, snapshot=json.dumps(snapshot), return_tab=entity,
                expires_at=web.now_utc() + timedelta(seconds=PREVIEW_SECONDS),
            )
            db.add(batch)
            db.commit()
            token = _serializer().dumps({"id": batch.id, "owner": owner})
            status_rows = [
                {"entity": key, "state": state, "count": count,
                 "group": STATUS_LABELS[classify_status(state, entity=key)]}
                for key in targets for state, count in sorted(parsed["states"][key].items())
            ]
            response = make_response(render_template(
                "order_import_preview.html", parsed=parsed, previous=snapshot, filename=batch.filename,
                confirmation_token=token, csrf_token=session["hotel_orders_csrf"],
                columns=SAMTECH_USER_FIELDS if entity == "samtech_qr" else ORDER_IMPORT_FIELDS,
                report_note=ORDER_REPORT_NOTE, status_rows=status_rows, qr_import=entity == "samtech_qr",
                tab=entity, titles=ORDER_TITLES,
            ))
            response.headers["Cache-Control"] = "no-store"
            return response
        except SQLAlchemyError:
            db.rollback()
            current_app.logger.exception("No se pudo preparar la importación de órdenes")
            return _error("No se pudo preparar la vista previa. Los registros siguen intactos. Inténtalo nuevamente.", entity, 503)


def _payload():
    if not _csrf_valid():
        raise ValueError("La sesión expiró. Vuelve a cargar el archivo para revisar una nueva vista previa.")
    try:
        payload = _serializer().loads(request.form.get("confirmation_token", ""), max_age=PREVIEW_SECONDS)
    except BadData:
        raise ValueError("La confirmación es inválida o expiró. Vuelve a cargar el archivo.") from None
    if payload.get("owner") != _owner():
        raise ValueError("Esta vista previa pertenece a otra sesión. Vuelve a cargar el archivo.")
    return payload


def _insert_groups(db, models, groups):
    # Lotes acotados: el Excel adjunto contiene más de 28.000 órdenes.
    for entity in groups:
        for offset in range(0, len(groups[entity]), 500):
            db.execute(models[entity].__table__.insert(), groups[entity][offset:offset + 500])


@order_imports.post("/import/ordenes/confirm")
def confirm():
    try:
        payload = _payload()
    except ValueError as exc:
        return _error(str(exc))
    if request.form.get("confirm_replace") != "1":
        return _error("Marca la confirmación para reemplazar los registros indicados en la vista previa. Todavía no se ha modificado ningún registro.")
    web = _web()
    tab = "solicitud_ot"
    with web.SessionLocal() as db:
        try:
            if db.get_bind().dialect.name == "sqlite":
                db.execute(text("BEGIN IMMEDIATE"))
            batch = db.query(web.OrderImportBatch).filter_by(id=payload["id"], owner=_owner()).with_for_update().first()
            if not batch or batch.expires_at < web.now_utc():
                return _error("Esta vista previa ya se utilizó, se canceló o expiró. Vuelve a cargar el archivo.", status=409)
            tab = batch.return_tab
            targets = order_import_targets(tab)
            parsed = parse_orders(batch.content, entity=tab)
            if db.get_bind().dialect.name == "postgresql":
                # Bloquea también nuevas inserciones hasta terminar la transacción.
                tables = ", ".join(web.ENTITY_MODEL[key].__tablename__ for key in targets)
                db.execute(text(f"LOCK TABLE {tables} IN SHARE ROW EXCLUSIVE MODE"))
            if orders_snapshot(db, targets) != json.loads(batch.snapshot):
                return _error("Los registros cambiaron desde la vista previa. No se reemplazó nada. Vuelve a cargar el archivo para revisar las cantidades actuales.", tab, 409)
            for entity in targets:
                db.query(web.ENTITY_MODEL[entity]).delete(synchronize_session=False)
            _insert_groups(db, web.ENTITY_MODEL, parsed["groups"])
            db.delete(batch)  # La misma confirmación no puede ejecutarse dos veces.
            db.commit()
        except ValueError as exc:
            db.rollback()
            return _error(str(exc), tab)
        except SQLAlchemyError:
            db.rollback()
            current_app.logger.exception("Se revirtió el reemplazo de órdenes")
            return _error("No se pudo completar la importación. Se conservaron todos los registros anteriores. Vuelve a cargar el archivo e inténtalo nuevamente.", tab, 503)
    summary = " y ".join(f"{len(parsed['groups'][key]):,} {ORDER_TITLES[key]}" for key in targets).replace(",", ".")
    flash(f"Reemplazo completado: {summary}. El reporte de Gestión de usuarios ya utiliza los datos nuevos.", "success")
    return redirect(url_for("registros", vista=tab))


@order_imports.post("/import/ordenes/cancel")
def cancel():
    try:
        payload = _payload()
    except ValueError as exc:
        return _error(str(exc))
    web = _web()
    tab = "solicitud_ot"
    with web.SessionLocal() as db:
        try:
            if db.get_bind().dialect.name == "sqlite":
                db.execute(text("BEGIN IMMEDIATE"))
            batch = db.query(web.OrderImportBatch).filter_by(id=payload["id"], owner=_owner()).with_for_update().first()
            if batch:
                tab = batch.return_tab
                db.delete(batch)
                db.commit()
        except SQLAlchemyError:
            db.rollback()
            return _error("No se pudo cancelar la vista previa. Las bases de registros no se modificaron.", tab, 503)
    flash("Importación cancelada. Los registros anteriores se conservaron.", "info")
    return redirect(url_for("panel", tab=tab))


@order_imports.after_request
def no_store(response):
    response.headers["Cache-Control"] = "no-store"
    return response
