"""Búsqueda literal en los campos vigentes de una categoría de registros."""

import math

from sqlalchemy import BigInteger, Date, DateTime, Float, Integer, Numeric, String, case, cast, func, literal, or_

from gestion5s.editing import list_fields
from gestion5s.alarm_times import MAX_ALARM_HOURS


MAX_SEARCH_LENGTH = 200
ACCENTED = "ÁÉÍÓÚÜÑáéíóúüñ\u0301\u0303\u0308"
PLAIN = "AEIOUUNaeiouun"
SEARCH_TRANSLATION = {ord(char): PLAIN[index] if index < len(PLAIN) else ""
                      for index, char in enumerate(ACCENTED)}


def clean_search_term(value):
    if not isinstance(value, str) or "\x00" in value or len(value.strip()) > MAX_SEARCH_LENGTH:
        raise ValueError("La búsqueda no es válida. Escribe hasta 200 caracteres.")
    value = value.strip()
    if value and not value.translate(SEARCH_TRANSLATION).strip():
        raise ValueError("Escribe una palabra o número para buscar.")
    return value


def normalized_search_text(expression, dialect):
    if dialect == "postgresql":
        return func.lower(func.translate(expression, ACCENTED, PLAIN))
    # SQLite no necesita extensiones ni funciones instaladas en la base de datos.
    for char in ACCENTED:
        expression = func.replace(expression, char, SEARCH_TRANSLATION[ord(char)])
    return func.lower(expression)


def duration_text(column):
    seconds = column % 60
    minutes = cast((column - seconds) / 60, Integer)

    def padded(value):
        text = cast(value, String)
        return case((value < 10, literal("0") + text), else_=text)

    return padded(minutes) + literal(":") + padded(seconds)


def alarm_hours_text(column):
    # El límite evita convertir valores históricos no finitos a enteros en PostgreSQL.
    safe = case(((column >= 0) & (column < MAX_ALARM_HOURS + 1), column), else_=None)
    # PostgreSQL redondea NUMERIC con la misma regla que format_alarm_hours.
    total = cast(func.round(cast(safe, Numeric) * 3600), BigInteger)
    seconds = total % 60
    minutes_total = cast((total - seconds) / 60, BigInteger)
    minutes = minutes_total % 60
    hours = cast((minutes_total - minutes) / 60, BigInteger)

    def padded(value):
        text = cast(value, String)
        return case((value < 10, literal("0") + text), else_=text)

    return padded(hours) + literal(":") + padded(minutes) + literal(":") + padded(seconds)


def record_search_condition(entity, model, search, dialect):
    """Devuelve la misma condición para listado, exportación y borrado confirmado."""
    needle = search.translate(SEARCH_TRANSLATION).lower()
    conditions = []
    for field in list_fields(entity, model()):
        column = getattr(model, field["name"])
        if field["kind"] == "duration":
            value = duration_text(column)
        elif field["kind"] == "hours_hms":
            value = alarm_hours_text(column)
        else:
            value = cast(column, String)
        if isinstance(column.type, String):
            value = normalized_search_text(value, dialect)
        # %, _ y / se buscan como caracteres, nunca como comodines SQL.
        conditions.append(value.contains(needle, autoescape=True))
        if isinstance(column.type, (Date, DateTime)):
            formatted = (func.to_char(column, "DD/MM/YYYY") if dialect == "postgresql"
                         else func.strftime("%d/%m/%Y", column))
            conditions.append(formatted.contains(needle, autoescape=True))
        elif isinstance(column.type, Float):
            try:
                number = float(needle.replace(",", "."))
            except ValueError:
                continue
            if math.isfinite(number):
                conditions.append(column == number)
    return or_(*conditions)
