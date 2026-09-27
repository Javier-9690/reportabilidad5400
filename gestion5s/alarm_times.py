"""HH:MM:SS para alarmas, conservando las columnas históricas en horas."""

from datetime import datetime, time, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import math
import re


ALARM_HOUR_FIELDS = (
    "aviso_mantencion_h", "llegada_mantencion_h", "aviso_lider_h", "llegada_lider_h",
)
ALARM_TIME_FIELDS = (*ALARM_HOUR_FIELDS, "hora_reporte_salfa")
MAX_ALARM_HOURS = 999999999
TIME_ERROR = "Usa horas:minutos:segundos, por ejemplo 08:15:30. Minutos y segundos deben estar entre 00 y 59."


def format_alarm_hours(value):
    if value is None or value == "":
        return ""
    hours = Decimal(str(value))
    if not hours.is_finite() or hours < 0:
        return str(value)  # Un dato histórico inválido permanece visible para corregirlo.
    seconds = int((hours * 3600).to_integral_value(rounding=ROUND_HALF_UP))
    hours, rest = divmod(seconds, 3600)
    minutes, seconds = divmod(rest, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def parse_alarm_hours(value, allow_decimal=False):
    """Acepta horas Excel y HH:MM[:SS]; los decimales antiguos siguen siendo horas."""
    if value is None or isinstance(value, str) and not value.strip():
        return None
    if isinstance(value, bool):
        raise ValueError(TIME_ERROR)
    if isinstance(value, datetime):
        value = value.time()
    if isinstance(value, time):
        if value.tzinfo is not None:
            raise ValueError(TIME_ERROR)
        seconds = value.hour * 3600 + value.minute * 60 + value.second
        seconds += value.microsecond / 1000000
        hours = seconds / 3600
    elif isinstance(value, timedelta):
        hours = value.total_seconds() / 3600
    else:
        text = str(value).strip()
        match = re.fullmatch(r"([0-9]{1,9}):([0-5][0-9])(?::([0-5][0-9]))?", text)
        if match:
            h, m, s = match.groups()
            hours = (int(h) * 3600 + int(m) * 60 + int(s or 0)) / 3600
        elif allow_decimal:
            try:
                hours = float(Decimal(text.replace(",", ".")))
            except (InvalidOperation, ValueError, OverflowError):
                raise ValueError(TIME_ERROR) from None
        else:
            raise ValueError(TIME_ERROR)
    if not math.isfinite(hours) or not 0 <= hours < MAX_ALARM_HOURS + 1:
        raise ValueError(TIME_ERROR)
    return hours


def parse_alarm_clock(value):
    """Hora del día para SALFA, con segundos y compatibilidad con texto AM/PM."""
    text = str(value).strip()
    try:
        result = time.fromisoformat(text)
    except ValueError:
        result = None
        for pattern in ("%H:%M:%S", "%H:%M", "%I:%M:%S %p", "%I:%M %p", "%I:%M:%S%p", "%I:%M%p"):
            try:
                result = datetime.strptime(text, pattern).time()
                break
            except ValueError:
                continue
        if result is None:
            raise ValueError(TIME_ERROR) from None
    if result.tzinfo is not None:
        raise ValueError("Ingresa una hora local.")
    return result
