"""Lectura del CSV de registros, incluido su indicador de separador para Excel."""

import csv
from io import StringIO


def excel_csv_reader(payload):
    stream = StringIO(payload.decode("utf-8-sig"), newline="")
    if stream.readline() != "sep=;\r\n":
        raise AssertionError("La descarga debe declarar el separador para Excel")
    return csv.DictReader(stream, delimiter=";")
