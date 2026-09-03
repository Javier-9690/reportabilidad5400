"""Migra una base 5S antigua a la base de Reportabilidad 5400.

Variables requeridas:
    SOURCE_5S_DATABASE_URL  Base de datos del sistema 5S anterior.
    DATABASE_URL            Base de datos del sistema integrado.

La migración se detiene antes de escribir si encuentra datos 5S en el destino.
"""

import os
import sys
from datetime import datetime, timezone
from pathlib import Path

from sqlalchemy import MetaData, Table, create_engine, func, inspect, select, text
from sqlalchemy.engine import make_url


source_url = os.environ.get("SOURCE_5S_DATABASE_URL", "").strip()
target_url = os.environ.get("DATABASE_URL", "").strip()

if not source_url or not target_url:
    sys.exit("Debes definir SOURCE_5S_DATABASE_URL y DATABASE_URL.")

if make_url(source_url).render_as_string(hide_password=False) == make_url(target_url).render_as_string(hide_password=False):
    sys.exit("El origen y el destino no pueden ser la misma base de datos.")

# El módulo crea las tablas 5S faltantes en DATABASE_URL al importarse.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from gestion5s.web import Base, ENGINE as target_engine, normalize_db_url


source_engine = create_engine(normalize_db_url(source_url), pool_pre_ping=True)
source_tables = set(inspect(source_engine).get_table_names())


def target_has_data(connection):
    populated = []
    for table in Base.metadata.sorted_tables:
        count = connection.scalar(select(func.count()).select_from(table)) or 0
        if count:
            populated.append(f"{table.name} ({count})")
    return populated


def prepare_row(table_name, row):
    data = dict(row)
    if table_name == "cumplimiento_eecc" and not data.get("fecha"):
        created = data.get("creado")
        data["fecha"] = created.date() if isinstance(created, datetime) else datetime.now(timezone.utc).date()
    return data


with target_engine.begin() as target_connection:
    populated = target_has_data(target_connection)
    if populated:
        sys.exit(
            "Migración cancelada: el destino ya contiene datos 5S: "
            + ", ".join(populated)
        )

    migrated = {}
    with source_engine.connect() as source_connection:
        for target_table in Base.metadata.sorted_tables:
            if target_table.name not in source_tables:
                migrated[target_table.name] = 0
                continue

            source_metadata = MetaData()
            source_table = Table(target_table.name, source_metadata, autoload_with=source_engine)
            valid_columns = {column.name for column in target_table.columns}
            result = source_connection.execution_options(stream_results=True).execute(select(source_table))

            count = 0
            while True:
                batch = result.mappings().fetchmany(1000)
                if not batch:
                    break
                values = [
                    {key: value for key, value in prepare_row(target_table.name, row).items() if key in valid_columns}
                    for row in batch
                ]
                if values:
                    target_connection.execute(target_table.insert(), values)
                    count += len(values)
            migrated[target_table.name] = count

    if target_engine.dialect.name == "postgresql":
        for table in Base.metadata.sorted_tables:
            if "id" not in table.c:
                continue
            target_connection.execute(text(
                f"SELECT setval(pg_get_serial_sequence('{table.name}', 'id'), "
                f"COALESCE((SELECT MAX(id) FROM {table.name}), 1), "
                f"EXISTS(SELECT 1 FROM {table.name}))"
            ))

print("Migración 5S completada:")
for table_name, count in migrated.items():
    print(f"- {table_name}: {count} registro(s)")
