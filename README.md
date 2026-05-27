# Dotación Reportes

Aplicación Flask para importar una curva planificada y censos diarios, cruzar IDs contra la curva activa y generar el reporte **Resumen de Dotación por Gerencia**.

## Stack

- Flask
- SQLAlchemy
- PostgreSQL en Render
- Pandas + openpyxl + pyxlsb para importación Excel
- Bootstrap + Bootstrap Icons

## Desarrollo local

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
python app.py
```

## Render + PostgreSQL

El proyecto incluye `render.yaml`, `Procfile` y soporte para `DATABASE_URL`. Los archivos importados se guardan como binario en PostgreSQL para no depender del filesystem efímero del servidor.

## Flujo

1. Importar curva planificada usando hoja `Fcst_5400`.
2. Importar censos diarios `.xlsb` o `.xlsx`.
3. Generar reporte por rango de fechas.
4. Exportar Excel con estética roja similar al reporte enviado.


## Corrección importante para Render

Si en los logs aparece `sqlite3.OperationalError: unable to open database file`, significa que la app no recibió `DATABASE_URL` de PostgreSQL y cayó al modo SQLite local.

Para Render usa una de estas opciones:

### Opción recomendada: Blueprint

1. Sube este proyecto a GitHub.
2. En Render selecciona **New +** → **Blueprint**.
3. Selecciona el repositorio.
4. Render leerá `render.yaml`, creará el servicio web y también la base PostgreSQL `dotacion-reportes-db`.
5. Verifica que el servicio web tenga la variable `DATABASE_URL`.

### Opción manual

1. Crea una base de datos PostgreSQL en Render.
2. Copia su **Internal Database URL**.
3. En el Web Service agrega la variable de entorno `DATABASE_URL` con ese valor.
4. Reinicia el servicio.

El fallback SQLite queda solo para desarrollo local. En Render, PostgreSQL es obligatorio si quieres conservar los censos y curvas importadas.
