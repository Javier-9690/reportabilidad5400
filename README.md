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
