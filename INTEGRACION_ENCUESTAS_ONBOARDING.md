# Reportabilidad 5400: módulos públicos integrados

El ZIP incluye **todos los archivos** de Reportabilidad 5400, Camp Checklist, el motor original de reportabilidad de Encuesta 5400 y las 12 imágenes originales del reglamento Onboarding.

## Entradas sin contraseña
- `/encuesta`: cinco preguntas con escala 1 a 5 + comentarios.
- `/sugerencias`: sugerencias y reclamos.
- `/onboarding`: 12 páginas del reglamento, aceptación, formulario y firma táctil. Genera PDF y lo guarda en PostgreSQL.

## Acceso administrativo
- `/encuestas`: dashboard histórico, reportes por período, Excel/PDF/CSV, y agregación de encuestas importadas.
- `/encuestas/registros`, `/sugerencias/registros`, `/onboarding/registros`: consultas y borrado con clave por operación.
- `/onboarding/pdf/<id>`: comprobantes almacenados como binarios en PostgreSQL.

## Almacenamiento
Todas las tablas nuevas se crean con `CREATE TABLE IF NOT EXISTS` vía `db.create_all()`; **no se eliminan ni migran datos antiguos**.
- `public_survey_5400`
- `public_feedback_5400`
- `public_onboarding_5400`
Checklist conserva prioridad por `CHECKLIST_DATABASE_URL` y, si no existe, utiliza `DATABASE_URL`. Esto permite usar una sola base `kpi-db` para los tres módulos si `DATABASE_URL` apunta a la base existente. Los datos nuevos se guardan en `DATABASE_URL`.

## Variables Render
Conservar `SECRET_KEY` y `ADMIN_PASSWORD`. Configurar `DATABASE_URL` con la cadena existente de `kpi-db` si es la única base. `CHECKLIST_DATABASE_URL` es opcional cuando usa la misma conexión, pero debe conservarse si Checklist tiene una base distinta. **No sustituir** `DATABASE_URL` por `kpi-db` sin comprobar el esquema y el histórico. Para aprovechar la misma `kpi-db` en ambos servicios, se puede establecer la misma cadena en ambas variables tras verificar compatibilidad.

## Consideraciones
Las respuestas del nuevo formulario se procesan con los indicadores históricos del proyecto Encuesta 5400 (semanas cerradas lunes-domingo; cumplimiento 4.5; excelencia 4.8). Las encuestas existentes en SQLite/CSV de otras instalaciones no se migran automáticamente: se pueden importar en el panel con Excel, en modo **solo agregar**. Respaldar antes de desplegar. Verificar procesamiento en Render y reglas de protección de datos personales antes de producción.
