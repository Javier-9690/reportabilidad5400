# Camp Checklist integrado — datos existentes intactos

## Integración
- Nueva navegación **Checklist campamento**, ruta `/camp-checklist/`.
- Conserva todos los módulos originales de Reportabilidad 5400 y Registros Hotelería.
- Camp Checklist mantiene sus propias tablas `rooms`, `receptionists` y `checklists` en **su base de datos PostgreSQL original**, en lugar de usar la BD principal.
- No se incluye ni se modifica el contenido de la base de datos: el ZIP de código no trae los datos guardados en Render.
- El importador de Camp Checklist no borra registros: `db.create_all()` solo crea tablas si faltan; la carga de habitaciones iniciales solo se hace si no existen habitaciones.
- El módulo adopta el inicio de sesión administrativo de Reportabilidad y solicita la clave `ADMIN_PASSWORD` también en las acciones de eliminación (checklists, módulos y desactivación de recepcionistas).

## Configuración OBLIGATORIA en Render, antes del deploy
1. En servicio **camp-checklist** original, encuentra la variable `DATABASE_URL` y/o el `Internal Database URL` de PostgreSQL (`kpi-db`). Copia la URL de la **misma base existente**; no crees otra. Mantén el servicio y la BD originales.
2. En **dotacion-reportes** > Environment, agrega `CHECKLIST_DATABASE_URL` con esa URL. Nunca la publiques en GitHub.
3. Mantén `DATABASE_URL` original del servicio dotacion-reportes (para las tablas de reportabilidad). Mantén `SECRET_KEY` estable y agrega `ADMIN_PASSWORD`.
4. Despliega. La nueva ruta es `/camp-checklist/` y forma parte de la navegación.

**No omitas CHECKLIST_DATABASE_URL**. Si falta, la aplicación falla intencionalmente al iniciar para evitar crear una base vacía que aparentaría pérdida de registros.

La URL interna de PostgreSQL solo sirve si ambos servicios pueden alcanzar la base de datos dentro de la misma red/entorno de Render. Si no, usa la URL externa de la base original con SSL, según la configuración del proveedor.

## Respaldo
Antes de cambiar el despliegue, exporta o respalda la base original de Camp Checklist. No borres ni reemplaces la base `kpi-db`.

## Comportamiento de rutas
La app Camp Checklist se monta como una aplicación WSGI independiente en `/camp-checklist`. Sus plantillas usan `request.script_root` para que fetch/API/navegación sigan funcionando con el nuevo prefijo.
