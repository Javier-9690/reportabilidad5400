# Reportabilidad 5400 — acceso administrativo

Este repositorio ya contiene todos los archivos integrados. No necesitas editar código ni ejecutar un instalador.

## Configuración obligatoria en Render

- **ADMIN_PASSWORD**: define la contraseña de administrador acordada, en **Environment** (no la publiques en GitHub).
- **SECRET_KEY**: conserva un valor secreto largo, aleatorio y estable entre despliegues. El `render.yaml` genera este valor al crear el servicio.
- Mantén el comando actual: `gunicorn app:app --workers 1 --threads 2 --timeout 300 --graceful-timeout 300`.

Sin `ADMIN_PASSWORD` o con `SECRET_KEY` insegura el sistema no inicia, por seguridad.

## Uso

Ingresa a `/login` con usuario `admin` y la contraseña definida en Render. Las páginas y API requieren sesión. Antes de eliminar archivos importados, registros individuales, lotes, o reemplazar órdenes, el sistema vuelve a pedir la contraseña. Se valida en servidor, además de mostrar el diálogo en el navegador. El cierre de sesión aparece en ambos módulos.

## Precauciones

Respalda PostgreSQL antes de sustituir el sistema. Este paquete no incluye la base de datos ni credenciales de producción.
