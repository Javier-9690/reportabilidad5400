# Dotación Reportes 5400

Aplicación Flask preparada para Render.com + PostgreSQL.

## Funcionalidades

- Importar curva planificada desde hoja `Fcst_5400`.
- Importar censos `.xlsb` o `.xlsx`.
- Guardar archivos y datos derivados en PostgreSQL.
- Eliminar archivos subidos y sus datos derivados.
- Revisar registros de censo sin match contra la curva.
- Proponer correcciones automáticas de ID usando similitud, dígitos, sufijos y coincidencias de empresa/gerencia.
- Aplicar corrección automática o corrección manual por ID exacto de curva.
- Generar reporte de dotación por gerencia.
- Exportar reporte a Excel.

## Render

Build command:

```bash
pip install -r requirements.txt
```

Start command:

```bash
gunicorn app:app --workers 1 --threads 2 --timeout 300 --graceful-timeout 300
```

Variables recomendadas:

```text
DATABASE_URL=<Internal Database URL de PostgreSQL>
SECRET_KEY=<valor-secreto>
MAX_CONTENT_LENGTH_MB=80
PYTHON_VERSION=3.12.7
```

Después de cambiar `requirements.txt`, usar:

```text
Manual Deploy -> Clear build cache & deploy
```

## Nota sobre eliminación

- Eliminar un censo borra sus registros asociados.
- Eliminar una curva borra sus IDs y planificación. Los registros de censo que estaban cruzados con esa curva quedan como `sin match`, para que puedan corregirse contra otra curva activa o histórica.

## Optimización Sin Match

- La pantalla **Sin match** ahora agrupa por ID de censo para evitar repetir el mismo error muchas veces.
- Las propuestas ya no comparan cada registro contra toda la curva. Se usa un índice por ID normalizado, dígitos y sufijos.
- Al aplicar una propuesta o una corrección manual, se actualizan todos los registros sin match que tengan el mismo ID de censo.

## Actualización: visualización y carga manual de curva

Esta versión agrega:

- Pantalla `/curva` para visualizar los datos de la curva por pantalla.
- Filtros por versión de curva, gerencia y búsqueda libre.
- Paginación de IDs de curva.
- Acción para editar un ID existente de la curva.
- Pantalla `/curva/nuevo` para agregar un nuevo ID manualmente.
- Opción para cargar planificación diaria simple por rango de fechas.
- Desde `Sin match`, botón para crear el ID no encontrado directamente en la curva.
- Al crear o editar un ID, se puede corregir automáticamente todos los registros sin match que tengan ese mismo ID.

## Cambio: Detalle Censos con ID corregido

En la exportación Excel, la pestaña **Detalle Censos** conserva el formato base del censo importado, pero ahora muestra en la columna **Id** el ID corregido/aplicado en la curva cuando exista una corrección realizada desde **Sin match**. Además, cuando el registro tiene match con curva, se alinean Gerencia, AREA y Empresa con la curva para que el detalle concuerde con el reporte gerencial.


## Corrección filtros EGP/F&A
Se corrigió el parseo de fechas HTML (`YYYY-MM-DD`) para evitar inversión de día/mes en rangos como junio 2026.
