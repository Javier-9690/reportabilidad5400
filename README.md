# Reportabilidad 5400 integrada

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
- Acceder al módulo Registros hotelería desde la misma cabecera y el mismo dominio.
- Registrar, importar, consultar, editar, eliminar y exportar los 15 tipos de registros operacionales de hotelería.
- Consultar el dashboard KPI consolidado de hotelería por fechas o semanas.

## Menú principal

Importar · Censos · Sin Match · Curva · Nuevo ID · Reportes · Registros hotelería.

- **Reportes**: Dotación, Ocupabilidad, EGP y F&A.
- **Registros hotelería**: Ingresar registros, Consultar registros y Dashboard KPI.

El Dashboard general no figura en el menú; sigue disponible desde el encabezado.

## Registros hotelería

El módulo antes llamado Gestión 5S se encuentra en estas rutas dentro de la aplicación principal:

- `/gestion-5s/panel`: ingreso manual e importación Excel.
- `/gestion-5s/registros`: consulta, eliminación y descarga CSV.
- `/gestion-5s/edit/<entidad>/<id>`: edición del registro desde el icono de lápiz junto a la papelera.
- `/gestion-5s/dashboard`: indicadores y gráficos.

En **Consultar registros**, el lápiz abre un formulario con los datos actuales.
**Guardar cambios** actualiza el mismo registro y vuelve al listado con sus
filtros; **Cancelar** vuelve sin guardar. El censo permite calcular el total
automáticamente o mantener un total manual. Las encuestas recalculan su total
y promedio. Fechas, horas, cantidades y tiempos `mm:ss` se validan antes de
guardar; los errores conservan lo escrito para corregirlo. Si otra persona
modificó el registro mientras estaba abierto, se solicita cargar su versión
actual para evitar sobrescribirla. La edición no requiere migrar tablas.

Después de eliminar un registro, la aplicación vuelve al listado del mismo
módulo con sus filtros y muestra la confirmación. Se conserva el prefijo
`/gestion-5s` en la dirección de regreso, incluso si la eliminación se envía
desde una página abierta antes de esta corrección.

### Eliminar varios registros o todo el listado

Los 15 tipos de registros de hotelería incluyen una casilla en cada fila y
la opción **Seleccionar todos los registros del listado**. El contador muestra
cuántos están marcados; **Eliminar seleccionados** actúa sobre esas filas.

**Eliminar todos** abarca los registros del módulo actual que cumplen los
filtros de fechas o semana. Para eliminar todo el historial de un módulo,
abre ese módulo sin filtros. Ambas opciones muestran una pantalla para
confirmar la cantidad y el rango de fechas antes de ejecutar el borrado.
**Cancelar** regresa sin eliminar registros.

La confirmación guarda la selección exacta durante 15 minutos. Si alguno de
esos registros cambia, debe revisarse otra vez; los registros nuevos que se
creen después de confirmar la selección no se incluyen en el borrado.
La operación se guarda en una sola transacción y vuelve al mismo listado.

### Acciones en desviaciones

El formulario de **Desviaciones** incluye el campo opcional **Acciones**.
También se puede editar desde el lápiz y se muestra completo en la columna
**Acciones** del listado, con sus saltos de línea. Los iconos de editar y
eliminar aparecen en la columna **Opciones**.

La plantilla Excel agrega **ACCIONES** al final, después de **CORREO_DESTINO**,
y la descarga CSV incluye el mismo dato. Las plantillas anteriores de
desviaciones se siguen aceptando y dejan Acciones vacío.

Al iniciar la aplicación se añade automáticamente la columna de texto
`desviaciones.acciones` si falta, tanto en PostgreSQL como en SQLite. Los
registros existentes se conservan; el nuevo campo queda vacío hasta completarlo.

Ambos módulos utilizan la misma variable `DATABASE_URL`. Las tablas existentes de
los dos proyectos conservan sus nombres, por lo que el despliegue no elimina ni
sobrescribe información previa.

### Migrar los registros del antiguo despliegue 5S

Si los dos sistemas utilizaban bases PostgreSQL separadas, ejecuta una sola vez:

```bash
SOURCE_5S_DATABASE_URL="<base-antigua-5S>" \
DATABASE_URL="<base-reportabilidad>" \
python scripts/migrate_5s_database.py
```

La utilidad no borra información y cancela la operación si las tablas 5S del
destino ya contienen registros, para evitar duplicados.

## Pruebas

```bash
python -m unittest discover -v
```

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
