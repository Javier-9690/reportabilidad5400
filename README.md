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
- Registrar, importar, consultar, editar, eliminar y exportar los 20 tipos de registros operacionales de hotelería.
- Consultar el dashboard KPI consolidado de hotelería por fechas o semanas.

## Menú principal

Importar · Censos · Sin Match · Curva · Nuevo ID · Reportes · Registros hotelería.

- **Reportes**: Dotación, Ocupabilidad, EGP, F&A, Gestión de usuarios y Reporte de desviaciones, solicitudes y reclamos.
- **Registros hotelería**: Ingresar registros, Consultar registros y Dashboard KPI.

El Dashboard general no figura en el menú; sigue disponible desde el encabezado.

## Reporte de desviaciones, solicitudes y reclamos

Disponible en **Reportes > Reporte de desviaciones, solicitudes y reclamos**
(`/reports/desviaciones-solicitudes-reclamos`). Añade una vista basada en las
siete pestañas de registros del archivo de referencia `reporte.xlsx`:

| Pestaña | Base consultada | Fecha utilizada |
| --- | --- | --- |
| Misceláneos | Misceláneos | Fecha creación; Fecha inicio si falta |
| Solicitudes y OT de usuario | Solicitudes OT, carga general Samtech | Fecha creación; Fecha inicio si falta |
| Reclamos de usuarios | Reclamos de usuarios | Fecha |
| Robos y hurtos | Robos y hurtos | Fecha |
| Doble asignación | Duplicidades | Fecha |
| Habitaciones bloqueadas | Habitaciones bloqueadas | Fecha de bloqueo |
| Desviaciones | Desviaciones | Fecha |

Selecciona **Desde** y **Hasta**, escribe opcionalmente una palabra o número y
pulsa **Aplicar filtros**. Inicialmente muestra el mes actual hasta hoy. La
búsqueda usa todos los campos y admite tildes y mayúsculas.
**Incluir también registros sin fecha** permite consultar y descargar esas filas
sin asignarles una fecha ficticia; se separan de la tendencia diaria.

- Resumen con cantidad de registros, abiertos, cerrados, porcentaje de cierre,
  estados por revisar, evolución, estados por categoría y conclusiones calculadas.
- Una pestaña por categoría, indicadores, principales empresas, distribución por
  módulo o habitación y todos los campos disponibles. La tabla tiene desplazamiento
  horizontal y vertical, y páginas de 50 registros.
- Seguimiento de acciones en Desviaciones, receptor de denuncia en Robos y hurtos
  y fechas de liberación por área en Habitaciones bloqueadas.
- **Descargar esta pestaña**: Excel con resumen, evolución diaria y el detalle
  completo de la categoría. También se descarga desde el resumen general.
- **Descargar reporte general**: Excel con resumen, gráficos, evolución diaria y
  siete hojas de detalle. Ambas descargas incluyen todos los resultados de los
  filtros, independientemente de la página consultada.

Los identificadores se conservan como texto y las fechas/horas como valores de
Excel. Las hojas tienen encabezados, filtros y paneles inmovilizados. El resumen
y la evolución usan fórmulas sobre los detalles exportados.

El cierre se calcula como cerrados / (abiertos + cerrados), usando el estado
actual y la fecha de origen. Las categorías sin campo de estado (Desviaciones,
Robos y hurtos y Habitaciones bloqueadas) no participan en esa tasa. Un comentario
o una fecha de liberación de un área no implica cierre ni liberación final de la
habitación. Los estados desconocidos permanecen separados. El total suma filas
de las siete bases, no personas ni tickets únicos.

El reporte usa las tablas vigentes, sin cargar automáticamente el Excel de
referencia ni modificar registros. Muestra todos los campos que almacena el
sistema, incluidas las columnas históricas de Solicitudes OT. La hoja auxiliar
oculta `Desplegables` contiene listas de validación, no una base de registros.
No requiere migraciones, nuevas dependencias ni cambios de configuración.
La consulta usa una única transacción para que resúmenes y detalles no mezclen
dos cargas si se confirma una importación simultánea.

**Gestión de usuarios** conserva sus fuentes y el seguimiento de Solicitudes de
usuarios Samtech QR, separado de la carga general.

Pruebas del nuevo reporte:

```bash
python -m unittest tests.test_operations_reports -q
node tests/test_operations_report_interface.cjs
```

## Reporte Gestión de usuarios

Abre **Reportes > Gestión de usuarios**, completa **Desde** y **Hasta** y pulsa
**Generar reporte**. La ruta es `/reports/usuarios`. El período incluye ambos
extremos y todos los días intermedios, aunque no tengan registros.

| Fila del reporte | Registro de origen | Fecha utilizada | Campo de estado |
| --- | --- | --- | --- |
| Doble asignación | Duplicidades | Fecha | Estatus |
| Reclamos usuarios | Reclamos de usuarios | Fecha | Estatus |
| Solicitudes de usuarios | Solicitudes de usuarios Samtech QR | Fecha creación; Fecha inicio si falta | Estado |
| Solicitudes totales Samtech | Misceláneos + Solicitudes OT (carga general completa) | Fecha creación; Fecha inicio si falta | Estado |
| Desviaciones clientes | Desviaciones | Fecha | No tiene campo de estado |

El reporte diario reproduce las secciones de la referencia: registros, casos
abiertos, casos cerrados y comparación con porcentaje de cierre. Tiene encabezados
rojos, totales amarillos, una columna **Total período**, etiquetas fijas y barras
horizontales sincronizadas. Los casos abiertos se desglosan por categoría y estado;
se añaden filas cuando aparecen otros estados abiertos reconocidos. Cada fila
guardada cuenta una vez; los números de ticket repetidos cuentan por separado.

El **Dashboard de conclusiones** incluye conteos, porcentaje de cierre, resumen por
categoría y dos gráficos: registros diarios y estado actual por categoría. Sus
conclusiones se calculan con los datos del rango: volumen, categoría con mayor
carga, abiertos para seguimiento, días de mayor ingreso y datos por completar.
Los enlaces **Ver registros** abren cada listado con el mismo período.

### Criterios de estado y fechas

Los conteos muestran el **estado actual de los registros fechados en el período**.
Un caso que después se cierra permanece en su fecha de origen. El informe no
reconstruye estados históricos ni cuenta los cierres por el día en que ocurrieron,
porque estos registros no comparten un historial de cambios de estado.
La Fecha cierre de Duplicidades no reemplaza su Fecha del registro.
Las fechas de término y aprobación de Samtech no reemplazan su Fecha creación.

Se ignoran mayúsculas, tildes, espacios repetidos, guiones y guiones bajos en los
estados. Se aplican estas equivalencias:

| Grupo | Estados reconocidos |
| --- | --- |
| Abierto | Abierto, Abierta |
| No iniciada | No iniciada, No iniciado |
| En progreso | En progreso, En proceso, En curso, En ejecución, Iniciado, Iniciada |
| Pendiente | Pendiente, En espera |
| Cerrado | Cerrado, Cerrada, Resuelto, Resuelta, Finalizado, Finalizada, Completado, Completada, Terminado, Terminada |

Los primeros cuatro grupos cuentan como abiertos. Los estados vacíos u otros
valores, incluidos Cancelado y Rechazado, quedan **Sin clasificar**. Cuando existen,
se muestran en una sección adicional y en las conclusiones. Así, el total de
registros siempre coincide con abiertos + cerrados + sin clasificar.

**Desviaciones clientes** participa en las cinco filas de volumen, sus totales,
conclusiones y ambos gráficos. Como su formulario no tiene campo de estado, se
incluye en **Sin clasificar**, sin atribuirle un estado abierto o cerrado a partir
de Acciones. Su porcentaje de cierre por categoría muestra **No aplica**. Las
conclusiones distinguen estos casos de los estados vacíos o no reconocidos de
los otros registros; no solicitan completar un campo que no existe.

**% cerrado = cerrados / (abiertos + cerrados)**. El total del período se calcula
con los conteos acumulados, sin promediar los porcentajes diarios. Sin casos
clasificables se muestra 0% y se explica esa ausencia. Los casos sin clasificar
no se incluyen en el denominador.

Las solicitudes QR se cargan en **Solicitudes de usuarios Samtech QR** y alimentan
exclusivamente **Solicitudes de usuarios**. La importación conjunta de órdenes
alimenta **Solicitudes totales Samtech** sumando Misceláneos y Solicitudes OT;
CARPINTERIA MENOR sigue guardándose en Misceláneos. La tabla anterior Samtech
usuarios conserva sus datos y funciones, pero ya no alimenta este reporte.

Ambos indicadores pueden compartir tickets. **Total registros (suma de categorías)**
suma las filas del reporte y no representa tickets únicos. La advertencia figura
en pantalla y en Excel. Los estados de cada indicador provienen de su propia
carga, por lo que conviene actualizar ambos archivos al comparar el período.

En QR, Misceláneos y Solicitudes OT, **Aprobada/Aprobado** cuenta como Cerrado;
Completada también equivale a Cerrado. Eliminado y Felicitaciones quedan Sin
clasificar. Las cuatro fechas originales se conservan. Se usa Fecha creación
y, si falta, Fecha inicio; término, aprobación y fecha de importación no reemplazan
esta referencia. El Excel añade **Fecha para reporte** calculada con la misma regla.
El detalle de Solicitudes totales Samtech incluye **Registro de origen** para
distinguir Misceláneos de Solicitudes OT, con todas sus columnas vigentes.

El reporte lee una versión coherente de todas las categorías aunque se confirme
una importación en otra conexión mientras se genera la pantalla o el Excel.

Los registros sin la fecha de referencia quedan fuera del período. Se informa
cuántos existen en el histórico de las cinco categorías, para completar sus fechas;
ese conteo no se atribuye al rango solicitado. Los criterios y estados pendientes
pueden consultarse al pie del reporte.

### Exportación para envío

**Exportar Excel** descarga `/reports/usuarios.xlsx` con el rango aplicado.
Si cambias las fechas del formulario, debes generar el reporte antes de exportar.
El archivo se llama `reporte_usuarios_AAAA-MM-DD_AAAA-MM-DD.xlsx` e incluye:

- **Reporte diario**: estructura de la referencia, fechas reales, conteos, totales
  y porcentajes mediante fórmulas. Las fechas y la columna de ítems quedan fijas.
- **Conclusiones**: indicadores, resumen por categoría, conclusiones, dos gráficos
  de Excel y criterios del informe.
- **Doble asignación**, **Reclamos usuarios**, **Solicitudes usuarios QR**, **Solicitudes totales Samtech**
  y **Desviaciones clientes**: todos los campos
  vigentes de cada registro del rango, encabezados fijos, autofiltro y Estado
  agrupado. Los textos e identificadores conservan su valor, incluidos ceros iniciales.
  Desviaciones conserva el campo Acciones y muestra **Sin campo de estado** en
  Estado agrupado. En total, el archivo incluye siete hojas.

Los conteos diarios usan `COUNTIFS` sobre las hojas de detalle. Los totales y
porcentajes usan fórmulas con resultados guardados para que puedan verse también
en visores. Los gráficos se enlazan a esos datos. El Excel consulta nuevamente los
registros al exportar e identifica ese momento en UTC; las conclusiones corresponden
a esa exportación. Tras modificar registros en el programa, genera un archivo nuevo.

Las fechas inválidas, incompletas o invertidas muestran un error sin ampliar el
período. Se respetan los límites propios de Excel: 16.382 días más las columnas
Ítem y Total, 1.048.576 filas por hoja y 32.767 caracteres por celda. Un texto que
excede ese límite se informa sin truncarlo silenciosamente. La exportación usa
las dependencias existentes del programa y no requiere migraciones ni nuevas tablas.

## Registros hotelería

### Descarga CSV en columnas para Excel

Todos los registros, incluido el histórico, descargan CSV con punto y coma (`;`) como separador,
la indicación inicial `sep=;` para Excel y codificación UTF-8 con BOM para
conservar tildes y eñes. Los campos quedan separados al abrir la descarga en
Excel, incluso si la configuración regional del equipo utiliza otro separador.
Las comas, puntos y coma, comillas y saltos de línea dentro de un campo se
escapan como CSV y conservan su contenido en una sola celda. Los campos vacíos
mantienen su posición. La búsqueda y los filtros de fecha siguen aplicándose
a toda la descarga, sin limitarla a la página visible.

Para procesar estas descargas con otra herramienta, omite la primera línea
`sep=;`, usa `;` como delimitador y decodifica el archivo como `utf-8-sig`.
Este ajuste se aplica a los archivos descargados después de actualizar el
programa; los CSV descargados anteriormente deben generarse de nuevo.

### Activación de alarmas: horas, minutos y segundos

Los siguientes cinco campos permiten ingresar y consultar **HH:MM:SS**, por
ejemplo `08:15:30`, tanto en el formulario de ingreso como en la edición:

- Aviso a mantención.
- Llegada de mantención.
- Aviso al líder de emergencias.
- Llegada del líder.
- Hora de reporte a SALFA.

El listado, buscador y exportación CSV conservan los segundos. Los cinco campos
pueden quedar vacíos: un vacío sigue siendo un vacío y `00:00:00` se conserva
como un valor distinto. También se admite `HH:MM` y se completa con segundos `00`.
Minutos y segundos deben estar entre 00 y 59; SALFA es una hora del día de 00 a 23.

La plantilla Excel aplica formato de hora a esas cinco columnas y conserva los
encabezados anteriores. Incluye ayuda en los encabezados y queda vacía, para no
importar accidentalmente una fila de ejemplo. Acepta texto `HH:MM:SS` y horas
reales de Excel. Una hora inválida muestra la fila y cancela toda la importación;
no se convierte silenciosamente en vacío ni en medianoche.

Las cuatro columnas históricas de aviso/llegada siguen almacenadas en horas
decimales, por lo que no se necesita migrar la base. Por ejemplo, `0.5` horas se
muestra como `00:30:00` y `1.25` como `01:15:00`; se conservan también valores
acumulados mayores de 24 horas. Editar otro campo no redondea el valor guardado.
Los números de plantillas antiguas sin formato de hora mantienen esa unidad;
las celdas con formato de hora Excel se interpretan como horas de Excel.

### Importación conjunta de órdenes: Misceláneos y Solicitudes OT

Desde cualquiera de las dos categorías, selecciona el Excel de órdenes y pulsa
**Importar órdenes · Revisar**. Ambas plantillas tienen las mismas 15 columnas:
Ticket, División, Área, Lugar, Ubicación, Disciplina, Especialidad, Falla, Empresa,
Fecha creación, Fecha inicio, Fecha término, Fecha aprobación, Estado y Comentario.
Se acepta el formato del archivo `ordenes.xlsx`, con fechas `dd/mm/aaaa` o fechas
reales de Excel, y también `aaaa-mm-dd`. Ticket se guarda en el campo OT existente.

- **CARPINTERIA MENOR** en Especialidad → Misceláneos.
- Todas las demás especialidades, incluidas las vacías → Solicitudes OT.
- La comparación ignora tildes, mayúsculas y espacios adicionales. **CARPINTERIA**
  y **CARPINTERIA MENOR EXTRA** no son CARPINTERIA MENOR y van a Solicitudes OT.

La vista previa muestra los registros actuales que se eliminarán, las órdenes
nuevas por destino, los estados, las fechas y hasta cinco ejemplos por categoría.
Para ejecutar hay que marcar la confirmación y pulsar **Confirmar y reemplazar
ambas bases**. **Cancelar** conserva los datos. No se acumula sobre las filas
existentes: al confirmar, ambas categorías se reemplazan en una sola transacción.
Si falla una parte, se revierten todos los cambios. Si alguna categoría recibe
cero filas, la vista previa avisa que quedará vacía. Una plantilla totalmente
vacía nunca permite borrar las bases.

Los campos y fechas vacíos se admiten. Las filas completamente vacías se omiten.
Un encabezado, fecha o valor inválido rechaza todo el archivo antes de reemplazar
datos. Las fórmulas de origen deben convertirse a valores. Los tickets repetidos
se señalan y cada fila cuenta una vez, sin deduplicar datos silenciosamente.
Los campos largos, como Falla y Comentario, se conservan completos.

Se valida que las bases no hayan cambiado desde la vista previa: una alta, edición
o eliminación posterior obliga a revisarla nuevamente. La confirmación pertenece
a la sesión que cargó el archivo, vence a los 30 minutos y se utiliza una sola vez.
La última vista previa sustituye a las anteriores de esa sesión para la misma
carga. La carga general y la QR mantienen confirmaciones independientes. El archivo
temporal se elimina al confirmar, cancelar o limpiar vistas expiradas en una carga
posterior. Los límites son 25 MB por archivo y 150.000 órdenes.

Los 15 campos importados aparecen en formularios, edición, listados, búsqueda,
CSV y detalle Excel. Solicitudes OT conserva también sus campos históricos.
Sus listados y Misceláneos muestran 100 filas por página con todas las columnas.
La búsqueda, CSV y **Eliminar todos** abarcan todo el resultado filtrado; las
casillas seleccionan la página visible. El reporte de Gestión de usuarios consulta
inmediatamente la carga general completa como Solicitudes totales Samtech.
La importación general conserva las solicitudes QR y los demás registros. Los registros sin ambas fechas permanecen guardados y quedan fuera
de los reportes por fecha; esta cantidad se informa.

La actualización añade columnas opcionales a Solicitudes OT y una tabla temporal
para las vistas previas, sin borrar los registros existentes. El reemplazo de datos
ocurre únicamente al confirmar dentro del programa. No requiere nuevas dependencias.

El módulo antes llamado Gestión 5S se encuentra en estas rutas dentro de la aplicación principal:

- `/gestion-5s/panel`: ingreso manual e importación Excel.
- `/gestion-5s/registros`: consulta con buscador por categoría, eliminación y descarga CSV.
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

### Formularios y listados completos

Los 20 formularios incluyen **Ver registros** junto a **Descargar Plantilla**
e **Importar**. El enlace abre directamente el listado del módulo activo;
la encuesta también funciona al acceder con `tab=encuesta` o `tab=encuestas`.

El ingreso y la edición comparten los mismos campos, tipos, límites y reglas
de validación. Se comprueban los datos obligatorios, las fechas y horas, los
enteros, las duraciones `mm:ss`, los puntajes y los límites de texto antes de
guardar. Las duraciones admiten más de 99 minutos y segundos de 00 a 59.
Los errores aparecen junto al campo y conservan lo escrito para corregirlo.
Un fallo de base de datos revierte el intento de guardado y conserva el formulario.
Se mantienen los campos opcionales y la compatibilidad con formularios abiertos
antes de esta actualización, incluidos los nombres anteriores de ID y tiempos.

Cada listado muestra todos los campos vigentes del registro, los textos completos
con sus saltos de línea y la fecha de creación. La encuesta incluye también todas
las respuestas, sus puntajes, comentarios, total y promedio. Los valores de cero
se muestran como cero; una duración de cero se muestra como `00:00`.
El campo histórico Proyecto de Extensiones permanece guardado internamente,
como se solicitó al sustituir las columnas de ese módulo.

Cuando las columnas exceden el ancho de la pantalla, la tabla tiene desplazamiento
horizontal. Una segunda barra sincronizada aparece encima para poder desplazarse
sin recorrer todas las filas. La tabla también permite desplazamiento vertical,
mantiene visibles los encabezados y admite navegación por teclado. **Limpiar filtros**
conserva el módulo actual. Se mantienen el lápiz y todas las opciones de eliminación.

En Onboarding, **Archivo PDF (nombre)** permite escribir el nombre o completarlo
seleccionando un PDF. Se registra únicamente el nombre, que puede revisarse o
corregirse antes de guardar; este campo no almacena el contenido del documento.

### Buscar dentro de una categoría

En **Consultar registros / Ver registros**, elige la **Categoría**, escribe una
palabra o número en **Buscar dentro de esta categoría** y pulsa **Buscar** o
Enter. Funciona en los 20 tipos de registros visibles, incluidas las solicitudes QR.
La búsqueda del histórico Samtech usuarios sigue disponible mediante su enlace anterior.
Encuentra coincidencias parciales en cualquiera de los campos vigentes de esa
categoría, incluidos observaciones, acciones, ID, RUT, habitación y ticket.
No distingue mayúsculas ni tildes: `mantencion` encuentra `Mantención`.
Los caracteres `%`, `_` y `/` se buscan literalmente.

La búsqueda puede combinarse con fechas o semana. También admite fechas como
`2026-09-09` o `09/09/2026`, horas y duraciones `mm:ss`. Los campos vacíos no
impiden buscar por los demás datos; los registros sin fecha aparecen cuando no
hay un filtro de fecha o semana. Se muestran la búsqueda activa y la cantidad
de coincidencias. Si no hay coincidencias, el listado queda vacío.

**Descargar CSV**, **Eliminar seleccionados** y **Eliminar todos** respetan la
misma búsqueda. Al editar, cancelar o eliminar, se conserva la categoría y sus
filtros. **Quitar búsqueda** conserva las fechas y semana; **Limpiar filtros**
quita todos los filtros y mantiene la categoría. Dejar el buscador vacío equivale
a no filtrar por texto. La búsqueda admite hasta 200 caracteres y se envía en
el parámetro `q`. No requiere cambios en las tablas ni extensiones de base de datos.

### Eliminar varios registros o todo el listado

Los 20 tipos de registros de hotelería incluyen una casilla en cada fila y
la opción **Seleccionar todos los registros del listado**. El contador muestra
cuántos están marcados; **Eliminar seleccionados** actúa sobre esas filas.

**Eliminar todos** abarca los registros del módulo actual que cumplen los
filtros de fechas, semana y búsqueda. Para eliminar todo el historial de un módulo,
abre ese módulo sin filtros. Ambas opciones muestran una pantalla para
confirmar la cantidad, el rango de fechas y la búsqueda antes de ejecutar el borrado.
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

### Campos de extensiones y excepciones

El formulario de ingreso, la edición, la plantilla Excel, el listado y la
descarga CSV utilizan estos 12 campos, en este orden:

1. FECHA DE SOLICITUD
2. ID
3. EMPRESA
4. CO
5. GERENCIA
6. CENTRO COSTOS
7. CANT. CLIENTES
8. TIPO DE SOLICITUD
9. DESDE
10. HASTA
11. APROBADOR
12. OBSERVACION

**Tipo de solicitud** es texto libre. El listado muestra también el ID
ingresado, la gerencia y la observación completa, y conserva las opciones
de editar, seleccionar y eliminar. Una cantidad de cero se muestra como cero.

Al iniciar, se agregan automáticamente `centro_costos` y `tipo_solicitud` a
`extension_excepcion` si faltan. Los registros anteriores y su dato histórico
`proyecto` se conservan en la base; Proyecto deja de aparecer en los formularios
y reportes de este registro. Los campos nuevos quedan vacíos hasta completarlos
desde el lápiz, sin asignarles datos que correspondían a otro concepto.
Para importar, descarga la nueva plantilla de 12 columnas: la anterior de
11 columnas se rechaza sin guardar filas, para evitar desplazar fechas y datos.

### Registro de entradas y salidas

Disponible en **Registros hotelería > Ingresar registros > Entradas y salidas**
y en el selector de **Consultar registros**. El formulario, la edición, la
plantilla Excel y el listado utilizan estos 16 campos, en este orden:

1. FECHA INGRESO
2. FECHA SALIDA
3. HORA ENTRADA
4. HORA SALIDA
5. EMPRESA
6. ID
7. NOMBRE
8. RUT
9. TURNO
10. PABELLON
11. HABITACIÓN
12. MOTIVO
13. AUTORIZADO
14. PENDULO
15. N° TARJETA
16. DEVOLUCIÓN

Los 16 campos son opcionales: se pueden importar, ingresar y editar registros
con fechas, horas u otros datos vacíos. Basta con que la fila tenga al menos
un dato; las filas totalmente vacías se omiten. Los vacíos se conservan como
tales, sin inventar fechas ni horas. Se validan los datos que sí están
informados y el orden cronológico cuando hay información para compararlo.
Se admiten salidas al día siguiente y horas a medianoche. Los demás campos permiten texto libre;
ID, RUT y N° tarjeta conservan ceros iniciales y guiones cuando se ingresan
como texto. Motivo conserva los saltos de línea.

**Descargar Plantilla** genera el XLSX vacío con encabezados, formatos de
fecha/hora y celdas de texto para los identificadores. **Importar** acepta
las fechas y horas de Excel y texto como `08/09/2026` y `23:45:00`.
Si una fila contiene un dato inválido, se indica su número y no se guarda
ninguna fila de ese archivo. La descarga CSV incluye los 16 campos.

El lápiz, la eliminación individual, la selección múltiple y **Eliminar todos**
funcionan como en los otros registros. El listado, CSV y borrado aplican sus
filtros de fechas o semana a **FECHA INGRESO**.
Los registros sin esa fecha aparecen al abrir el módulo sin filtros y pueden
exportarse, completarse desde el lápiz o eliminarse como los demás.

El dashboard muestra la tarjeta y el gráfico **Entrada a habitaciones**, con
el conteo de entradas por FECHA INGRESO dentro del período seleccionado.
Se retiraron la tarjeta y la serie de salidas. Las fechas de salida tampoco
añaden días al eje del dashboard. Los datos de salida siguen disponibles en
el formulario, el listado, la edición y la exportación del registro.
Las entradas sin fecha no se asignan a otro día. Las cifras representan
registros, no personas únicas.

La tabla `entradas_salidas` se crea automáticamente al iniciar si no existe.
Si se instaló la versión anterior, se retiran las restricciones de obligatoriedad
de `fecha_ingreso` y `hora_entrada`, conservando los registros. PostgreSQL
actualiza las restricciones; SQLite realiza la adaptación en una transacción
que conserva los datos, índices y triggers existentes.

### Habitaciones bloqueadas

Disponible en **Registros hotelería > Ingresar registros > Habitaciones bloqueadas**
y en el selector de **Consultar registros**. Incluye formulario, plantilla Excel,
importación, edición, listado, descarga CSV y eliminación individual o múltiple.
Los 12 campos aparecen en este orden:

1. FECHA DE BLOQUEO
2. HABITACIÓN
3. OT
4. EMPRESA
5. ID
6. MOTIVO
7. COMUNICADO
8. FECHA LIBERADA POR INGECLEAN
9. FECHA LIBERADA POR FACILITY
10. FECHA LIBERADA POR MANTENCION
11. FECHA LIBERADA PROCESO INVESTIGACION
12. OBSERVACIÓN

Todos son opcionales y pueden completarse por separado. Las filas de Excel
completamente vacías se omiten. Habitación, OT e ID se guardan como texto para
conservar ceros iniciales; Motivo, Comunicado y Observación admiten varias líneas.
Las fechas informadas se validan y, cuando se conoce la fecha de bloqueo, ninguna
liberación puede ser anterior a ella. No se exige un orden entre las cuatro
liberaciones. Una fila inválida cancela toda la importación indicando su número.

La tarjeta y el gráfico **Habitaciones bloqueadas** cuentan registros por
**FECHA DE BLOQUEO**, respetando el período o semana seleccionados. Este conteo
representa los bloqueos registrados en el período; no calcula cuántas habitaciones
siguen bloqueadas ni requiere que se completen todas las fechas de liberación.
El listado, la exportación y el borrado usan la misma fecha de bloqueo para filtrar.
Los registros sin fecha se consultan y exportan sin filtros.

La tabla `habitaciones_bloqueadas` se crea automáticamente al iniciar la aplicación.

### Ordenamiento y Habitaciones liberadas

Ambos registros están disponibles en **Registros hotelería > Ingresar registros**
y en el selector de **Consultar registros**. Cada uno incluye formulario,
descarga de plantilla Excel, importación, listado, exportación CSV, edición
con el lápiz y eliminación individual, por selección o de todos los registros.

**Ordenamiento** utiliza estas 10 columnas, en el mismo orden en el formulario,
la edición, la plantilla y el listado:

1. FECHA EJECUCIÓN
2. EMPRESA
3. HABITACIÓN
4. NOMBRE
5. RUT
6. TURNO
7. REASIGNACIÓN
8. MOTIVO CAMBIO
9. PROCESADO
10. PENDIENTE

**Habitaciones liberadas** utiliza estas 7 columnas:

1. Habitacion
2. Empresa
3. Entrega /devolucion
4. Fecha de devolucion
5. Comentario
6. USO A PARTIR DE
7. Observación

Todos los campos son opcionales, incluidas las fechas. Basta con un dato para
guardar o importar una fila; las filas totalmente vacías se omiten. Los campos
vacíos se conservan y pueden completarse después. Habitación, RUT y Reasignación
admiten códigos con ceros iniciales como texto; Motivo cambio, Comentario y
Observación conservan varias líneas. Procesado, Pendiente y Entrega /devolucion
son campos de texto libre.

Las plantillas admiten fechas de Excel y texto `DD/MM/AAAA` o `AAAA-MM-DD`.
Las fechas informadas se validan; si alguna fila contiene un dato inválido,
se indica su número y no se guarda ninguna fila del archivo.

Cada módulo tiene una tarjeta y un gráfico en el dashboard. **Ordenamiento**
cuenta registros por **FECHA EJECUCIÓN**; **Habitaciones liberadas**, por
**Fecha de devolucion**. Los filtros de período o semana del listado, CSV y
borrado utilizan esas mismas fechas. **USO A PARTIR DE** se conserva como
fecha independiente y no sustituye la fecha de devolución. Los registros
sin la fecha principal se consultan y exportan sin filtros, y no se asignan
a un día en el gráfico. Las cifras cuentan registros, no habitaciones únicas.

Las tablas `ordenamiento` y `habitaciones_liberadas` se crean automáticamente
al iniciar la aplicación, conservando las tablas y registros existentes.

### Solicitudes de usuarios Samtech QR

Nuevo registro independiente en **Registros hotelería > Ingresar registros**
(`/gestion-5s/panel?tab=samtech_qr`). Usa las mismas 15 columnas del archivo
Samtech: Ticket, División, Área, Lugar, Ubicación, Disciplina, Especialidad,
Falla, Empresa, Fecha creación, Fecha inicio, Fecha término, Fecha aprobación,
Estado y Comentario.

Incluye formulario, **Ver registros**, plantilla Excel, importación confirmada,
exportación CSV de todo el resultado filtrado, buscador por palabra o número,
edición con lápiz, eliminación individual, múltiple y **Eliminar todos**.
El listado muestra todas las columnas, barras de desplazamiento y 100 filas por
página. El dashboard de registros tiene su tarjeta y gráfico propios.

Para cargar un archivo QR, selecciona el Excel y pulsa **Importar órdenes · Revisar**.
La vista previa muestra cuántos registros QR se reemplazarán y una muestra.
Marca la confirmación y pulsa **Confirmar y reemplazar solicitudes QR**.
Todas las especialidades, incluida CARPINTERIA MENOR, permanecen en el registro QR.
Esta operación reemplaza exclusivamente la tabla QR; no modifica Misceláneos,
Solicitudes OT ni el antiguo Samtech usuarios. La carga general tampoco modifica QR.

Se admiten campos y fechas vacíos. La plantilla vacía no puede eliminar la base;
un dato inválido cancela el archivo completo. Si falla el guardado, se recuperan
los registros anteriores. Una edición posterior a la vista previa obliga a volver
a revisarla. Reimportar el mismo archivo reemplaza los datos y no acumula duplicados.

El listado, CSV, dashboard y reporte usan Fecha creación, o Fecha inicio si falta.
Los registros sin ambas fechas quedan guardados y fuera de los filtros por fecha.
La tabla `samtech_usuarios_qr` se crea al iniciar la aplicación actualizada, sin
mover ni borrar datos anteriores y sin nuevas dependencias.

### Samtech usuarios: histórico conservado

**Solicitudes de usuarios Samtech QR** es la única opción visible para esta carga
de solicitudes. Conserva el formulario, la plantilla, la importación directa
con vista previa y confirmación, la consulta, edición, búsqueda y exportación.
Alimenta **Solicitudes de usuarios** en Gestión de usuarios.

El antiguo **Samtech usuarios** se retiró del menú de ingreso, del selector de
categorías habitual y de las tarjetas y gráficos del dashboard. Sus filas tampoco
crean días vacíos en los gráficos ni se suman a los indicadores de QR.

La tabla `samtech_usuarios` y sus registros se conservan. Por compatibilidad, los
enlaces anteriores permiten acceder al histórico y descargarlo, por ejemplo
`/gestion-5s/registros?vista=samtech_usuarios`. Al abrirlo directamente se
identifica como histórico y se ofrece un enlace al registro QR. No se realiza una
migración automática de esas filas a QR, porque no se puede asumir que todas
provengan de solicitudes de usuarios por QR.

La carga general de **Misceláneos + Solicitudes OT** sigue alimentando
**Solicitudes totales Samtech**, como se describe arriba.

Reportabilidad y Registros hotelería utilizan la misma variable `DATABASE_URL`. Las tablas existentes de
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

Las pruebas de búsqueda verifican las 21 categorías, coincidencias parciales,
tildes, números, campos opcionales, exportación, filtros combinados y alcance del
borrado confirmado. Se ejecutan con SQLite temporal; la consulta de PostgreSQL
también se comprueba mediante compilación SQL.

Los eventos de selección, desplazamiento horizontal y ayuda de formularios
tienen una comprobación adicional que puede ejecutarse con Node.js:

```bash
node tests/test_hotel_interface.cjs
node tests/test_user_report_interface.cjs
```

El reporte de usuarios tiene pruebas para los conteos y estados de sus cinco indicadores,
rangos inclusivos, días vacíos, porcentaje ponderado, fechas faltantes, filtros,
exportación, datos completos, fórmulas, gráficos y navegación. La interfaz comprueba
que cambiar las fechas pida generar el reporte de nuevo antes de descargarlo.

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

### Corrección de `ModuleNotFoundError: No module named 'psycopg'`

SQLAlchemy 2.1 cambió el controlador predeterminado de PostgreSQL a **Psycopg 3**.
Una instalación nueva puede seleccionar esa versión con la regla `SQLAlchemy>=2.0,<3`
del proyecto, y solicitar `psycopg` incluso usando una URL `postgresql://`.
La conexión explícita `postgresql+psycopg://` también requiere ese módulo.
El paquete `psycopg2-binary` instala otro módulo y no cubre esas conexiones. `requirements.txt`
incluye ahora ambos controladores: `psycopg[binary]>=3.2.10,<4` y
`psycopg2-binary>=2.9.11,<3`. Estas versiones mínimas ya admiten Python 3.14,
además del Python 3.12 configurado en `render.yaml`.

Para aplicar la corrección al servicio existente:

1. Subir el `requirements.txt` actualizado a la raíz del repositorio que usa Render.
2. Comprobar que **Build Command** sea `pip install -r requirements.txt`.
3. Ejecutar **Manual Deploy → Clear build cache & deploy** para instalar las dependencias.
4. Revisar que la instalación incluya `psycopg` y `psycopg-binary`, y que el nuevo
   despliegue finalice correctamente.

Se conserva la `DATABASE_URL` existente y el comando de inicio. Esta corrección
no modifica tablas ni registros y no requiere crear otra base de datos.
Si se fija una versión de Python, Render utiliza `PYTHON_VERSION` o un archivo
`.python-version`; `render.yaml` solo configura el servicio si se aplica como Blueprint.

Referencias: [instalación de Psycopg](https://www.psycopg.org/psycopg3/docs/basic/install.html),
[notas de Psycopg 3](https://www.psycopg.org/psycopg3/docs/news.html),
[notas de Psycopg 2](https://www.psycopg.org/docs/news.html) y
[versión de Python en Render](https://render.com/docs/python-version).
El cambio de controlador está documentado en la
[migración a SQLAlchemy 2.1](https://docs.sqlalchemy.org/en/21/changelog/migration_21.html).

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


## Corrección de picos de memoria (30-09-2026)

Las exportaciones Excel de registros y CSV recorren los datos por lotes y usan
archivos temporales; las hojas, fórmulas y gráficos se conservan. Las listas de
registros tienen 100 filas por página y sus descargas incluyen todos los resultados.
Los binarios originales y los Excel terminados se cargan de la base de datos
solo cuando se necesita su contenido, no en consultas de estado o de historial.

Las importaciones de registros usan lectura de Excel por filas y lotes dentro de
una única transacción. La carga confirmada Samtech/QR también prepara las filas
en disco. Los lectores de censos y curvas descartan XML procesado y filas que
solo tienen formato. Los censos insertan lotes sin acumular todos los modelos.

`memory_utils.py` comparte un límite de una operación pesada por proceso entre
las importaciones y exportaciones. Los trabajos en segundo plano usan una cola
con un trabajador; las operaciones HTTP directas reciben 503 con `Retry-After`
si el servicio está ocupado y deben volver a intentarse. No se aceptan cambios de
datos en ese caso. Mantener `--workers 1 --threads 2`: aumentar procesos multiplica
la memoria y cada proceso tendría su propio límite. La cola vive en el proceso;
si hay un reinicio durante una exportación pendiente, debe solicitarse de nuevo.

Prueba local con 30.000 solicitudes en el reporte general: pico RSS del proceso
374,5 MiB antes y 125,4 MiB después (aproximadamente 67% menos). Con 90.000 filas,
148,1 MiB. Se comprobó que las descargas contienen todas las filas. Son mediciones
con datos sintéticos, Python 3.12 y SQLite, no métricas del servicio en producción
ni una garantía de memoria máxima. Los archivos temporales también requieren
espacio de disco; la lectura de XLSX puede conservar su tabla de textos/estilos.

Ver `LEEME_ACTUALIZACION.txt` para aplicar la actualización y comprobar las métricas.
La causa exacta del reinicio requiere los logs privados y su correlación con el
pico de memoria; no puede deducirse solamente del aviso de Render.

Referencias técnicas:
[Excel por filas](https://xlsxwriter.readthedocs.io/working_with_memory.html),
[consultas por lotes](https://docs.sqlalchemy.org/en/20/orm/queryguide/api.html#fetching-large-result-sets-with-yield-per),
[métricas de Render](https://render.com/docs/service-metrics).
