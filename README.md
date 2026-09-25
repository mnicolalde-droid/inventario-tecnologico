# Generador de actas de activos tecnológicos

Automatización en **Google Apps Script** vinculada a una hoja de cálculo de Google Sheets. El sistema permite al área responsable de TI seleccionar colaboradores y activos, emitir actas en PDF desde plantillas de Google Docs y reflejar la operación en el inventario de equipos y accesorios.

Su objetivo es estandarizar el registro de entregas, préstamos, recepciones, devoluciones, periféricos y cambios de equipos tecnológicos, evitando el diligenciamiento manual de documentos y manteniendo sincronizado el inventario.

> Los identificadores de plantillas y carpetas de Drive están definidos en el código. No deben publicarse ni reemplazarse con recursos que no tengan los permisos adecuados.

## Alcance funcional

El menú **📄 Actas** se incorpora automáticamente a la hoja de cálculo al abrirla. Desde la opción **Generar Acta** se abre un formulario modal de 700 × 800 px donde el usuario puede:

- Elegir uno de los seis tipos de acta disponibles.
- Buscar y seleccionar colaboradores activos por nombre, excepto en las devoluciones, donde el usuario se deriva del activo seleccionado.
- Consultar y seleccionar equipos y accesorios de acuerdo con su disponibilidad, estado y asignación.
- Registrar información adicional cuando aplica: proceso del préstamo, observaciones de recepción, estado de un equipo o motivo de cambio.
- Generar un PDF desde una plantilla de Google Docs, almacenarlo en la carpeta configurada y abrir su vista previa en una nueva pestaña.
- Actualizar las filas correspondientes de las hojas de inventario una vez creado el documento.

## Tipos de acta y reglas de negocio

| Tipo | Selección y datos requeridos | Efecto en inventario |
| --- | --- | --- |
| Entrega | Colaborador activo y al menos un equipo o accesorio disponible. | Asigna el activo al colaborador y lo marca como `Ocupado`. |
| Préstamo | Igual que entrega, más un proceso o motivo opcional. | Asigna el activo, lo marca como `Prestado` y guarda el proceso cuando se diligencia. |
| Recepción | Colaborador activo; muestra los activos `Ocupado` que tiene asignados. Permite indicar estado y observaciones. | Envía el activo a bodega (ID administrativo), lo marca `Disponible` y guarda observaciones; en equipos también actualiza el estado. |
| Entrega de periféricos | Colaborador activo y accesorios disponibles en buen estado o sin estado registrado. No muestra equipos. | Asigna el accesorio y lo marca `Ocupado`. |
| Devolución | No solicita selección previa de colaborador; presenta equipos y accesorios `Prestado` e identifica a su titular. | Envía el activo a bodega, lo marca `Disponible` y elimina el proceso asociado. |
| Cambio | Colaborador activo, equipo actual, nuevo equipo disponible, motivo y estado del equipo anterior. Se pueden devolver, conservar o asignar accesorios. | Libera el equipo anterior, asigna el nuevo y actualiza cada accesorio según la decisión tomada. |

Para operaciones distintas de recepción, devolución y cambio, los equipos elegibles son los que están sin colaborador o asignados al identificador administrativo de Sistemas, y que cumplen las condiciones de disponibilidad configuradas. El identificador de bodega usado por el proyecto es `1111111111` (`SISTEMAS_ADMIN_ID`).

## Arquitectura y archivos

| Archivo | Responsabilidad |
| --- | --- |
| [`code.js`](code.js) | Punto de entrada de Apps Script. Crea el menú de Sheets y abre el diálogo HTML. |
| [`index.html`](index.html) | Interfaz del formulario: estilos, selección de acta, búsqueda, validaciones del navegador y llamadas asíncronas a Apps Script mediante `google.script.run`. |
| [`data.js`](data.js) | Lógica de negocio: lectura y mapeo de hojas, filtros de activos, sustitución de marcadores de plantilla, creación de PDF y actualización del inventario. |
| [`appsscript.json`](appsscript.json) | Manifiesto del proyecto: zona horaria `America/Guayaquil`, runtime V8 y registro de excepciones en Stackdriver. |
| [`.clasp.json`](.clasp.json) | Vinculación local con el proyecto remoto de Apps Script para uso con `clasp`. |

El flujo principal es el siguiente:

```text
Google Sheets abre el archivo
        ↓
onOpen() agrega el menú “📄 Actas”
        ↓
abrirFormulario() carga index.html
        ↓
Formulario consulta colaboradores / equipos / accesorios
        ↓
generarActa(payload)
        ├─ copia la plantilla de Google Docs
        ├─ reemplaza marcadores con los datos seleccionados
        ├─ convierte la copia a PDF y la guarda en Drive
        └─ actualiza las filas de inventario involucradas
```

## Estructura requerida de la hoja de cálculo

El script espera las hojas `Colaboradores`, `Equipos` y `Accesorios`, con una fila inicial de encabezados. La posición de las columnas es relevante: el código trabaja por índice, no por nombre de encabezado.

### Hoja `Colaboradores`

| Columna | Campo esperado | Uso |
| ---: | --- | --- |
| A | Nombre | Lista y datos del acta. |
| B | Identificación | Clave de asignación del activo. |
| C | Cargo | Dato para el documento. |
| D | Estado | Solo se muestran valores `activo` (sin importar mayúsculas/minúsculas). |
| E | Ubicación | Ciudad/ubicación que se asigna a los activos y se imprime en el acta. |

### Hoja `Equipos`

| Columna | Campo esperado | Uso principal |
| ---: | --- | --- |
| A | Identificación del colaborador | Asignación; el ID de Sistemas representa bodega. |
| B | Ubicación del equipo | Ubicación del titular. |
| C | Tipo de equipo | Plantilla y lista del formulario. |
| D | Marca | Plantilla y lista del formulario. |
| E | Modelo | Plantilla y lista del formulario. |
| F | Número de serie | Identificación visible en el formulario. |
| H | Nombre del equipo | Mapeado para uso del sistema. |
| I | Memoria RAM | Plantilla y lista del formulario. |
| J | Almacenamiento | Plantilla. |
| K | Sistema operativo | Plantilla. |
| L | Estado del equipo | Filtros y recepción/cambio (`Bueno`, `Regular`, `Malo`). |
| M | Disponibilidad | `Disponible`, `Ocupado` o `Prestado`. |
| O | Observaciones | Se actualiza en recepción. |
| Q | Proceso | Se registra en préstamo y se limpia en devolución/cambio. |

### Hoja `Accesorios`

| Columna | Campo esperado | Uso principal |
| ---: | --- | --- |
| A | Identificador | Código único mostrado en el formulario. |
| B | Nombre del accesorio | Plantilla y lista. |
| C | Marca | Plantilla y lista. |
| D | Estado | Filtros, especialmente periféricos y cambio. |
| E | Disponibilidad | `Disponible`, `Ocupado` o `Prestado`. |
| F | Observaciones | Se actualiza en recepción. |
| H | Colaborador asignado | Identificación del titular o ID de Sistemas. |
| J | Proceso | Se registra en préstamo y se limpia en devolución/cambio. |

Las columnas no usadas por el script pueden conservarse para la administración interna del inventario. Si se mueve una columna documentada, deben actualizarse los índices de `mapearEquipo`, `mapearAccesorio` y las escrituras en `actualizarInventario`.

## Configuración de Google Drive y plantillas

La constante `ACTAS_CONFIG` en [`data.js`](data.js) centraliza, por tipo de acta:

- Nombre mostrado en el formulario.
- ID de la plantilla de Google Docs.
- ID de la carpeta de destino en Google Drive.
- Prefijo del archivo PDF.

El archivo generado usa el patrón:

```text
<prefijo>_<nombre del colaborador>_<AAAAMMDD>.pdf
```

El proceso crea una copia temporal de la plantilla, la convierte a PDF, guarda el PDF en la carpeta de destino y mueve la copia temporal a la papelera. La cuenta que ejecuta el script necesita permisos de lectura sobre las plantillas y de creación de archivos en las carpetas de destino.

### Marcadores admitidos en las plantillas

Las plantillas de Google Docs deben contener los marcadores literalmente, entre dobles llaves. El sistema reemplaza los siguientes valores:

```text
{{nombre}}                    {{identificacion}}             {{cargo}}
{{ciudad}}                    {{dia}}                        {{mes}}
{{anio}}                      {{equipos}}                    {{accesorios}}
{{tipo_equipo}}               {{marca_equipo}}               {{modelo}}
{{modelo_equipo}}             {{memoria_ram}}                {{almacenamiento}}
{{sistema_operativo}}         {{estado_equipo}}              {{proceso}}
{{observaciones_equipo}}      {{nombre_accesorio}}            {{marca_accesorio}}
{{observaciones_accesorio}}   {{motivo}}                     {{estado}}
```

En el acta de cambio, algunos marcadores de especificaciones pueden repetirse para representar el equipo anterior y el nuevo; el reemplazo conserva, cuando es posible, el estilo de texto contiguo. Para accesorios de cambio se admite el bloque combinado `{{accesorios}} - {{estado}}`.

## Instalación y despliegue

### Requisitos

- Una cuenta de Google con acceso al archivo de Sheets, las plantillas de Docs y las carpetas de Drive configuradas.
- Node.js y la CLI de Apps Script (`@google/clasp`) si se trabajará desde el repositorio local.
- Una hoja de cálculo con las tres hojas y estructura descritas anteriormente.

### Publicación desde el repositorio

1. Instale la CLI, si aún no está disponible: `npm install -g @google/clasp`.
2. Autentíquese con `clasp login`.
3. Revise que el `scriptId` de [`.clasp.json`](.clasp.json) corresponda al proyecto de Apps Script objetivo.
4. Sincronice los archivos con `clasp push`.
5. Abra o recargue la hoja de cálculo vinculada y autorice los permisos solicitados por Apps Script.
6. Verifique que aparezca el menú **📄 Actas → Generar Acta**.

El proyecto es un script vinculado a una hoja de cálculo; por ello se ejecuta desde el menú del spreadsheet y no requiere publicar una aplicación web.

## Uso operativo

1. Abra la hoja de cálculo y seleccione **📄 Actas → Generar Acta**.
2. Elija el tipo de acta.
3. Seleccione el colaborador cuando el flujo lo requiera; use el buscador para filtrar por nombre.
4. Seleccione uno o más activos. En **Cambio**, seleccione exactamente un equipo actual y un equipo nuevo, registre motivo y estado del equipo anterior, y clasifique los accesorios.
5. Complete los campos adicionales que aparezcan para el tipo de operación.
6. Pulse **Generar Acta**. El sistema valida los datos mínimos, abre una pestaña de vista previa y luego presenta la confirmación para generar otra acta o cerrar el formulario.
7. Confirme que el PDF esté en la carpeta configurada y que los estados del inventario hayan quedado actualizados.

## Validaciones y comportamiento relevante

- No se permite generar un acta sin tipo ni sin algún equipo o accesorio seleccionado.
- Los flujos que requieren colaborador no continúan hasta que se seleccione uno.
- El cambio exige equipo anterior, equipo nuevo, motivo y estado del equipo anterior.
- El tipo `cambio` permite un solo equipo actual y un solo nuevo; los accesorios pueden ser múltiples.
- La devolución toma los datos del colaborador del primer activo seleccionado. Por control operativo, conviene seleccionar en una misma acta únicamente activos del mismo titular.
- La disponibilidad y el estado se normalizan a minúsculas al leerlos; las escrituras usan valores con inicial en mayúscula (`Disponible`, `Ocupado`, `Prestado`).

## Pruebas recomendadas

Antes de usar el sistema en producción, pruebe cada flujo con registros de prueba y compruebe tanto el PDF como las columnas modificadas:

- Entrega de equipo y accesorio disponibles.
- Préstamo con proceso registrado y posterior devolución.
- Recepción con estado `Regular` y observaciones.
- Entrega exclusiva de periféricos.
- Cambio con accesorio conservado, devuelto y uno nuevo asignado.
- Validaciones: sin colaborador, sin activos, y cambio incompleto.
- Plantillas con uno y con varios activos para verificar los marcadores y el formato final.

## Evidencias visuales por adjuntar

Guarde las imágenes en `docs/capturas/` (crear la carpeta al adjuntarlas) y sustituya los enlaces siguientes por las capturas reales. Se recomienda ocultar identificaciones personales, números de serie y enlaces de Drive antes de publicar el repositorio.

| Evidencia requerida | Archivo sugerido | Inserción en este README |
| --- | --- | --- |
| Menú de Sheets con la opción del sistema | `docs/capturas/01-menu-actas.png` | `![Menú de Actas](docs/capturas/01-menu-actas.png)` |
| Formulario inicial y selector de tipo | `docs/capturas/02-formulario-inicial.png` | `![Formulario inicial](docs/capturas/02-formulario-inicial.png)` |
| Selección de activos para préstamo o entrega | `docs/capturas/03-seleccion-activos.png` | `![Selección de activos](docs/capturas/03-seleccion-activos.png)` |
| Flujo de cambio con asignaciones actuales | `docs/capturas/04-cambio-equipo.png` | `![Cambio de equipo](docs/capturas/04-cambio-equipo.png)` |
| PDF generado desde una plantilla | `docs/capturas/05-pdf-generado.png` | `![PDF generado](docs/capturas/05-pdf-generado.png)` |
| Inventario antes y después de una operación | `docs/capturas/06-inventario-actualizado.png` | `![Inventario actualizado](docs/capturas/06-inventario-actualizado.png)` |

También es aconsejable adjuntar enlaces internos o controlados a: las seis plantillas, las carpetas de destino, el spreadsheet de inventario, una matriz de pruebas y el procedimiento institucional de custodia de activos.

## Mantenimiento y consideraciones

- Mantenga los nombres de hojas, los valores de disponibilidad y el ID administrativo alineados con el código.
- Al crear un nuevo tipo de acta, agregue su configuración en `ACTAS_CONFIG`, su plantilla/carpeta, la interfaz necesaria y sus reglas de filtrado y actualización.
- Compruebe los permisos de Drive después de cambiar de cuenta propietaria, copiar el spreadsheet o cambiar una plantilla.
- `appsscript.json` utiliza la zona horaria `America/Guayaquil`; la fecha incluida en los PDFs y sus nombres se calcula con esa configuración.
- Los IDs de Google Drive y el `scriptId` son configuraciones sensibles del entorno. Para distribuir el proyecto a otro entorno, reemplace esos valores por recursos propios y revise los permisos antes del despliegue.
- El proyecto no incluye pruebas automatizadas ni un mecanismo de auditoría histórico independiente; la trazabilidad actual depende del PDF emitido y de las actualizaciones realizadas en las hojas. Si se requiere auditoría formal, se recomienda registrar cada operación en una hoja de historial con fecha, usuario ejecutor, tipo de acta, activos y enlace al PDF.

## Estado del proyecto

El repositorio contiene la implementación funcional del generador y su configuración de Apps Script. La última revisión registrada incorporó validaciones de actas y la visualización de códigos únicos de activos en el formulario. Las capturas y la evidencia de pruebas quedan pendientes de adjuntar en los espacios señalados arriba.
