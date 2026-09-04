# Historial de cambios

Todas las versiones desde la 2.1.0. Las causas raíz están anotadas porque varias fueron
difíciles de encontrar y conviene no volver a investigarlas.

## 2.1.10.0 — Aislamiento de dependencias

Corrige que este plugin provocara fallos en **otros** plugins de XrmToolBox.

- **Causa raíz:** el handler `AssemblyResolve` se registraba en el constructor de `Plugin` y
  nunca se quitaba. Ese evento es de todo el proceso, así que nuestros handlers interceptaban
  las resoluciones de assembly de todos los demás plugins cargados, con riesgo de entregarles
  nuestra versión de una librería compartida y romperlos.
- El handler ahora: se registra una sola vez por proceso; ignora peticiones cuyo
  `RequestingAssembly` no sea el propio; solo atiende dependencias declaradas por este
  assembly; busca únicamente en su propia subcarpeta; nunca devuelve una versión anterior a la
  pedida; y no propaga excepciones.
- **Corregido `"$argName.dll"`**: interpolación de PowerShell dentro de un literal de C#. El
  handler buscaba un archivo con ese nombre literal y nunca resolvía nada, lo que obligaba a
  copiar las dependencias a la raíz de `Plugins` como parche.
- **El instalador ya no copia dependencias a la raíz de `Plugins`**, solo a la subcarpeta del
  plugin. Además detecta copias sueltas dejadas por instalaciones anteriores y ofrece
  eliminarlas.
- `AssemblyVersion` / `AssemblyFileVersion` estaban congeladas en 2.1.0.0 mientras el diálogo
  "Acerca de" avanzaba solo; ahora se mantienen sincronizadas.
- Verificado por ejecución en un layout de carpetas idéntico al de XrmToolBox: no responde a
  peticiones de otros assemblies, no entrega EPPlus 4.5.3.3 a quien pide EPPlus 7, no responde
  por librerías ajenas, resuelve su propia dependencia desde la subcarpeta, y no lanza
  excepción ante un nombre de assembly corrupto.

## 2.1.9.0 — Buscador de tablas

- Búsqueda **multi-término**: acepta varios nombres separados por espacio, coma, punto y coma,
  tabulación, salto de línea o barra vertical, con coincidencia OR. Permite pegar una lista
  completa de tablas y verlas todas a la vez, en lugar de buscarlas de una en una.
- Búsqueda **sin acentos**: "admision" encuentra "Admisión", "accion" encuentra "Acción".
- La búsqueda cubre también el **nombre de esquema**, además del visible y el lógico.
- Nuevo filtro rápido **"Selected"**: muestra solo las tablas marcadas, para revisar la
  selección antes de exportar.
- Contador ampliado: `N seleccionada(s) | N mostrada(s) de N cargada(s)`.
- Corregido: cada clic en los botones de filtro creaba tres objetos `Font` sin liberarlos.

## 2.1.8.0 — Progreso en todas las operaciones

- Progreso "X de Y" en Export ERD, Show Relationships y ambos exports de matriz de relaciones
  (las dos matrices no mostraban ningún progreso).
- Corregido: el export de matriz truncaba el nombre lógico a 31 caracteres para nombrar la hoja
  **sin verificar unicidad**, y `ExcelWorksheets.Add` lanza excepción ante duplicados: con miles
  de tablas habría abortado el export completo.
- Corregido: `ws.Dimension.Address` sin comprobación de nulo reventaba con
  `NullReferenceException` en cuanto una tabla no tenía relaciones 1:N.

## 2.1.7.0 — Enlaces y archivo válido

- Corregido el diálogo **"Hemos encontrado un problema con contenido"**: la columna de metadata
  concatena todas las opciones de un OptionSet en una celda y superaba el límite duro de Excel
  de 32.767 caracteres. EPPlus escribe cadenas más largas sin protestar, pero Excel declara el
  archivo dañado y al repararlo **descarta el formato**.
- Corregido que los hipervínculos no se vieran como enlaces en filas alternas: el rayado cebra
  se aplicaba después del enlace y `StyleName` sobre un rango reemplaza el estilo completo de la
  celda. Ahora hay estilos con nombre `LinkCell` / `LinkCellZebra` aplicados después del cebra.

## 2.1.6.0 — Navegación y referencias

- Enlace **"◄ Volver al Índice"** en la parte superior de cada hoja de tabla.
- Relaciones **navegables**: en 1:N, N:1 y N:N el nombre de la tabla relacionada es un enlace a
  su propia hoja. Degrada a texto plano si esa tabla no está en el mismo archivo.
- Corregida la colisión masiva de nombres de hoja: se conservan los dos extremos del nombre
  lógico en lugar de truncar al inicio.
- Línea de cobertura en el índice: cuántas tablas se documentaron de las solicitadas.

## 2.1.5.0 — Causa raíz del archivo vacío

- **El defecto de fondo de toda la serie.** `worker.ReportProgress` lanzaba
  `InvalidOperationException` ("Este BackgroundWorker indica que no notifica el progreso")
  porque XrmToolBox solo activa `WorkerReportsProgress` cuando el `WorkAsyncInfo` declara un
  handler `ProgressChanged`, y el plugin no lo declaraba en ninguna de sus operaciones. Como esa
  llamada era la primera línea del `try` por tabla, las 2.735 tablas se saltaban una por una y
  el `.xlsx` salía solo con la portada del índice. No tenía relación con memoria, con Dataverse
  ni con el tamaño del entorno, y venía desde la versión original.
- `ReportProgress` ahora verifica `WorkerReportsProgress` y captura excepciones: reportar
  progreso es cosmético y no puede romper el export.
- Se agregó el handler `ProgressChanged` a las tres operaciones asincrónicas.

## 2.1.4.0 — Opción de archivo único

- Casilla **"Single file"** (marcada por defecto): vuelve a exportar todo en un solo `.xlsx`,
  el comportamiento original. Al desmarcarla, divide en archivos de 20 tablas.
- El instalador detecta si XrmToolBox está abierto (mantiene el DLL bloqueado) y ofrece
  cerrarlo, en lugar de fallar con un error de acceso.

## 2.1.3.0 — Reversión de regresión

- Revertido el timeout por tabla de la 2.1.2.0: se había implementado moviendo cada llamada a
  Dataverse a otro hilo con `Task.Run`, lo que rompía la conexión para prácticamente todas las
  tablas. Se conservó solo la reducción del tamaño de lote.

## 2.1.2.0 — (retirada)

- Introdujo un timeout de 30 s por tabla mediante `Task.Run`, que resultó ser una regresión.
  No usar.

## 2.1.1.0 — Primer fix del crash

- `Show Relationships` creaba ~8 controles WinForms por tabla seleccionada sin límite, con
  riesgo de agotar los handles USER/GDI de Windows: limitado a 50 tablas por vista interactiva.

## 2.1.0.0 — Correcciones de estabilidad del export masivo

- `EntityFilters` acotado a lo que el exportador realmente lee (antes pedía `All`, incluyendo
  `Privileges`, que nunca se usa).
- Filtro de tablas de sistema con casilla en la UI (desmarcada por defecto).
- Estilos con nombre reutilizados en lugar de estilos "inline" repetidos por fila, anchos de
  columna fijos en exports grandes, y división automática en varios archivos.
- Logging real (`LogInfo`/`LogWarning`/`LogError`) y `try/catch` de nivel superior.
- Diálogo de confirmación con conteo de tablas y estimación de tiempo.
- Cableado real de `ExportOptions` / `Settings` a la interfaz, y soporte de cancelación.
