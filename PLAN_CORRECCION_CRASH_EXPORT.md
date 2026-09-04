# Plan de corrección — Crash de XrmToolBox al "Exportar todo el modelo"

**Plugin:** Metadata Dataverse Document (v2.1.0.0)
**Ruta analizada:** `MetadataDataverseDocument-Source/`
**Flujo afectado:** `UI/MetadataDocumentControl.cs` → botón *Export Data Dictionary* con "exportar TODAS las tablas cargadas" → `Exporters/MetadataExcelExporter.cs`

## 1. Resumen del problema

Cuando cargas todas las tablas de la organización ("Default Solution (All Entities)") y luego usas *Export Data Dictionary* respondiendo "Sí" a "¿exportar todas las tablas cargadas?", el plugin recorre **cada tabla del entorno** (incluyendo entidades estándar, de sistema, de actividades, auditoría, etc. — normalmente varios cientos en un Dataverse típico), pide su metadata completa y arma un único libro Excel en memoria con una hoja por tabla. Esa combinación es la que hace que el proceso de XrmToolBox se quede sin memoria y se cierre de golpe, sin mostrar un cuadro de error — que es exactamente el síntoma que describes (no es una excepción "manejada", es el propio proceso muriendo).

No es un solo bug puntual: son **tres decisiones de diseño que se potencian entre sí**. Corrigiendo la primera ya deberías ver una mejora enorme; las tres juntas dejan el export en un estado robusto.

## 2. Diagnóstico técnico (causas, ordenadas por impacto)

### 2.1. Sobre-solicitud de metadata por tabla — `EntityFilters.All` (impacto alto)

En `MetadataExcelExporter.RetrieveFullEntity` (línea ~111-122) cada tabla se recupera así:

```csharp
var req = new RetrieveEntityRequest
{
    LogicalName = logicalName,
    EntityFilters = EntityFilters.All,   // <-- trae TODO
    RetrieveAsIfPublished = true
};
```

`EntityFilters.All` incluye `Entity | Attributes | Privileges | Relationships`. El exporter **nunca usa los privilegios** (`Privileges`), pero igual los pide y los deserializa para cada una de las tablas. Esto multiplica innecesariamente el tamaño de cada respuesta SOAP/WCF y la memoria retenida en el `EntityMetadata` de cada tabla, para un dato que después se descarta. El mismo patrón se repite en `MetadataDocumentControl.cs` en `ExecuteShowRelationships`, `PromptExportErd`, `PromptExportRelationshipsExcel` y `PromptExportRelationshipsHtml` (ahí al menos ya usan `Relationships | Entity`, sin `Privileges`, así que el peor caso está concentrado en el exporter de Excel).

### 2.2. Sin filtro de alcance — "todo el modelo" es literalmente todo (impacto alto)

`ExecuteLoadTables(null)` (línea ~202-208) usa:

```csharp
var req = new RetrieveAllEntitiesRequest
{
    EntityFilters = EntityFilters.Entity,
    RetrieveAsIfPublished = true
};
```

sin excluir entidades de sistema, no personalizables, de auditoría, colas, logs de plugin, etc. Cuando el usuario elige "Default Solution (All Entities)" y luego "exportar todas", `PromptExportDataDictionary` hace `selected = _allEntities` (línea ~624) — es decir, se documentan también las tablas internas de la plataforma que casi nunca se necesitan en un diccionario de datos. En un Dataverse actual eso fácilmente son varios cientos de tablas adicionales que no aportan valor al documento pero sí consumen memoria y tiempo.

### 2.3. Todo el libro Excel se construye en memoria, con una hoja por tabla (impacto alto)

En `MetadataExcelExporter.Export` se crea **una `ExcelWorksheet` por cada tabla** (línea ~79), con formato, merges, `AutoFitColumns()` y coloreado de filas alternas por cada sección (claves, atributos, 1:N, N:1, N:N). EPPlus (la versión usada aquí es la 4.5.x LGPL, de 2020) mantiene **todo el paquete en memoria hasta el `SaveAs` final**, que además comprime todo el ZIP de una sola vez de forma síncrona. Con cientos de hojas, cada una con su propio formato aplicado celda por celda, el consumo de memoria crece de forma no lineal y es la causa más probable del `OutOfMemoryException` que termina matando el proceso (una `OutOfMemoryException` fuera del hilo de UI, o durante una operación nativa de compresión/GDI+, muchas veces no puede ser "atrapada" limpiamente por .NET y tira abajo todo el host — de ahí que XrmToolBox se cierre en vez de mostrarte un mensaje de error).

### 2.4. Falta de manejo de errores a nivel de operación completa (impacto medio)

Dentro de `Export()`, el `try/catch` (línea 63-101) sólo protege el procesamiento de **una** tabla a la vez — si falla la construcción del índice (`PopulateIndexSheet`) o el propio `SavePackage`/`package.SaveAs(...)` (por ejemplo, por quedarse sin memoria durante la compresión final, o por un problema de disco/ruta), esa falla no queda contenida y se propaga sin un mensaje claro. Esto no es la causa raíz, pero hace que cuando algo sí falla, el usuario no reciba ningún diagnóstico útil — solo ve que la herramienta "se cierra".

### 2.5. Sin confirmación ni indicación de magnitud antes de lanzar la operación (impacto medio)

`PromptExportDataDictionary` sólo pregunta "¿exportar TODAS las tablas cargadas?" (Sí/No) pero no informa cuántas tablas son ni el tiempo/memoria estimados, y no hay forma de cancelar una vez iniciado el `WorkAsync`. El usuario no tiene forma de saber, antes de lanzar la exportación, que está a punto de pedir metadata completa de 600+ tablas.

### 2.6. Detalle menor — fuentes GDI no liberadas (impacto bajo)

En `BuildModernHeaderButtons` se crean objetos `Font` (`IDisposable`) por cada botón sin `Dispose()`, y el método se puede volver a ejecutar (`MetadataDocumentControl_Load`). No es la causa del crash puntual que describes, pero es un leak de handles GDI que conviene limpiar ya que estás tocando este archivo.

## 3. Plan de corrección (en orden de prioridad)

### Prioridad 1 — Reducir el filtro de metadata al mínimo necesario

En `MetadataExcelExporter.RetrieveFullEntity`, cambiar:

```csharp
EntityFilters = EntityFilters.All,
```

por:

```csharp
EntityFilters = EntityFilters.Entity | EntityFilters.Attributes | EntityFilters.Relationships,
```

Esto elimina `Privileges` (dato que nunca se usa) del payload de cada tabla sin cambiar el resultado del documento. Aplica el mismo criterio en los otros tres lugares del `MetadataDocumentControl.cs` que ya usan `Relationships | Entity` — para el diccionario completo agrega `Attributes` si en algún momento decides mostrar más detalle ahí también.

### Prioridad 2 — Excluir del "cargar todo" las tablas que no aportan al diccionario

Al construir `entities` en `ExecuteLoadTables`, filtrar por defecto las que no son útiles en un diccionario de datos de negocio, por ejemplo:

```csharp
entities = entities
    .Where(e => e.IsValidForAdvancedFind.GetValueOrDefault(false) || e.IsCustomEntity.GetValueOrDefault(false))
    .ToList();
```

y ofrecer en la UI un checkbox "Incluir tablas de sistema" (desmarcado por defecto) para quien realmente las necesite. Esto reduce drásticamente el universo de tablas en el escenario típico de "exportar todo", que es el que hoy revienta el proceso.

### Prioridad 3 — No mantener todo el libro en memoria de una sola vez

Aquí hay dos caminos, de menor a mayor esfuerzo:

- **Rápido (mitigación):** antes de lanzar el export cuando `selected.Count` supera un umbral (por ejemplo 150-200 tablas), mostrar una advertencia con el número real de tablas y tiempo estimado, y ofrecer **dividir automáticamente la salida en varios archivos .xlsx** (por ejemplo, uno cada 100 tablas). Esto es un cambio contenido en `PromptExportDataDictionary` + `MetadataExcelExporter.Export` (agregar un parámetter `int? maxSheetsPerFile` y generar `archivo_1.xlsx`, `archivo_2.xlsx`, etc.).
- **De fondo (correcto a mediano plazo):** reemplazar los estilos "inline" repetidos (`range.Style.Fill...SetColor(...)` en cada fila/sección) por **estilos con nombre reutilizados** (`package.Workbook.Styles.NamedStyles.Add(...)`), definidos una sola vez y aplicados por referencia. Esto reduce sustancialmente la tabla de estilos interna de EPPlus y el pico de memoria, además de acercarte al límite práctico de estilos distintos por libro de Excel (~64.000) que un export de cientos de tablas con estilos ad-hoc puede llegar a rozar.

En cualquiera de los dos caminos, evita llamar `AutoFitColumns()` cuando el número de tablas a exportar sea grande — es una operación cara (usa medición de texto vía GDI) y se puede reemplazar por anchos de columna fijos para el modo "exportar todo".

### Prioridad 4 — Envolver toda la operación en un try/catch de nivel superior con mensaje claro

En `MetadataExcelExporter.Export`, envolver el bloque completo (no solo el `foreach`) para que cualquier falla en `PopulateIndexSheet` o `SavePackage` también llegue como `args.Error` legible en el `PostWorkCallBack`, en vez de un fallo silencioso o un mensaje genérico. Adicionalmente, registrar en un log de texto (junto al .xlsx de salida) qué tabla se estaba procesando cuando ocurrió el problema — hoy ese dato solo va a `Debug.WriteLine`, que no es visible fuera de un depurador adjunto.

### Prioridad 5 — Confirmar alcance real antes de ejecutar

Cambiar el diálogo de `PromptExportDataDictionary` para que, cuando no hay tablas marcadas, muestre el número real de tablas que se exportarán (`_allEntities.Count`) y una estimación de tiempo (por ejemplo, "≈X segundos por tabla × N tablas"), en vez de solo "¿exportar todas?". Esto le da al usuario información para decidir si de verdad quiere ese alcance o prefiere filtrar primero por solución.

### Prioridad 6 (opcional, buena práctica) — Liberar los `Font` creados en `BuildModernHeaderButtons`

Guardar las instancias de `Font` en campos y hacer `Dispose()` de las anteriores antes de crear nuevas, o reutilizar una única instancia de `Font` compartida para todos los botones, ya que hoy se crea una nueva cada vez que se reconstruyen los botones.

## 4. Cómo validar que quedó resuelto

1. Reproducir primero el crash actual contra una copia de prueba (mismo número de tablas) para tener una línea base de memoria — puedes abrir el Administrador de Tareas y ver el consumo de `XrmToolBox.exe` justo antes de que se cierre.
2. Aplicar la Prioridad 1 y 2 solas, y repetir el export completo — deberían bajar drásticamente tanto el tiempo total como el pico de memoria, incluso sin tocar el resto.
3. Si con eso ya no se cae, aplicar igualmente la Prioridad 3 (al menos el corte automático en archivos múltiples) como red de seguridad para organizaciones más grandes o futuras.
4. Probar el caso límite real: cargar "Default Solution (All Entities)" sin ningún filtro y lanzar el export tal como lo hace hoy el usuario que reporta el problema, confirmando que ya no cierra XrmToolBox y que, si algo falla, aparece un `MessageBox` con el detalle del error en vez de un cierre silencioso.
5. Revisar en el Administrador de Tareas si `XrmToolBox.exe` corre en modo x86 o x64 (columna "Nombre" no lo indica; hay que mirar en la pestaña Detalles si aparece "*32" junto al proceso). Si corre en 32 bits, migrar a la versión 64 bits de XrmToolBox eliminaría por completo el techo de ~2-4 GB de memoria direccionable, que es el límite más probable que se está golpeando.

## 5. Mejoras a mediano plazo (no bloquean el fix, pero vale la pena planificarlas)

- Exponer en la UI las opciones que ya existen en el modelo `ExportOptions` (hoy siempre se instancia con `new ExportOptions()` por defecto, y las de `Settings.cs` — `IncludeAlternateKeys`, `IncludeAttributes`, `IncludeRelationships` — se cargan pero nunca se pasan al exporter ni se guardan con `SaveSettings()`). Conectar esas preferencias reales daría control fino sin tocar código cada vez.
- Agregar soporte de cancelación al `WorkAsync` (usar el `worker.CancellationPending` que ya expone `BackgroundWorker`) para que una exportación larga se pueda abortar limpiamente en vez de forzar el cierre de XrmToolBox por impaciencia.
- Migrar de EPPlus 4.5.x (LGPL, 2020) a una versión más reciente o a una librería con escritura en streaming (por ejemplo, `ClosedXML` con SAX writer, o el propio EPPlus 7 bajo licencia comercial) si el volumen de tablas de la Universidad sigue creciendo — el enfoque "todo en memoria" tiene un techo estructural independientemente de las optimizaciones anteriores.
