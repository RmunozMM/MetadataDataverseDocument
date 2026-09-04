# Metadata Dataverse Document

Plugin para **XrmToolBox** que documenta la metadata y las relaciones de un entorno de
Microsoft Dataverse / Dynamics 365: diccionario de datos técnico en Excel, diagramas ERD en
Mermaid y matrices de relaciones en Excel/HTML.

- **Autor:** Rogelio Muñoz — [www.rogeliomunoz.cl](http://www.rogeliomunoz.cl)
- **Versión actual:** 2.1.9.0
- **Plataforma:** .NET Framework 4.8, WinForms
- **Repositorio privado.** Todos los derechos reservados.

---

## Qué hace

| Botón | Qué genera |
|---|---|
| **Export Data Dictionary** | Un `.xlsx` con una hoja por tabla: ficha técnica, atributos (tipo, requerimiento, auditoría, FLS, OptionSets), claves alternas y relaciones 1:N / N:1 / N:N con comportamientos en cascada. Incluye hoja de índice con enlaces a cada tabla, enlace de vuelta al índice en cada hoja, y referencias cruzadas navegables entre tablas. |
| **Show Relationships** | Vista interactiva en pestañas, hasta 50 tablas (más que eso agota los handles de Windows). |
| **Export ERD (Mermaid)** | Diagrama entidad-relación en formato Mermaid para Markdown, GitHub, Azure DevOps o Notion. |
| **Export Rel. Matrix** | Matriz de relaciones en `.xlsx` o reporte `.html`. |

El panel izquierdo permite buscar tablas por varios términos a la vez (espacio, coma, punto y
coma o salto de línea), sin distinguir acentos, y filtrar por All / Custom / Standard / Selected.

---

## Cómo compilar

Requiere **Visual Studio 2019 o superior** con soporte de .NET Framework 4.8.

```
git clone <url-del-repo>
cd MetadataDataverseDocument-Source
```

Abre `MetadataDataverseDocument.sln` y compila en **Release**. El resultado queda en
`bin\Release\MetadataDataverseDocument.dll`.

### Sobre las referencias en `lib/`

Las 9 DLLs de terceros necesarias para compilar (XrmToolBox, SDK de Dataverse, EPPlus,
McTools) están versionadas en **`lib/`**, así que un clon nuevo compila sin pasos previos.

> Originalmente el `.csproj` apuntaba a `bin\Release\*.dll`, es decir a la **misma carpeta de
> salida del build**. Eso hacía que las referencias fueran borrables por un *Clean* y que no se
> pudiera ignorar `bin/` en git sin romper la compilación. Se movieron a `lib/`.

`lib/` incluye además `System.Resources.Extensions.dll`, que no es referencia de compilación
pero **sí** es dependencia de ejecución del plugin instalado.

---

## Cómo empaquetar el instalador

El contenido de `dist/` (scripts de instalación) está versionado; los `.zip` y `.dll`
generados **no** — se publican como *GitHub Releases*.

Un instalador es un `.zip` con:

```
MetadataDataverseDocument.dll          <- compilado desde bin\Release
EPPlus.dll                             <- desde lib/
System.Resources.Extensions.dll        <- desde lib/
Instalar-Plugin.bat                    <- desde dist/
Install-MetadataDataverseDocument.ps1  <- desde dist/
README.txt                             <- desde dist/
```

El script instala en `%APPDATA%\MscrmTools\XrmToolBox\Plugins\` y **exige que XrmToolBox esté
cerrado** (si está abierto mantiene el DLL bloqueado; desde la 2.1.4 el script lo detecta y
ofrece cerrarlo).

---

## Estructura

```
Plugin.cs                         Punto de entrada (IXrmToolBoxPlugin, metadata de XrmToolBox)
Settings.cs                       Preferencias persistidas por SettingsManager
Models/ExportOptions.cs           Flags de qué incluir en el diccionario
UI/MetadataDocumentControl.cs     Control principal: carga, búsqueda, filtros, orquestación
UI/*.designer.cs                  Layout generado por el diseñador
Exporters/MetadataExcelExporter.cs    Diccionario de datos en Excel (EPPlus)
Exporters/MermaidErdExporter.cs       Diagramas ERD
Exporters/HtmlDocumentationExporter.cs Reporte HTML
lib/                              DLLs de terceros para compilar (versionadas)
dist/                             Scripts de instalación (los .zip van a Releases)
Resources/                        Iconos del plugin
```

---

## Trampas ya pagadas (leer antes de tocar el exportador)

Cuatro defectos reales que costaron varias versiones de encontrar. Están documentados aquí para
no repetirlos.

**1. `worker.ReportProgress` lanza excepción si el `WorkAsyncInfo` no declara `ProgressChanged`.**
XrmToolBox solo activa `BackgroundWorker.WorkerReportsProgress` cuando se le pasa un handler
`ProgressChanged`. Sin él, cada llamada a `ReportProgress` lanza `InvalidOperationException`.
Como esa llamada era la primera línea del `try` de cada tabla, **las 2.735 tablas se saltaban
una por una** y el archivo salía con la portada del índice y cero contenido. Toda operación con
`WorkAsync` que reporte progreso debe declarar su `ProgressChanged`.

**2. Excel no admite más de 32.767 caracteres por celda; EPPlus los escribe igual.**
La columna de metadata concatena todas las opciones de un OptionSet en una celda. Un OptionSet
grande la desborda, y Excel entonces declara el archivo dañado, lo repara y **descarta el
formato** (por eso los hipervínculos dejaban de verse como enlaces). Todo texto de celda pasa
por `ClampCellText`.

**3. Asignar `StyleName` a un rango reemplaza el estilo COMPLETO de la celda.**
El rayado cebra se aplicaba después de escribir el hipervínculo y le borraba el azul y el
subrayado, en filas alternas. El estilo del enlace debe aplicarse **después** del cebra, con los
estilos con nombre `LinkCell` / `LinkCellZebra`.

**4. Los nombres de hoja de Excel se limitan a 31 caracteres y deben ser únicos.**
Truncar al inicio colapsa familias enteras de tablas (`wit_<x>_wit_<y>`) en nombres casi
idénticos, y `Worksheets.Add` **lanza excepción** ante un duplicado. Se conservan los dos
extremos del nombre lógico (20 primeros + `~` + 10 últimos) y se resuelve unicidad con sufijo.

---

## Diagnóstico

El plugin escribe en el log de XrmToolBox:

```
%APPDATA%\MscrmTools\XrmToolBox\Logs\MetadataDataverseDocument.log
```

Cada tabla que no se pudo documentar queda ahí como `Warning` con el motivo. Además, la fila 3
de la hoja de índice indica cuántas tablas se documentaron de las solicitadas (verde si están
todas, ámbar si falta alguna).

---

## Mejoras pendientes

- EPPlus está fijado en 4.5.3.3 (LGPL) y construye el libro completo en memoria: no permite
  escritura incremental. Migrar a una librería con escritura en streaming permitiría exportar
  un único archivo muy grande con memoria acotada.
- Las DLLs de `lib/` podrían venir de paquetes NuGet (`XrmToolBox.Extensibility`,
  `Microsoft.CrmSdk.XrmTooling.CoreAssembly`) en lugar de estar versionadas.
