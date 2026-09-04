================================================================================
Metadata Dataverse Document - Plugin para XrmToolBox
Versión: 2.1.0.0 (Release)
Desarrollador: Rogelio Muñoz
Sitio Web: http://www.rogeliomunoz.cl
Contacto: rmunoz1612@gmail.com
Copyright © Rogelio Muñoz 2026. Todos los derechos reservados.
================================================================================

DESCRIPCIÓN
--------------------------------------------------------------------------------
Metadata Dataverse Document es un plugin profesional de alta velocidad para XrmToolBox
diseñado para documentar soluciones, esquemas, diccionarios de datos y relaciones de 
Microsoft Dataverse / Dynamics 365.

CARACTERÍSTICAS PRINCIPALES
--------------------------------------------------------------------------------
1. Diccionario de Datos Técnico Integral en Excel:
   - Resumen y ficha técnica completa por tabla.
   - Tabla detallada de Atributos (Nombre lógico, esquema, tipo, requerimiento,
     auditoría, seguridad FLS, opciones OptionSet con código y etiqueta, motivos de
     estado StatusCode/StateCode, rangos y longitud).
   - Tabla de Claves Alternas (Alternate Keys) con estado y campos compuestos.
   - Tablas de Relaciones 1:N con comportamientos en cascada (Delete, Assign, Share,
     Unshare, Reparent, Rollup).
   - Tablas de Relaciones N:1 con atributos de búsqueda Lookup.
   - Tablas de Relaciones N:N con tablas intermedias de cruce.
   - Pestaña de índice interactivo con hipervínculos directos.

2. Generador de Diagramas ERD (Mermaid / Markdown):
   - Exporta diagramas de Entidad-Relación en formato estándar Mermaid para visualizar
     la arquitectura de datos en Markdown, GitHub, Azure DevOps, Notion o VS Code.

3. Filtros Rápidos en Interfaz (UI):
   - Filtros instantáneos: [Todas], [Solo Personalizadas], [Solo Estándar].
   - Búsqueda en vivo por nombre visible o lógico.
   - Selección y deselección masiva con contador dinámico.

4. Exportación de Matrices de Relaciones:
   - Exportación de matrices relacionales tanto en Excel (.xlsx) como en formato web HTML.

REQUISITOS
--------------------------------------------------------------------------------
- XrmToolBox v1.2021 o superior
- .NET Framework 4.8
- Conexión activa a un entorno de Microsoft Dataverse / Dynamics 365

INSTALACIÓN RÁPIDA
--------------------------------------------------------------------------------
Haga clic derecho en 'Install-MetadataDataverseDocument.ps1' y seleccione:
"Ejecutar con PowerShell" (Run with PowerShell).

El script desbloqueará automáticamente los archivos descargados y los copiará al
directorio de plugins de XrmToolBox (%APPDATA%\MscrmTools\XrmToolBox\Plugins\).

INSTALACIÓN MANUAL
--------------------------------------------------------------------------------
1. Cierre XrmToolBox si está abierto.
2. Desbloquee los archivos DLL (Clic derecho -> Propiedades -> Desbloquear / Unblock).
3. Copie 'MetadataDataverseDocument.dll' a:
   %APPDATA%\MscrmTools\XrmToolBox\Plugins\
4. Cree una subcarpeta llamada 'MetadataDataverseDocument' en Plugins y copie allí 'EPPlus.dll':
   %APPDATA%\MscrmTools\XrmToolBox\Plugins\MetadataDataverseDocument\EPPlus.dll
5. Abra XrmToolBox y busque "Metadata Dataverse Document".
