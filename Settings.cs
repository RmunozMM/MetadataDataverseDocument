using System;

namespace MetadataDataverseDocument
{
    public class Settings
    {
        public string LastUsedOrganizationWebappUrl { get; set; }
        public string LastSelectedSolution { get; set; }
        public bool IncludeAlternateKeys { get; set; } = true;
        public bool IncludeAttributes { get; set; } = true;
        public bool IncludeOneToMany { get; set; } = true;
        public bool IncludeManyToOne { get; set; } = true;
        public bool IncludeManyToMany { get; set; } = true;
        public bool GenerateIndexSheet { get; set; } = true;

        // Default true: matches the plugin's original, expected behavior of a single consolidated
        // .xlsx for the whole data dictionary. When false, the export is split into multiple
        // smaller files (MetadataExcelExporter.DefaultMaxSheetsPerFile tables each) as a safer
        // fallback for very large organizations where a single huge in-memory workbook risks
        // running out of memory.
        public bool CombineIntoSingleFile { get; set; } = true;

        // Ancho en pixeles del panel izquierdo (lista de tablas). Se guarda cada vez que el
        // usuario mueve el separador, para que su ajuste sobreviva al cierre del plugin.
        // 380 es el valor por defecto: alcanza para los cuatro botones de filtro en una fila
        // y para leer los nombres de tabla sin que se corten.
        public int LeftPanelWidth { get; set; } = 380;

        public string LastExportFolder { get; set; }
    }
}
