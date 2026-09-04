using System;
using System.Collections.Generic;

namespace MetadataDataverseDocument.Models
{
    public sealed class ExportOptions
    {
        public bool IncludeAttributes { get; set; } = true;
        public bool IncludeAlternateKeys { get; set; } = true;
        public bool IncludeOneToMany { get; set; } = true;
        public bool IncludeManyToOne { get; set; } = true;
        public bool IncludeManyToMany { get; set; } = true;
        public bool GenerateIndexSheet { get; set; } = true;
    }

    public sealed class TableSummaryInfo
    {
        public string SheetName { get; set; }
        public string DisplayName { get; set; }
        public string LogicalName { get; set; }
        public string SchemaName { get; set; }
        public bool IsCustom { get; set; }
        public int AttributeCount { get; set; }
        public int KeyCount { get; set; }
        public int OneToManyCount { get; set; }
        public int ManyToOneCount { get; set; }
        public int ManyToManyCount { get; set; }
    }
}
