using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk.Metadata;

namespace MetadataDataverseDocument.Exporters
{
    public sealed class HtmlDocumentationExporter
    {
        public static void ExportToHtmlFile(string filePath, List<EntityMetadata> entities)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang=\"en\">");
            sb.AppendLine("<head>");
            sb.AppendLine("    <meta charset=\"UTF-8\">");
            sb.AppendLine("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
            sb.AppendLine("    <title>Metadata Dataverse Document - Relationship Report</title>");
            sb.AppendLine("    <style>");
            sb.AppendLine("        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8fafc; color: #1e293b; margin: 0; padding: 20px; }");
            sb.AppendLine("        .header { background-color: #1e293b; color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }");
            sb.AppendLine("        .header h1 { margin: 0 0 5px 0; font-size: 24px; }");
            sb.AppendLine("        .header p { margin: 0; font-size: 13px; color: #94a3b8; }");
            sb.AppendLine("        .card { background: white; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); padding: 20px; margin-bottom: 25px; border: 1px solid #e2e8f0; }");
            sb.AppendLine("        .card h2 { margin-top: 0; color: #0f172a; font-size: 18px; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; }");
            sb.AppendLine("        table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 13px; }");
            sb.AppendLine("        th { background-color: #334155; color: white; padding: 10px; text-align: left; }");
            sb.AppendLine("        td { padding: 8px 10px; border-bottom: 1px solid #e2e8f0; }");
            sb.AppendLine("        tr:nth-child(even) { background-color: #f8fafc; }");
            sb.AppendLine("        .badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: bold; }");
            sb.AppendLine("        .badge-custom { background-color: #dbeafe; color: #1e40af; }");
            sb.AppendLine("        .badge-standard { background-color: #f1f5f9; color: #475569; }");
            sb.AppendLine("    </style>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine("    <div class=\"header\">");
            sb.AppendLine("        <h1>Metadata Dataverse Document</h1>");
            sb.AppendLine($"        <p>Generated on {DateTime.Now:yyyy-MM-dd HH:mm:ss} | Author: Rogelio Muñoz (www.rogeliomunoz.cl)</p>");
            sb.AppendLine("    </div>");

            foreach (var entity in entities.OrderBy(e => e.DisplayName?.UserLocalizedLabel?.Label ?? e.LogicalName))
            {
                string disp = entity.DisplayName?.UserLocalizedLabel?.Label ?? entity.LogicalName;
                bool isCustom = entity.IsCustomEntity.GetValueOrDefault(false);

                sb.AppendLine("    <div class=\"card\">");
                sb.AppendLine($"        <h2>{disp} <small>({entity.LogicalName})</small> <span class=\"badge {(isCustom ? "badge-custom" : "badge-standard")}\">{(isCustom ? "Custom" : "Standard")}</span></h2>");

                // 1:N
                if (entity.OneToManyRelationships != null && entity.OneToManyRelationships.Length > 0)
                {
                    sb.AppendLine("        <h3>1:N Relationships</h3>");
                    sb.AppendLine("        <table>");
                    sb.AppendLine("            <tr><th>Relationship Name</th><th>Related Table</th><th>Foreign Key</th><th>Delete</th><th>Assign</th></tr>");
                    foreach (var r in entity.OneToManyRelationships.OrderBy(x => x.SchemaName))
                    {
                        sb.AppendLine($"            <tr><td>{r.SchemaName}</td><td>{r.ReferencingEntity}</td><td>{r.ReferencingAttribute}</td><td>{r.CascadeConfiguration?.Delete}</td><td>{r.CascadeConfiguration?.Assign}</td></tr>");
                    }
                    sb.AppendLine("        </table>");
                }

                // N:1
                if (entity.ManyToOneRelationships != null && entity.ManyToOneRelationships.Length > 0)
                {
                    sb.AppendLine("        <h3>N:1 Relationships (Lookups)</h3>");
                    sb.AppendLine("        <table>");
                    sb.AppendLine("            <tr><th>Relationship Name</th><th>Parent Table</th><th>Lookup Attribute</th><th>Referenced PK</th></tr>");
                    foreach (var r in entity.ManyToOneRelationships.OrderBy(x => x.SchemaName))
                    {
                        sb.AppendLine($"            <tr><td>{r.SchemaName}</td><td>{r.ReferencedEntity}</td><td>{r.ReferencingAttribute}</td><td>{r.ReferencedAttribute}</td></tr>");
                    }
                    sb.AppendLine("        </table>");
                }

                // N:N
                if (entity.ManyToManyRelationships != null && entity.ManyToManyRelationships.Length > 0)
                {
                    sb.AppendLine("        <h3>N:N Relationships</h3>");
                    sb.AppendLine("        <table>");
                    sb.AppendLine("            <tr><th>Relationship Name</th><th>Intersect Table</th><th>Associated Table</th></tr>");
                    foreach (var r in entity.ManyToManyRelationships.OrderBy(x => x.SchemaName))
                    {
                        string other = r.Entity1LogicalName.Equals(entity.LogicalName, StringComparison.OrdinalIgnoreCase) ? r.Entity2LogicalName : r.Entity1LogicalName;
                        sb.AppendLine($"            <tr><td>{r.SchemaName}</td><td>{r.IntersectEntityName}</td><td>{other}</td></tr>");
                    }
                    sb.AppendLine("        </table>");
                }

                sb.AppendLine("    </div>");
            }

            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }
    }
}
