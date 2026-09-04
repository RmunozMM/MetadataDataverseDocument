using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk.Metadata;

namespace MetadataDataverseDocument.Exporters
{
    public sealed class MermaidErdExporter
    {
        public static void ExportToMarkdownFile(string filePath, List<EntityMetadata> entities)
        {
            if (entities == null || entities.Count == 0)
            {
                File.WriteAllText(filePath, "# Dataverse ERD Diagram\n\nNo entities selected.\n", Encoding.UTF8);
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine("# Dataverse Entity Relationship Diagram (ERD)");
            sb.AppendLine();
            sb.AppendLine($"> Generated on: {DateTime.Now:yyyy-MM-dd HH:mm:ss} | Documented Entities: {entities.Count} | Author: Rogelio Muñoz (www.rogeliomunoz.cl)");
            sb.AppendLine();
            sb.AppendLine("```mermaid");
            sb.AppendLine("erDiagram");

            var entitySet = new HashSet<string>(entities.Select(e => e.LogicalName), StringComparer.OrdinalIgnoreCase);
            var writtenRelationships = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 1. Write Entity definitions with key attributes
            foreach (var entity in entities.OrderBy(e => e.LogicalName))
            {
                string entityAlias = CleanIdentifier(entity.LogicalName);
                string displayName = entity.DisplayName?.UserLocalizedLabel?.Label ?? entity.LogicalName;

                sb.AppendLine($"    {entityAlias} {{");
                
                string pk = entity.PrimaryIdAttribute ?? $"{entity.LogicalName}id";
                sb.AppendLine($"        string {CleanIdentifier(pk)} PK \"Primary Key\"");

                if (!string.IsNullOrEmpty(entity.PrimaryNameAttribute) && !entity.PrimaryNameAttribute.Equals(pk, StringComparison.OrdinalIgnoreCase))
                {
                    sb.AppendLine($"        string {CleanIdentifier(entity.PrimaryNameAttribute)} \"Primary Name\"");
                }

                sb.AppendLine("    }");
            }

            sb.AppendLine();

            // 2. Write 1:N Relationships
            foreach (var entity in entities)
            {
                string fromEntity = CleanIdentifier(entity.LogicalName);

                if (entity.OneToManyRelationships != null)
                {
                    foreach (var rel in entity.OneToManyRelationships)
                    {
                        if (string.IsNullOrEmpty(rel.ReferencingEntity)) continue;

                        // Only connect if target is in the exported set
                        if (entitySet.Contains(rel.ReferencingEntity))
                        {
                            string toEntity = CleanIdentifier(rel.ReferencingEntity);
                            string relKey = $"{fromEntity}_{toEntity}_{rel.SchemaName}";

                            if (!writtenRelationships.Contains(relKey))
                            {
                                writtenRelationships.Add(relKey);
                                string label = string.IsNullOrEmpty(rel.ReferencingAttribute) ? rel.SchemaName : rel.ReferencingAttribute;
                                sb.AppendLine($"    {fromEntity} ||--o{{ {toEntity} : \"{CleanLabel(label)}\"");
                            }
                        }
                    }
                }

                // 3. Write N:N Relationships
                if (entity.ManyToManyRelationships != null)
                {
                    foreach (var nn in entity.ManyToManyRelationships)
                    {
                        string otherEntity = nn.Entity1LogicalName.Equals(entity.LogicalName, StringComparison.OrdinalIgnoreCase) 
                            ? nn.Entity2LogicalName 
                            : nn.Entity1LogicalName;

                        if (!string.IsNullOrEmpty(otherEntity) && entitySet.Contains(otherEntity))
                        {
                            string ent1 = CleanIdentifier(entity.LogicalName);
                            string ent2 = CleanIdentifier(otherEntity);
                            
                            // Normalise key to avoid duplicate reverse lines
                            string relKey = string.Compare(ent1, ent2, StringComparison.OrdinalIgnoreCase) < 0
                                ? $"NN_{ent1}_{ent2}_{nn.SchemaName}"
                                : $"NN_{ent2}_{ent1}_{nn.SchemaName}";

                            if (!writtenRelationships.Contains(relKey))
                            {
                                writtenRelationships.Add(relKey);
                                sb.AppendLine($"    {ent1} }}o--o{{ {ent2} : \"{CleanLabel(nn.SchemaName)}\"");
                            }
                        }
                    }
                }
            }

            sb.AppendLine("```");
            sb.AppendLine();
            sb.AppendLine("---");
            sb.AppendLine("### Tables Included in Diagram:");
            sb.AppendLine();
            sb.AppendLine("| Display Name | Schema Name | Logical Name | Type |");
            sb.AppendLine("| :--- | :--- | :--- | :--- |");
            foreach (var e in entities.OrderBy(e => e.DisplayName?.UserLocalizedLabel?.Label ?? e.LogicalName))
            {
                string disp = e.DisplayName?.UserLocalizedLabel?.Label ?? e.LogicalName;
                string type = e.IsCustomEntity.GetValueOrDefault(false) ? "Custom" : "Standard";
                sb.AppendLine($"| {disp} | {e.SchemaName} | `{e.LogicalName}` | {type} |");
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        private static string CleanIdentifier(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Unknown";
            string cleaned = Regex.Replace(name, @"[^a-zA-Z0-9_]", "_");
            if (char.IsDigit(cleaned[0])) cleaned = "_" + cleaned;
            return cleaned;
        }

        private static string CleanLabel(string label)
        {
            if (string.IsNullOrEmpty(label)) return string.Empty;
            return label.Replace("\"", "'").Trim();
        }
    }
}
