using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using MetadataDataverseDocument.Models;

namespace MetadataDataverseDocument.Exporters
{
    public sealed class MetadataExcelExporter
    {
        private readonly IOrganizationService service;
        private readonly BackgroundWorker worker;
        private readonly HashSet<string> usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Maps a table's logical name to the worksheet name it gets in the CURRENT output file.
        // Computed for the whole batch before any sheet is written, so that while documenting one
        // table we already know the sheet name of every other table in the same file and can turn
        // its relationship rows into real clickable cross-references (a table's related tables were
        // previously just plain text, which in a 2000+ table dictionary makes the relationships
        // effectively un-navigable). Reset per output file, like usedSheetNames.
        private readonly Dictionary<string, string> sheetNameByLogicalName =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Name of the index worksheet in the current output file, or null when the user turned the
        // index off - drives the "back to index" link written at the top of every entity sheet.
        private string indexSheetName;

        // P4: this class does not inherit XrmToolBox.Extensibility.PluginControlBase (only
        // MetadataDocumentControl does, and only it has LogInfo/LogWarning/LogError), so the caller
        // hands us plain delegates pointed at its own logging methods. Defaulted to no-ops so this
        // class stays usable (e.g. from tests) without a logging host.
        private readonly Action<string> logInfo;
        private readonly Action<string> logWarning;
        private readonly Action<string> logError;

        // Round 4 - "Cancelación real": checked at the top of the per-file batch loop and the
        // per-table loop inside Export(), so a user-requested cancellation (via the XrmToolBox
        // progress dialog's Cancel button, which sets BackgroundWorker.CancellationPending) stops
        // further work at the next safe boundary instead of running to completion. `worker` can be
        // null (e.g. this class used outside a WorkAsync context), so this defaults to "not
        // cancelled" in that case rather than throwing.
        private bool IsCancellationRequested => worker != null && worker.CancellationPending;

        // Color scheme
        private static readonly Color HeaderNavy = Color.FromArgb(30, 41, 59);        // #1E293B
        private static readonly Color SectionKeys = Color.FromArgb(15, 118, 110);      // #0F766E (Teal)
        private static readonly Color SectionAttributes = Color.FromArgb(3, 105, 161); // #0369A1 (Sky blue)
        private static readonly Color SectionOneToMany = Color.FromArgb(180, 83, 9);   // #B45309 (Amber)
        private static readonly Color SectionManyToOne = Color.FromArgb(109, 40, 217); // #6D28D9 (Purple)
        private static readonly Color SectionManyToMany = Color.FromArgb(190, 24, 93); // #BE185D (Pink/Magenta)
        private static readonly Color ZebraStripe = Color.FromArgb(248, 250, 252);     // #F8FAFC
        private static readonly Color LightBorder = Color.FromArgb(226, 232, 240);     // #E2E8F0

        // P3(b): above this many tables in the export, AutoFitColumns() (which measures every
        // cell's rendered text with a Graphics object, per sheet) is skipped in favor of fixed,
        // "good enough" column widths. AutoFitColumns gives a nicer result for a handful of
        // tables, but its cost scales with sheet count and is one of the contributors to the
        // OutOfMemoryException described in BRIEFING.md P3 when exporting hundreds of tables.
        private const int AutoFitColumnsThreshold = 50;

        // P3(c): above this many tables, the workbook is split into several .xlsx files instead
        // of one workbook with hundreds of sheets, each carrying its own style table, shared
        // strings, etc. 150 is the threshold suggested in BRIEFING.md P3(c) and keeps each
        // individual package's in-memory footprint (and AutoFitColumns/style cost, see above)
        // bounded regardless of how many tables are exported in total.
        // Lowered from 150 -> 20 after a real crash report on a ~2700-table org: smaller files
        // finish (and get saved to disk) much faster and more often, so a problem late in a huge
        // run only loses the current small file's worth of work instead of hours of progress, and
        // each individual SaveAs/workbook stays small regardless of how many tables the org has.
        public const int DefaultMaxSheetsPerFile = 20;

        // P3(a): named styles (OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml, created via
        // ExcelStyles.CreateNamedStyle and applied via ExcelRangeBase.StyleName) replace the
        // "inline" range.Style.Font/.Fill... assignments that used to be repeated for every
        // section header, column header row and zebra-striped row of every exported table -
        // each inline assignment allocates a new style record in the workbook's style table, and
        // with hundreds of tables x several sections x many rows each, those records were a
        // major contributor to the memory blow-up. Each named style below is created and
        // configured exactly once per output package (see CreateNamedStyles), then referenced by
        // name everywhere it used to be set inline.
        private const string StyleZebraRow = "ZebraRow";
        private const string StyleEntityTitleBanner = "EntityTitleBanner";
        private const string StyleOverviewLabel = "OverviewLabel";
        private const string StyleSectionHeaderKeys = "SectionHeaderKeys";
        private const string StyleSectionHeaderAttributes = "SectionHeaderAttributes";
        private const string StyleSectionHeaderOneToMany = "SectionHeaderOneToMany";
        private const string StyleSectionHeaderManyToOne = "SectionHeaderManyToOne";
        private const string StyleSectionHeaderManyToMany = "SectionHeaderManyToMany";
        private const string StyleColumnHeaderKeys = "ColumnHeaderKeys";
        private const string StyleColumnHeaderAttributes = "ColumnHeaderAttributes";
        private const string StyleColumnHeaderOneToMany = "ColumnHeaderOneToMany";
        private const string StyleColumnHeaderManyToOne = "ColumnHeaderManyToOne";
        private const string StyleColumnHeaderManyToMany = "ColumnHeaderManyToMany";

        // Link cells need their own NAMED styles rather than the inline
        // Font.UnderLine/Font.Color they used to get. The zebra-stripe style was applied to the
        // whole row AFTER the link cell was written, and assigning ExcelRangeBase.StyleName
        // replaces a cell's entire style - so on every zebra row the blue underline was wiped and
        // the hyperlink rendered as ordinary black text (working, but not looking like a link).
        // Two variants so a link keeps its appearance on plain and striped rows alike.
        private const string StyleLinkCell = "LinkCell";
        private const string StyleLinkCellZebra = "LinkCellZebra";

        /// <summary>
        /// Creates the exporter. <paramref name="logInfo"/>/<paramref name="logWarning"/>/
        /// <paramref name="logError"/> let a host that isn't a PluginControlBase-derived class (see
        /// P4 remarks below) still route this class's logging through its own logging methods.
        /// </summary>
        /// <param name="service">Organization service used to retrieve full entity metadata.</param>
        /// <param name="worker">Background worker used to report progress; may be null.</param>
        /// <param name="logInfo">
        /// P4: called with a human-readable message once the export finishes successfully. Pass the
        /// host's <c>LogInfo</c> (e.g. <c>msg =&gt; LogInfo(msg)</c>). Optional - defaults to a no-op.
        /// </param>
        /// <param name="logWarning">
        /// P4: called once per table that failed to export but did not abort the whole file (the
        /// export continues with the remaining tables). Pass the host's <c>LogWarning</c>. Optional -
        /// defaults to a no-op.
        /// </param>
        /// <param name="logError">
        /// P4: called when generating an entire output file fails (package/style/save-level
        /// failure, as opposed to a single table). Pass the host's <c>LogError</c>. Optional -
        /// defaults to a no-op.
        /// </param>
        public MetadataExcelExporter(
            IOrganizationService service,
            BackgroundWorker worker,
            Action<string> logInfo = null,
            Action<string> logWarning = null,
            Action<string> logError = null)
        {
            this.service = service ?? throw new ArgumentNullException(nameof(service));
            this.worker = worker;
            this.logInfo = logInfo ?? (_ => { });
            this.logWarning = logWarning ?? (_ => { });
            this.logError = logError ?? (_ => { });
        }

        /// <summary>
        /// Generates one or more .xlsx data dictionaries for <paramref name="entitiesToExport"/>.
        /// Returns the list of file paths actually written - normally just <paramref name="filePath"/>,
        /// but when the number of tables exceeds <paramref name="maxSheetsPerFile"/> the output is
        /// split into <c>{name}_1.xlsx</c>, <c>{name}_2.xlsx</c>, etc. (see P3(c) in BRIEFING.md),
        /// each one a self-contained workbook with its own Index sheet.
        /// </summary>
        public List<string> Export(string filePath, List<EntityMetadata> entitiesToExport, ExportOptions options = null, int maxSheetsPerFile = DefaultMaxSheetsPerFile)
        {
            if (options == null) options = new ExportOptions();
            var generatedFiles = new List<string>();

            if (entitiesToExport == null || entitiesToExport.Count == 0)
            {
                // P4: package creation/styling/save were previously unprotected here - any failure
                // (e.g. an invalid filePath, a locked file, a disk/IO error) propagated as a bare
                // exception with no indication of what was being generated. Wrap it, log it, and
                // rethrow with context so it still reaches PostWorkCallBack as a legible args.Error.
                try
                {
                    using (var package = new ExcelPackage())
                    {
                        usedSheetNames.Clear();
                        CreateNamedStyles(package);

                        // Same options.GenerateIndexSheet gating as the main path below, for
                        // consistency - this branch is not reachable from the current caller
                        // (PromptExportDataDictionary always passes at least one entity or returns
                        // early), but it should still honor the option rather than special-case it.
                        if (options.GenerateIndexSheet)
                        {
                            var indexSheet = package.Workbook.Worksheets.Add("Index");
                            usedSheetNames.Add("Index");
                            BuildIndexHeader(indexSheet, null);
                            indexSheet.Cells[4, 1].Value = "No tables were provided for export.";
                        }
                        else
                        {
                            // Defensive fallback: EPPlus cannot save a workbook with zero
                            // worksheets, and there are no entity sheets in this branch either.
                            var placeholder = package.Workbook.Worksheets.Add("Index");
                            usedSheetNames.Add("Index");
                            placeholder.Cells[1, 1].Value = "No tables were provided for export.";
                        }
                        SavePackage(package, filePath);
                    }
                }
                catch (Exception ex)
                {
                    logError($"Failed to generate '{filePath}': {ex.Message}");
                    throw new InvalidOperationException($"Failed to generate '{filePath}': {ex.Message}", ex);
                }

                logInfo($"Data dictionary export finished: 1 file generated ('{filePath}'), no tables were provided.");
                generatedFiles.Add(filePath);
                return generatedFiles;
            }

            var orderedEntities = entitiesToExport
                .OrderBy(e => e.DisplayName?.UserLocalizedLabel?.Label ?? e.LogicalName)
                .ToList();

            int overallTotal = orderedEntities.Count;
            // Fixed widths avoid the per-sheet AutoFitColumns() cost across a large export (see
            // AutoFitColumnsThreshold above); the decision is based on the overall table count,
            // not the per-file batch size, since splitting into several files does not change how
            // expensive the export as a whole is.
            bool useFixedColumnWidths = overallTotal > AutoFitColumnsThreshold;

            var batches = ChunkList(orderedEntities, Math.Max(1, maxSheetsPerFile));
            bool multiFile = batches.Count > 1;
            int overallProcessed = 0;

            bool cancelledEarly = false;

            for (int batchIndex = 0; batchIndex < batches.Count; batchIndex++)
            {
                // Round 4 - "Cancelación real": checked before starting a new output file at all,
                // so a cancellation between batches never creates a package that would just be
                // thrown away. Whatever files were already saved in generatedFiles are kept.
                if (IsCancellationRequested)
                {
                    logInfo($"Data dictionary export cancelled by user before starting file {batchIndex + 1} of {batches.Count}.");
                    cancelledEarly = true;
                    break;
                }

                var batch = batches[batchIndex];
                string batchFilePath = multiFile ? BuildBatchFilePath(filePath, batchIndex + 1, batches.Count) : filePath;
                string batchInfo = multiFile
                    ? $"File {batchIndex + 1} of {batches.Count} - {batch.Count} of {overallTotal} tables in this file."
                    : null;

                // P4: everything below - package creation, named styles, index population and the
                // final SaveAs - used to run with no try/catch of its own; only the per-table loop
                // was protected. A failure anywhere else here (e.g. SavePackage hitting a locked or
                // invalid path) previously aborted the whole Export() call with a bare exception,
                // discarding the in-memory `generatedFiles` list built so far even though earlier
                // batches had already been written to disk successfully. Wrapping each batch keeps
                // that context: which file/batch failed, and how many files had already succeeded.
                try
                {
                    using (var package = new ExcelPackage())
                    {
                        // Sheet names only need to be unique within a single workbook, so this is
                        // reset for every output file rather than accumulated across the whole export.
                        usedSheetNames.Clear();
                        CreateNamedStyles(package);

                        // Round 4 - "Configurabilidad real": options.GenerateIndexSheet used to be
                        // declared on ExportOptions but never actually consulted anywhere in this
                        // class (a "phantom" option) - it now really gates whether the Index sheet
                        // is created at all.
                        ExcelWorksheet indexSheet = null;
                        indexSheetName = null;
                        if (options.GenerateIndexSheet)
                        {
                            indexSheet = package.Workbook.Worksheets.Add("Index");
                            usedSheetNames.Add("Index");
                            indexSheetName = indexSheet.Name;
                            BuildIndexHeader(indexSheet, batchInfo);
                        }

                        // Reserve the worksheet name of every table in this file UP FRONT, before
                        // documenting any of them. Two reasons: (1) each entity sheet can link to
                        // the sheets of its related tables even though those sheets do not exist
                        // yet at the time its relationship rows are written, and (2) sheet-name
                        // collisions are resolved deterministically over the whole batch instead of
                        // depending on the order tables happen to complete in.
                        sheetNameByLogicalName.Clear();
                        foreach (var entity in batch)
                        {
                            if (string.IsNullOrEmpty(entity.LogicalName) ||
                                sheetNameByLogicalName.ContainsKey(entity.LogicalName))
                            {
                                continue;
                            }

                            string reserved = BuildSafeSheetName(
                                GetLocalizedLabel(entity.DisplayName, entity.LogicalName),
                                entity.LogicalName);
                            usedSheetNames.Add(reserved);
                            sheetNameByLogicalName[entity.LogicalName] = reserved;
                        }

                        var tableSummaries = new List<TableSummaryInfo>();

                        foreach (var entity in batch)
                        {
                            // Round 4 - "Cancelación real": checked at the top of the per-table
                            // loop so a cancellation mid-batch stops documenting further tables.
                            // The current batch's index sheet (if any) and file are still populated
                            // and saved below with whatever tables were completed so far, rather
                            // than being discarded - the outer batch loop's own check (above) then
                            // stops before starting any further output file.
                            if (IsCancellationRequested)
                            {
                                logInfo($"Data dictionary export cancelled by user while documenting file {batchIndex + 1} of {batches.Count} (completed {tableSummaries.Count} of {batch.Count} tables in this file).");
                                cancelledEarly = true;
                                break;
                            }

                            try
                            {
                                overallProcessed++;
                                ReportProgress((int)((overallProcessed / (double)overallTotal) * 100),
                                    $"Documenting table {overallProcessed} of {overallTotal}: {entity.LogicalName}...");

                                var fullEntity = RetrieveFullEntity(entity.LogicalName);
                                if (fullEntity == null)
                                {
                                    continue;
                                }

                                // Use the name reserved for this table before the loop (see above),
                                // so the cross-reference links other sheets already wrote to it
                                // resolve. Fall back to computing one only if this table somehow
                                // was not in the reservation pass.
                                string safeSheetName;
                                if (!sheetNameByLogicalName.TryGetValue(fullEntity.LogicalName, out safeSheetName))
                                {
                                    safeSheetName = BuildSafeSheetName(
                                        GetLocalizedLabel(fullEntity.DisplayName, fullEntity.LogicalName),
                                        fullEntity.LogicalName);
                                    usedSheetNames.Add(safeSheetName);
                                    sheetNameByLogicalName[fullEntity.LogicalName] = safeSheetName;
                                }

                                var sheet = package.Workbook.Worksheets.Add(safeSheetName);

                                WriteEntitySheet(sheet, fullEntity, options, useFixedColumnWidths);

                                tableSummaries.Add(new TableSummaryInfo
                                {
                                    SheetName = safeSheetName,
                                    DisplayName = GetLocalizedLabel(fullEntity.DisplayName, fullEntity.LogicalName),
                                    LogicalName = fullEntity.LogicalName,
                                    SchemaName = fullEntity.SchemaName,
                                    IsCustom = fullEntity.IsCustomEntity.GetValueOrDefault(false),
                                    AttributeCount = fullEntity.Attributes?.Length ?? 0,
                                    KeyCount = fullEntity.Keys?.Length ?? 0,
                                    OneToManyCount = fullEntity.OneToManyRelationships?.Length ?? 0,
                                    ManyToOneCount = fullEntity.ManyToOneRelationships?.Length ?? 0,
                                    ManyToManyCount = fullEntity.ManyToManyRelationships?.Length ?? 0
                                });
                            }
                            catch (Exception ex)
                            {
                                // A single table failing to document does not abort the file - log it
                                // as a warning (not an error) and move on to the next table, same
                                // behavior as before this change, just with real logging instead of
                                // Debug.WriteLine (which is silently compiled out of Release builds).
                                logWarning($"[Exporter] Skipped table '{entity.LogicalName}' (file {batchIndex + 1} of {batches.Count}): {ex.Message}");
                            }
                        }

                        if (options.GenerateIndexSheet)
                        {
                            // Coverage line: states how many of the tables requested for this file
                            // actually made it in. Previously a table that failed to document was
                            // only mentioned in the plugin log, so a workbook could come out
                            // missing tables with nothing in the file itself saying so.
                            int missing = batch.Count - tableSummaries.Count;
                            string coverage = missing == 0
                                ? $"{tableSummaries.Count} table(s) documented in this file - all requested tables were documented successfully."
                                : $"{tableSummaries.Count} of {batch.Count} requested table(s) documented in this file - {missing} could not be documented (see the plugin log: XrmToolBox > Help > Open plugin log).";
                            if (!string.IsNullOrEmpty(batchInfo))
                            {
                                coverage = $"{batchInfo} {coverage}";
                            }

                            indexSheet.Cells[3, 1].Value = coverage;
                            using (var range = indexSheet.Cells[3, 1, 3, 10])
                            {
                                range.Merge = true;
                                range.Style.Font.Bold = true;
                                range.Style.Font.Size = 10;
                                range.Style.Font.Color.SetColor(missing == 0
                                    ? Color.FromArgb(21, 128, 61)    // green - complete
                                    : Color.FromArgb(180, 83, 9));   // amber - something missing
                            }

                            // Populate index sheet rows
                            PopulateIndexSheet(indexSheet, tableSummaries, useFixedColumnWidths);
                        }
                        else if (package.Workbook.Worksheets.Count == 0)
                        {
                            // Defensive fallback: with GenerateIndexSheet off there is no Index
                            // sheet to fall back on, and EPPlus cannot save a workbook with zero
                            // worksheets. This only triggers in the edge case where every table in
                            // this batch failed to generate a sheet (see the per-table catch below)
                            // or the batch was cancelled before documenting any table.
                            var placeholder = package.Workbook.Worksheets.Add("Index");
                            usedSheetNames.Add("Index");
                            placeholder.Cells[1, 1].Value = "No sheets were generated for this file.";
                        }

                        SavePackage(package, batchFilePath);
                    }

                    generatedFiles.Add(batchFilePath);
                }
                catch (Exception ex)
                {
                    string context = multiFile
                        ? $"Failed to generate file {batchIndex + 1} of {batches.Count} ('{batchFilePath}')"
                        : $"Failed to generate '{batchFilePath}'";
                    string alreadyGenerated = generatedFiles.Count > 0
                        ? $" {generatedFiles.Count} file(s) had already been generated successfully before this failure."
                        : string.Empty;
                    string fullMessage = $"{context}: {ex.Message}.{alreadyGenerated}";

                    logError(fullMessage);
                    throw new InvalidOperationException(fullMessage, ex);
                }
            }

            if (cancelledEarly)
            {
                logInfo($"Data dictionary export cancelled by user: {generatedFiles.Count} file(s) generated for {overallProcessed} of {overallTotal} table(s) before cancellation.");
            }
            else
            {
                logInfo($"Data dictionary export finished: {generatedFiles.Count} file(s) generated for {overallTotal} table(s).");
            }
            return generatedFiles;
        }

        /// <summary>
        /// Splits <paramref name="source"/> into consecutive chunks of at most <paramref name="size"/>
        /// items each. Manual implementation because this project targets .NET Framework 4.8, where
        /// Enumerable.Chunk (added in .NET 6) is not available.
        /// </summary>
        private static List<List<T>> ChunkList<T>(List<T> source, int size)
        {
            var result = new List<List<T>>();
            for (int i = 0; i < source.Count; i += size)
            {
                result.Add(source.GetRange(i, Math.Min(size, source.Count - i)));
            }
            return result;
        }

        /// <summary>
        /// Inserts a "_{fileNumber}" suffix before the extension of <paramref name="basePath"/>,
        /// e.g. "C:\Dict.xlsx" -&gt; "C:\Dict_2.xlsx". Used only when the export is split across
        /// multiple files (P3(c)).
        /// </summary>
        private static string BuildBatchFilePath(string basePath, int fileNumber, int totalFiles)
        {
            string dir = Path.GetDirectoryName(basePath);
            string nameNoExt = Path.GetFileNameWithoutExtension(basePath);
            string ext = Path.GetExtension(basePath);
            string newName = $"{nameNoExt}_{fileNumber}{ext}";
            return string.IsNullOrEmpty(dir) ? newName : Path.Combine(dir, newName);
        }

        /// <summary>
        /// Creates and configures (once per output package) every named style referenced via
        /// ExcelRangeBase.StyleName elsewhere in this class. Must run before any worksheet in
        /// <paramref name="package"/> is populated. See the Style* constants above for what each
        /// name is used for.
        /// </summary>
        private void CreateNamedStyles(ExcelPackage package)
        {
            var styles = package.Workbook.Styles;

            var zebra = styles.CreateNamedStyle(StyleZebraRow);
            zebra.Style.Fill.PatternType = ExcelFillStyle.Solid;
            zebra.Style.Fill.BackgroundColor.SetColor(ZebraStripe);

            var titleBanner = styles.CreateNamedStyle(StyleEntityTitleBanner);
            titleBanner.Style.Font.Bold = true;
            titleBanner.Style.Font.Size = 14;
            titleBanner.Style.Font.Color.SetColor(Color.White);
            titleBanner.Style.Fill.PatternType = ExcelFillStyle.Solid;
            titleBanner.Style.Fill.BackgroundColor.SetColor(HeaderNavy);
            titleBanner.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            var overviewLabel = styles.CreateNamedStyle(StyleOverviewLabel);
            overviewLabel.Style.Font.Bold = true;
            overviewLabel.Style.Fill.PatternType = ExcelFillStyle.Solid;
            overviewLabel.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(241, 245, 249));

            CreateSectionHeaderStyle(styles, StyleSectionHeaderKeys, SectionKeys);
            CreateSectionHeaderStyle(styles, StyleSectionHeaderAttributes, SectionAttributes);
            CreateSectionHeaderStyle(styles, StyleSectionHeaderOneToMany, SectionOneToMany);
            CreateSectionHeaderStyle(styles, StyleSectionHeaderManyToOne, SectionManyToOne);
            CreateSectionHeaderStyle(styles, StyleSectionHeaderManyToMany, SectionManyToMany);

            CreateColumnHeaderStyle(styles, StyleColumnHeaderKeys, Color.FromArgb(204, 251, 241));       // Light teal
            CreateColumnHeaderStyle(styles, StyleColumnHeaderAttributes, Color.FromArgb(224, 242, 254)); // Light sky
            CreateColumnHeaderStyle(styles, StyleColumnHeaderOneToMany, Color.FromArgb(254, 243, 199));  // Light amber
            CreateColumnHeaderStyle(styles, StyleColumnHeaderManyToOne, Color.FromArgb(243, 232, 255));  // Light purple
            CreateColumnHeaderStyle(styles, StyleColumnHeaderManyToMany, Color.FromArgb(252, 231, 243)); // Light pink

            var link = styles.CreateNamedStyle(StyleLinkCell);
            link.Style.Font.UnderLine = true;
            link.Style.Font.Color.SetColor(Color.Blue);

            var linkZebra = styles.CreateNamedStyle(StyleLinkCellZebra);
            linkZebra.Style.Font.UnderLine = true;
            linkZebra.Style.Font.Color.SetColor(Color.Blue);
            linkZebra.Style.Fill.PatternType = ExcelFillStyle.Solid;
            linkZebra.Style.Fill.BackgroundColor.SetColor(ZebraStripe);
        }

        private static void CreateSectionHeaderStyle(ExcelStyles styles, string name, Color backColor)
        {
            var namedStyle = styles.CreateNamedStyle(name);
            namedStyle.Style.Font.Bold = true;
            namedStyle.Style.Font.Color.SetColor(Color.White);
            namedStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
            namedStyle.Style.Fill.BackgroundColor.SetColor(backColor);
        }

        private static void CreateColumnHeaderStyle(ExcelStyles styles, string name, Color backColor)
        {
            var namedStyle = styles.CreateNamedStyle(name);
            namedStyle.Style.Font.Bold = true;
            namedStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
            namedStyle.Style.Fill.BackgroundColor.SetColor(backColor);
        }

        private EntityMetadata RetrieveFullEntity(string logicalName)
        {
            var req = new RetrieveEntityRequest
            {
                LogicalName = logicalName,
                // Only request the filters this exporter actually reads (Entity, Attributes,
                // Relationships). EntityFilters.All also pulls Privileges metadata, which is never
                // used below and needlessly increases the payload/memory for every table retrieved.
                EntityFilters = EntityFilters.Entity | EntityFilters.Attributes | EntityFilters.Relationships,
                RetrieveAsIfPublished = true
            };

            var resp = (RetrieveEntityResponse)service.Execute(req);
            return resp.EntityMetadata;
        }

        private void BuildIndexHeader(ExcelWorksheet sheet, string batchInfo)
        {
            // This banner and the column header row below are built once per output file (not
            // once per table/row), so they are left as one-off "inline" styles rather than named
            // styles - P3(a) targets the styles that get repeated per table/row, which this is not.
            sheet.Cells[1, 1].Value = "Metadata Dataverse Document - Index & Data Dictionary";
            using (var range = sheet.Cells[1, 1, 1, 10])
            {
                range.Merge = true;
                range.Style.Font.Bold = true;
                range.Style.Font.Size = 16;
                range.Style.Font.Color.SetColor(Color.White);
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(HeaderNavy);
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }
            sheet.Row(1).Height = 35;

            sheet.Cells[2, 1].Value = $"Generated on: {DateTime.Now:yyyy-MM-dd HH:mm:ss} | Author: Rogelio Muñoz (www.rogeliomunoz.cl)";
            using (var range = sheet.Cells[2, 1, 2, 10])
            {
                range.Merge = true;
                range.Style.Font.Italic = true;
                range.Style.Font.Size = 10;
                range.Style.Font.Color.SetColor(Color.DarkSlateGray);
            }

            // Row 3 is written by Export() AFTER the per-table loop (the coverage line: how many
            // tables were actually documented, plus the batch info when the export is split), so
            // that it can state the real outcome rather than only the intent. Nothing is written
            // here to avoid merging the same range twice.

            string[] headers = {
                "Entity Display Name",
                "Schema Name",
                "Logical Name",
                "Type",
                "Attributes",
                "Alternate Keys",
                "1:N Rel.",
                "N:1 Rel.",
                "N:N Rel.",
                "Go to Sheet"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                sheet.Cells[4, i + 1].Value = headers[i];
            }

            using (var range = sheet.Cells[4, 1, 4, headers.Length])
            {
                range.Style.Font.Bold = true;
                range.Style.Font.Color.SetColor(Color.White);
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(HeaderNavy);
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            sheet.Row(4).Height = 24;
        }

        private void PopulateIndexSheet(ExcelWorksheet sheet, List<TableSummaryInfo> summaries, bool useFixedColumnWidths)
        {
            int row = 5;
            foreach (var info in summaries)
            {
                sheet.Cells[row, 1].Value = info.DisplayName;
                sheet.Cells[row, 2].Value = info.SchemaName;
                sheet.Cells[row, 3].Value = info.LogicalName;
                sheet.Cells[row, 4].Value = info.IsCustom ? "Custom" : "Standard";
                sheet.Cells[row, 5].Value = info.AttributeCount;
                sheet.Cells[row, 6].Value = info.KeyCount;
                sheet.Cells[row, 7].Value = info.OneToManyCount;
                sheet.Cells[row, 8].Value = info.ManyToOneCount;
                sheet.Cells[row, 9].Value = info.ManyToManyCount;

                // Clickable hyperlink to sheet
                sheet.Cells[row, 10].Formula = $"HYPERLINK(\"#'{EscapeSheetFormula(info.SheetName)}'!A1\", \"View Sheet\")";

                // ORDER MATTERS: the zebra style is applied to the whole row FIRST, and only then
                // the link cell gets its own style. Doing it the other way round (as before) made
                // StyleName on the row overwrite the link cell's entire style, so on every striped
                // row the link lost its blue underline and looked like plain text.
                bool zebra = row % 2 == 0;
                if (zebra)
                {
                    using (var range = sheet.Cells[row, 1, row, 10])
                    {
                        range.StyleName = StyleZebraRow;
                    }
                }
                sheet.Cells[row, 10].StyleName = zebra ? StyleLinkCellZebra : StyleLinkCell;

                row++;
            }

            if (sheet.Dimension != null)
            {
                if (useFixedColumnWidths)
                {
                    SetFixedIndexColumnWidths(sheet);
                }
                else
                {
                    sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                }
                sheet.View.FreezePanes(5, 1);
            }
        }

        /// <summary>
        /// P3(b): fixed column widths for the Index sheet, used instead of AutoFitColumns() once
        /// the export is large enough (see AutoFitColumnsThreshold). Sized to comfortably fit the
        /// header labels defined in BuildIndexHeader without measuring actual cell content.
        /// </summary>
        private static void SetFixedIndexColumnWidths(ExcelWorksheet sheet)
        {
            double[] widths = { 32, 26, 26, 11, 11, 12, 9, 9, 9, 13 };
            for (int i = 0; i < widths.Length; i++)
            {
                sheet.Column(i + 1).Width = widths[i];
            }
        }

        private void WriteEntitySheet(ExcelWorksheet sheet, EntityMetadata entity, ExportOptions options, bool useFixedColumnWidths)
        {
            int row = 1;
            string displayName = GetLocalizedLabel(entity.DisplayName, entity.LogicalName);

            // Title Banner
            sheet.Cells[row, 1].Value = $"{displayName} ({entity.LogicalName})";
            using (var range = sheet.Cells[row, 1, row, 11])
            {
                range.Merge = true;
                range.StyleName = StyleEntityTitleBanner;
            }
            sheet.Row(row).Height = 30;
            row++;

            // "Back to index" link: with thousands of sheets, scrolling the tab bar back to the
            // Index is impractical, so every entity sheet gets a one-click way home directly under
            // its title. Written only when an Index sheet actually exists in this file
            // (options.GenerateIndexSheet).
            if (!string.IsNullOrEmpty(indexSheetName))
            {
                sheet.Cells[row, 1].Formula =
                    $"HYPERLINK(\"#'{EscapeSheetFormula(indexSheetName)}'!A1\", \"◄ Volver al Índice\")";
                sheet.Cells[row, 1].StyleName = StyleLinkCell;
            }
            row += 2;

            // Overview Key-Value Block
            row = WriteOverviewItem(sheet, row, "Display Name", displayName);
            row = WriteOverviewItem(sheet, row, "Plural Name", GetLocalizedLabel(entity.DisplayCollectionName, entity.SchemaName));
            row = WriteOverviewItem(sheet, row, "Schema Name", entity.SchemaName);
            row = WriteOverviewItem(sheet, row, "Logical Name", entity.LogicalName);
            row = WriteOverviewItem(sheet, row, "Description", ClampCellText(GetLocalizedLabel(entity.Description, "N/A")));
            row = WriteOverviewItem(sheet, row, "Primary ID Attribute", entity.PrimaryIdAttribute ?? "N/A");
            row = WriteOverviewItem(sheet, row, "Primary Name Attribute", entity.PrimaryNameAttribute ?? "N/A");
            row = WriteOverviewItem(sheet, row, "Object Type Code", entity.ObjectTypeCode?.ToString(CultureInfo.InvariantCulture) ?? "N/A");
            row = WriteOverviewItem(sheet, row, "Is Custom Entity", entity.IsCustomEntity.GetValueOrDefault(false) ? "Yes" : "No");
            row = WriteOverviewItem(sheet, row, "Ownership Type", entity.OwnershipType?.ToString() ?? "N/A");
            row = WriteOverviewItem(sheet, row, "Total Attributes", (entity.Attributes?.Length ?? 0).ToString());
            row = WriteOverviewItem(sheet, row, "Alternate Keys", (entity.Keys?.Length ?? 0).ToString());
            row = WriteOverviewItem(sheet, row, "1:N Relationships", (entity.OneToManyRelationships?.Length ?? 0).ToString());
            row = WriteOverviewItem(sheet, row, "N:1 Relationships", (entity.ManyToOneRelationships?.Length ?? 0).ToString());
            row = WriteOverviewItem(sheet, row, "N:N Relationships", (entity.ManyToManyRelationships?.Length ?? 0).ToString());

            row += 1;

            // Section 1: Alternate Keys (if any or enabled)
            if (options.IncludeAlternateKeys)
            {
                row = WriteAlternateKeysSection(sheet, row, entity);
                row += 2;
            }

            // Section 2: Attributes
            if (options.IncludeAttributes)
            {
                row = WriteAttributesSection(sheet, row, entity);
                row += 2;
            }

            // Section 3: 1:N Relationships
            if (options.IncludeOneToMany)
            {
                row = WriteOneToManySection(sheet, row, entity);
                row += 2;
            }

            // Section 4: N:1 Relationships
            if (options.IncludeManyToOne)
            {
                row = WriteManyToOneSection(sheet, row, entity);
                row += 2;
            }

            // Section 5: N:N Relationships
            if (options.IncludeManyToMany)
            {
                row = WriteManyToManySection(sheet, row, entity);
            }

            if (sheet.Dimension != null)
            {
                if (useFixedColumnWidths)
                {
                    SetFixedEntityColumnWidths(sheet);
                }
                else
                {
                    sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                }
                sheet.View.FreezePanes(18, 1);
            }
        }

        /// <summary>
        /// P3(b): fixed column widths for an entity sheet, used instead of AutoFitColumns() once
        /// the export is large enough (see AutoFitColumnsThreshold). Columns 1-11 are shared by
        /// several sections (Overview, Alternate Keys, Attributes, 1:N/N:1/N:N Relationships)
        /// with different meanings each, so these widths are a reasonable general-purpose fit
        /// rather than a per-section exact measurement.
        /// </summary>
        private static void SetFixedEntityColumnWidths(ExcelWorksheet sheet)
        {
            double[] widths = { 28, 30, 26, 16, 34, 14, 12, 14, 14, 12, 50 };
            for (int i = 0; i < widths.Length; i++)
            {
                sheet.Column(i + 1).Width = widths[i];
            }
        }

        private int WriteOverviewItem(ExcelWorksheet sheet, int row, string label, string value)
        {
            sheet.Cells[row, 1].Value = label;
            sheet.Cells[row, 1].StyleName = StyleOverviewLabel;

            sheet.Cells[row, 2].Value = value;
            using (var range = sheet.Cells[row, 2, row, 11])
            {
                range.Merge = true;
            }
            sheet.Row(row).Height = 20;
            return row + 1;
        }

        private int WriteAlternateKeysSection(ExcelWorksheet sheet, int row, EntityMetadata entity)
        {
            // Section Header
            sheet.Cells[row, 1].Value = "ALTERNATE KEYS (ENTITY KEYS)";
            using (var range = sheet.Cells[row, 1, row, 5])
            {
                range.Merge = true;
                range.StyleName = StyleSectionHeaderKeys;
            }
            sheet.Row(row).Height = 22;
            row++;

            string[] headers = { "Key Name", "Display Name", "Schema Name", "Key Attributes", "Index Status" };
            for (int i = 0; i < headers.Length; i++) sheet.Cells[row, i + 1].Value = headers[i];

            using (var range = sheet.Cells[row, 1, row, headers.Length])
            {
                range.StyleName = StyleColumnHeaderKeys;
            }
            row++;

            if (entity.Keys == null || entity.Keys.Length == 0)
            {
                sheet.Cells[row, 1].Value = "No alternate keys defined for this table.";
                sheet.Cells[row, 1].Style.Font.Italic = true;
                return row + 1;
            }

            foreach (var key in entity.Keys)
            {
                sheet.Cells[row, 1].Value = key.LogicalName;
                sheet.Cells[row, 2].Value = GetLocalizedLabel(key.DisplayName, key.LogicalName);
                sheet.Cells[row, 3].Value = key.SchemaName;
                sheet.Cells[row, 4].Value = key.KeyAttributes != null ? ClampCellText(string.Join(", ", key.KeyAttributes)) : string.Empty;
                sheet.Cells[row, 5].Value = key.EntityKeyIndexStatus.ToString();

                if (row % 2 == 0)
                {
                    using (var range = sheet.Cells[row, 1, row, 5])
                    {
                        range.StyleName = StyleZebraRow;
                    }
                }
                row++;
            }
            return row;
        }

        private int WriteAttributesSection(ExcelWorksheet sheet, int row, EntityMetadata entity)
        {
            sheet.Cells[row, 1].Value = "ATTRIBUTES & SCHEMA DEFINITION";
            using (var range = sheet.Cells[row, 1, row, 11])
            {
                range.Merge = true;
                range.StyleName = StyleSectionHeaderAttributes;
            }
            sheet.Row(row).Height = 22;
            row++;

            string[] headers = {
                "Logical Name",
                "Schema Name",
                "Display Name",
                "Type",
                "Description",
                "Required Level",
                "Is Custom",
                "Is Secured (FLS)",
                "Is Audit Enabled",
                "Source Type",
                "Additional Metadata & OptionSet Values"
            };

            for (int i = 0; i < headers.Length; i++) sheet.Cells[row, i + 1].Value = headers[i];

            using (var range = sheet.Cells[row, 1, row, headers.Length])
            {
                range.StyleName = StyleColumnHeaderAttributes;
            }
            row++;

            if (entity.Attributes == null || entity.Attributes.Length == 0)
            {
                sheet.Cells[row, 1].Value = "No attributes found for this table.";
                sheet.Cells[row, 1].Style.Font.Italic = true;
                return row + 1;
            }

            foreach (var attr in entity.Attributes.OrderBy(a => a.LogicalName))
            {
                sheet.Cells[row, 1].Value = attr.LogicalName;
                sheet.Cells[row, 2].Value = attr.SchemaName;
                sheet.Cells[row, 3].Value = GetLocalizedLabel(attr.DisplayName, string.Empty);
                sheet.Cells[row, 4].Value = GetAttributeTypeName(attr);
                sheet.Cells[row, 5].Value = GetLocalizedLabel(attr.Description, string.Empty);
                sheet.Cells[row, 6].Value = attr.RequiredLevel?.Value.ToString() ?? "None";
                sheet.Cells[row, 7].Value = attr.IsCustomAttribute.GetValueOrDefault(false) ? "Yes" : "No";
                sheet.Cells[row, 8].Value = attr.IsSecured.GetValueOrDefault(false) ? "Yes" : "No";
                sheet.Cells[row, 9].Value = attr.IsAuditEnabled?.Value.ToString() ?? "N/A";
                sheet.Cells[row, 10].Value = GetSourceType(attr);
                sheet.Cells[row, 11].Value = GetAdditionalMetadata(attr);

                if (row % 2 == 0)
                {
                    using (var range = sheet.Cells[row, 1, row, 11])
                    {
                        range.StyleName = StyleZebraRow;
                    }
                }
                row++;
            }

            return row;
        }

        private int WriteOneToManySection(ExcelWorksheet sheet, int row, EntityMetadata entity)
        {
            sheet.Cells[row, 1].Value = "1:N RELATIONSHIPS (ONE-TO-MANY)";
            using (var range = sheet.Cells[row, 1, row, 9])
            {
                range.Merge = true;
                range.StyleName = StyleSectionHeaderOneToMany;
            }
            sheet.Row(row).Height = 22;
            row++;

            string[] headers = {
                "Relationship Schema Name",
                "Related Table (Referencing)",
                "Foreign Key Attribute",
                "Cascade Delete",
                "Cascade Assign",
                "Cascade Share",
                "Cascade Unshare",
                "Cascade Reparent",
                "Rollup View"
            };

            for (int i = 0; i < headers.Length; i++) sheet.Cells[row, i + 1].Value = headers[i];
            using (var range = sheet.Cells[row, 1, row, headers.Length])
            {
                range.StyleName = StyleColumnHeaderOneToMany;
            }
            row++;

            if (entity.OneToManyRelationships == null || entity.OneToManyRelationships.Length == 0)
            {
                sheet.Cells[row, 1].Value = "No 1:N relationships found.";
                sheet.Cells[row, 1].Style.Font.Italic = true;
                return row + 1;
            }

            foreach (var rel in entity.OneToManyRelationships.OrderBy(r => r.SchemaName))
            {
                sheet.Cells[row, 1].Value = rel.SchemaName;
                // Clickable cross-reference to the related table's own sheet (1:N child table).
                // Written after the zebra block below (see there).
                sheet.Cells[row, 3].Value = rel.ReferencingAttribute;
                sheet.Cells[row, 4].Value = rel.CascadeConfiguration?.Delete?.ToString() ?? "N/A";
                sheet.Cells[row, 5].Value = rel.CascadeConfiguration?.Assign?.ToString() ?? "N/A";
                sheet.Cells[row, 6].Value = rel.CascadeConfiguration?.Share?.ToString() ?? "N/A";
                sheet.Cells[row, 7].Value = rel.CascadeConfiguration?.Unshare?.ToString() ?? "N/A";
                sheet.Cells[row, 8].Value = rel.CascadeConfiguration?.Reparent?.ToString() ?? "N/A";
                sheet.Cells[row, 9].Value = rel.CascadeConfiguration?.RollupView?.ToString() ?? "N/A";

                bool zebra1N = row % 2 == 0;
                if (zebra1N)
                {
                    using (var range = sheet.Cells[row, 1, row, 9])
                    {
                        range.StyleName = StyleZebraRow;
                    }
                }
                // Clickable cross-reference, styled last so the zebra fill above cannot erase it.
                WriteEntityReferenceCell(sheet, row, 2, rel.ReferencingEntity, zebra1N);
                row++;
            }
            return row;
        }

        private int WriteManyToOneSection(ExcelWorksheet sheet, int row, EntityMetadata entity)
        {
            sheet.Cells[row, 1].Value = "N:1 RELATIONSHIPS (MANY-TO-ONE / LOOKUPS)";
            using (var range = sheet.Cells[row, 1, row, 5])
            {
                range.Merge = true;
                range.StyleName = StyleSectionHeaderManyToOne;
            }
            sheet.Row(row).Height = 22;
            row++;

            string[] headers = {
                "Relationship Schema Name",
                "Lookup Attribute",
                "Referenced Table (Parent)",
                "Referenced Attribute (PK)",
                "Is Custom"
            };

            for (int i = 0; i < headers.Length; i++) sheet.Cells[row, i + 1].Value = headers[i];
            using (var range = sheet.Cells[row, 1, row, headers.Length])
            {
                range.StyleName = StyleColumnHeaderManyToOne;
            }
            row++;

            if (entity.ManyToOneRelationships == null || entity.ManyToOneRelationships.Length == 0)
            {
                sheet.Cells[row, 1].Value = "No N:1 relationships found.";
                sheet.Cells[row, 1].Style.Font.Italic = true;
                return row + 1;
            }

            foreach (var rel in entity.ManyToOneRelationships.OrderBy(r => r.SchemaName))
            {
                sheet.Cells[row, 1].Value = rel.SchemaName;
                sheet.Cells[row, 2].Value = rel.ReferencingAttribute;
                sheet.Cells[row, 4].Value = rel.ReferencedAttribute;
                sheet.Cells[row, 5].Value = rel.IsCustomRelationship.GetValueOrDefault(false) ? "Yes" : "No";

                bool zebraN1 = row % 2 == 0;
                if (zebraN1)
                {
                    using (var range = sheet.Cells[row, 1, row, 5])
                    {
                        range.StyleName = StyleZebraRow;
                    }
                }
                // Clickable cross-reference to the related table's own sheet (N:1 parent table),
                // styled last so the zebra fill above cannot erase its link appearance.
                WriteEntityReferenceCell(sheet, row, 3, rel.ReferencedEntity, zebraN1);
                row++;
            }
            return row;
        }

        private int WriteManyToManySection(ExcelWorksheet sheet, int row, EntityMetadata entity)
        {
            sheet.Cells[row, 1].Value = "N:N RELATIONSHIPS (MANY-TO-MANY)";
            using (var range = sheet.Cells[row, 1, row, 5])
            {
                range.Merge = true;
                range.StyleName = StyleSectionHeaderManyToMany;
            }
            sheet.Row(row).Height = 22;
            row++;

            string[] headers = {
                "Relationship Schema Name",
                "Intersect Table",
                "Related Table",
                "Entity 1 Attribute",
                "Entity 2 Attribute"
            };

            for (int i = 0; i < headers.Length; i++) sheet.Cells[row, i + 1].Value = headers[i];
            using (var range = sheet.Cells[row, 1, row, headers.Length])
            {
                range.StyleName = StyleColumnHeaderManyToMany;
            }
            row++;

            if (entity.ManyToManyRelationships == null || entity.ManyToManyRelationships.Length == 0)
            {
                sheet.Cells[row, 1].Value = "No N:N relationships found.";
                sheet.Cells[row, 1].Style.Font.Italic = true;
                return row + 1;
            }

            foreach (var rel in entity.ManyToManyRelationships.OrderBy(r => r.SchemaName))
            {
                string otherEntity = rel.Entity1LogicalName.Equals(entity.LogicalName, StringComparison.OrdinalIgnoreCase)
                    ? rel.Entity2LogicalName
                    : rel.Entity1LogicalName;

                sheet.Cells[row, 1].Value = rel.SchemaName;
                sheet.Cells[row, 2].Value = rel.IntersectEntityName;
                sheet.Cells[row, 4].Value = rel.Entity1IntersectAttribute;
                sheet.Cells[row, 5].Value = rel.Entity2IntersectAttribute;

                bool zebraNN = row % 2 == 0;
                if (zebraNN)
                {
                    using (var range = sheet.Cells[row, 1, row, 5])
                    {
                        range.StyleName = StyleZebraRow;
                    }
                }
                // Clickable cross-reference to the other side of the N:N relationship, styled last
                // so the zebra fill above cannot erase its link appearance.
                WriteEntityReferenceCell(sheet, row, 3, otherEntity, zebraNN);
                row++;
            }
            return row;
        }

        private string GetAttributeTypeName(AttributeMetadata attribute)
        {
            if (attribute == null) return "Unknown";
            string typeName = attribute.AttributeTypeName?.Value;
            if (!string.IsNullOrEmpty(typeName)) return typeName;
            return attribute.AttributeType?.ToString() ?? "Unknown";
        }

        private string GetSourceType(AttributeMetadata attribute)
        {
            if (attribute?.SourceType == null) return "N/A";
            switch (attribute.SourceType.Value)
            {
                case 0: return "Standard";
                case 1: return "Calculated";
                case 2: return "Rollup";
                default: return attribute.SourceType.Value.ToString(CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        /// Hard cap Excel enforces on the text length of a single cell. EPPlus writes a longer
        /// string without complaining, but Excel then refuses to open the workbook normally and
        /// shows "We found a problem with some content in '&lt;file&gt;.xlsx'", repairing the file and
        /// discarding formatting (which is what made the "Go to Sheet" links lose their link
        /// appearance). Reached in practice by the OptionSet column: every option of a picklist is
        /// concatenated into one cell, and large global option sets blow past 32767 characters.
        /// </summary>
        private const int MaxExcelCellTextLength = 32767;

        /// <summary>
        /// Truncates <paramref name="text"/> so it can never exceed <see cref="MaxExcelCellTextLength"/>,
        /// leaving a visible marker so the reader knows the value was cut rather than silently
        /// losing content.
        /// </summary>
        private static string ClampCellText(string text)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= MaxExcelCellTextLength)
            {
                return text;
            }

            const string marker = " […texto truncado: excede el límite de 32.767 caracteres por celda de Excel]";
            return text.Substring(0, MaxExcelCellTextLength - marker.Length) + marker;
        }

        private string GetAdditionalMetadata(AttributeMetadata attribute)
        {
            if (attribute == null) return string.Empty;

            var sb = new StringBuilder();

            if (attribute is StringAttributeMetadata str)
            {
                sb.Append($"Max Length: {str.MaxLength}; Format: {str.FormatName?.Value ?? str.Format?.ToString()}");
            }
            else if (attribute is MemoAttributeMetadata memo)
            {
                sb.Append($"Max Length: {memo.MaxLength}; Format: {memo.Format?.ToString()}");
            }
            else if (attribute is IntegerAttributeMetadata intAttr)
            {
                sb.Append($"Range: [{intAttr.MinValue} to {intAttr.MaxValue}]; Format: {intAttr.Format?.ToString()}");
            }
            else if (attribute is DecimalAttributeMetadata decAttr)
            {
                sb.Append($"Range: [{decAttr.MinValue} to {decAttr.MaxValue}]; Precision: {decAttr.Precision}");
            }
            else if (attribute is DoubleAttributeMetadata dblAttr)
            {
                sb.Append($"Range: [{dblAttr.MinValue} to {dblAttr.MaxValue}]; Precision: {dblAttr.Precision}");
            }
            else if (attribute is MoneyAttributeMetadata money)
            {
                sb.Append($"Range: [{money.MinValue} to {money.MaxValue}]; Precision: {money.Precision}");
            }
            else if (attribute is DateTimeAttributeMetadata dt)
            {
                sb.Append($"Format: {dt.Format?.ToString()}; DateTimeBehavior: {dt.DateTimeBehavior?.Value}");
            }
            else if (attribute is LookupAttributeMetadata lookup)
            {
                sb.Append($"Targets: {(lookup.Targets != null ? string.Join(", ", lookup.Targets) : "N/A")}");
            }
            else if (attribute is StateAttributeMetadata stateAttr)
            {
                if (stateAttr.OptionSet?.Options != null)
                {
                    var states = stateAttr.OptionSet.Options.Select(o => $"{o.Value}:{GetLocalizedLabel(o.Label, "State")}");
                    sb.Append("States: " + string.Join(" | ", states));
                }
            }
            else if (attribute is StatusAttributeMetadata statusAttr)
            {
                if (statusAttr.OptionSet?.Options != null)
                {
                    var statuses = statusAttr.OptionSet.Options.Select(o => {
                        var statusOpt = o as StatusOptionMetadata;
                        string stateInfo = statusOpt?.State != null ? $"(State:{statusOpt.State}) " : string.Empty;
                        return $"{stateInfo}{o.Value}:{GetLocalizedLabel(o.Label, "Status")}";
                    });
                    sb.Append("Status Reasons: " + string.Join(" | ", statuses));
                }
            }
            else if (attribute is EnumAttributeMetadata enumAttr)
            {
                if (enumAttr.OptionSet?.Options != null)
                {
                    var options = enumAttr.OptionSet.Options.Select(o => $"{o.Value}:{GetLocalizedLabel(o.Label, "Option")}");
                    sb.Append("Options: " + string.Join(" | ", options));
                }
            }
            else if (attribute is BooleanAttributeMetadata boolAttr)
            {
                string trueLabel = GetLocalizedLabel(boolAttr.OptionSet?.TrueOption?.Label, "True");
                string falseLabel = GetLocalizedLabel(boolAttr.OptionSet?.FalseOption?.Label, "False");
                sb.Append($"True: {trueLabel} ({boolAttr.OptionSet?.TrueOption?.Value}); False: {falseLabel} ({boolAttr.OptionSet?.FalseOption?.Value})");
            }

            // Clamped at the source: this is the value that in practice exceeds Excel's
            // per-cell text limit (large OptionSets concatenate every option into one string).
            return ClampCellText(sb.ToString());
        }

        private string GetLocalizedLabel(Label label, string fallback)
        {
            if (label == null) return fallback;
            return label.UserLocalizedLabel?.Label 
                ?? (label.LocalizedLabels != null && label.LocalizedLabels.Count > 0 ? label.LocalizedLabels[0].Label : fallback);
        }

        private string BuildSafeSheetName(string displayName, string logicalName)
        {
            string disp = SanitizeSheetName(displayName);
            string logic = SanitizeSheetName(logicalName);

            string candidate = $"{disp} ({logic})";
            if (candidate.Length > 31)
            {
                string tag = $"({logic})";
                int avail = 31 - tag.Length - 3;
                if (avail > 3)
                {
                    candidate = disp.Substring(0, Math.Min(disp.Length, avail)) + "..." + tag;
                }
                else if (logic.Length <= 31)
                {
                    candidate = logic;
                }
                else
                {
                    // Keeping only the FIRST 31 characters of a long logical name collapses whole
                    // families of tables onto the same name (e.g. every
                    // "wit_<something>_wit_<something>" N:N intersect table shares a long prefix),
                    // which then only differ by the "_1"/"_2" dedupe suffix and become impossible
                    // to tell apart. Keep both ends instead - the head identifies the family, the
                    // tail is what actually distinguishes one table from another.
                    candidate = logic.Substring(0, 20) + "~" + logic.Substring(logic.Length - 10);
                }
            }

            string safe = candidate;
            int counter = 1;
            while (usedSheetNames.Contains(safe))
            {
                string suffix = $"_{counter}";
                int maxLen = 31 - suffix.Length;
                safe = candidate.Length > maxLen ? candidate.Substring(0, maxLen) + suffix : candidate + suffix;
                counter++;
            }
            return safe;
        }

        private string SanitizeSheetName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Table";
            char[] invalidChars = { ':', '\\', '/', '?', '*', '[', ']' };
            string result = name;
            foreach (var c in invalidChars)
            {
                result = result.Replace(c, '_');
            }
            return result.Trim();
        }

        /// <summary>
        /// Writes a related table's logical name into a cell as a clickable link to that table's
        /// own sheet when it is documented in the SAME output file, and as plain text otherwise
        /// (the table was filtered out of this export, or landed in a different file when the
        /// export is split). This is what makes the 1:N / N:1 / N:N sections navigable instead of
        /// being a list of names the reader has to hunt for by hand.
        /// </summary>
        private void WriteEntityReferenceCell(ExcelWorksheet sheet, int row, int col, string relatedLogicalName, bool zebraRow)
        {
            if (string.IsNullOrEmpty(relatedLogicalName))
            {
                sheet.Cells[row, col].Value = "N/A";
                return;
            }

            string targetSheet;
            if (sheetNameByLogicalName.TryGetValue(relatedLogicalName, out targetSheet) &&
                !string.Equals(targetSheet, sheet.Name, StringComparison.OrdinalIgnoreCase))
            {
                sheet.Cells[row, col].Formula =
                    $"HYPERLINK(\"#'{EscapeSheetFormula(targetSheet)}'!A1\", \"{EscapeFormulaText(relatedLogicalName)}\")";
                // Named style applied AFTER the caller has already striped the row, so the zebra
                // fill cannot wipe the link's appearance (see PopulateIndexSheet for the same note).
                sheet.Cells[row, col].StyleName = zebraRow ? StyleLinkCellZebra : StyleLinkCell;
            }
            else
            {
                sheet.Cells[row, col].Value = relatedLogicalName;
            }
        }

        /// <summary>
        /// Escapes a literal string for use inside an Excel formula (double quotes are doubled).
        /// </summary>
        private static string EscapeFormulaText(string text)
        {
            return (text ?? string.Empty).Replace("\"", "\"\"");
        }

        private string EscapeSheetFormula(string sheetName)
        {
            return sheetName.Replace("'", "''");
        }

        /// <summary>
        /// Reports progress ONLY when the host actually enabled progress reporting on the worker.
        /// BackgroundWorker.ReportProgress throws InvalidOperationException when
        /// WorkerReportsProgress is false, and XrmToolBox only sets that flag when the
        /// WorkAsyncInfo supplies a ProgressChanged handler. Because this call sits at the very top
        /// of the per-table loop, an unguarded throw here made EVERY table fail on its first
        /// statement and get skipped by the per-table catch - producing an export with a correct
        /// Index header and zero documented tables. Progress reporting is cosmetic, so it must
        /// never be able to break the export.
        /// </summary>
        private void ReportProgress(int percentage, string message)
        {
            if (worker == null || !worker.WorkerReportsProgress)
            {
                return;
            }

            try
            {
                worker.ReportProgress(Math.Min(100, Math.Max(0, percentage)), message);
            }
            catch (InvalidOperationException)
            {
                // Host changed its mind about progress reporting mid-run - ignore, never abort a table.
            }
        }

        private void SavePackage(ExcelPackage package, string filePath)
        {
            string dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            FileInfo fi = new FileInfo(filePath);
            if (fi.Exists)
            {
                fi.Delete();
            }

            package.SaveAs(fi);
        }
    }
}
