using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using WinLabel = System.Windows.Forms.Label;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using OfficeOpenXml;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Args;
using XrmToolBox.Extensibility.Interfaces;
using MetadataDataverseDocument.Exporters;
using MetadataDataverseDocument.Models;

namespace MetadataDataverseDocument.UI
{
    public partial class MetadataDocumentControl : PluginControlBase, IAboutPlugin
    {
        private enum QuickFilterMode { All, Custom, Standard, Selected }

        private List<EntityMetadata> _allEntities = new List<EntityMetadata>();
        private readonly HashSet<string> _checkedTableLogicalNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private string _currentSolutionUniqueName = null;
        private QuickFilterMode _currentFilterMode = QuickFilterMode.All;
        private Settings _settings = new Settings();

        private Button _btnLoad;
        private Button _btnShowRelationships;
        private Button _btnExportDictionary;
        private Button _btnExportErd;
        private Button _btnExportRelationships;
        private Button _btnAbout;
        private CheckBox _chkIncludeSystemTables;

        // Cuarto filtro rapido ("Selected"): muestra SOLO las tablas ya marcadas. Con miles de
        // tablas cargadas era imposible revisar que 20 quedaron seleccionadas antes de exportar,
        // porque las marcadas quedaban dispersas en una lista de 2500 elementos.
        private Button _btnFilterSelected;

        // Cache de nombres normalizados (minusculas y sin acentos) por nombre logico, para que la
        // busqueda no tenga que normalizar 2500 cadenas en cada pulsacion de tecla. Se reconstruye
        // cuando cambia el conjunto de tablas cargadas.
        private readonly Dictionary<string, string[]> _searchIndex =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        private int _searchIndexBuiltFor = -1;

        // Export-options checkboxes (Round 4 - "Configurabilidad real"). Built once by
        // BuildExportOptionsPanel and reflected from _settings; read on the UI thread by
        // PromptExportDataDictionary before starting the background export.
        private Panel _exportOptionsPanel;
        private CheckBox _chkIncludeAlternateKeys;
        private CheckBox _chkIncludeAttributes;
        private CheckBox _chkIncludeOneToMany;
        private CheckBox _chkIncludeManyToOne;
        private CheckBox _chkIncludeManyToMany;
        private CheckBox _chkGenerateIndexSheet;
        private CheckBox _chkCombineIntoSingleFile;

        // El .designer asigna SplitterDistance = 300 en el constructor, cuando el control aun
        // no tiene su ancho definitivo. WinForms RECORTA ese valor al espacio disponible en ese
        // momento (el panel quedaba en ~120 px) y nunca lo recupera. Hay que aplicarlo cuando el
        // control ya fue dimensionado; este flag evita reaplicarlo despues y pisar el ajuste
        // manual del usuario.
        private bool _splitterAplicado;
        private const int DefaultLeftPanelWidth = 380;
        private const int MinLeftPanelWidth = 300;   // los 4 botones de filtro en una fila
        private const int MinRightPanelWidth = 320;

        // Shared Font for the header buttons, built once and reused on every call to
        // BuildModernHeaderButtons (which can run more than once - see constructor and
        // MetadataDocumentControl_Load) instead of allocating a new IDisposable Font per button
        // per call, which used to leak.
        private static readonly Font HeaderButtonFont = new Font("Segoe UI", 9F, FontStyle.Bold);

        public MetadataDocumentControl()
        {
            InitializeComponent();
            LoadSettings();
            BuildModernHeaderButtons();
            BuildExportOptionsPanel();
            BuildTableSearchEnhancements();
            ConectarPersistenciaDelSeparador();
            AplicarAnchoPanelIzquierdo();
            UpdateSelectionStats();
        }

        private void MetadataDocumentControl_Load(object sender, EventArgs e)
        {
            if (headerButtonsPanel.Controls.Count == 0)
            {
                BuildModernHeaderButtons();
            }
            if (_exportOptionsPanel == null)
            {
                BuildExportOptionsPanel();
            }
            if (_btnFilterSelected == null)
            {
                BuildTableSearchEnhancements();
            }
            AplicarAnchoPanelIzquierdo();
            UpdateSelectionStats();
        }

        /// <summary>
        /// Aplica el ancho guardado del panel izquierdo, una sola vez y solo cuando el
        /// SplitContainer ya tiene un ancho real. Llamado desde Load y desde SizeChanged porque
        /// XrmToolBox crea el control antes de darle su tamaño definitivo.
        /// </summary>
        private void AplicarAnchoPanelIzquierdo()
        {
            if (_splitterAplicado || splitContainerMain == null) return;

            int disponible = splitContainerMain.Width;
            if (disponible < MinLeftPanelWidth + MinRightPanelWidth + splitContainerMain.SplitterWidth)
            {
                return;   // todavia no hay espacio real: reintentar en el proximo SizeChanged
            }

            splitContainerMain.Panel1MinSize = MinLeftPanelWidth;
            splitContainerMain.Panel2MinSize = MinRightPanelWidth;

            int deseado = _settings.LeftPanelWidth > 0 ? _settings.LeftPanelWidth : DefaultLeftPanelWidth;
            int maximo = disponible - MinRightPanelWidth - splitContainerMain.SplitterWidth;
            int valor = Math.Max(MinLeftPanelWidth, Math.Min(deseado, maximo));

            try
            {
                splitContainerMain.SplitterDistance = valor;
                _splitterAplicado = true;
            }
            catch (InvalidOperationException)
            {
                // El contenedor cambio de tamaño mientras asignabamos: se reintenta luego.
            }
        }

        /// <summary>
        /// Guarda la posicion del separador cuando el usuario termina de arrastrarlo, para que
        /// su ajuste se conserve entre sesiones. SplitterMoved se dispara al soltar, no durante
        /// el arrastre, asi que no genera escrituras excesivas.
        /// </summary>
        private void ConectarPersistenciaDelSeparador()
        {
            if (splitContainerMain == null) return;

            splitContainerMain.SizeChanged += (s, e) => AplicarAnchoPanelIzquierdo();
            splitContainerMain.SplitterMoved += (s, e) =>
            {
                if (!_splitterAplicado) return;   // movimientos automaticos previos al ajuste
                _settings.LeftPanelWidth = splitContainerMain.SplitterDistance;
                SaveSettings();
            };
        }

        private void LoadSettings()
        {
            if (SettingsManager.Instance.TryLoad(GetType(), out Settings s))
            {
                _settings = s;
            }
            ApplyExportOptionCheckboxesFromSettings();
        }

        private void SaveSettings()
        {
            SettingsManager.Instance.Save(GetType(), _settings);
        }

        // Pushes the current _settings values into the export-options checkboxes, if they have
        // already been created. Called from LoadSettings() (constructor order: LoadSettings()
        // runs before BuildExportOptionsPanel(), so the checkboxes do not exist yet at that first
        // call - BuildExportOptionsPanel() applies the same values itself right after creating
        // them) and is safe to call again any time afterwards (e.g. if settings were reloaded).
        private void ApplyExportOptionCheckboxesFromSettings()
        {
            if (_chkIncludeAlternateKeys == null)
            {
                return;
            }
            _chkIncludeAlternateKeys.Checked = _settings.IncludeAlternateKeys;
            _chkIncludeAttributes.Checked = _settings.IncludeAttributes;
            _chkIncludeOneToMany.Checked = _settings.IncludeOneToMany;
            _chkIncludeManyToOne.Checked = _settings.IncludeManyToOne;
            _chkIncludeManyToMany.Checked = _settings.IncludeManyToMany;
            _chkGenerateIndexSheet.Checked = _settings.GenerateIndexSheet;
            if (_chkCombineIntoSingleFile != null)
            {
                _chkCombineIntoSingleFile.Checked = _settings.CombineIntoSingleFile;
            }
        }

        // Builds a dedicated row of checkboxes controlling ExportOptions for the Data Dictionary
        // export (Round 4 - "Configurabilidad real"). Created once, added to panelRight AFTER its
        // designer-declared children (tabControlTables, labelInfo) so that, per this project's
        // existing Dock=Top ordering (see labelInfo, which is added after tabControlTables and
        // ends up above it), this panel docks above labelInfo as the topmost strip of panelRight -
        // all without touching MetadataDocumentControl.designer.cs, matching the pattern used for
        // _chkIncludeSystemTables in Round 1.
        /// <summary>
        /// Añade el cuarto filtro rapido ("Selected") junto a All/Custom/Standard y aclara en la
        /// etiqueta del buscador que acepta VARIOS terminos. Se construye en tiempo de ejecucion,
        /// igual que el resto de los controles añadidos, para no tocar el archivo .designer.cs.
        /// </summary>
        private void BuildTableSearchEnhancements()
        {
            if (pnlQuickFilters == null) return;

            // Los cuatro filtros tienen que caber en UNA fila dentro de los 300 px del panel
            // izquierdo. pnlQuickFilters es un FlowLayoutPanel de altura fija (28 px): si el
            // ultimo boton no cabe, se envuelve a una segunda fila y queda recortado, invisible.
            // Se reajustan los tres del designer y se dimensiona el nuevo para que la suma
            // (44+62+70+72 = 248, mas ~24 de margenes) entre con holgura en los 288 px utiles.
            btnFilterAll.Width = 44;
            btnFilterCustom.Width = 62;
            btnFilterStandard.Width = 70;

            _btnFilterSelected = new Button
            {
                Text = "Selected",
                Width = 72,
                Height = 24,
                Font = FilterFontInactive
            };
            _btnFilterSelected.Click += (s, ev) => ApplyQuickFilter(QuickFilterMode.Selected);
            pnlQuickFilters.Controls.Add(_btnFilterSelected);

            // El panel izquierdo mide 300 px (SplitterDistance en el .designer), asi que la
            // etiqueta tiene que caber en ese ancho: un texto largo se recorta y ademas
            // desordena el layout. La explicacion del multi-termino va en el tooltip.
            if (labelSearchTable != null)
            {
                labelSearchTable.Text = "Buscar tabla(s):";
            }

            // Ayuda visible al pasar el mouse, para que el flujo "pegar lista -> Select All" se
            // descubra sin tener que leer documentacion.
            var tip = new ToolTip { AutoPopDelay = 15000, InitialDelay = 400, ReshowDelay = 200 };
            if (txtFilterTables != null)
            {
                tip.SetToolTip(txtFilterTables,
                    "Escriba o pegue varios nombres separados por espacio, coma, punto y coma o salto de linea." + Environment.NewLine +
                    "Se muestran todas las tablas que coincidan con CUALQUIER termino (busca en nombre visible, logico y de esquema, sin distinguir acentos)." + Environment.NewLine +
                    "Ejemplo: account contact wit_admision" + Environment.NewLine + Environment.NewLine +
                    "Luego use 'Select All' para marcar todo lo mostrado, y el filtro 'Selected' para revisar su seleccion.");
            }
            tip.SetToolTip(_btnFilterSelected, "Muestra solo las tablas ya marcadas, para revisar la seleccion antes de exportar.");
        }

        private void BuildExportOptionsPanel()
        {
            _exportOptionsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                BackColor = Color.FromArgb(241, 245, 249),
                Padding = new Padding(10, 6, 10, 4)
            };

            var label = new WinLabel
            {
                Text = "Export Data Dictionary options:",
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(51, 65, 85),
                Margin = new Padding(0, 4, 14, 0)
            };

            Func<string, CheckBox> makeChk = text => new CheckBox
            {
                Text = text,
                AutoSize = true,
                Checked = true,
                Margin = new Padding(0, 3, 14, 3)
            };

            _chkIncludeAlternateKeys = makeChk("Alternate keys");
            _chkIncludeAttributes = makeChk("Attributes");
            _chkIncludeOneToMany = makeChk("1:N relationships");
            _chkIncludeManyToOne = makeChk("N:1 relationships");
            _chkIncludeManyToMany = makeChk("N:N relationships");
            _chkGenerateIndexSheet = makeChk("Index sheet");
            // Default checked: a single consolidated .xlsx is the plugin's original, expected
            // behavior. When unchecked, the export falls back to the safer multi-file split
            // (MetadataExcelExporter.DefaultMaxSheetsPerFile tables per file) for organizations
            // large enough that one huge in-memory workbook risks running out of memory.
            _chkCombineIntoSingleFile = makeChk("Single file (uncheck if it crashes on very large orgs)");

            _exportOptionsPanel.Controls.Add(label);
            _exportOptionsPanel.Controls.Add(_chkIncludeAlternateKeys);
            _exportOptionsPanel.Controls.Add(_chkIncludeAttributes);
            _exportOptionsPanel.Controls.Add(_chkIncludeOneToMany);
            _exportOptionsPanel.Controls.Add(_chkIncludeManyToOne);
            _exportOptionsPanel.Controls.Add(_chkIncludeManyToMany);
            _exportOptionsPanel.Controls.Add(_chkGenerateIndexSheet);
            _exportOptionsPanel.Controls.Add(_chkCombineIntoSingleFile);

            panelRight.Controls.Add(_exportOptionsPanel);

            // The checkboxes above were just created with Checked = true regardless of what is
            // in _settings (which was already loaded by LoadSettings() earlier in the
            // constructor) - apply the real saved values now that they exist.
            ApplyExportOptionCheckboxesFromSettings();
        }

        private void BuildModernHeaderButtons()
        {
            headerButtonsPanel.Controls.Clear();

            Action<Button, Color, string, int> styleBtn = (btn, backColor, text, width) =>
            {
                btn.Text = text;
                btn.BackColor = backColor;
                btn.ForeColor = Color.White;
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 0;
                btn.Font = HeaderButtonFont;
                btn.Height = 32;
                btn.Width = width;
                btn.Cursor = Cursors.Hand;
                btn.Margin = new Padding(4, 0, 4, 0);
            };

            // 0. Include system tables checkbox (default: unchecked - only Advanced Find-valid /
            // custom entities are loaded when loading the whole organization). Must exist before
            // the user clicks "Load Tables", since its value is read at that point.
            _chkIncludeSystemTables = new CheckBox
            {
                Text = "Incluir tablas de sistema",
                AutoSize = true,
                Checked = false,
                ForeColor = Color.White,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(10, 9, 12, 0)
            };

            // 1. Load Tables Button
            _btnLoad = new Button();
            styleBtn(_btnLoad, Color.FromArgb(15, 118, 110), "Load Tables", 120);
            var cmsLoad = new ContextMenuStrip();
            cmsLoad.Items.Add("Default Solution (All Entities)", null, (s, ev) => {
                _currentSolutionUniqueName = null;
                ExecuteMethod(() => ExecuteLoadTables(null));
            });
            cmsLoad.Items.Add("Select Solution...", null, (s, ev) => PromptSelectSolution());
            _btnLoad.Click += (s, ev) => cmsLoad.Show(_btnLoad, new Point(0, _btnLoad.Height));

            // 2. Show Relationships Button
            _btnShowRelationships = new Button();
            styleBtn(_btnShowRelationships, Color.FromArgb(51, 65, 85), "Show Relationships", 150);
            _btnShowRelationships.Enabled = false;
            _btnShowRelationships.Click += (s, ev) => ExecuteMethod(ExecuteShowRelationships);

            // 3. Export Data Dictionary Button (Excel - Full Technical Document)
            _btnExportDictionary = new Button();
            styleBtn(_btnExportDictionary, Color.FromArgb(217, 119, 6), "Export Data Dictionary", 160);
            _btnExportDictionary.Click += (s, ev) => ExecuteMethod(PromptExportDataDictionary);

            // 4. Export ERD Button (Mermaid Markdown)
            _btnExportErd = new Button();
            styleBtn(_btnExportErd, Color.FromArgb(124, 58, 237), "Export ERD (Mermaid)", 150);
            _btnExportErd.Click += (s, ev) => ExecuteMethod(PromptExportErd);

            // 5. Export Relationships Button (Excel / HTML)
            _btnExportRelationships = new Button();
            styleBtn(_btnExportRelationships, Color.FromArgb(71, 85, 105), "Export Rel. Matrix", 140);
            var cmsExportRel = new ContextMenuStrip();
            cmsExportRel.Items.Add("Relationships to Excel (.xlsx)...", null, (s, ev) => PromptExportRelationshipsExcel());
            cmsExportRel.Items.Add("Relationships to HTML (.html)...", null, (s, ev) => PromptExportRelationshipsHtml());
            _btnExportRelationships.Click += (s, ev) => cmsExportRel.Show(_btnExportRelationships, new Point(0, _btnExportRelationships.Height));

            // 6. About Button
            _btnAbout = new Button();
            styleBtn(_btnAbout, Color.FromArgb(15, 23, 42), "About", 75);
            _btnAbout.Click += (s, ev) => ShowAboutDialog();

            headerButtonsPanel.Controls.Add(_chkIncludeSystemTables);
            headerButtonsPanel.Controls.Add(_btnLoad);
            headerButtonsPanel.Controls.Add(_btnShowRelationships);
            headerButtonsPanel.Controls.Add(_btnExportDictionary);
            headerButtonsPanel.Controls.Add(_btnExportErd);
            headerButtonsPanel.Controls.Add(_btnExportRelationships);
            headerButtonsPanel.Controls.Add(_btnAbout);
        }

        #region Solution Loading & Table Retrieval

        private void PromptSelectSolution()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading Dataverse solutions...",
                Work = (worker, args) =>
                {
                    var query = new QueryExpression("solution")
                    {
                        ColumnSet = new ColumnSet("uniquename", "friendlyname", "version", "publisherid", "isvisible"),
                        Criteria = new FilterExpression
                        {
                            Conditions = { new ConditionExpression("isvisible", ConditionOperator.Equal, true) }
                        },
                        Orders = { new OrderExpression("friendlyname", OrderType.Ascending) }
                    };
                    var linkPub = query.AddLink("publisher", "publisherid", "publisherid", JoinOperator.LeftOuter);
                    linkPub.Columns = new ColumnSet("friendlyname");
                    linkPub.EntityAlias = "pub";

                    args.Result = Service.RetrieveMultiple(query).Entities;
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        MessageBox.Show(this, $"Error loading solutions:\n{args.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var solutions = (DataCollection<Entity>)args.Result;
                    if (solutions == null || solutions.Count == 0)
                    {
                        MessageBox.Show(this, "No accessible solutions were found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    using (var picker = new SolutionPickerDialog(solutions))
                    {
                        if (picker.ShowDialog(this) == DialogResult.OK && picker.SelectedSolution != null)
                        {
                            _currentSolutionUniqueName = picker.SelectedSolution.UniqueName;
                            ExecuteMethod(() => ExecuteLoadTables(_currentSolutionUniqueName));
                        }
                    }
                }
            });
        }

        private void ExecuteLoadTables(string solutionUniqueName)
        {
            _checkedTableLogicalNames.Clear();
            tabControlTables.TabPages.Clear();
            _btnShowRelationships.Enabled = false;

            // Read the checkbox value now, on the UI thread, before the background Work delegate
            // runs - the checkbox must not be touched from the worker thread.
            bool includeSystemTables = _chkIncludeSystemTables != null && _chkIncludeSystemTables.Checked;

            string statusMsg = solutionUniqueName == null
                ? "Loading tables from organization..."
                : $"Loading tables from solution '{solutionUniqueName}'...";

            WorkAsync(new WorkAsyncInfo
            {
                Message = statusMsg,
                Work = (worker, args) =>
                {
                    var req = new RetrieveAllEntitiesRequest
                    {
                        EntityFilters = EntityFilters.Entity,
                        RetrieveAsIfPublished = true
                    };
                    var resp = (RetrieveAllEntitiesResponse)Service.Execute(req);
                    var entities = resp.EntityMetadata.ToList();

                    if (!string.IsNullOrEmpty(solutionUniqueName))
                    {
                        var compQuery = new QueryExpression("solutioncomponent")
                        {
                            ColumnSet = new ColumnSet("objectid"),
                            Criteria = new FilterExpression
                            {
                                Conditions = { new ConditionExpression("componenttype", ConditionOperator.Equal, 1) }
                            }
                        };
                        var solLink = compQuery.AddLink("solution", "solutionid", "solutionid");
                        solLink.LinkCriteria.AddCondition("uniquename", ConditionOperator.Equal, solutionUniqueName);

                        var entityIds = new HashSet<Guid>(Service.RetrieveMultiple(compQuery).Entities
                            .Select(e => e.GetAttributeValue<Guid>("objectid")));

                        entities = entities.Where(e => entityIds.Contains(e.MetadataId.GetValueOrDefault())).ToList();
                    }
                    else if (!includeSystemTables)
                    {
                        // Loading the whole organization ("Default Solution / All Entities") without
                        // a system-tables opt-in: keep only entities that make sense in a business
                        // data dictionary - Advanced Find-valid tables, or custom entities (custom
                        // tables are frequently not Advanced Find-enabled but are still relevant).
                        // This is the default (checkbox unchecked); checking "Incluir tablas de
                        // sistema" restores the previous behavior of loading every entity.
                        entities = entities.Where(e =>
                            e.IsValidForAdvancedFind.GetValueOrDefault(false) ||
                            e.IsCustomEntity.GetValueOrDefault(false)).ToList();
                    }

                    args.Result = entities;
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        MessageBox.Show(this, $"Error loading tables:\n{args.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    _allEntities = (List<EntityMetadata>)args.Result;
                    labelInfo.Text = $"Loaded {_allEntities.Count} tables from {(solutionUniqueName ?? "Default Solution")}.";
                    RefreshTableList();
                }
            });
        }

        #endregion

        #region Filtering and Selection

        /// <summary>
        /// Normaliza un texto para busqueda: minusculas y sin acentos, de modo que escribir
        /// "admision" encuentre "Admisión" y "accion" encuentre "Acción" (los nombres visibles del
        /// org estan en español y antes habia que escribir el acento exacto para encontrarlos).
        /// </summary>
        private static string NormalizeForSearch(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            string decomposed = text.Normalize(System.Text.NormalizationForm.FormD);
            var sb = new System.Text.StringBuilder(decomposed.Length);
            foreach (char c in decomposed)
            {
                if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c) !=
                    System.Globalization.UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(c);
                }
            }
            return sb.ToString().Normalize(System.Text.NormalizationForm.FormC).ToLowerInvariant();
        }

        /// <summary>
        /// (Re)construye el indice de busqueda si cambio el conjunto de tablas cargadas.
        /// </summary>
        private void EnsureSearchIndex()
        {
            if (_searchIndexBuiltFor == _allEntities.Count) return;

            _searchIndex.Clear();
            foreach (var e in _allEntities)
            {
                if (string.IsNullOrEmpty(e.LogicalName)) continue;
                _searchIndex[e.LogicalName] = new[]
                {
                    NormalizeForSearch(e.DisplayName?.UserLocalizedLabel?.Label ?? string.Empty),
                    NormalizeForSearch(e.LogicalName),
                    NormalizeForSearch(e.SchemaName ?? string.Empty)
                };
            }
            _searchIndexBuiltFor = _allEntities.Count;
        }

        /// <summary>
        /// Parte el texto de busqueda en varios terminos. Permite pegar una lista de tablas
        /// separada por espacios, comas, punto y coma, tabulaciones o saltos de linea y verlas
        /// todas a la vez (coincidencia OR) en lugar de buscarlas de una en una: con eso, armar un
        /// modelo acotado de 20 tablas es pegar la lista, pulsar "Select All" y exportar.
        /// </summary>
        private static string[] ParseSearchTerms(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return new string[0];

            var terms = new List<string>();
            foreach (var part in raw.Split(new[] { ' ', ',', ';', '\t', '\r', '\n', '|' },
                                           StringSplitOptions.RemoveEmptyEntries))
            {
                string norm = NormalizeForSearch(part.Trim());
                if (norm.Length > 0) terms.Add(norm);
            }
            return terms.ToArray();
        }

        private void RefreshTableList()
        {
            EnsureSearchIndex();

            checkedListBoxTables.BeginUpdate();
            checkedListBoxTables.Items.Clear();

            string[] terms = ParseSearchTerms(txtFilterTables.Text);

            var filtered = _allEntities.AsEnumerable();

            // Quick Filter mode
            if (_currentFilterMode == QuickFilterMode.Custom)
            {
                filtered = filtered.Where(e => e.IsCustomEntity.GetValueOrDefault(false));
            }
            else if (_currentFilterMode == QuickFilterMode.Standard)
            {
                filtered = filtered.Where(e => !e.IsCustomEntity.GetValueOrDefault(false));
            }
            else if (_currentFilterMode == QuickFilterMode.Selected)
            {
                // Solo lo ya marcado, para revisar la seleccion antes de exportar.
                filtered = filtered.Where(e => _checkedTableLogicalNames.Contains(e.LogicalName));
            }

            // Busqueda de texto: multi-termino, OR, sin acentos, contra nombre visible, logico y de esquema.
            if (terms.Length > 0)
            {
                filtered = filtered.Where(e =>
                {
                    string[] haystack;
                    if (!_searchIndex.TryGetValue(e.LogicalName ?? string.Empty, out haystack))
                    {
                        return false;
                    }
                    foreach (var term in terms)
                    {
                        for (int h = 0; h < haystack.Length; h++)
                        {
                            if (haystack[h].Contains(term)) return true;
                        }
                    }
                    return false;
                });
            }

            foreach (var entity in filtered.OrderBy(e => e.DisplayName?.UserLocalizedLabel?.Label ?? e.LogicalName))
            {
                string disp = entity.DisplayName?.UserLocalizedLabel?.Label ?? entity.LogicalName;
                var item = new TableListItem(entity.LogicalName, disp, entity.IsCustomEntity.GetValueOrDefault(false));
                bool isChecked = _checkedTableLogicalNames.Contains(entity.LogicalName);
                checkedListBoxTables.Items.Add(item, isChecked);
            }

            checkedListBoxTables.EndUpdate();
            UpdateSelectionStats();
        }

        private void txtFilterTables_TextChanged(object sender, EventArgs e)
        {
            RefreshTableList();
        }

        // Fuentes compartidas para el resaltado del filtro activo. Antes cada handler creaba tres
        // objetos Font nuevos (IDisposable) en cada clic, sin liberarlos nunca.
        private static readonly Font FilterFontActive = new Font("Segoe UI", 8F, FontStyle.Bold);
        private static readonly Font FilterFontInactive = new Font("Segoe UI", 8F, FontStyle.Regular);

        private void ApplyQuickFilter(QuickFilterMode mode)
        {
            _currentFilterMode = mode;
            btnFilterAll.Font = mode == QuickFilterMode.All ? FilterFontActive : FilterFontInactive;
            btnFilterCustom.Font = mode == QuickFilterMode.Custom ? FilterFontActive : FilterFontInactive;
            btnFilterStandard.Font = mode == QuickFilterMode.Standard ? FilterFontActive : FilterFontInactive;
            if (_btnFilterSelected != null)
            {
                _btnFilterSelected.Font = mode == QuickFilterMode.Selected ? FilterFontActive : FilterFontInactive;
            }
            RefreshTableList();
        }

        private void btnFilterAll_Click(object sender, EventArgs e)
        {
            ApplyQuickFilter(QuickFilterMode.All);
        }

        private void btnFilterCustom_Click(object sender, EventArgs e)
        {
            ApplyQuickFilter(QuickFilterMode.Custom);
        }

        private void btnFilterStandard_Click(object sender, EventArgs e)
        {
            ApplyQuickFilter(QuickFilterMode.Standard);
        }

        private bool _isUpdatingCheckboxes = false;

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            _isUpdatingCheckboxes = true;
            checkedListBoxTables.BeginUpdate();
            try
            {
                for (int i = 0; i < checkedListBoxTables.Items.Count; i++)
                {
                    checkedListBoxTables.SetItemChecked(i, true);
                    var item = checkedListBoxTables.Items[i] as TableListItem;
                    if (item != null) _checkedTableLogicalNames.Add(item.LogicalName);
                }
            }
            finally
            {
                checkedListBoxTables.EndUpdate();
                _isUpdatingCheckboxes = false;
            }
            UpdateSelectionStats();
        }

        private void btnDeselectAll_Click(object sender, EventArgs e)
        {
            _isUpdatingCheckboxes = true;
            checkedListBoxTables.BeginUpdate();
            try
            {
                for (int i = 0; i < checkedListBoxTables.Items.Count; i++)
                {
                    checkedListBoxTables.SetItemChecked(i, false);
                }
                _checkedTableLogicalNames.Clear();
            }
            finally
            {
                checkedListBoxTables.EndUpdate();
                _isUpdatingCheckboxes = false;
            }
            UpdateSelectionStats();
        }

        private void checkedListBoxTables_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (_isUpdatingCheckboxes) return;

            var item = checkedListBoxTables.Items[e.Index] as TableListItem;
            if (item != null)
            {
                if (e.NewValue == CheckState.Checked)
                {
                    _checkedTableLogicalNames.Add(item.LogicalName);
                }
                else
                {
                    _checkedTableLogicalNames.Remove(item.LogicalName);
                }
            }

            BeginInvoke(new Action(UpdateSelectionStats));
        }

        private void UpdateSelectionStats()
        {
            if (lblSelectionStats != null)
            {
                // Muestra las tres cifras que importan al armar un modelo acotado: cuantas hay
                // marcadas en total (aunque no esten visibles por el filtro), cuantas se estan
                // mostrando y cuantas hay cargadas. Al pegar una lista de 20 nombres, comparar
                // "mostradas" con lo pegado delata de inmediato los nombres que no coincidieron.
                lblSelectionStats.Text =
                    $"{_checkedTableLogicalNames.Count} seleccionada(s) | {checkedListBoxTables.Items.Count} mostrada(s) de {_allEntities.Count} cargada(s)";
            }
            bool hasSelection = _checkedTableLogicalNames.Count > 0;
            if (_btnShowRelationships != null)
            {
                _btnShowRelationships.Enabled = hasSelection;
            }
            if (_btnExportDictionary != null)
            {
                _btnExportDictionary.Enabled = hasSelection;
            }
            if (_btnExportErd != null)
            {
                _btnExportErd.Enabled = hasSelection;
            }
            if (_btnExportRelationships != null)
            {
                _btnExportRelationships.Enabled = hasSelection;
            }
        }

        private List<EntityMetadata> GetSelectedEntities()
        {
            return _allEntities.Where(e => _checkedTableLogicalNames.Contains(e.LogicalName)).ToList();
        }

        #endregion

        #region Interactive UI Relationships Viewer

        // Round 5 (post-implementation fix): CreateInteractiveTableTab builds ~8 live WinForms
        // controls per table (an outer TabPage + a nested TabControl + 3 child TabPages + 3
        // DataGridViews), all created synchronously on the UI thread with no upper bound. A large
        // selection (e.g. "Select All" after loading an org with thousands of tables - the actual
        // crash report that prompted this fix involved 2735 tables) can freeze the UI for a long
        // time while building tens of thousands of controls, and risks exhausting the per-process
        // USER/GDI object limits (10,000 each by default on Windows) - a well-known cause of
        // WinForms crashes/instability. Unlike the Excel/HTML exporters (Export Data Dictionary,
        // Export Rel. Matrix), which write to a file and scale to however many tables are selected,
        // this view keeps every table's controls alive in the UI at once, so it cannot scale the
        // same way. This is deliberately a hard block, not a confirmation prompt like P5's "export
        // all tables" dialog - there is no reasonable way for this interactive view to be usable
        // with thousands of simultaneous tabs, so letting the user proceed anyway would just move
        // the freeze/crash a few seconds later instead of preventing it.
        private const int MaxInteractiveRelationshipTables = 50;

        private void ExecuteShowRelationships()
        {
            var selected = GetSelectedEntities();
            if (selected.Count == 0)
            {
                MessageBox.Show(this, "Please check at least one table from the list.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (selected.Count > MaxInteractiveRelationshipTables)
            {
                MessageBox.Show(
                    this,
                    $"You have {selected.Count} tables checked.\n\n" +
                    $"\"Show Relationships\" opens an interactive tab per table and is meant for exploring a manageable number of tables at a time (up to {MaxInteractiveRelationshipTables}). " +
                    "With very large selections it can freeze the UI for a long time or crash XrmToolBox while building that many controls at once.\n\n" +
                    "Please reduce your selection and try again. To document many tables at once instead, use \"Export Data Dictionary\" or \"Export Rel. Matrix\", which write to a file rather than opening live tabs.",
                    "Too Many Tables Selected",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            tabControlTables.TabPages.Clear();

            WorkAsync(new WorkAsyncInfo
            {
                Message = $"Retrieving relationships for {selected.Count} tables...",
                // Required for the worker.ReportProgress(...) call below: XrmToolBox only enables
                // BackgroundWorker.WorkerReportsProgress when a ProgressChanged handler is
                // supplied, and calling ReportProgress without it throws InvalidOperationException
                // on the very first table.
                ProgressChanged = (args) =>
                {
                    SetWorkingMessage(args.UserState as string ?? $"{args.ProgressPercentage}%");
                },
                Work = (worker, args) =>
                {
                    var detailed = new List<EntityMetadata>();
                    for (int i = 0; i < selected.Count; i++)
                    {
                        var e = selected[i];
                        worker.ReportProgress((int)(((i + 1) / (double)selected.Count) * 100), $"Loading relationships {i + 1} of {selected.Count}: {e.LogicalName}...");
                        var req = new RetrieveEntityRequest
                        {
                            LogicalName = e.LogicalName,
                            EntityFilters = EntityFilters.Relationships | EntityFilters.Entity,
                            RetrieveAsIfPublished = true
                        };
                        var resp = (RetrieveEntityResponse)Service.Execute(req);
                        detailed.Add(resp.EntityMetadata);
                    }
                    args.Result = detailed;
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        MessageBox.Show(this, $"Error loading relationships:\n{args.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var list = (List<EntityMetadata>)args.Result;
                    foreach (var ent in list)
                    {
                        CreateInteractiveTableTab(ent);
                    }

                    labelInfo.Text = $"Displaying relationships for {list.Count} tables.";
                }
            });
        }

        private void CreateInteractiveTableTab(EntityMetadata entity)
        {
            string disp = entity.DisplayName?.UserLocalizedLabel?.Label ?? entity.LogicalName;
            var tab = new TabPage($"{disp} ({entity.LogicalName})");

            var subTabs = new TabControl { Dock = DockStyle.Fill };

            // 1:N
            var tab1N = new TabPage("1:N Relationships");
            var dgv1N = CreateStyledGrid();
            Populate1NGrid(entity, dgv1N, tab1N);
            tab1N.Controls.Add(dgv1N);

            // N:1
            var tabN1 = new TabPage("N:1 Relationships (Lookups)");
            var dgvN1 = CreateStyledGrid();
            PopulateN1Grid(entity, dgvN1, tabN1);
            tabN1.Controls.Add(dgvN1);

            // N:N
            var tabNN = new TabPage("N:N Relationships");
            var dgvNN = CreateStyledGrid();
            PopulateNNGrid(entity, dgvNN, tabNN);
            tabNN.Controls.Add(dgvNN);

            subTabs.TabPages.Add(tab1N);
            subTabs.TabPages.Add(tabN1);
            subTabs.TabPages.Add(tabNN);

            tab.Controls.Add(subTabs);
            tabControlTables.TabPages.Add(tab);
        }

        private DataGridView CreateStyledGrid()
        {
            return new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(248, 250, 252) },
                EnableHeadersVisualStyles = false,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                    BackColor = Color.FromArgb(30, 41, 59),
                    ForeColor = Color.White
                },
                Font = new Font("Segoe UI", 9F)
            };
        }

        private void Populate1NGrid(EntityMetadata entity, DataGridView dgv, TabPage tab)
        {
            var dt = new DataTable();
            dt.Columns.Add("Relationship Schema Name");
            dt.Columns.Add("Related Table (Referencing)");
            dt.Columns.Add("Foreign Key Attribute");
            dt.Columns.Add("Cascade Delete");
            dt.Columns.Add("Cascade Assign");
            dt.Columns.Add("Cascade Share");
            dt.Columns.Add("Cascade Unshare");
            dt.Columns.Add("Cascade Reparent");

            if (entity.OneToManyRelationships != null)
            {
                foreach (var r in entity.OneToManyRelationships.OrderBy(x => x.SchemaName))
                {
                    dt.Rows.Add(
                        r.SchemaName,
                        r.ReferencingEntity,
                        r.ReferencingAttribute,
                        r.CascadeConfiguration?.Delete?.ToString(),
                        r.CascadeConfiguration?.Assign?.ToString(),
                        r.CascadeConfiguration?.Share?.ToString(),
                        r.CascadeConfiguration?.Unshare?.ToString(),
                        r.CascadeConfiguration?.Reparent?.ToString());
                }
            }

            dgv.DataSource = dt;
            tab.Text = $"1:N Relationships ({dt.Rows.Count})";
        }

        private void PopulateN1Grid(EntityMetadata entity, DataGridView dgv, TabPage tab)
        {
            var dt = new DataTable();
            dt.Columns.Add("Relationship Schema Name");
            dt.Columns.Add("Referenced Table (Parent)");
            dt.Columns.Add("Lookup Attribute");
            dt.Columns.Add("Referenced Attribute (PK)");
            dt.Columns.Add("Is Custom");

            if (entity.ManyToOneRelationships != null)
            {
                foreach (var r in entity.ManyToOneRelationships.OrderBy(x => x.SchemaName))
                {
                    dt.Rows.Add(
                        r.SchemaName,
                        r.ReferencedEntity,
                        r.ReferencingAttribute,
                        r.ReferencedAttribute,
                        r.IsCustomRelationship.GetValueOrDefault(false) ? "Yes" : "No");
                }
            }

            dgv.DataSource = dt;
            tab.Text = $"N:1 Relationships ({dt.Rows.Count})";
        }

        private void PopulateNNGrid(EntityMetadata entity, DataGridView dgv, TabPage tab)
        {
            var dt = new DataTable();
            dt.Columns.Add("Relationship Schema Name");
            dt.Columns.Add("Intersect Table");
            dt.Columns.Add("Associated Table");
            dt.Columns.Add("Entity 1 Attribute");
            dt.Columns.Add("Entity 2 Attribute");

            if (entity.ManyToManyRelationships != null)
            {
                foreach (var r in entity.ManyToManyRelationships.OrderBy(x => x.SchemaName))
                {
                    string other = r.Entity1LogicalName.Equals(entity.LogicalName, StringComparison.OrdinalIgnoreCase)
                        ? r.Entity2LogicalName
                        : r.Entity1LogicalName;

                    dt.Rows.Add(
                        r.SchemaName,
                        r.IntersectEntityName,
                        other,
                        r.Entity1IntersectAttribute,
                        r.Entity2IntersectAttribute);
                }
            }

            dgv.DataSource = dt;
            tab.Text = $"N:N Relationships ({dt.Rows.Count})";
        }

        #endregion

        #region Export Actions

        private void PromptExportDataDictionary()
        {
            var selected = GetSelectedEntities();
            if (selected.Count == 0)
            {
                // P5: previously this asked a generic Yes/No question with no indication of how
                // many tables that actually means or how long it might take - for an org-wide load
                // (potentially hundreds of tables) that made it too easy to kick off a very long
                // export by accident. Show the real count plus a rough, honestly-labeled time
                // estimate (and whether the output will be split into multiple files) before asking
                // for confirmation.
                int totalCount = _allEntities.Count;
                if (totalCount == 0)
                {
                    MessageBox.Show(this, "No tables are loaded. Load tables first, then check the ones you want to export.", "Export All Tables", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                bool combineIntoSingleFilePreview = _chkCombineIntoSingleFile == null || _chkCombineIntoSingleFile.Checked;
                string prompt = $"No tables are currently checked.\n\n" +
                    $"Do you want to export ALL {totalCount} loaded table(s)?\n\n" +
                    $"{EstimateExportDuration(totalCount)}" +
                    BuildSplitFileNote(totalCount, combineIntoSingleFilePreview);

                var answer = MessageBox.Show(this, prompt, "Export All Tables", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (answer == DialogResult.Yes) selected = _allEntities;
                else return;
            }

            using (var sfd = new SaveFileDialog())
            {
                sfd.Title = "Save Metadata Dataverse Technical Dictionary";
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                sfd.FileName = $"Dataverse_Metadata_Dictionary_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";

                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    string path = sfd.FileName;

                    // Round 4 - "Configurabilidad real": read the export-options checkboxes on the
                    // UI thread (same pattern as _chkIncludeSystemTables in Round 1 - these controls
                    // must not be touched from the background worker) and build a real ExportOptions
                    // from them instead of the previous hard-coded `new ExportOptions()`. Persist the
                    // choice for next time via _settings/SaveSettings(), the same way LoadSettings()
                    // already restores it on startup.
                    var exportOptions = new ExportOptions
                    {
                        IncludeAlternateKeys = _chkIncludeAlternateKeys == null || _chkIncludeAlternateKeys.Checked,
                        IncludeAttributes = _chkIncludeAttributes == null || _chkIncludeAttributes.Checked,
                        IncludeOneToMany = _chkIncludeOneToMany == null || _chkIncludeOneToMany.Checked,
                        IncludeManyToOne = _chkIncludeManyToOne == null || _chkIncludeManyToOne.Checked,
                        IncludeManyToMany = _chkIncludeManyToMany == null || _chkIncludeManyToMany.Checked,
                        GenerateIndexSheet = _chkGenerateIndexSheet == null || _chkGenerateIndexSheet.Checked
                    };

                    bool combineIntoSingleFile = _chkCombineIntoSingleFile == null || _chkCombineIntoSingleFile.Checked;
                    int maxSheetsPerFile = combineIntoSingleFile
                        ? int.MaxValue
                        : MetadataExcelExporter.DefaultMaxSheetsPerFile;

                    _settings.IncludeAlternateKeys = exportOptions.IncludeAlternateKeys;
                    _settings.IncludeAttributes = exportOptions.IncludeAttributes;
                    _settings.IncludeOneToMany = exportOptions.IncludeOneToMany;
                    _settings.IncludeManyToOne = exportOptions.IncludeManyToOne;
                    _settings.IncludeManyToMany = exportOptions.IncludeManyToMany;
                    _settings.GenerateIndexSheet = exportOptions.GenerateIndexSheet;
                    _settings.CombineIntoSingleFile = combineIntoSingleFile;
                    SaveSettings();

                    WorkAsync(new WorkAsyncInfo
                    {
                        Message = $"Generating Data Dictionary Excel for {selected.Count} tables...",
                        // Round 4 - "Cancelación real": lets XrmToolBox show a Cancel affordance on
                        // the progress dialog. BackgroundWorker.CancellationPending is checked both
                        // inside MetadataExcelExporter.Export() (to stop early and still keep
                        // whatever was already generated) and again right below, right after Export()
                        // returns, so the completion callback can tell a genuine cancellation apart
                        // from a normal finish.
                        IsCancelable = true,
                        // Supplying this handler is what makes XrmToolBox set
                        // BackgroundWorker.WorkerReportsProgress = true. Without it, the exporter's
                        // per-table worker.ReportProgress(...) call throws InvalidOperationException
                        // ("This BackgroundWorker states that it does not report progress"), and
                        // since that call is the first statement inside the per-table try block,
                        // every single table was skipped and the workbook came out with only an
                        // Index header and no tables at all.
                        ProgressChanged = (args) =>
                        {
                            SetWorkingMessage(args.UserState as string ?? $"{args.ProgressPercentage}%");
                        },
                        Work = (worker, args) =>
                        {
                            // P4: MetadataExcelExporter does not inherit PluginControlBase, so it
                            // cannot call LogInfo/LogWarning/LogError itself - hand it delegates
                            // pointed at this control's own (inherited) logging methods instead.
                            var exporter = new MetadataExcelExporter(
                                Service,
                                worker,
                                logInfo: msg => LogInfo(msg),
                                logWarning: msg => LogWarning(msg),
                                logError: msg => LogError(msg));
                            // Export() now returns every file it actually wrote: normally just
                            // `path`, but it is split into path_1.xlsx, path_2.xlsx, etc. when the
                            // number of tables exceeds MetadataExcelExporter.DefaultMaxSheetsPerFile
                            // (see P3(c) in BRIEFING.md) - splitting keeps each individual workbook's
                            // memory footprint bounded for very large exports. Export() itself checks
                            // worker.CancellationPending in its main per-table loop and returns early
                            // (with whatever files it had already saved) instead of throwing, so we
                            // still need to flag the cancellation back to WorkAsync here.
                            args.Result = exporter.Export(path, selected, exportOptions, maxSheetsPerFile);
                            if (worker.CancellationPending)
                            {
                                args.Cancel = true;
                            }
                        },
                        PostWorkCallBack = (args) =>
                        {
                            // Must be checked before touching args.Result: RunWorkerCompletedEventArgs
                            // throws InvalidOperationException from its Result getter when Cancelled is
                            // true, and a cancellation is not an error, so it gets its own message
                            // distinct from both the error path below and the success path.
                            if (args.Cancelled)
                            {
                                MessageBox.Show(this, "Export cancelled by user. Any file(s) already completed before the cancellation were kept - check the log for details.", "Export Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }

                            if (args.Error != null)
                            {
                                MessageBox.Show(this, $"Error generating document:\n{args.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            var generatedFiles = args.Result as List<string> ?? new List<string>();
                            string message = generatedFiles.Count <= 1
                                ? "Metadata Dataverse Document generated successfully!\n\nDo you want to open the file?"
                                : $"Metadata Dataverse Document generated successfully in {generatedFiles.Count} files " +
                                  "(the export was split because it exceeded the per-file table limit).\n\n" +
                                  "Do you want to open the first file?";

                            var open = MessageBox.Show(this, message, "Export Complete", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                            if (open == DialogResult.Yes && generatedFiles.Count > 0)
                            {
                                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(generatedFiles[0]) { UseShellExecute = true });
                            }
                        }
                    });
                }
            }
        }

        /// <summary>
        /// P5: rough, honestly-worded time estimate shown before exporting every loaded table
        /// without a per-table selection. This is a ballpark heuristic (a Dataverse metadata
        /// round-trip plus writing a sheet per table), not a measured prediction - it exists only to
        /// give the user a realistic sense of scale ("a few minutes" vs. "a couple hours") before
        /// they commit to a potentially very large export.
        /// </summary>
        private static string EstimateExportDuration(int tableCount)
        {
            const double secondsPerTable = 0.75;
            double totalSeconds = tableCount * secondsPerTable;

            if (totalSeconds < 45)
            {
                return "This should take well under a minute.";
            }
            if (totalSeconds < 90)
            {
                return "This may take about a minute.";
            }

            int minutes = (int)Math.Ceiling(totalSeconds / 60.0);
            return $"This may take approximately {minutes} minutes - please be patient and avoid closing XrmToolBox while it runs.";
        }

        /// <summary>
        /// P5: mentions the P3(c) automatic file-splitting behavior when it will actually kick in
        /// for this many tables, so the user isn't surprised to find several .xlsx files afterwards -
        /// or, when "Single file" is checked, warns that a very large single workbook carries more
        /// memory risk than the split alternative, since this plugin cannot predict how much memory
        /// is actually available on the user's machine.
        /// </summary>
        private static string BuildSplitFileNote(int tableCount, bool combineIntoSingleFile)
        {
            if (!combineIntoSingleFile)
            {
                if (tableCount <= MetadataExcelExporter.DefaultMaxSheetsPerFile)
                {
                    return string.Empty;
                }

                int estimatedFiles = (int)Math.Ceiling(tableCount / (double)MetadataExcelExporter.DefaultMaxSheetsPerFile);
                return $"\n\nBecause this exceeds {MetadataExcelExporter.DefaultMaxSheetsPerFile} tables, the output will automatically be split into approximately {estimatedFiles} .xlsx files.";
            }

            const int SingleFileWarningThreshold = 500;
            if (tableCount > SingleFileWarningThreshold)
            {
                return "\n\nNote: \"Single file\" is checked, so all of these tables will go into one .xlsx workbook. " +
                    "For very large organizations this uses more memory than the multi-file option; if the export " +
                    "crashes, uncheck \"Single file\" and try again to get several smaller files instead.";
            }

            return string.Empty;
        }

        private void PromptExportErd()
        {
            var selected = GetSelectedEntities();
            if (selected.Count == 0)
            {
                var answer = MessageBox.Show(this, "No tables are checked. Do you want to export ERD for ALL loaded tables?", "Export ERD", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (answer == DialogResult.Yes) selected = _allEntities;
                else return;
            }

            using (var sfd = new SaveFileDialog())
            {
                sfd.Title = "Save Dataverse ERD (Mermaid Diagram)";
                sfd.Filter = "Markdown File (*.md)|*.md|Mermaid File (*.mmd)|*.mmd";
                sfd.FileName = $"Dataverse_ERD_{DateTime.Now:yyyyMMdd_HHmm}.md";

                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    string path = sfd.FileName;
                    WorkAsync(new WorkAsyncInfo
                    {
                        Message = "Retrieving metadata for ERD generation...",
                        // Required for the worker.ReportProgress(...) call below - see the same
                        // note on the relationships and data dictionary WorkAsyncInfo blocks.
                        ProgressChanged = (args) =>
                        {
                            SetWorkingMessage(args.UserState as string ?? $"{args.ProgressPercentage}%");
                        },
                        Work = (worker, args) =>
                        {
                            var fullEntities = new List<EntityMetadata>();
                            for (int i = 0; i < selected.Count; i++)
                            {
                                var e = selected[i];
                                worker.ReportProgress((int)(((i + 1) / (double)selected.Count) * 100), $"Loading ERD data {i + 1} of {selected.Count}: {e.LogicalName}...");
                                var req = new RetrieveEntityRequest
                                {
                                    LogicalName = e.LogicalName,
                                    EntityFilters = EntityFilters.Relationships | EntityFilters.Entity,
                                    RetrieveAsIfPublished = true
                                };
                                var resp = (RetrieveEntityResponse)Service.Execute(req);
                                fullEntities.Add(resp.EntityMetadata);
                            }

                            MermaidErdExporter.ExportToMarkdownFile(path, fullEntities);
                        },
                        PostWorkCallBack = (args) =>
                        {
                            if (args.Error != null)
                            {
                                MessageBox.Show(this, $"Error generating ERD:\n{args.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            var open = MessageBox.Show(this, "ERD Diagram exported successfully!\n\nDo you want to open the file?", "Export Complete", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                            if (open == DialogResult.Yes)
                            {
                                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
                            }
                        }
                    });
                }
            }
        }

        private void PromptExportRelationshipsExcel()
        {
            var selected = GetSelectedEntities();
            if (selected.Count == 0) selected = _allEntities;

            using (var sfd = new SaveFileDialog())
            {
                sfd.Title = "Save Relationships Matrix";
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                sfd.FileName = $"Dataverse_Relationships_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";

                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    string path = sfd.FileName;
                    WorkAsync(new WorkAsyncInfo
                    {
                        Message = $"Exporting relationships matrix for {selected.Count} tables...",
                        // Enables BackgroundWorker.WorkerReportsProgress so the ReportProgress
                        // calls below actually work instead of throwing.
                        ProgressChanged = (args) =>
                        {
                            SetWorkingMessage(args.UserState as string ?? $"{args.ProgressPercentage}%");
                        },
                        Work = (worker, args) =>
                        {
                            var fullEntities = new List<EntityMetadata>();
                            for (int i = 0; i < selected.Count; i++)
                            {
                                worker.ReportProgress((int)(((i + 1) / (double)selected.Count) * 100), $"Loading relationships {i + 1} of {selected.Count}: {selected[i].LogicalName}...");
                                var req = new RetrieveEntityRequest
                                {
                                    LogicalName = selected[i].LogicalName,
                                    EntityFilters = EntityFilters.Relationships | EntityFilters.Entity,
                                    RetrieveAsIfPublished = true
                                };
                                fullEntities.Add(((RetrieveEntityResponse)Service.Execute(req)).EntityMetadata);
                            }

                            using (var pkg = new ExcelPackage())
                            {
                                // Worksheet names must be unique within a workbook and are capped
                                // at 31 characters: truncating long logical names collides for
                                // whole families of tables, and ExcelWorksheets.Add THROWS on a
                                // duplicate name, which would abort the entire export.
                                var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                                int written = 0;
                                foreach (var ent in fullEntities)
                                {
                                    worker.ReportProgress((int)((++written / (double)fullEntities.Count) * 100), $"Writing sheet {written} of {fullEntities.Count}: {ent.LogicalName}...");
                                    string sheetName = BuildUniqueMatrixSheetName(ent.LogicalName, usedNames);
                                    var ws = pkg.Workbook.Worksheets.Add(sheetName);
                                    ws.Cells[1, 1].Value = "Relationship";
                                    ws.Cells[1, 2].Value = "Target / Parent";
                                    ws.Cells[1, 3].Value = "Foreign Key";
                                    int row = 2;
                                    if (ent.OneToManyRelationships != null)
                                    {
                                        foreach (var r in ent.OneToManyRelationships)
                                        {
                                            ws.Cells[row, 1].Value = r.SchemaName;
                                            ws.Cells[row, 2].Value = r.ReferencingEntity;
                                            ws.Cells[row, 3].Value = r.ReferencingAttribute;
                                            row++;
                                        }
                                    }
                                    // ws.Dimension is null when the sheet ended up with no rows at
                                    // all, and AutoFitColumns measures every cell (too costly for a
                                    // very large export), so both are guarded here.
                                    if (ws.Dimension != null && fullEntities.Count <= 50)
                                    {
                                        ws.Cells[ws.Dimension.Address].AutoFitColumns();
                                    }
                                }
                                pkg.SaveAs(new FileInfo(path));
                            }
                        },
                        PostWorkCallBack = (args) =>
                        {
                            if (args.Error != null)
                            {
                                MessageBox.Show(this, $"Error exporting relationships:\n{args.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            MessageBox.Show(this, "Relationships exported successfully!", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    });
                }
            }
        }

        /// <summary>
        /// Builds a worksheet name for the relationships-matrix export that is at most 31
        /// characters (Excel's limit) AND unique within the workbook. Truncating a long logical
        /// name to its first 31 characters collides for whole families of tables (the org has many
        /// names sharing a long prefix), and ExcelWorksheets.Add throws "A worksheet with this
        /// name already exists" on a duplicate, which aborted the whole export. Long names keep
        /// both ends, since the tail is what distinguishes one table from another.
        /// </summary>
        private static string BuildUniqueMatrixSheetName(string logicalName, HashSet<string> usedNames)
        {
            string baseName = string.IsNullOrEmpty(logicalName) ? "Table" : logicalName;
            foreach (var c in new[] { ':', '\\', '/', '?', '*', '[', ']' })
            {
                baseName = baseName.Replace(c, '_');
            }
            if (baseName.Length > 31)
            {
                baseName = baseName.Substring(0, 20) + "~" + baseName.Substring(baseName.Length - 10);
            }

            string candidate = baseName;
            int counter = 1;
            while (usedNames.Contains(candidate))
            {
                string suffix = $"_{counter++}";
                int maxLen = 31 - suffix.Length;
                candidate = (baseName.Length > maxLen ? baseName.Substring(0, maxLen) : baseName) + suffix;
            }
            usedNames.Add(candidate);
            return candidate;
        }

        private void PromptExportRelationshipsHtml()
        {
            var selected = GetSelectedEntities();
            if (selected.Count == 0) selected = _allEntities;

            using (var sfd = new SaveFileDialog())
            {
                sfd.Title = "Save Relationships HTML Report";
                sfd.Filter = "HTML Document (*.html)|*.html";
                sfd.FileName = $"Dataverse_Relationships_{DateTime.Now:yyyyMMdd_HHmm}.html";

                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    string path = sfd.FileName;
                    WorkAsync(new WorkAsyncInfo
                    {
                        Message = $"Generating HTML report for {selected.Count} tables...",
                        // Enables BackgroundWorker.WorkerReportsProgress so the ReportProgress
                        // call below actually works instead of throwing.
                        ProgressChanged = (args) =>
                        {
                            SetWorkingMessage(args.UserState as string ?? $"{args.ProgressPercentage}%");
                        },
                        Work = (worker, args) =>
                        {
                            var fullEntities = new List<EntityMetadata>();
                            int loaded = 0;
                            foreach (var e in selected)
                            {
                                worker.ReportProgress((int)((++loaded / (double)selected.Count) * 100), $"Loading relationships {loaded} of {selected.Count}: {e.LogicalName}...");
                                var req = new RetrieveEntityRequest
                                {
                                    LogicalName = e.LogicalName,
                                    EntityFilters = EntityFilters.Relationships | EntityFilters.Entity,
                                    RetrieveAsIfPublished = true
                                };
                                fullEntities.Add(((RetrieveEntityResponse)Service.Execute(req)).EntityMetadata);
                            }
                            HtmlDocumentationExporter.ExportToHtmlFile(path, fullEntities);
                        },
                        PostWorkCallBack = (args) =>
                        {
                            if (args.Error != null)
                            {
                                MessageBox.Show(this, $"Error generating HTML:\n{args.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
                        }
                    });
                }
            }
        }

        #endregion

        #region About Dialog

        public void ShowAboutDialog()
        {
            using (var aboutForm = new Form
            {
                Text = "About Metadata Dataverse Document",
                Size = new Size(560, 400),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                Font = new Font("Segoe UI", 9F),
                BackColor = Color.White
            })
            {
                var banner = new Panel
                {
                    Dock = DockStyle.Top,
                    Height = 85,
                    BackColor = Color.FromArgb(30, 41, 59),
                    Padding = new Padding(20, 15, 20, 10)
                };

                var lblTitle = new WinLabel
                {
                    Text = "Metadata Dataverse Document",
                    Font = new Font("Segoe UI", 15F, FontStyle.Bold),
                    ForeColor = Color.White,
                    AutoSize = true,
                    Location = new Point(15, 12)
                };

                var lblVersion = new WinLabel
                {
                    Text = "Version 2.1.14.0 | Enlace al repositorio",
                    Font = new Font("Segoe UI", 9F),
                    ForeColor = Color.FromArgb(148, 163, 184),
                    AutoSize = true,
                    Location = new Point(16, 45)
                };

                banner.Controls.Add(lblTitle);
                banner.Controls.Add(lblVersion);

                var pnlContent = new Panel
                {
                    Dock = DockStyle.Fill,
                    Padding = new Padding(25, 20, 25, 10)
                };

                var lblDesc = new WinLabel
                {
                    Text = "Generador técnico de diccionario de datos, atributos, claves alternas, relaciones (1:N, N:1, N:N) y diagramas relacionales ERD para Microsoft Dataverse y Dynamics 365.",
                    AutoSize = false,
                    Size = new Size(470, 45),
                    Location = new Point(25, 100)
                };

                var lblDev = new WinLabel
                {
                    Text = "Desarrollador: Rogelio Muñoz\nCopyright © 2026 Rogelio Muñoz. Todos los derechos reservados.",
                    AutoSize = true,
                    Location = new Point(25, 155),
                    Font = new Font("Segoe UI", 9F, FontStyle.Bold)
                };

                var linkWeb = new LinkLabel
                {
                    Text = "Sitio Web: www.rogeliomunoz.cl",
                    AutoSize = true,
                    Location = new Point(25, 200)
                };
                linkWeb.LinkClicked += (s, e) => System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("http://www.rogeliomunoz.cl") { UseShellExecute = true });

                var linkEmail = new LinkLabel
                {
                    Text = "Contacto: rmunoz1612@gmail.com",
                    AutoSize = true,
                    Location = new Point(25, 225)
                };
                linkEmail.LinkClicked += (s, e) => System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("mailto:rmunoz1612@gmail.com") { UseShellExecute = true });

                var linkRepo = new LinkLabel
                {
                    Text = "Repositorio: github.com/RmunozMM/MetadataDataverseDocument",
                    AutoSize = true,
                    Location = new Point(25, 250)
                };
                linkRepo.LinkClicked += (s, e) => System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("https://github.com/RmunozMM/MetadataDataverseDocument") { UseShellExecute = true });

                var btnClose = new Button
                {
                    Text = "Cerrar",
                    DialogResult = DialogResult.OK,
                    Size = new Size(85, 30),
                    Location = new Point(430, 305),
                    BackColor = Color.FromArgb(30, 41, 59),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnClose.FlatAppearance.BorderSize = 0;

                aboutForm.Controls.Add(btnClose);
                aboutForm.Controls.Add(linkRepo);
                aboutForm.Controls.Add(linkEmail);
                aboutForm.Controls.Add(linkWeb);
                aboutForm.Controls.Add(lblDev);
                aboutForm.Controls.Add(lblDesc);
                aboutForm.Controls.Add(banner);

                aboutForm.AcceptButton = btnClose;
                aboutForm.ShowDialog(this);
            }
        }

        #endregion

        #region Helper Classes

        private sealed class TableListItem
        {
            public string LogicalName { get; }
            public string DisplayName { get; }
            public bool IsCustom { get; }

            public TableListItem(string logicalName, string displayName, bool isCustom)
            {
                LogicalName = logicalName;
                DisplayName = displayName;
                IsCustom = isCustom;
            }

            public override string ToString()
            {
                string tag = IsCustom ? "[Custom]" : "[Standard]";
                return $"{DisplayName} ({LogicalName})  {tag}";
            }
        }

        private sealed class SolutionItem
        {
            public string UniqueName { get; set; }
            public string FriendlyName { get; set; }
            public string Version { get; set; }
            public string Publisher { get; set; }
        }

        private sealed class SolutionPickerDialog : Form
        {
            public SolutionItem SelectedSolution { get; private set; }
            private ListView lvSolutions;
            private Button btnOk;
            private Button btnCancel;

            public SolutionPickerDialog(DataCollection<Entity> solutions)
            {
                BuildLayout();
                Populate(solutions);
            }

            private void BuildLayout()
            {
                Text = "Select Dataverse Solution";
                Size = new Size(680, 480);
                StartPosition = FormStartPosition.CenterParent;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                Font = new Font("Segoe UI", 9F);

                var lblHeader = new WinLabel
                {
                    Text = "Select a solution to load and document tables:",
                    Dock = DockStyle.Top,
                    Height = 35,
                    Padding = new Padding(12, 10, 0, 0),
                    Font = new Font("Segoe UI", 9.5F, FontStyle.Bold)
                };

                lvSolutions = new ListView
                {
                    Dock = DockStyle.Fill,
                    View = View.Details,
                    FullRowSelect = true,
                    GridLines = true,
                    MultiSelect = false
                };
                lvSolutions.Columns.Add("Display Name", 280);
                lvSolutions.Columns.Add("Unique Name", 180);
                lvSolutions.Columns.Add("Version", 90);
                lvSolutions.Columns.Add("Publisher", 160);
                lvSolutions.DoubleClick += (s, e) => Accept();

                var pnlBottom = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    Height = 45,
                    FlowDirection = FlowDirection.RightToLeft,
                    Padding = new Padding(8)
                };

                btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 80, Height = 28 };
                btnOk = new Button { Text = "Select", Width = 80, Height = 28 };
                btnOk.Click += (s, e) => Accept();

                pnlBottom.Controls.Add(btnCancel);
                pnlBottom.Controls.Add(btnOk);

                Controls.Add(lvSolutions);
                Controls.Add(pnlBottom);
                Controls.Add(lblHeader);

                AcceptButton = btnOk;
                CancelButton = btnCancel;
            }

            private void Populate(DataCollection<Entity> solutions)
            {
                foreach (var sol in solutions)
                {
                    string unique = sol.GetAttributeValue<string>("uniquename");
                    string friendly = sol.GetAttributeValue<string>("friendlyname") ?? unique;
                    string version = sol.GetAttributeValue<string>("version") ?? "1.0.0.0";
                    string pub = sol.Contains("pub.friendlyname")
                        ? ((AliasedValue)sol["pub.friendlyname"]).Value?.ToString()
                        : "Default Publisher";

                    var lvi = new ListViewItem(friendly);
                    lvi.SubItems.Add(unique);
                    lvi.SubItems.Add(version);
                    lvi.SubItems.Add(pub);
                    lvi.Tag = new SolutionItem
                    {
                        UniqueName = unique,
                        FriendlyName = friendly,
                        Version = version,
                        Publisher = pub
                    };
                    lvSolutions.Items.Add(lvi);
                }

                if (lvSolutions.Items.Count > 0)
                {
                    lvSolutions.Items[0].Selected = true;
                }
            }

            private void Accept()
            {
                if (lvSolutions.SelectedItems.Count > 0)
                {
                    SelectedSolution = lvSolutions.SelectedItems[0].Tag as SolutionItem;
                    DialogResult = DialogResult.OK;
                    Close();
                }
            }
        }

        #endregion
    }
}
