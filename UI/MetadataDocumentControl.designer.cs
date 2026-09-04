namespace MetadataDataverseDocument.UI
{
    partial class MetadataDocumentControl
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.splitContainerMain = new System.Windows.Forms.SplitContainer();
            this.panelLeft = new System.Windows.Forms.Panel();
            this.checkedListBoxTables = new System.Windows.Forms.CheckedListBox();
            this.panelLeftTop = new System.Windows.Forms.Panel();
            this.lblSelectionStats = new System.Windows.Forms.Label();
            this.pnlSelectionButtons = new System.Windows.Forms.FlowLayoutPanel();
            this.btnSelectAll = new System.Windows.Forms.Button();
            this.btnDeselectAll = new System.Windows.Forms.Button();
            this.pnlQuickFilters = new System.Windows.Forms.FlowLayoutPanel();
            this.btnFilterAll = new System.Windows.Forms.Button();
            this.btnFilterCustom = new System.Windows.Forms.Button();
            this.btnFilterStandard = new System.Windows.Forms.Button();
            this.txtFilterTables = new System.Windows.Forms.TextBox();
            this.labelSearchTable = new System.Windows.Forms.Label();
            this.panelRight = new System.Windows.Forms.Panel();
            this.tabControlTables = new System.Windows.Forms.TabControl();
            this.labelInfo = new System.Windows.Forms.Label();
            this.panelHeader = new System.Windows.Forms.Panel();
            this.lblHeaderTitle = new System.Windows.Forms.Label();
            this.headerButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerMain)).BeginInit();
            this.splitContainerMain.Panel1.SuspendLayout();
            this.splitContainerMain.Panel2.SuspendLayout();
            this.splitContainerMain.SuspendLayout();
            this.panelLeft.SuspendLayout();
            this.panelLeftTop.SuspendLayout();
            this.pnlSelectionButtons.SuspendLayout();
            this.pnlQuickFilters.SuspendLayout();
            this.panelRight.SuspendLayout();
            this.panelHeader.SuspendLayout();
            this.panelLeftBottom = new System.Windows.Forms.Panel();
            this.SuspendLayout();

            // 
            // panelHeader
            // 
            this.panelHeader.BackColor = System.Drawing.Color.FromArgb(30, 41, 59);
            this.panelHeader.Controls.Add(this.lblHeaderTitle);
            this.panelHeader.Controls.Add(this.headerButtonsPanel);
            this.panelHeader.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelHeader.Height = 55;
            this.panelHeader.Padding = new System.Windows.Forms.Padding(15, 8, 15, 8);

            // 
            // lblHeaderTitle
            // 
            this.lblHeaderTitle.AutoSize = true;
            this.lblHeaderTitle.Dock = System.Windows.Forms.DockStyle.Left;
            this.lblHeaderTitle.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold);
            this.lblHeaderTitle.ForeColor = System.Drawing.Color.White;
            this.lblHeaderTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblHeaderTitle.Text = "METADATA DATAVERSE DOCUMENT";

            // 
            // headerButtonsPanel
            // 
            this.headerButtonsPanel.AutoSize = true;
            this.headerButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.headerButtonsPanel.Dock = System.Windows.Forms.DockStyle.Right;
            this.headerButtonsPanel.FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight;
            this.headerButtonsPanel.Padding = new System.Windows.Forms.Padding(0, 3, 0, 0);
            this.headerButtonsPanel.WrapContents = false;

            // 
            // splitContainerMain
            // 
            this.splitContainerMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainerMain.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainerMain.Location = new System.Drawing.Point(0, 55);
            this.splitContainerMain.SplitterDistance = 300;
            this.splitContainerMain.SplitterWidth = 5;

            // 
            // splitContainerMain.Panel1 (panelLeft)
            // 
            this.splitContainerMain.Panel1.Controls.Add(this.panelLeft);
            this.panelLeft.Controls.Add(this.checkedListBoxTables);
            this.panelLeft.Controls.Add(this.panelLeftBottom);
            this.panelLeft.Controls.Add(this.panelLeftTop);
            this.panelLeft.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelLeftTop.SendToBack();
            this.panelLeftBottom.SendToBack();
            this.checkedListBoxTables.BringToFront();

            // 
            // panelLeftTop
            // 
            this.panelLeftTop.Controls.Add(this.pnlQuickFilters);
            this.panelLeftTop.Controls.Add(this.txtFilterTables);
            this.panelLeftTop.Controls.Add(this.labelSearchTable);
            this.panelLeftTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelLeftTop.Height = 78;
            this.panelLeftTop.Padding = new System.Windows.Forms.Padding(6, 6, 6, 2);

            // 
            // labelSearchTable
            // 
            this.labelSearchTable.Dock = System.Windows.Forms.DockStyle.Top;
            this.labelSearchTable.Font = new System.Drawing.Font("Segoe UI", 8.5F, System.Drawing.FontStyle.Bold);
            this.labelSearchTable.Height = 18;
            this.labelSearchTable.Text = "Search Table(s):";

            // 
            // txtFilterTables
            // 
            this.txtFilterTables.Dock = System.Windows.Forms.DockStyle.Top;
            this.txtFilterTables.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtFilterTables.Height = 22;
            this.txtFilterTables.TextChanged += new System.EventHandler(this.txtFilterTables_TextChanged);

            // 
            // pnlQuickFilters
            // 
            this.pnlQuickFilters.Controls.Add(this.btnFilterAll);
            this.pnlQuickFilters.Controls.Add(this.btnFilterCustom);
            this.pnlQuickFilters.Controls.Add(this.btnFilterStandard);
            this.pnlQuickFilters.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlQuickFilters.Height = 28;
            this.pnlQuickFilters.Padding = new System.Windows.Forms.Padding(0, 2, 0, 0);

            // 
            // btnFilterAll
            // 
            this.btnFilterAll.Text = "All";
            this.btnFilterAll.Width = 60;
            this.btnFilterAll.Height = 24;
            this.btnFilterAll.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            this.btnFilterAll.Click += new System.EventHandler(this.btnFilterAll_Click);

            // 
            // btnFilterCustom
            // 
            this.btnFilterCustom.Text = "Custom";
            this.btnFilterCustom.Width = 75;
            this.btnFilterCustom.Height = 24;
            this.btnFilterCustom.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnFilterCustom.Click += new System.EventHandler(this.btnFilterCustom_Click);

            // 
            // btnFilterStandard
            // 
            this.btnFilterStandard.Text = "Standard";
            this.btnFilterStandard.Width = 75;
            this.btnFilterStandard.Height = 24;
            this.btnFilterStandard.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnFilterStandard.Click += new System.EventHandler(this.btnFilterStandard_Click);

            // 
            // checkedListBoxTables
            // 
            this.checkedListBoxTables.CheckOnClick = true;
            this.checkedListBoxTables.Dock = System.Windows.Forms.DockStyle.Fill;
            this.checkedListBoxTables.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.checkedListBoxTables.FormattingEnabled = true;
            this.checkedListBoxTables.IntegralHeight = false;
            this.checkedListBoxTables.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.checkedListBoxTables_ItemCheck);

            // 
            // panelLeftBottom
            // 
            this.panelLeftBottom.Controls.Add(this.pnlSelectionButtons);
            this.panelLeftBottom.Controls.Add(this.lblSelectionStats);
            this.panelLeftBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelLeftBottom.Height = 56;
            this.panelLeftBottom.Padding = new System.Windows.Forms.Padding(6, 2, 6, 4);

            // 
            // lblSelectionStats
            // 
            this.lblSelectionStats.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblSelectionStats.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Italic);
            this.lblSelectionStats.ForeColor = System.Drawing.Color.DarkSlateGray;
            this.lblSelectionStats.Height = 18;
            this.lblSelectionStats.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblSelectionStats.Text = "0 table(s) selected";

            // 
            // pnlSelectionButtons
            // 
            this.pnlSelectionButtons.Controls.Add(this.btnSelectAll);
            this.pnlSelectionButtons.Controls.Add(this.btnDeselectAll);
            this.pnlSelectionButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnlSelectionButtons.Height = 28;
            this.pnlSelectionButtons.Padding = new System.Windows.Forms.Padding(0);

            // 
            // btnSelectAll
            // 
            this.btnSelectAll.Text = "Select All";
            this.btnSelectAll.Width = 110;
            this.btnSelectAll.Height = 26;
            this.btnSelectAll.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.btnSelectAll.Click += new System.EventHandler(this.btnSelectAll_Click);

            // 
            // btnDeselectAll
            // 
            this.btnDeselectAll.Text = "Deselect All";
            this.btnDeselectAll.Width = 110;
            this.btnDeselectAll.Height = 26;
            this.btnDeselectAll.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.btnDeselectAll.Click += new System.EventHandler(this.btnDeselectAll_Click);

            // 
            // splitContainerMain.Panel2 (panelRight)
            // 
            this.splitContainerMain.Panel2.Controls.Add(this.panelRight);
            this.panelRight.Controls.Add(this.tabControlTables);
            this.panelRight.Controls.Add(this.labelInfo);
            this.panelRight.Dock = System.Windows.Forms.DockStyle.Fill;

            // 
            // labelInfo
            // 
            this.labelInfo.Dock = System.Windows.Forms.DockStyle.Top;
            this.labelInfo.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.labelInfo.ForeColor = System.Drawing.Color.FromArgb(100, 116, 139);
            this.labelInfo.Location = new System.Drawing.Point(0, 0);
            this.labelInfo.Height = 35;
            this.labelInfo.Padding = new System.Windows.Forms.Padding(10, 8, 10, 0);
            this.labelInfo.Text = "Please load tables from a solution or default solution to get started.";

            // 
            // tabControlTables
            // 
            this.tabControlTables.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlTables.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.tabControlTables.Location = new System.Drawing.Point(0, 35);

            // 
            // MetadataDocumentControl
            // 
            this.Controls.Add(this.splitContainerMain);
            this.Controls.Add(this.panelHeader);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Size = new System.Drawing.Size(1200, 750);
            this.Load += new System.EventHandler(this.MetadataDocumentControl_Load);

            this.splitContainerMain.Panel1.ResumeLayout(false);
            this.splitContainerMain.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerMain)).EndInit();
            this.splitContainerMain.ResumeLayout(false);
            this.panelLeft.ResumeLayout(false);
            this.panelLeftTop.ResumeLayout(false);
            this.panelLeftTop.PerformLayout();
            this.pnlSelectionButtons.ResumeLayout(false);
            this.pnlQuickFilters.ResumeLayout(false);
            this.panelRight.ResumeLayout(false);
            this.panelHeader.ResumeLayout(false);
            this.panelHeader.PerformLayout();
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.Label lblHeaderTitle;
        private System.Windows.Forms.FlowLayoutPanel headerButtonsPanel;
        private System.Windows.Forms.SplitContainer splitContainerMain;
        private System.Windows.Forms.Panel panelLeft;
        private System.Windows.Forms.Panel panelLeftTop;
        private System.Windows.Forms.Label labelSearchTable;
        private System.Windows.Forms.TextBox txtFilterTables;
        private System.Windows.Forms.FlowLayoutPanel pnlQuickFilters;
        private System.Windows.Forms.Button btnFilterAll;
        private System.Windows.Forms.Button btnFilterCustom;
        private System.Windows.Forms.Button btnFilterStandard;
        private System.Windows.Forms.FlowLayoutPanel pnlSelectionButtons;
        private System.Windows.Forms.Button btnSelectAll;
        private System.Windows.Forms.Button btnDeselectAll;
        private System.Windows.Forms.Label lblSelectionStats;
        private System.Windows.Forms.CheckedListBox checkedListBoxTables;
        private System.Windows.Forms.Panel panelRight;
        private System.Windows.Forms.Panel panelLeftBottom;
        private System.Windows.Forms.Label labelInfo;
        private System.Windows.Forms.TabControl tabControlTables;
    }
}
