namespace COMPGIS1
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            //Ensures that any ESRI libraries that have been used are unloaded in the correct order. 
            //Failure to do this may result in random crashes on exit due to the operating system unloading 
            //the libraries in the incorrect order. 
            ESRI.ArcGIS.ADF.COMSupport.AOUninitialize.Shutdown();

            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle14 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle12 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle13 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.menuFile = new System.Windows.Forms.ToolStripMenuItem();
            this.menuNewDoc = new System.Windows.Forms.ToolStripMenuItem();
            this.menuOpenDoc = new System.Windows.Forms.ToolStripMenuItem();
            this.menuSaveDoc = new System.Windows.Forms.ToolStripMenuItem();
            this.menuSaveAs = new System.Windows.Forms.ToolStripMenuItem();
            this.menuSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.menuExitApp = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.viewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mapLegendToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lineSectionRiskRankingToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.utilitiesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.synchronizeFeatureTablesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.layerConfiguration = new System.Windows.Forms.ToolStripMenuItem();
            this.runDatabaseUpdatesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.findToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.linesBPToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.linesInASystemToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lineSectionsBPToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.inspectedLineSectionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lineSectionsAssociatedWithLCOsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lineSectionsWithAHistoryOfLeaksToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
            this.lineSectionsByRiskRankingValuesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.boringBPExcavationLocationsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.corrosionBPProbesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.axMapControl1 = new ESRI.ArcGIS.Controls.AxMapControl();
            this.axToolbarControl1 = new ESRI.ArcGIS.Controls.AxToolbarControl();
            this.axTOCControl1 = new ESRI.ArcGIS.Controls.AxTOCControl();
            this.axLicenseControl1 = new ESRI.ArcGIS.Controls.AxLicenseControl();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.tsFileName = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusBarXY = new System.Windows.Forms.ToolStripStatusLabel();
            this.tsVersion = new System.Windows.Forms.ToolStripStatusLabel();
            this.scMain = new System.Windows.Forms.SplitContainer();
            this.tcTOC = new System.Windows.Forms.TabControl();
            this.tabTOC = new System.Windows.Forms.TabPage();
            this.tabFilter = new System.Windows.Forms.TabPage();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lbFiltersApplied = new System.Windows.Forms.ListBox();
            this.cbApplyToSelection = new System.Windows.Forms.CheckBox();
            this.btnRemoveFilter = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.btnFilterSave = new System.Windows.Forms.Button();
            this.btnFilterEdit = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.cbActiveFilter = new System.Windows.Forms.ComboBox();
            this.cbFilterLayerSelect = new System.Windows.Forms.ComboBox();
            this.btnCreateFilter = new System.Windows.Forms.Button();
            this.btnResetFilter = new System.Windows.Forms.Button();
            this.tabFind = new System.Windows.Forms.TabPage();
            this.pnlFindBanner = new System.Windows.Forms.Panel();
            this.lblQueryDesc = new System.Windows.Forms.Label();
            this.pnlFindBG = new System.Windows.Forms.Panel();
            this.pnlFindDetails = new System.Windows.Forms.Panel();
            this.lblQueryParamNote = new System.Windows.Forms.Label();
            this.tabDisplay = new System.Windows.Forms.TabPage();
            this.btnApply = new System.Windows.Forms.Button();
            this.rbByRiskRankingColor = new System.Windows.Forms.RadioButton();
            this.label2 = new System.Windows.Forms.Label();
            this.cbRiskRankingChoices = new System.Windows.Forms.ComboBox();
            this.axSymbologyControl1 = new ESRI.ArcGIS.Controls.AxSymbologyControl();
            this.rbByRiskRankingColorRamp = new System.Windows.Forms.RadioButton();
            this.rbByLayer = new System.Windows.Forms.RadioButton();
            this.scMap = new System.Windows.Forms.SplitContainer();
            this.scMapLegend = new System.Windows.Forms.SplitContainer();
            this.tcMapLegend = new System.Windows.Forms.TabControl();
            this.tabMapLegendRiskRanking = new System.Windows.Forms.TabPage();
            this.btnMapLegendRiskRankingClose = new System.Windows.Forms.Button();
            this.rbMapLegendDisplayLSByRisk = new System.Windows.Forms.RadioButton();
            this.rbMapLegendDisplayLSByLayer = new System.Windows.Forms.RadioButton();
            this.lblMapLegendRiskRankingBanner = new System.Windows.Forms.Label();
            this.cbMapLegendRiskFieldToUse = new System.Windows.Forms.ComboBox();
            this.gridRiskRankingLegend = new System.Windows.Forms.DataGridView();
            this.ColLow = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColLowMed = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColMed = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColMedHigh = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColHigh = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.scSelections = new System.Windows.Forms.SplitContainer();
            this.btnExportTreeViewToExcel = new System.Windows.Forms.Button();
            this.btnExpandAll = new System.Windows.Forms.Button();
            this.btnCollapseAll = new System.Windows.Forms.Button();
            this.treeViewMain = new System.Windows.Forms.TreeView();
            this.cmsTreeViewGeneral = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mnuExpandAll = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuCollapseAll = new System.Windows.Forms.ToolStripMenuItem();
            this.lblSelectedFeaturesTitle = new System.Windows.Forms.Label();
            this.pnlCommands = new System.Windows.Forms.Panel();
            this.tcCommandModes = new System.Windows.Forms.TabControl();
            this.tabPrebuiltSearch = new System.Windows.Forms.TabPage();
            this.tabBE_LineSection_Association = new System.Windows.Forms.TabPage();
            this.tbRadius = new System.Windows.Forms.TextBox();
            this.Reset = new System.Windows.Forms.Button();
            this.Cancel = new System.Windows.Forms.Button();
            this.btnLineSectionsQuery = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cmsBPNodeLineSection = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mnuLineSectionProperties = new System.Windows.Forms.ToolStripMenuItem();
            this.inspectionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cmsBPRootLineSections = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mnuFilterLineSections = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuCustomFilterSeparatorLineSections = new System.Windows.Forms.ToolStripSeparator();
            this.lineSectionsAssociatedWithLCSsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lineSectionsThatHaveBeenInspectedToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lineSectionsWithAHistoryOfLeaksToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lineSectionsWithGuidedWaveInspectionsByLevelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ilTreeViewMain = new System.Windows.Forms.ImageList(this.components);
            this.cmsBPRootLines = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mnuFilterLines = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuCustomFilterSeparatorLines = new System.Windows.Forms.ToolStripSeparator();
            this.selectLinesInASystemToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.inspectedLinesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.linesContainingLicensedMaterialradioactiveFluidsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cmsBPRootBoringLocations = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.cmsBPRootCorrosionProbes = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.cmsBPNodeLines = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mnuLineProperties = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.cmsBPNodeBoringLocations = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mnuBoringLocationProperties = new System.Windows.Forms.ToolStripMenuItem();
            this.spatialQueryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.linesNearbyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lineSectionsNearbyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.associatedLineSectionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cmsBPNodeCorrosionProbes = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mnuCPProperties = new System.Windows.Forms.ToolStripMenuItem();
            this.spatialQueryToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.linesNearbyToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.lineSectionsNearbyToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.tooltipMain = new System.Windows.Forms.ToolTip(this.components);
            this.saveExcelDialog = new System.Windows.Forms.SaveFileDialog();
            this.fileSystemWatcher1 = new System.IO.FileSystemWatcher();
            this.cmsSWRootLines = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem5 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenuItem6 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem7 = new System.Windows.Forms.ToolStripMenuItem();
            this.cmsSWNodeLines = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem8 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.cmsSWRootFlowSegments = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem9 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.cmsSWNodeFlowSegments = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem10 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem11 = new System.Windows.Forms.ToolStripMenuItem();
            this.cmsSWRootComponents = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem12 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.cmsSWNodeComponents = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem13 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem14 = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.axMapControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.axToolbarControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.axTOCControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.axLicenseControl1)).BeginInit();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.scMain)).BeginInit();
            this.scMain.Panel1.SuspendLayout();
            this.scMain.Panel2.SuspendLayout();
            this.scMain.SuspendLayout();
            this.tcTOC.SuspendLayout();
            this.tabTOC.SuspendLayout();
            this.tabFilter.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tabFind.SuspendLayout();
            this.pnlFindBanner.SuspendLayout();
            this.pnlFindBG.SuspendLayout();
            this.pnlFindDetails.SuspendLayout();
            this.tabDisplay.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.axSymbologyControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.scMap)).BeginInit();
            this.scMap.Panel1.SuspendLayout();
            this.scMap.Panel2.SuspendLayout();
            this.scMap.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.scMapLegend)).BeginInit();
            this.scMapLegend.Panel1.SuspendLayout();
            this.scMapLegend.Panel2.SuspendLayout();
            this.scMapLegend.SuspendLayout();
            this.tcMapLegend.SuspendLayout();
            this.tabMapLegendRiskRanking.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridRiskRankingLegend)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.scSelections)).BeginInit();
            this.scSelections.Panel1.SuspendLayout();
            this.scSelections.Panel2.SuspendLayout();
            this.scSelections.SuspendLayout();
            this.cmsTreeViewGeneral.SuspendLayout();
            this.pnlCommands.SuspendLayout();
            this.tcCommandModes.SuspendLayout();
            this.tabBE_LineSection_Association.SuspendLayout();
            this.cmsBPNodeLineSection.SuspendLayout();
            this.cmsBPRootLineSections.SuspendLayout();
            this.cmsBPRootLines.SuspendLayout();
            this.cmsBPNodeLines.SuspendLayout();
            this.cmsBPNodeBoringLocations.SuspendLayout();
            this.cmsBPNodeCorrosionProbes.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).BeginInit();
            this.cmsSWRootLines.SuspendLayout();
            this.cmsSWNodeLines.SuspendLayout();
            this.cmsSWRootFlowSegments.SuspendLayout();
            this.cmsSWNodeFlowSegments.SuspendLayout();
            this.cmsSWRootComponents.SuspendLayout();
            this.cmsSWNodeComponents.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuFile,
            this.toolStripMenuItem2,
            this.viewToolStripMenuItem,
            this.utilitiesToolStripMenuItem,
            this.toolStripMenuItem3,
            this.findToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1336, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            this.menuStrip1.Visible = false;
            this.menuStrip1.MenuActivate += new System.EventHandler(this.menuStrip1_MenuActivate);
            // 
            // menuFile
            // 
            this.menuFile.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuNewDoc,
            this.menuOpenDoc,
            this.menuSaveDoc,
            this.menuSaveAs,
            this.menuSeparator,
            this.menuExitApp});
            this.menuFile.Name = "menuFile";
            this.menuFile.Size = new System.Drawing.Size(37, 20);
            this.menuFile.Text = "File";
            // 
            // menuNewDoc
            // 
            this.menuNewDoc.Image = ((System.Drawing.Image)(resources.GetObject("menuNewDoc.Image")));
            this.menuNewDoc.ImageTransparentColor = System.Drawing.Color.White;
            this.menuNewDoc.Name = "menuNewDoc";
            this.menuNewDoc.Size = new System.Drawing.Size(171, 22);
            this.menuNewDoc.Text = "New Document";
            this.menuNewDoc.Visible = false;
            this.menuNewDoc.Click += new System.EventHandler(this.menuNewDoc_Click);
            // 
            // menuOpenDoc
            // 
            this.menuOpenDoc.Image = ((System.Drawing.Image)(resources.GetObject("menuOpenDoc.Image")));
            this.menuOpenDoc.ImageTransparentColor = System.Drawing.Color.White;
            this.menuOpenDoc.Name = "menuOpenDoc";
            this.menuOpenDoc.Size = new System.Drawing.Size(171, 22);
            this.menuOpenDoc.Text = "Open Document...";
            this.menuOpenDoc.Visible = false;
            this.menuOpenDoc.Click += new System.EventHandler(this.menuOpenDoc_Click);
            // 
            // menuSaveDoc
            // 
            this.menuSaveDoc.Image = ((System.Drawing.Image)(resources.GetObject("menuSaveDoc.Image")));
            this.menuSaveDoc.ImageTransparentColor = System.Drawing.Color.White;
            this.menuSaveDoc.Name = "menuSaveDoc";
            this.menuSaveDoc.Size = new System.Drawing.Size(171, 22);
            this.menuSaveDoc.Text = "SaveDocument";
            this.menuSaveDoc.Visible = false;
            this.menuSaveDoc.Click += new System.EventHandler(this.menuSaveDoc_Click);
            // 
            // menuSaveAs
            // 
            this.menuSaveAs.Name = "menuSaveAs";
            this.menuSaveAs.Size = new System.Drawing.Size(171, 22);
            this.menuSaveAs.Text = "Save As...";
            this.menuSaveAs.Visible = false;
            this.menuSaveAs.Click += new System.EventHandler(this.menuSaveAs_Click);
            // 
            // menuSeparator
            // 
            this.menuSeparator.Name = "menuSeparator";
            this.menuSeparator.Size = new System.Drawing.Size(168, 6);
            this.menuSeparator.Visible = false;
            // 
            // menuExitApp
            // 
            this.menuExitApp.Name = "menuExitApp";
            this.menuExitApp.Size = new System.Drawing.Size(171, 22);
            this.menuExitApp.Text = "Exit";
            this.menuExitApp.Click += new System.EventHandler(this.menuExitApp_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(12, 20);
            // 
            // viewToolStripMenuItem
            // 
            this.viewToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mapLegendToolStripMenuItem});
            this.viewToolStripMenuItem.Name = "viewToolStripMenuItem";
            this.viewToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.viewToolStripMenuItem.Text = "View";
            // 
            // mapLegendToolStripMenuItem
            // 
            this.mapLegendToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lineSectionRiskRankingToolStripMenuItem});
            this.mapLegendToolStripMenuItem.Name = "mapLegendToolStripMenuItem";
            this.mapLegendToolStripMenuItem.Size = new System.Drawing.Size(140, 22);
            this.mapLegendToolStripMenuItem.Text = "Map Legend";
            // 
            // lineSectionRiskRankingToolStripMenuItem
            // 
            this.lineSectionRiskRankingToolStripMenuItem.Name = "lineSectionRiskRankingToolStripMenuItem";
            this.lineSectionRiskRankingToolStripMenuItem.Size = new System.Drawing.Size(208, 22);
            this.lineSectionRiskRankingToolStripMenuItem.Text = "Line Section Risk Ranking";
            this.lineSectionRiskRankingToolStripMenuItem.Click += new System.EventHandler(this.lineSectionRiskRankingToolStripMenuItem_Click);
            // 
            // utilitiesToolStripMenuItem
            // 
            this.utilitiesToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.synchronizeFeatureTablesToolStripMenuItem,
            this.layerConfiguration,
            this.runDatabaseUpdatesToolStripMenuItem});
            this.utilitiesToolStripMenuItem.Name = "utilitiesToolStripMenuItem";
            this.utilitiesToolStripMenuItem.Size = new System.Drawing.Size(58, 20);
            this.utilitiesToolStripMenuItem.Text = "Utilities";
            // 
            // synchronizeFeatureTablesToolStripMenuItem
            // 
            this.synchronizeFeatureTablesToolStripMenuItem.Name = "synchronizeFeatureTablesToolStripMenuItem";
            this.synchronizeFeatureTablesToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.synchronizeFeatureTablesToolStripMenuItem.Text = "Feature Configuration...";
            this.synchronizeFeatureTablesToolStripMenuItem.Click += new System.EventHandler(this.synchronizeFeatureTablesToolStripMenuItem_Click);
            // 
            // layerConfiguration
            // 
            this.layerConfiguration.Name = "layerConfiguration";
            this.layerConfiguration.Size = new System.Drawing.Size(201, 22);
            this.layerConfiguration.Text = "Layer Configuration...";
            this.layerConfiguration.Click += new System.EventHandler(this.layerConfiguration_Click);
            // 
            // runDatabaseUpdatesToolStripMenuItem
            // 
            this.runDatabaseUpdatesToolStripMenuItem.Name = "runDatabaseUpdatesToolStripMenuItem";
            this.runDatabaseUpdatesToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.runDatabaseUpdatesToolStripMenuItem.Text = "Run Database Updates...";
            this.runDatabaseUpdatesToolStripMenuItem.Click += new System.EventHandler(this.runDatabaseUpdatesToolStripMenuItem_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(12, 20);
            // 
            // findToolStripMenuItem
            // 
            this.findToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.linesBPToolStripMenuItem,
            this.lineSectionsBPToolStripMenuItem,
            this.boringBPExcavationLocationsToolStripMenuItem,
            this.corrosionBPProbesToolStripMenuItem});
            this.findToolStripMenuItem.Name = "findToolStripMenuItem";
            this.findToolStripMenuItem.Size = new System.Drawing.Size(42, 20);
            this.findToolStripMenuItem.Text = "Find";
            // 
            // linesBPToolStripMenuItem
            // 
            this.linesBPToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem1,
            this.linesInASystemToolStripMenuItem});
            this.linesBPToolStripMenuItem.Name = "linesBPToolStripMenuItem";
            this.linesBPToolStripMenuItem.Size = new System.Drawing.Size(230, 22);
            this.linesBPToolStripMenuItem.Text = "Lines";
            this.linesBPToolStripMenuItem.Visible = false;
            // 
            // lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem1
            // 
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem1.Name = "lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem1";
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem1.Size = new System.Drawing.Size(359, 22);
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem1.Text = "Lines Containing Licensed Material (radioactive fluids)";
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem1.Click += new System.EventHandler(this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem1_Click);
            // 
            // linesInASystemToolStripMenuItem
            // 
            this.linesInASystemToolStripMenuItem.Name = "linesInASystemToolStripMenuItem";
            this.linesInASystemToolStripMenuItem.Size = new System.Drawing.Size(359, 22);
            this.linesInASystemToolStripMenuItem.Text = "Lines in a System...";
            this.linesInASystemToolStripMenuItem.Click += new System.EventHandler(this.linesInASystemToolStripMenuItem_Click);
            // 
            // lineSectionsBPToolStripMenuItem
            // 
            this.lineSectionsBPToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.inspectedLineSectionsToolStripMenuItem,
            this.lineSectionsAssociatedWithLCOsToolStripMenuItem,
            this.lineSectionsWithAHistoryOfLeaksToolStripMenuItem1,
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem2,
            this.toolStripMenuItem4,
            this.lineSectionsByRiskRankingValuesToolStripMenuItem});
            this.lineSectionsBPToolStripMenuItem.Name = "lineSectionsBPToolStripMenuItem";
            this.lineSectionsBPToolStripMenuItem.Size = new System.Drawing.Size(230, 22);
            this.lineSectionsBPToolStripMenuItem.Text = "Line Sections";
            this.lineSectionsBPToolStripMenuItem.Visible = false;
            // 
            // inspectedLineSectionsToolStripMenuItem
            // 
            this.inspectedLineSectionsToolStripMenuItem.Name = "inspectedLineSectionsToolStripMenuItem";
            this.inspectedLineSectionsToolStripMenuItem.Size = new System.Drawing.Size(401, 22);
            this.inspectedLineSectionsToolStripMenuItem.Text = "Inspected Line Sections...";
            this.inspectedLineSectionsToolStripMenuItem.Click += new System.EventHandler(this.lineSectionsThatHaveBeenInspectedToolStripMenuItem_Click);
            // 
            // lineSectionsAssociatedWithLCOsToolStripMenuItem
            // 
            this.lineSectionsAssociatedWithLCOsToolStripMenuItem.Name = "lineSectionsAssociatedWithLCOsToolStripMenuItem";
            this.lineSectionsAssociatedWithLCOsToolStripMenuItem.Size = new System.Drawing.Size(401, 22);
            this.lineSectionsAssociatedWithLCOsToolStripMenuItem.Text = "Line Sections Associated with LCO\'s...";
            this.lineSectionsAssociatedWithLCOsToolStripMenuItem.Click += new System.EventHandler(this.lineSectionsAssociatedWithLCOsToolStripMenuItem_Click);
            // 
            // lineSectionsWithAHistoryOfLeaksToolStripMenuItem1
            // 
            this.lineSectionsWithAHistoryOfLeaksToolStripMenuItem1.Enabled = false;
            this.lineSectionsWithAHistoryOfLeaksToolStripMenuItem1.Name = "lineSectionsWithAHistoryOfLeaksToolStripMenuItem1";
            this.lineSectionsWithAHistoryOfLeaksToolStripMenuItem1.Size = new System.Drawing.Size(401, 22);
            this.lineSectionsWithAHistoryOfLeaksToolStripMenuItem1.Text = "Line Sections with a history of leaks";
            this.lineSectionsWithAHistoryOfLeaksToolStripMenuItem1.Visible = false;
            // 
            // lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem2
            // 
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem2.Enabled = false;
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem2.Name = "lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem2";
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem2.Size = new System.Drawing.Size(401, 22);
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem2.Text = "Line Sections Containing Licensed Material (radioactive fluids)";
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem2.Visible = false;
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(401, 22);
            this.toolStripMenuItem4.Text = "Line Sections with Guided Wave Inspections by Level...";
            this.toolStripMenuItem4.Click += new System.EventHandler(this.lineSectionsWithGuidedWaveInspectionsByLevelToolStripMenuItem_Click);
            // 
            // lineSectionsByRiskRankingValuesToolStripMenuItem
            // 
            this.lineSectionsByRiskRankingValuesToolStripMenuItem.Name = "lineSectionsByRiskRankingValuesToolStripMenuItem";
            this.lineSectionsByRiskRankingValuesToolStripMenuItem.Size = new System.Drawing.Size(401, 22);
            this.lineSectionsByRiskRankingValuesToolStripMenuItem.Text = "Line Sections by Risk Ranking Values";
            this.lineSectionsByRiskRankingValuesToolStripMenuItem.Click += new System.EventHandler(this.lineSectionsByRiskRankingValuesToolStripMenuItem_Click);
            // 
            // boringBPExcavationLocationsToolStripMenuItem
            // 
            this.boringBPExcavationLocationsToolStripMenuItem.Name = "boringBPExcavationLocationsToolStripMenuItem";
            this.boringBPExcavationLocationsToolStripMenuItem.Size = new System.Drawing.Size(230, 22);
            this.boringBPExcavationLocationsToolStripMenuItem.Text = "Boring / Excavation Locations";
            this.boringBPExcavationLocationsToolStripMenuItem.Visible = false;
            // 
            // corrosionBPProbesToolStripMenuItem
            // 
            this.corrosionBPProbesToolStripMenuItem.Name = "corrosionBPProbesToolStripMenuItem";
            this.corrosionBPProbesToolStripMenuItem.Size = new System.Drawing.Size(230, 22);
            this.corrosionBPProbesToolStripMenuItem.Text = "Corrosion Probes";
            this.corrosionBPProbesToolStripMenuItem.Visible = false;
            // 
            // axMapControl1
            // 
            this.axMapControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.axMapControl1.Location = new System.Drawing.Point(0, 0);
            this.axMapControl1.Name = "axMapControl1";
            this.axMapControl1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axMapControl1.OcxState")));
            this.axMapControl1.Size = new System.Drawing.Size(789, 694);
            this.axMapControl1.TabIndex = 2;
            this.axMapControl1.OnMouseDown += new ESRI.ArcGIS.Controls.IMapControlEvents2_Ax_OnMouseDownEventHandler(this.axMapControl1_OnMouseDown);
            this.axMapControl1.OnMouseMove += new ESRI.ArcGIS.Controls.IMapControlEvents2_Ax_OnMouseMoveEventHandler(this.axMapControl1_OnMouseMove);
            this.axMapControl1.OnSelectionChanged += new System.EventHandler(this.axMapControl1_OnSelectionChanged);
            this.axMapControl1.OnMapReplaced += new ESRI.ArcGIS.Controls.IMapControlEvents2_Ax_OnMapReplacedEventHandler(this.axMapControl1_OnMapReplaced);
            // 
            // axToolbarControl1
            // 
            this.axToolbarControl1.Dock = System.Windows.Forms.DockStyle.Top;
            this.axToolbarControl1.Location = new System.Drawing.Point(0, 0);
            this.axToolbarControl1.Name = "axToolbarControl1";
            this.axToolbarControl1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axToolbarControl1.OcxState")));
            this.axToolbarControl1.Size = new System.Drawing.Size(1336, 28);
            this.axToolbarControl1.TabIndex = 3;
            // 
            // axTOCControl1
            // 
            this.axTOCControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.axTOCControl1.Location = new System.Drawing.Point(3, 3);
            this.axTOCControl1.Name = "axTOCControl1";
            this.axTOCControl1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axTOCControl1.OcxState")));
            this.axTOCControl1.Size = new System.Drawing.Size(205, 745);
            this.axTOCControl1.TabIndex = 4;
            this.axTOCControl1.OnDoubleClick += new ESRI.ArcGIS.Controls.ITOCControlEvents_Ax_OnDoubleClickEventHandler(this.axTOCControl1_OnDoubleClick);
            // 
            // axLicenseControl1
            // 
            this.axLicenseControl1.Enabled = true;
            this.axLicenseControl1.Location = new System.Drawing.Point(973, 54);
            this.axLicenseControl1.Name = "axLicenseControl1";
            this.axLicenseControl1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axLicenseControl1.OcxState")));
            this.axLicenseControl1.Size = new System.Drawing.Size(32, 32);
            this.axLicenseControl1.TabIndex = 5;
            // 
            // splitter1
            // 
            this.splitter1.Location = new System.Drawing.Point(0, 28);
            this.splitter1.Name = "splitter1";
            this.splitter1.Size = new System.Drawing.Size(3, 815);
            this.splitter1.TabIndex = 6;
            this.splitter1.TabStop = false;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsFileName,
            this.statusBarXY,
            this.tsVersion});
            this.statusStrip1.Location = new System.Drawing.Point(3, 821);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1333, 22);
            this.statusStrip1.Stretch = false;
            this.statusStrip1.TabIndex = 7;
            this.statusStrip1.Text = "statusBar1";
            // 
            // tsFileName
            // 
            this.tsFileName.AutoSize = false;
            this.tsFileName.BackColor = System.Drawing.SystemColors.Control;
            this.tsFileName.BorderStyle = System.Windows.Forms.Border3DStyle.RaisedInner;
            this.tsFileName.Name = "tsFileName";
            this.tsFileName.Size = new System.Drawing.Size(200, 17);
            this.tsFileName.Text = "-";
            // 
            // statusBarXY
            // 
            this.statusBarXY.AutoSize = false;
            this.statusBarXY.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter;
            this.statusBarXY.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.statusBarXY.Name = "statusBarXY";
            this.statusBarXY.Size = new System.Drawing.Size(200, 17);
            this.statusBarXY.Text = "-";
            // 
            // tsVersion
            // 
            this.tsVersion.AutoSize = false;
            this.tsVersion.Name = "tsVersion";
            this.tsVersion.Size = new System.Drawing.Size(120, 17);
            this.tsVersion.Text = "v1.1.3x - 04.07.2016x";
            // 
            // scMain
            // 
            this.scMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.scMain.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.scMain.Location = new System.Drawing.Point(3, 28);
            this.scMain.Name = "scMain";
            // 
            // scMain.Panel1
            // 
            this.scMain.Panel1.Controls.Add(this.tcTOC);
            this.scMain.Panel1.Padding = new System.Windows.Forms.Padding(8);
            this.scMain.Panel1MinSize = 235;
            // 
            // scMain.Panel2
            // 
            this.scMain.Panel2.Controls.Add(this.scMap);
            this.scMain.Panel2.Padding = new System.Windows.Forms.Padding(4);
            this.scMain.Size = new System.Drawing.Size(1333, 793);
            this.scMain.SplitterDistance = 235;
            this.scMain.TabIndex = 18;
            // 
            // tcTOC
            // 
            this.tcTOC.Controls.Add(this.tabTOC);
            this.tcTOC.Controls.Add(this.tabFilter);
            this.tcTOC.Controls.Add(this.tabFind);
            this.tcTOC.Controls.Add(this.tabDisplay);
            this.tcTOC.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tcTOC.Location = new System.Drawing.Point(8, 8);
            this.tcTOC.Name = "tcTOC";
            this.tcTOC.SelectedIndex = 0;
            this.tcTOC.Size = new System.Drawing.Size(219, 777);
            this.tcTOC.TabIndex = 5;
            // 
            // tabTOC
            // 
            this.tabTOC.Controls.Add(this.axTOCControl1);
            this.tabTOC.Location = new System.Drawing.Point(4, 22);
            this.tabTOC.Name = "tabTOC";
            this.tabTOC.Padding = new System.Windows.Forms.Padding(3);
            this.tabTOC.Size = new System.Drawing.Size(211, 751);
            this.tabTOC.TabIndex = 0;
            this.tabTOC.Text = "ToC";
            this.tabTOC.UseVisualStyleBackColor = true;
            // 
            // tabFilter
            // 
            this.tabFilter.Controls.Add(this.panel1);
            this.tabFilter.Controls.Add(this.cbApplyToSelection);
            this.tabFilter.Controls.Add(this.btnRemoveFilter);
            this.tabFilter.Controls.Add(this.label4);
            this.tabFilter.Controls.Add(this.btnFilterSave);
            this.tabFilter.Controls.Add(this.btnFilterEdit);
            this.tabFilter.Controls.Add(this.label3);
            this.tabFilter.Controls.Add(this.cbActiveFilter);
            this.tabFilter.Controls.Add(this.cbFilterLayerSelect);
            this.tabFilter.Controls.Add(this.btnCreateFilter);
            this.tabFilter.Controls.Add(this.btnResetFilter);
            this.tabFilter.Location = new System.Drawing.Point(4, 22);
            this.tabFilter.Name = "tabFilter";
            this.tabFilter.Size = new System.Drawing.Size(211, 751);
            this.tabFilter.TabIndex = 1;
            this.tabFilter.Text = "Filter";
            this.tabFilter.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.SystemColors.Info;
            this.panel1.Controls.Add(this.lbFiltersApplied);
            this.panel1.Location = new System.Drawing.Point(0, 210);
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding(4);
            this.panel1.Size = new System.Drawing.Size(209, 539);
            this.panel1.TabIndex = 25;
            // 
            // lbFiltersApplied
            // 
            this.lbFiltersApplied.BackColor = System.Drawing.SystemColors.Info;
            this.lbFiltersApplied.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.lbFiltersApplied.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbFiltersApplied.FormattingEnabled = true;
            this.lbFiltersApplied.Location = new System.Drawing.Point(4, 4);
            this.lbFiltersApplied.Name = "lbFiltersApplied";
            this.lbFiltersApplied.SelectionMode = System.Windows.Forms.SelectionMode.None;
            this.lbFiltersApplied.Size = new System.Drawing.Size(201, 531);
            this.lbFiltersApplied.TabIndex = 26;
            this.lbFiltersApplied.Visible = false;
            // 
            // cbApplyToSelection
            // 
            this.cbApplyToSelection.AutoSize = true;
            this.cbApplyToSelection.Location = new System.Drawing.Point(9, 103);
            this.cbApplyToSelection.Name = "cbApplyToSelection";
            this.cbApplyToSelection.Size = new System.Drawing.Size(148, 17);
            this.cbApplyToSelection.TabIndex = 24;
            this.cbApplyToSelection.Text = "Apply to Current Selection";
            this.cbApplyToSelection.UseVisualStyleBackColor = true;
            // 
            // btnRemoveFilter
            // 
            this.btnRemoveFilter.Location = new System.Drawing.Point(145, 183);
            this.btnRemoveFilter.Name = "btnRemoveFilter";
            this.btnRemoveFilter.Size = new System.Drawing.Size(58, 23);
            this.btnRemoveFilter.TabIndex = 22;
            this.btnRemoveFilter.Text = "Remove";
            this.btnRemoveFilter.UseVisualStyleBackColor = true;
            this.btnRemoveFilter.Click += new System.EventHandler(this.btnRemoveFilter_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 60);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(74, 13);
            this.label4.TabIndex = 19;
            this.label4.Text = "Selected Filter";
            // 
            // btnFilterSave
            // 
            this.btnFilterSave.Location = new System.Drawing.Point(9, 135);
            this.btnFilterSave.Name = "btnFilterSave";
            this.btnFilterSave.Size = new System.Drawing.Size(58, 23);
            this.btnFilterSave.TabIndex = 21;
            this.btnFilterSave.Text = "Save";
            this.btnFilterSave.UseVisualStyleBackColor = true;
            this.btnFilterSave.Visible = false;
            this.btnFilterSave.Click += new System.EventHandler(this.btnFilterSave_Click);
            // 
            // btnFilterEdit
            // 
            this.btnFilterEdit.Location = new System.Drawing.Point(9, 136);
            this.btnFilterEdit.Name = "btnFilterEdit";
            this.btnFilterEdit.Size = new System.Drawing.Size(58, 23);
            this.btnFilterEdit.TabIndex = 23;
            this.btnFilterEdit.Text = "Edit";
            this.btnFilterEdit.UseVisualStyleBackColor = true;
            this.btnFilterEdit.Visible = false;
            this.btnFilterEdit.Click += new System.EventHandler(this.btnFilterEdit_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 10);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(33, 13);
            this.label3.TabIndex = 18;
            this.label3.Text = "Layer";
            // 
            // cbActiveFilter
            // 
            this.cbActiveFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbActiveFilter.FormattingEnabled = true;
            this.cbActiveFilter.Location = new System.Drawing.Point(9, 76);
            this.cbActiveFilter.Name = "cbActiveFilter";
            this.cbActiveFilter.Size = new System.Drawing.Size(194, 21);
            this.cbActiveFilter.TabIndex = 14;
            this.cbActiveFilter.SelectedIndexChanged += new System.EventHandler(this.cbActiveLayerFilter_SelectedIndexChanged);
            // 
            // cbFilterLayerSelect
            // 
            this.cbFilterLayerSelect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFilterLayerSelect.FormattingEnabled = true;
            this.cbFilterLayerSelect.Location = new System.Drawing.Point(9, 26);
            this.cbFilterLayerSelect.Name = "cbFilterLayerSelect";
            this.cbFilterLayerSelect.Size = new System.Drawing.Size(194, 21);
            this.cbFilterLayerSelect.TabIndex = 17;
            this.cbFilterLayerSelect.SelectedIndexChanged += new System.EventHandler(this.cbFilterLayerSelect_SelectedIndexChanged);
            // 
            // btnCreateFilter
            // 
            this.btnCreateFilter.Location = new System.Drawing.Point(145, 135);
            this.btnCreateFilter.Name = "btnCreateFilter";
            this.btnCreateFilter.Size = new System.Drawing.Size(58, 23);
            this.btnCreateFilter.TabIndex = 15;
            this.btnCreateFilter.Text = "Create";
            this.btnCreateFilter.UseVisualStyleBackColor = true;
            this.btnCreateFilter.Click += new System.EventHandler(this.btnCreateFilter_Click);
            // 
            // btnResetFilter
            // 
            this.btnResetFilter.Location = new System.Drawing.Point(145, 159);
            this.btnResetFilter.Name = "btnResetFilter";
            this.btnResetFilter.Size = new System.Drawing.Size(58, 23);
            this.btnResetFilter.TabIndex = 16;
            this.btnResetFilter.Text = "Clear";
            this.btnResetFilter.UseVisualStyleBackColor = true;
            this.btnResetFilter.Click += new System.EventHandler(this.btnResetFilter_Click);
            // 
            // tabFind
            // 
            this.tabFind.Controls.Add(this.pnlFindBanner);
            this.tabFind.Controls.Add(this.pnlFindBG);
            this.tabFind.Location = new System.Drawing.Point(4, 22);
            this.tabFind.Name = "tabFind";
            this.tabFind.Size = new System.Drawing.Size(211, 751);
            this.tabFind.TabIndex = 3;
            this.tabFind.Text = "Find";
            this.tabFind.UseVisualStyleBackColor = true;
            // 
            // pnlFindBanner
            // 
            this.pnlFindBanner.Controls.Add(this.lblQueryDesc);
            this.pnlFindBanner.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlFindBanner.Location = new System.Drawing.Point(0, 0);
            this.pnlFindBanner.Name = "pnlFindBanner";
            this.pnlFindBanner.Size = new System.Drawing.Size(211, 28);
            this.pnlFindBanner.TabIndex = 34;
            // 
            // lblQueryDesc
            // 
            this.lblQueryDesc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblQueryDesc.AutoEllipsis = true;
            this.lblQueryDesc.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.lblQueryDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQueryDesc.Location = new System.Drawing.Point(5, 4);
            this.lblQueryDesc.MinimumSize = new System.Drawing.Size(2, 20);
            this.lblQueryDesc.Name = "lblQueryDesc";
            this.lblQueryDesc.Size = new System.Drawing.Size(203, 20);
            this.lblQueryDesc.TabIndex = 23;
            this.lblQueryDesc.Text = "-";
            this.lblQueryDesc.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblQueryDesc.UseCompatibleTextRendering = true;
            // 
            // pnlFindBG
            // 
            this.pnlFindBG.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlFindBG.BackColor = System.Drawing.SystemColors.Window;
            this.pnlFindBG.Controls.Add(this.pnlFindDetails);
            this.pnlFindBG.Location = new System.Drawing.Point(0, 27);
            this.pnlFindBG.Name = "pnlFindBG";
            this.pnlFindBG.Padding = new System.Windows.Forms.Padding(3);
            this.pnlFindBG.Size = new System.Drawing.Size(211, 724);
            this.pnlFindBG.TabIndex = 25;
            // 
            // pnlFindDetails
            // 
            this.pnlFindDetails.BackColor = System.Drawing.SystemColors.Info;
            this.pnlFindDetails.Controls.Add(this.lblQueryParamNote);
            this.pnlFindDetails.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlFindDetails.Location = new System.Drawing.Point(3, 3);
            this.pnlFindDetails.Name = "pnlFindDetails";
            this.pnlFindDetails.Padding = new System.Windows.Forms.Padding(3);
            this.pnlFindDetails.Size = new System.Drawing.Size(205, 718);
            this.pnlFindDetails.TabIndex = 0;
            // 
            // lblQueryParamNote
            // 
            this.lblQueryParamNote.AutoEllipsis = true;
            this.lblQueryParamNote.BackColor = System.Drawing.SystemColors.Info;
            this.lblQueryParamNote.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblQueryParamNote.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQueryParamNote.ForeColor = System.Drawing.SystemColors.WindowText;
            this.lblQueryParamNote.Location = new System.Drawing.Point(3, 3);
            this.lblQueryParamNote.Name = "lblQueryParamNote";
            this.lblQueryParamNote.Size = new System.Drawing.Size(199, 712);
            this.lblQueryParamNote.TabIndex = 32;
            // 
            // tabDisplay
            // 
            this.tabDisplay.Controls.Add(this.btnApply);
            this.tabDisplay.Controls.Add(this.rbByRiskRankingColor);
            this.tabDisplay.Controls.Add(this.label2);
            this.tabDisplay.Controls.Add(this.cbRiskRankingChoices);
            this.tabDisplay.Controls.Add(this.axSymbologyControl1);
            this.tabDisplay.Controls.Add(this.rbByRiskRankingColorRamp);
            this.tabDisplay.Controls.Add(this.rbByLayer);
            this.tabDisplay.Location = new System.Drawing.Point(4, 22);
            this.tabDisplay.Name = "tabDisplay";
            this.tabDisplay.Size = new System.Drawing.Size(211, 751);
            this.tabDisplay.TabIndex = 4;
            this.tabDisplay.Text = "Display";
            this.tabDisplay.UseVisualStyleBackColor = true;
            // 
            // btnApply
            // 
            this.btnApply.Location = new System.Drawing.Point(14, 714);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(75, 23);
            this.btnApply.TabIndex = 18;
            this.btnApply.Text = "Apply";
            this.btnApply.UseVisualStyleBackColor = true;
            // 
            // rbByRiskRankingColor
            // 
            this.rbByRiskRankingColor.AutoSize = true;
            this.rbByRiskRankingColor.Location = new System.Drawing.Point(25, 59);
            this.rbByRiskRankingColor.Name = "rbByRiskRankingColor";
            this.rbByRiskRankingColor.Size = new System.Drawing.Size(131, 17);
            this.rbByRiskRankingColor.TabIndex = 17;
            this.rbByRiskRankingColor.Text = "By Risk Ranking Color";
            this.rbByRiskRankingColor.UseVisualStyleBackColor = true;
            this.rbByRiskRankingColor.CheckedChanged += new System.EventHandler(this.rbByLayer_CheckedChanged);
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.Dock = System.Windows.Forms.DockStyle.Top;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(0, 0);
            this.label2.MinimumSize = new System.Drawing.Size(2, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(211, 25);
            this.label2.TabIndex = 16;
            this.label2.Text = "Line Sections";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cbRiskRankingChoices
            // 
            this.cbRiskRankingChoices.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbRiskRankingChoices.FormattingEnabled = true;
            this.cbRiskRankingChoices.Items.AddRange(new object[] {
            "Overall Ranking",
            "Susceptibility Ranking",
            "Consequence Ranking"});
            this.cbRiskRankingChoices.Location = new System.Drawing.Point(25, 105);
            this.cbRiskRankingChoices.Name = "cbRiskRankingChoices";
            this.cbRiskRankingChoices.Size = new System.Drawing.Size(170, 21);
            this.cbRiskRankingChoices.TabIndex = 15;
            // 
            // axSymbologyControl1
            // 
            this.axSymbologyControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.axSymbologyControl1.Location = new System.Drawing.Point(14, 142);
            this.axSymbologyControl1.Name = "axSymbologyControl1";
            this.axSymbologyControl1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axSymbologyControl1.OcxState")));
            this.axSymbologyControl1.Size = new System.Drawing.Size(181, 566);
            this.axSymbologyControl1.TabIndex = 14;
            // 
            // rbByRiskRankingColorRamp
            // 
            this.rbByRiskRankingColorRamp.AutoSize = true;
            this.rbByRiskRankingColorRamp.Location = new System.Drawing.Point(25, 82);
            this.rbByRiskRankingColorRamp.Name = "rbByRiskRankingColorRamp";
            this.rbByRiskRankingColorRamp.Size = new System.Drawing.Size(161, 17);
            this.rbByRiskRankingColorRamp.TabIndex = 13;
            this.rbByRiskRankingColorRamp.Text = "Risk Ranking by Color Ramp";
            this.rbByRiskRankingColorRamp.UseVisualStyleBackColor = true;
            this.rbByRiskRankingColorRamp.CheckedChanged += new System.EventHandler(this.rbByLayer_CheckedChanged);
            // 
            // rbByLayer
            // 
            this.rbByLayer.AutoSize = true;
            this.rbByLayer.Checked = true;
            this.rbByLayer.Location = new System.Drawing.Point(14, 36);
            this.rbByLayer.Name = "rbByLayer";
            this.rbByLayer.Size = new System.Drawing.Size(129, 17);
            this.rbByLayer.TabIndex = 12;
            this.rbByLayer.TabStop = true;
            this.rbByLayer.Text = "Default Layer Settings";
            this.rbByLayer.UseVisualStyleBackColor = true;
            this.rbByLayer.CheckedChanged += new System.EventHandler(this.rbByLayer_CheckedChanged);
            // 
            // scMap
            // 
            this.scMap.Dock = System.Windows.Forms.DockStyle.Fill;
            this.scMap.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.scMap.IsSplitterFixed = true;
            this.scMap.Location = new System.Drawing.Point(4, 4);
            this.scMap.Name = "scMap";
            // 
            // scMap.Panel1
            // 
            this.scMap.Panel1.Controls.Add(this.scMapLegend);
            this.scMap.Panel1.Padding = new System.Windows.Forms.Padding(4);
            // 
            // scMap.Panel2
            // 
            this.scMap.Panel2.Controls.Add(this.scSelections);
            this.scMap.Panel2MinSize = 280;
            this.scMap.Size = new System.Drawing.Size(1086, 785);
            this.scMap.SplitterDistance = 797;
            this.scMap.TabIndex = 3;
            // 
            // scMapLegend
            // 
            this.scMapLegend.Dock = System.Windows.Forms.DockStyle.Fill;
            this.scMapLegend.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.scMapLegend.IsSplitterFixed = true;
            this.scMapLegend.Location = new System.Drawing.Point(4, 4);
            this.scMapLegend.Name = "scMapLegend";
            this.scMapLegend.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // scMapLegend.Panel1
            // 
            this.scMapLegend.Panel1.Controls.Add(this.axMapControl1);
            // 
            // scMapLegend.Panel2
            // 
            this.scMapLegend.Panel2.Controls.Add(this.tcMapLegend);
            this.scMapLegend.Size = new System.Drawing.Size(789, 777);
            this.scMapLegend.SplitterDistance = 694;
            this.scMapLegend.TabIndex = 3;
            // 
            // tcMapLegend
            // 
            this.tcMapLegend.Appearance = System.Windows.Forms.TabAppearance.Buttons;
            this.tcMapLegend.Controls.Add(this.tabMapLegendRiskRanking);
            this.tcMapLegend.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tcMapLegend.ItemSize = new System.Drawing.Size(0, 1);
            this.tcMapLegend.Location = new System.Drawing.Point(0, 0);
            this.tcMapLegend.Multiline = true;
            this.tcMapLegend.Name = "tcMapLegend";
            this.tcMapLegend.Padding = new System.Drawing.Point(0, 0);
            this.tcMapLegend.SelectedIndex = 0;
            this.tcMapLegend.Size = new System.Drawing.Size(789, 79);
            this.tcMapLegend.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tcMapLegend.TabIndex = 2;
            // 
            // tabMapLegendRiskRanking
            // 
            this.tabMapLegendRiskRanking.Controls.Add(this.btnMapLegendRiskRankingClose);
            this.tabMapLegendRiskRanking.Controls.Add(this.rbMapLegendDisplayLSByRisk);
            this.tabMapLegendRiskRanking.Controls.Add(this.rbMapLegendDisplayLSByLayer);
            this.tabMapLegendRiskRanking.Controls.Add(this.lblMapLegendRiskRankingBanner);
            this.tabMapLegendRiskRanking.Controls.Add(this.cbMapLegendRiskFieldToUse);
            this.tabMapLegendRiskRanking.Controls.Add(this.gridRiskRankingLegend);
            this.tabMapLegendRiskRanking.Location = new System.Drawing.Point(4, 5);
            this.tabMapLegendRiskRanking.Name = "tabMapLegendRiskRanking";
            this.tabMapLegendRiskRanking.Size = new System.Drawing.Size(781, 70);
            this.tabMapLegendRiskRanking.TabIndex = 2;
            this.tabMapLegendRiskRanking.Text = "-";
            this.tabMapLegendRiskRanking.UseVisualStyleBackColor = true;
            // 
            // btnMapLegendRiskRankingClose
            // 
            this.btnMapLegendRiskRankingClose.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnMapLegendRiskRankingClose.FlatAppearance.BorderSize = 0;
            this.btnMapLegendRiskRankingClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMapLegendRiskRankingClose.Image = ((System.Drawing.Image)(resources.GetObject("btnMapLegendRiskRankingClose.Image")));
            this.btnMapLegendRiskRankingClose.Location = new System.Drawing.Point(2, 3);
            this.btnMapLegendRiskRankingClose.Name = "btnMapLegendRiskRankingClose";
            this.btnMapLegendRiskRankingClose.Size = new System.Drawing.Size(15, 15);
            this.btnMapLegendRiskRankingClose.TabIndex = 20;
            this.btnMapLegendRiskRankingClose.UseVisualStyleBackColor = false;
            this.btnMapLegendRiskRankingClose.Click += new System.EventHandler(this.btnMapLegendRiskRankingClose_Click);
            // 
            // rbMapLegendDisplayLSByRisk
            // 
            this.rbMapLegendDisplayLSByRisk.AutoSize = true;
            this.rbMapLegendDisplayLSByRisk.Location = new System.Drawing.Point(514, 44);
            this.rbMapLegendDisplayLSByRisk.Name = "rbMapLegendDisplayLSByRisk";
            this.rbMapLegendDisplayLSByRisk.Size = new System.Drawing.Size(104, 17);
            this.rbMapLegendDisplayLSByRisk.TabIndex = 19;
            this.rbMapLegendDisplayLSByRisk.Text = "By Risk Ranking";
            this.rbMapLegendDisplayLSByRisk.UseVisualStyleBackColor = true;
            this.rbMapLegendDisplayLSByRisk.CheckedChanged += new System.EventHandler(this.rbMapLegendDisplayLSByRisk_CheckedChanged);
            // 
            // rbMapLegendDisplayLSByLayer
            // 
            this.rbMapLegendDisplayLSByLayer.AutoSize = true;
            this.rbMapLegendDisplayLSByLayer.Checked = true;
            this.rbMapLegendDisplayLSByLayer.Location = new System.Drawing.Point(514, 25);
            this.rbMapLegendDisplayLSByLayer.Name = "rbMapLegendDisplayLSByLayer";
            this.rbMapLegendDisplayLSByLayer.Size = new System.Drawing.Size(66, 17);
            this.rbMapLegendDisplayLSByLayer.TabIndex = 18;
            this.rbMapLegendDisplayLSByLayer.TabStop = true;
            this.rbMapLegendDisplayLSByLayer.Text = "By Layer";
            this.rbMapLegendDisplayLSByLayer.UseVisualStyleBackColor = true;
            this.rbMapLegendDisplayLSByLayer.CheckedChanged += new System.EventHandler(this.rbMapLegendDisplayLSByLayer_CheckedChanged);
            // 
            // lblMapLegendRiskRankingBanner
            // 
            this.lblMapLegendRiskRankingBanner.BackColor = System.Drawing.Color.Silver;
            this.lblMapLegendRiskRankingBanner.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblMapLegendRiskRankingBanner.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblMapLegendRiskRankingBanner.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMapLegendRiskRankingBanner.Location = new System.Drawing.Point(0, 0);
            this.lblMapLegendRiskRankingBanner.MinimumSize = new System.Drawing.Size(2, 20);
            this.lblMapLegendRiskRankingBanner.Name = "lblMapLegendRiskRankingBanner";
            this.lblMapLegendRiskRankingBanner.Size = new System.Drawing.Size(781, 20);
            this.lblMapLegendRiskRankingBanner.TabIndex = 17;
            this.lblMapLegendRiskRankingBanner.Text = "Display Line Sections by Risk Ranking";
            this.lblMapLegendRiskRankingBanner.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cbMapLegendRiskFieldToUse
            // 
            this.cbMapLegendRiskFieldToUse.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbMapLegendRiskFieldToUse.FormattingEnabled = true;
            this.cbMapLegendRiskFieldToUse.Items.AddRange(new object[] {
            "Overall Ranking",
            "Susceptibility Ranking",
            "Consequence Ranking",
            "ALL"});
            this.cbMapLegendRiskFieldToUse.Location = new System.Drawing.Point(621, 43);
            this.cbMapLegendRiskFieldToUse.Name = "cbMapLegendRiskFieldToUse";
            this.cbMapLegendRiskFieldToUse.Size = new System.Drawing.Size(159, 21);
            this.cbMapLegendRiskFieldToUse.TabIndex = 16;
            this.cbMapLegendRiskFieldToUse.SelectedIndexChanged += new System.EventHandler(this.cbMapLegendRiskFieldToUse_SelectedIndexChanged);
            // 
            // gridRiskRankingLegend
            // 
            this.gridRiskRankingLegend.BackgroundColor = System.Drawing.SystemColors.Control;
            this.gridRiskRankingLegend.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.gridRiskRankingLegend.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.gridRiskRankingLegend.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle8;
            this.gridRiskRankingLegend.ColumnHeadersHeight = 21;
            this.gridRiskRankingLegend.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.gridRiskRankingLegend.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColLow,
            this.ColLowMed,
            this.ColMed,
            this.ColMedHigh,
            this.ColHigh});
            this.gridRiskRankingLegend.GridColor = System.Drawing.SystemColors.Control;
            this.gridRiskRankingLegend.Location = new System.Drawing.Point(3, 28);
            this.gridRiskRankingLegend.MultiSelect = false;
            this.gridRiskRankingLegend.Name = "gridRiskRankingLegend";
            this.gridRiskRankingLegend.ReadOnly = true;
            this.gridRiskRankingLegend.RowHeadersVisible = false;
            this.gridRiskRankingLegend.RowHeadersWidth = 120;
            this.gridRiskRankingLegend.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.gridRiskRankingLegend.RowsDefaultCellStyle = dataGridViewCellStyle14;
            this.gridRiskRankingLegend.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridRiskRankingLegend.RowTemplate.DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            this.gridRiskRankingLegend.RowTemplate.Height = 10;
            this.gridRiskRankingLegend.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.gridRiskRankingLegend.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.gridRiskRankingLegend.ShowRowErrors = false;
            this.gridRiskRankingLegend.Size = new System.Drawing.Size(503, 34);
            this.gridRiskRankingLegend.TabIndex = 0;
            // 
            // ColLow
            // 
            dataGridViewCellStyle9.BackColor = System.Drawing.Color.Lime;
            this.ColLow.DefaultCellStyle = dataGridViewCellStyle9;
            this.ColLow.HeaderText = "Low";
            this.ColLow.Name = "ColLow";
            this.ColLow.ReadOnly = true;
            // 
            // ColLowMed
            // 
            dataGridViewCellStyle10.BackColor = System.Drawing.Color.Yellow;
            this.ColLowMed.DefaultCellStyle = dataGridViewCellStyle10;
            this.ColLowMed.HeaderText = "Low-Medium";
            this.ColLowMed.Name = "ColLowMed";
            this.ColLowMed.ReadOnly = true;
            // 
            // ColMed
            // 
            dataGridViewCellStyle11.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.ColMed.DefaultCellStyle = dataGridViewCellStyle11;
            this.ColMed.HeaderText = "Medium";
            this.ColMed.Name = "ColMed";
            this.ColMed.ReadOnly = true;
            // 
            // ColMedHigh
            // 
            dataGridViewCellStyle12.BackColor = System.Drawing.Color.Fuchsia;
            this.ColMedHigh.DefaultCellStyle = dataGridViewCellStyle12;
            this.ColMedHigh.HeaderText = "Medium-High";
            this.ColMedHigh.Name = "ColMedHigh";
            this.ColMedHigh.ReadOnly = true;
            // 
            // ColHigh
            // 
            dataGridViewCellStyle13.BackColor = System.Drawing.Color.Red;
            this.ColHigh.DefaultCellStyle = dataGridViewCellStyle13;
            this.ColHigh.HeaderText = "High";
            this.ColHigh.Name = "ColHigh";
            this.ColHigh.ReadOnly = true;
            // 
            // scSelections
            // 
            this.scSelections.Dock = System.Windows.Forms.DockStyle.Fill;
            this.scSelections.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.scSelections.IsSplitterFixed = true;
            this.scSelections.Location = new System.Drawing.Point(0, 0);
            this.scSelections.MinimumSize = new System.Drawing.Size(285, 0);
            this.scSelections.Name = "scSelections";
            this.scSelections.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // scSelections.Panel1
            // 
            this.scSelections.Panel1.Controls.Add(this.btnExportTreeViewToExcel);
            this.scSelections.Panel1.Controls.Add(this.btnExpandAll);
            this.scSelections.Panel1.Controls.Add(this.btnCollapseAll);
            this.scSelections.Panel1.Controls.Add(this.treeViewMain);
            this.scSelections.Panel1.Controls.Add(this.lblSelectedFeaturesTitle);
            this.scSelections.Panel1.Padding = new System.Windows.Forms.Padding(4);
            // 
            // scSelections.Panel2
            // 
            this.scSelections.Panel2.BackColor = System.Drawing.SystemColors.Control;
            this.scSelections.Panel2.Controls.Add(this.pnlCommands);
            this.scSelections.Panel2.Padding = new System.Windows.Forms.Padding(4);
            this.scSelections.Size = new System.Drawing.Size(285, 785);
            this.scSelections.SplitterDistance = 632;
            this.scSelections.TabIndex = 0;
            // 
            // btnExportTreeViewToExcel
            // 
            this.btnExportTreeViewToExcel.Image = global::COMPGIS1.Properties.Resources.xl;
            this.btnExportTreeViewToExcel.Location = new System.Drawing.Point(5, 5);
            this.btnExportTreeViewToExcel.Name = "btnExportTreeViewToExcel";
            this.btnExportTreeViewToExcel.Size = new System.Drawing.Size(29, 29);
            this.btnExportTreeViewToExcel.TabIndex = 24;
            this.btnExportTreeViewToExcel.UseVisualStyleBackColor = true;
            this.btnExportTreeViewToExcel.Click += new System.EventHandler(this.btnExportTreeViewToExcel_Click);
            // 
            // btnExpandAll
            // 
            this.btnExpandAll.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.btnExpandAll.FlatAppearance.BorderSize = 0;
            this.btnExpandAll.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExpandAll.Image = global::COMPGIS1.Properties.Resources.expand_small_blue_Shapes4FREE;
            this.btnExpandAll.Location = new System.Drawing.Point(245, 5);
            this.btnExpandAll.Name = "btnExpandAll";
            this.btnExpandAll.Size = new System.Drawing.Size(15, 15);
            this.btnExpandAll.TabIndex = 13;
            this.btnExpandAll.UseVisualStyleBackColor = false;
            this.btnExpandAll.Click += new System.EventHandler(this.mnuExpandAll_Click);
            // 
            // btnCollapseAll
            // 
            this.btnCollapseAll.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.btnCollapseAll.FlatAppearance.BorderSize = 0;
            this.btnCollapseAll.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCollapseAll.Image = global::COMPGIS1.Properties.Resources.collapse_small_blue_Shapes4FREE;
            this.btnCollapseAll.Location = new System.Drawing.Point(263, 5);
            this.btnCollapseAll.Name = "btnCollapseAll";
            this.btnCollapseAll.Size = new System.Drawing.Size(15, 15);
            this.btnCollapseAll.TabIndex = 12;
            this.btnCollapseAll.UseVisualStyleBackColor = false;
            this.btnCollapseAll.Click += new System.EventHandler(this.mnuCollapseAll_Click);
            // 
            // treeViewMain
            // 
            this.treeViewMain.ContextMenuStrip = this.cmsTreeViewGeneral;
            this.treeViewMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewMain.Location = new System.Drawing.Point(4, 35);
            this.treeViewMain.Name = "treeViewMain";
            this.treeViewMain.ShowNodeToolTips = true;
            this.treeViewMain.Size = new System.Drawing.Size(277, 593);
            this.treeViewMain.TabIndex = 11;
            this.treeViewMain.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeViewMain_AfterSelect);
            this.treeViewMain.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeViewMain_NodeMouseDoubleClick);
            // 
            // cmsTreeViewGeneral
            // 
            this.cmsTreeViewGeneral.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuExpandAll,
            this.mnuCollapseAll});
            this.cmsTreeViewGeneral.Name = "cmsTreeViewGeneral";
            this.cmsTreeViewGeneral.Size = new System.Drawing.Size(137, 48);
            // 
            // mnuExpandAll
            // 
            this.mnuExpandAll.Name = "mnuExpandAll";
            this.mnuExpandAll.Size = new System.Drawing.Size(136, 22);
            this.mnuExpandAll.Text = "Expand All";
            this.mnuExpandAll.Click += new System.EventHandler(this.mnuExpandAll_Click);
            // 
            // mnuCollapseAll
            // 
            this.mnuCollapseAll.Name = "mnuCollapseAll";
            this.mnuCollapseAll.Size = new System.Drawing.Size(136, 22);
            this.mnuCollapseAll.Text = "Collapse All";
            this.mnuCollapseAll.Click += new System.EventHandler(this.mnuCollapseAll_Click);
            // 
            // lblSelectedFeaturesTitle
            // 
            this.lblSelectedFeaturesTitle.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.lblSelectedFeaturesTitle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSelectedFeaturesTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblSelectedFeaturesTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSelectedFeaturesTitle.Location = new System.Drawing.Point(4, 4);
            this.lblSelectedFeaturesTitle.MinimumSize = new System.Drawing.Size(2, 20);
            this.lblSelectedFeaturesTitle.Name = "lblSelectedFeaturesTitle";
            this.lblSelectedFeaturesTitle.Size = new System.Drawing.Size(277, 31);
            this.lblSelectedFeaturesTitle.TabIndex = 10;
            this.lblSelectedFeaturesTitle.Text = "Selected Features";
            this.lblSelectedFeaturesTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pnlCommands
            // 
            this.pnlCommands.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlCommands.Controls.Add(this.tcCommandModes);
            this.pnlCommands.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlCommands.Location = new System.Drawing.Point(4, 4);
            this.pnlCommands.Name = "pnlCommands";
            this.pnlCommands.Size = new System.Drawing.Size(277, 141);
            this.pnlCommands.TabIndex = 1;
            // 
            // tcCommandModes
            // 
            this.tcCommandModes.Appearance = System.Windows.Forms.TabAppearance.Buttons;
            this.tcCommandModes.Controls.Add(this.tabPrebuiltSearch);
            this.tcCommandModes.Controls.Add(this.tabBE_LineSection_Association);
            this.tcCommandModes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tcCommandModes.ItemSize = new System.Drawing.Size(0, 1);
            this.tcCommandModes.Location = new System.Drawing.Point(0, 0);
            this.tcCommandModes.Multiline = true;
            this.tcCommandModes.Name = "tcCommandModes";
            this.tcCommandModes.Padding = new System.Drawing.Point(0, 0);
            this.tcCommandModes.SelectedIndex = 0;
            this.tcCommandModes.Size = new System.Drawing.Size(275, 139);
            this.tcCommandModes.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tcCommandModes.TabIndex = 1;
            // 
            // tabPrebuiltSearch
            // 
            this.tabPrebuiltSearch.Location = new System.Drawing.Point(4, 5);
            this.tabPrebuiltSearch.Name = "tabPrebuiltSearch";
            this.tabPrebuiltSearch.Size = new System.Drawing.Size(267, 130);
            this.tabPrebuiltSearch.TabIndex = 2;
            this.tabPrebuiltSearch.Text = "-";
            this.tabPrebuiltSearch.UseVisualStyleBackColor = true;
            // 
            // tabBE_LineSection_Association
            // 
            this.tabBE_LineSection_Association.Controls.Add(this.tbRadius);
            this.tabBE_LineSection_Association.Controls.Add(this.Reset);
            this.tabBE_LineSection_Association.Controls.Add(this.Cancel);
            this.tabBE_LineSection_Association.Controls.Add(this.btnLineSectionsQuery);
            this.tabBE_LineSection_Association.Controls.Add(this.label1);
            this.tabBE_LineSection_Association.Location = new System.Drawing.Point(4, 5);
            this.tabBE_LineSection_Association.Name = "tabBE_LineSection_Association";
            this.tabBE_LineSection_Association.Padding = new System.Windows.Forms.Padding(3);
            this.tabBE_LineSection_Association.Size = new System.Drawing.Size(267, 130);
            this.tabBE_LineSection_Association.TabIndex = 0;
            this.tabBE_LineSection_Association.Text = "SpatialQuery";
            this.tabBE_LineSection_Association.UseVisualStyleBackColor = true;
            // 
            // tbRadius
            // 
            this.tbRadius.Location = new System.Drawing.Point(136, 10);
            this.tbRadius.Name = "tbRadius";
            this.tbRadius.Size = new System.Drawing.Size(37, 20);
            this.tbRadius.TabIndex = 11;
            this.tbRadius.Text = "5";
            // 
            // Reset
            // 
            this.Reset.Location = new System.Drawing.Point(185, 6);
            this.Reset.Name = "Reset";
            this.Reset.Size = new System.Drawing.Size(83, 27);
            this.Reset.TabIndex = 10;
            this.Reset.Text = "Reset Query";
            this.Reset.UseVisualStyleBackColor = true;
            this.Reset.Click += new System.EventHandler(this.Reset_Click);
            // 
            // Cancel
            // 
            this.Cancel.Location = new System.Drawing.Point(184, 88);
            this.Cancel.Name = "Cancel";
            this.Cancel.Size = new System.Drawing.Size(83, 26);
            this.Cancel.TabIndex = 14;
            this.Cancel.Text = "Cancel";
            this.Cancel.UseVisualStyleBackColor = true;
            this.Cancel.Click += new System.EventHandler(this.AssociateLineSectionsToBoringLocationsCancel_Click);
            // 
            // btnLineSectionsQuery
            // 
            this.btnLineSectionsQuery.Location = new System.Drawing.Point(91, 88);
            this.btnLineSectionsQuery.Name = "btnLineSectionsQuery";
            this.btnLineSectionsQuery.Size = new System.Drawing.Size(87, 26);
            this.btnLineSectionsQuery.TabIndex = 13;
            this.btnLineSectionsQuery.Text = "OK";
            this.btnLineSectionsQuery.UseVisualStyleBackColor = true;
            this.btnLineSectionsQuery.Click += new System.EventHandler(this.AssociateLineSectionsToBoringLocations_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(75, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "Radius (ft)";
            // 
            // cmsBPNodeLineSection
            // 
            this.cmsBPNodeLineSection.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuLineSectionProperties,
            this.inspectionsToolStripMenuItem});
            this.cmsBPNodeLineSection.Name = "cmRootLineSections";
            this.cmsBPNodeLineSection.Size = new System.Drawing.Size(144, 48);
            // 
            // mnuLineSectionProperties
            // 
            this.mnuLineSectionProperties.Name = "mnuLineSectionProperties";
            this.mnuLineSectionProperties.Size = new System.Drawing.Size(143, 22);
            this.mnuLineSectionProperties.Text = "Properties...";
            this.mnuLineSectionProperties.Click += new System.EventHandler(this.mnuLineSectionProperties_Click);
            // 
            // inspectionsToolStripMenuItem
            // 
            this.inspectionsToolStripMenuItem.Name = "inspectionsToolStripMenuItem";
            this.inspectionsToolStripMenuItem.Size = new System.Drawing.Size(143, 22);
            this.inspectionsToolStripMenuItem.Text = "Inspections...";
            this.inspectionsToolStripMenuItem.Click += new System.EventHandler(this.mnuLineSectionInspections);
            // 
            // cmsBPRootLineSections
            // 
            this.cmsBPRootLineSections.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuFilterLineSections,
            this.mnuCustomFilterSeparatorLineSections,
            this.lineSectionsAssociatedWithLCSsToolStripMenuItem,
            this.lineSectionsThatHaveBeenInspectedToolStripMenuItem,
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem,
            this.lineSectionsWithAHistoryOfLeaksToolStripMenuItem,
            this.lineSectionsWithGuidedWaveInspectionsByLevelToolStripMenuItem});
            this.cmsBPRootLineSections.Name = "cmRootLineSections";
            this.cmsBPRootLineSections.Size = new System.Drawing.Size(402, 142);
            this.cmsBPRootLineSections.Opening += new System.ComponentModel.CancelEventHandler(this.cmsRootLineSections_Opening);
            // 
            // mnuFilterLineSections
            // 
            this.mnuFilterLineSections.Name = "mnuFilterLineSections";
            this.mnuFilterLineSections.Size = new System.Drawing.Size(401, 22);
            this.mnuFilterLineSections.Text = "Filters";
            // 
            // mnuCustomFilterSeparatorLineSections
            // 
            this.mnuCustomFilterSeparatorLineSections.Name = "mnuCustomFilterSeparatorLineSections";
            this.mnuCustomFilterSeparatorLineSections.Size = new System.Drawing.Size(398, 6);
            // 
            // lineSectionsAssociatedWithLCSsToolStripMenuItem
            // 
            this.lineSectionsAssociatedWithLCSsToolStripMenuItem.Name = "lineSectionsAssociatedWithLCSsToolStripMenuItem";
            this.lineSectionsAssociatedWithLCSsToolStripMenuItem.Size = new System.Drawing.Size(401, 22);
            this.lineSectionsAssociatedWithLCSsToolStripMenuItem.Text = "Line Sections Associated with LCO\'s";
            // 
            // lineSectionsThatHaveBeenInspectedToolStripMenuItem
            // 
            this.lineSectionsThatHaveBeenInspectedToolStripMenuItem.Name = "lineSectionsThatHaveBeenInspectedToolStripMenuItem";
            this.lineSectionsThatHaveBeenInspectedToolStripMenuItem.Size = new System.Drawing.Size(401, 22);
            this.lineSectionsThatHaveBeenInspectedToolStripMenuItem.Text = "Line Sections that have been inspected...";
            this.lineSectionsThatHaveBeenInspectedToolStripMenuItem.Click += new System.EventHandler(this.lineSectionsThatHaveBeenInspectedToolStripMenuItem_Click);
            // 
            // lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem
            // 
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem.Enabled = false;
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem.Name = "lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem";
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem.Size = new System.Drawing.Size(401, 22);
            this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem.Text = "Line Sections Containing Licensed Material (radioactive fluids)";
            // 
            // lineSectionsWithAHistoryOfLeaksToolStripMenuItem
            // 
            this.lineSectionsWithAHistoryOfLeaksToolStripMenuItem.Enabled = false;
            this.lineSectionsWithAHistoryOfLeaksToolStripMenuItem.Name = "lineSectionsWithAHistoryOfLeaksToolStripMenuItem";
            this.lineSectionsWithAHistoryOfLeaksToolStripMenuItem.Size = new System.Drawing.Size(401, 22);
            this.lineSectionsWithAHistoryOfLeaksToolStripMenuItem.Text = "Line Sections with a history of leaks";
            // 
            // lineSectionsWithGuidedWaveInspectionsByLevelToolStripMenuItem
            // 
            this.lineSectionsWithGuidedWaveInspectionsByLevelToolStripMenuItem.Name = "lineSectionsWithGuidedWaveInspectionsByLevelToolStripMenuItem";
            this.lineSectionsWithGuidedWaveInspectionsByLevelToolStripMenuItem.Size = new System.Drawing.Size(401, 22);
            this.lineSectionsWithGuidedWaveInspectionsByLevelToolStripMenuItem.Text = "Line Sections with Guided Wave Inspections by Level...";
            this.lineSectionsWithGuidedWaveInspectionsByLevelToolStripMenuItem.Click += new System.EventHandler(this.lineSectionsWithGuidedWaveInspectionsByLevelToolStripMenuItem_Click);
            // 
            // ilTreeViewMain
            // 
            this.ilTreeViewMain.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("ilTreeViewMain.ImageStream")));
            this.ilTreeViewMain.TransparentColor = System.Drawing.Color.Transparent;
            this.ilTreeViewMain.Images.SetKeyName(0, "ORB_Icons_by_025.png");
            this.ilTreeViewMain.Images.SetKeyName(1, "ORB_Icons_by_026.png");
            this.ilTreeViewMain.Images.SetKeyName(2, "ORB_Icons_by_027.png");
            // 
            // cmsBPRootLines
            // 
            this.cmsBPRootLines.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuFilterLines,
            this.mnuCustomFilterSeparatorLines,
            this.selectLinesInASystemToolStripMenuItem,
            this.inspectedLinesToolStripMenuItem,
            this.linesContainingLicensedMaterialradioactiveFluidsToolStripMenuItem});
            this.cmsBPRootLines.Name = "cmRootLineSections";
            this.cmsBPRootLines.Size = new System.Drawing.Size(360, 98);
            this.cmsBPRootLines.Opening += new System.ComponentModel.CancelEventHandler(this.cmsRootLines_Opening);
            // 
            // mnuFilterLines
            // 
            this.mnuFilterLines.Name = "mnuFilterLines";
            this.mnuFilterLines.Size = new System.Drawing.Size(359, 22);
            this.mnuFilterLines.Text = "Filters";
            // 
            // mnuCustomFilterSeparatorLines
            // 
            this.mnuCustomFilterSeparatorLines.Name = "mnuCustomFilterSeparatorLines";
            this.mnuCustomFilterSeparatorLines.Size = new System.Drawing.Size(356, 6);
            // 
            // selectLinesInASystemToolStripMenuItem
            // 
            this.selectLinesInASystemToolStripMenuItem.Enabled = false;
            this.selectLinesInASystemToolStripMenuItem.Name = "selectLinesInASystemToolStripMenuItem";
            this.selectLinesInASystemToolStripMenuItem.Size = new System.Drawing.Size(359, 22);
            this.selectLinesInASystemToolStripMenuItem.Text = "Lines in a System...";
            this.selectLinesInASystemToolStripMenuItem.Visible = false;
            // 
            // inspectedLinesToolStripMenuItem
            // 
            this.inspectedLinesToolStripMenuItem.Enabled = false;
            this.inspectedLinesToolStripMenuItem.Name = "inspectedLinesToolStripMenuItem";
            this.inspectedLinesToolStripMenuItem.Size = new System.Drawing.Size(359, 22);
            this.inspectedLinesToolStripMenuItem.Text = "Inspected Lines...";
            this.inspectedLinesToolStripMenuItem.Visible = false;
            // 
            // linesContainingLicensedMaterialradioactiveFluidsToolStripMenuItem
            // 
            this.linesContainingLicensedMaterialradioactiveFluidsToolStripMenuItem.Name = "linesContainingLicensedMaterialradioactiveFluidsToolStripMenuItem";
            this.linesContainingLicensedMaterialradioactiveFluidsToolStripMenuItem.Size = new System.Drawing.Size(359, 22);
            this.linesContainingLicensedMaterialradioactiveFluidsToolStripMenuItem.Text = "Lines Containing Licensed Material (radioactive fluids)";
            this.linesContainingLicensedMaterialradioactiveFluidsToolStripMenuItem.Click += new System.EventHandler(this.lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem1_Click);
            // 
            // cmsBPRootBoringLocations
            // 
            this.cmsBPRootBoringLocations.Name = "cmRootLineSections";
            this.cmsBPRootBoringLocations.Size = new System.Drawing.Size(61, 4);
            this.cmsBPRootBoringLocations.Opening += new System.ComponentModel.CancelEventHandler(this.cmsRootBoringLocations_Opening);
            // 
            // cmsBPRootCorrosionProbes
            // 
            this.cmsBPRootCorrosionProbes.Name = "cmRootLineSections";
            this.cmsBPRootCorrosionProbes.Size = new System.Drawing.Size(61, 4);
            this.cmsBPRootCorrosionProbes.Opening += new System.ComponentModel.CancelEventHandler(this.cmsRootCorrosionProbes_Opening);
            // 
            // cmsBPNodeLines
            // 
            this.cmsBPNodeLines.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuLineProperties,
            this.toolStripMenuItem1});
            this.cmsBPNodeLines.Name = "cmRootLineSections";
            this.cmsBPNodeLines.Size = new System.Drawing.Size(137, 32);
            // 
            // mnuLineProperties
            // 
            this.mnuLineProperties.Name = "mnuLineProperties";
            this.mnuLineProperties.Size = new System.Drawing.Size(136, 22);
            this.mnuLineProperties.Text = "Properties...";
            this.mnuLineProperties.Click += new System.EventHandler(this.mnuLineProperties_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(133, 6);
            // 
            // cmsBPNodeBoringLocations
            // 
            this.cmsBPNodeBoringLocations.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuBoringLocationProperties,
            this.spatialQueryToolStripMenuItem,
            this.associatedLineSectionsToolStripMenuItem});
            this.cmsBPNodeBoringLocations.Name = "cmRootLineSections";
            this.cmsBPNodeBoringLocations.Size = new System.Drawing.Size(213, 70);
            // 
            // mnuBoringLocationProperties
            // 
            this.mnuBoringLocationProperties.Name = "mnuBoringLocationProperties";
            this.mnuBoringLocationProperties.Size = new System.Drawing.Size(212, 22);
            this.mnuBoringLocationProperties.Text = "Properties...";
            this.mnuBoringLocationProperties.Click += new System.EventHandler(this.mnuBoringLocationProperties_Click);
            // 
            // spatialQueryToolStripMenuItem
            // 
            this.spatialQueryToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.linesNearbyToolStripMenuItem,
            this.lineSectionsNearbyToolStripMenuItem});
            this.spatialQueryToolStripMenuItem.Enabled = false;
            this.spatialQueryToolStripMenuItem.Name = "spatialQueryToolStripMenuItem";
            this.spatialQueryToolStripMenuItem.Size = new System.Drawing.Size(212, 22);
            this.spatialQueryToolStripMenuItem.Text = "Spatial Query";
            // 
            // linesNearbyToolStripMenuItem
            // 
            this.linesNearbyToolStripMenuItem.Name = "linesNearbyToolStripMenuItem";
            this.linesNearbyToolStripMenuItem.Size = new System.Drawing.Size(193, 22);
            this.linesNearbyToolStripMenuItem.Text = "Lines Nearby...";
            // 
            // lineSectionsNearbyToolStripMenuItem
            // 
            this.lineSectionsNearbyToolStripMenuItem.Name = "lineSectionsNearbyToolStripMenuItem";
            this.lineSectionsNearbyToolStripMenuItem.Size = new System.Drawing.Size(193, 22);
            this.lineSectionsNearbyToolStripMenuItem.Text = "Line Sections Nearby...";
            this.lineSectionsNearbyToolStripMenuItem.Click += new System.EventHandler(this.lineSectionsNearbyToolStripMenuItem_Click);
            // 
            // associatedLineSectionsToolStripMenuItem
            // 
            this.associatedLineSectionsToolStripMenuItem.Enabled = false;
            this.associatedLineSectionsToolStripMenuItem.Name = "associatedLineSectionsToolStripMenuItem";
            this.associatedLineSectionsToolStripMenuItem.Size = new System.Drawing.Size(212, 22);
            this.associatedLineSectionsToolStripMenuItem.Text = "Associated Line Sections...";
            // 
            // cmsBPNodeCorrosionProbes
            // 
            this.cmsBPNodeCorrosionProbes.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuCPProperties,
            this.spatialQueryToolStripMenuItem1});
            this.cmsBPNodeCorrosionProbes.Name = "cmRootLineSections";
            this.cmsBPNodeCorrosionProbes.Size = new System.Drawing.Size(145, 48);
            // 
            // mnuCPProperties
            // 
            this.mnuCPProperties.Name = "mnuCPProperties";
            this.mnuCPProperties.Size = new System.Drawing.Size(144, 22);
            this.mnuCPProperties.Text = "Properties...";
            this.mnuCPProperties.Click += new System.EventHandler(this.mnuCPProperties_Click);
            // 
            // spatialQueryToolStripMenuItem1
            // 
            this.spatialQueryToolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.linesNearbyToolStripMenuItem1,
            this.lineSectionsNearbyToolStripMenuItem1});
            this.spatialQueryToolStripMenuItem1.Name = "spatialQueryToolStripMenuItem1";
            this.spatialQueryToolStripMenuItem1.Size = new System.Drawing.Size(144, 22);
            this.spatialQueryToolStripMenuItem1.Text = "Spatial Query";
            // 
            // linesNearbyToolStripMenuItem1
            // 
            this.linesNearbyToolStripMenuItem1.Name = "linesNearbyToolStripMenuItem1";
            this.linesNearbyToolStripMenuItem1.Size = new System.Drawing.Size(193, 22);
            this.linesNearbyToolStripMenuItem1.Text = "Lines Nearby...";
            // 
            // lineSectionsNearbyToolStripMenuItem1
            // 
            this.lineSectionsNearbyToolStripMenuItem1.Name = "lineSectionsNearbyToolStripMenuItem1";
            this.lineSectionsNearbyToolStripMenuItem1.Size = new System.Drawing.Size(193, 22);
            this.lineSectionsNearbyToolStripMenuItem1.Text = "Line Sections Nearby...";
            // 
            // tooltipMain
            // 
            this.tooltipMain.ShowAlways = true;
            // 
            // fileSystemWatcher1
            // 
            this.fileSystemWatcher1.EnableRaisingEvents = true;
            this.fileSystemWatcher1.SynchronizingObject = this;
            // 
            // cmsSWRootLines
            // 
            this.cmsSWRootLines.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem5,
            this.toolStripSeparator1,
            this.toolStripMenuItem6,
            this.toolStripMenuItem7});
            this.cmsSWRootLines.Name = "cmRootLineSections";
            this.cmsSWRootLines.Size = new System.Drawing.Size(174, 76);
            // 
            // toolStripMenuItem5
            // 
            this.toolStripMenuItem5.Name = "toolStripMenuItem5";
            this.toolStripMenuItem5.Size = new System.Drawing.Size(173, 22);
            this.toolStripMenuItem5.Text = "Filters";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(170, 6);
            // 
            // toolStripMenuItem6
            // 
            this.toolStripMenuItem6.Enabled = false;
            this.toolStripMenuItem6.Name = "toolStripMenuItem6";
            this.toolStripMenuItem6.Size = new System.Drawing.Size(173, 22);
            this.toolStripMenuItem6.Text = "Lines in a System...";
            this.toolStripMenuItem6.Visible = false;
            // 
            // toolStripMenuItem7
            // 
            this.toolStripMenuItem7.Enabled = false;
            this.toolStripMenuItem7.Name = "toolStripMenuItem7";
            this.toolStripMenuItem7.Size = new System.Drawing.Size(173, 22);
            this.toolStripMenuItem7.Text = "Inspected Lines...";
            this.toolStripMenuItem7.Visible = false;
            // 
            // cmsSWNodeLines
            // 
            this.cmsSWNodeLines.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem8,
            this.toolStripSeparator2});
            this.cmsSWNodeLines.Name = "cmRootLineSections";
            this.cmsSWNodeLines.Size = new System.Drawing.Size(137, 32);
            // 
            // toolStripMenuItem8
            // 
            this.toolStripMenuItem8.Name = "toolStripMenuItem8";
            this.toolStripMenuItem8.Size = new System.Drawing.Size(136, 22);
            this.toolStripMenuItem8.Text = "Properties...";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(133, 6);
            // 
            // cmsSWRootFlowSegments
            // 
            this.cmsSWRootFlowSegments.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem9,
            this.toolStripSeparator3});
            this.cmsSWRootFlowSegments.Name = "cmRootLineSections";
            this.cmsSWRootFlowSegments.Size = new System.Drawing.Size(106, 32);
            // 
            // toolStripMenuItem9
            // 
            this.toolStripMenuItem9.Name = "toolStripMenuItem9";
            this.toolStripMenuItem9.Size = new System.Drawing.Size(105, 22);
            this.toolStripMenuItem9.Text = "Filters";
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(102, 6);
            // 
            // cmsSWNodeFlowSegments
            // 
            this.cmsSWNodeFlowSegments.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem10,
            this.toolStripMenuItem11});
            this.cmsSWNodeFlowSegments.Name = "cmRootLineSections";
            this.cmsSWNodeFlowSegments.Size = new System.Drawing.Size(144, 48);
            // 
            // toolStripMenuItem10
            // 
            this.toolStripMenuItem10.Name = "toolStripMenuItem10";
            this.toolStripMenuItem10.Size = new System.Drawing.Size(143, 22);
            this.toolStripMenuItem10.Text = "Properties...";
            // 
            // toolStripMenuItem11
            // 
            this.toolStripMenuItem11.Name = "toolStripMenuItem11";
            this.toolStripMenuItem11.Size = new System.Drawing.Size(143, 22);
            this.toolStripMenuItem11.Text = "Inspections...";
            // 
            // cmsSWRootComponents
            // 
            this.cmsSWRootComponents.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem12,
            this.toolStripSeparator4});
            this.cmsSWRootComponents.Name = "cmRootLineSections";
            this.cmsSWRootComponents.Size = new System.Drawing.Size(153, 54);
            // 
            // toolStripMenuItem12
            // 
            this.toolStripMenuItem12.Name = "toolStripMenuItem12";
            this.toolStripMenuItem12.Size = new System.Drawing.Size(152, 22);
            this.toolStripMenuItem12.Text = "Filters";
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(149, 6);
            // 
            // cmsSWNodeComponents
            // 
            this.cmsSWNodeComponents.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem13,
            this.toolStripMenuItem14});
            this.cmsSWNodeComponents.Name = "cmRootLineSections";
            this.cmsSWNodeComponents.Size = new System.Drawing.Size(144, 48);
            // 
            // toolStripMenuItem13
            // 
            this.toolStripMenuItem13.Name = "toolStripMenuItem13";
            this.toolStripMenuItem13.Size = new System.Drawing.Size(143, 22);
            this.toolStripMenuItem13.Text = "Properties...";
            // 
            // toolStripMenuItem14
            // 
            this.toolStripMenuItem14.Name = "toolStripMenuItem14";
            this.toolStripMenuItem14.Size = new System.Drawing.Size(143, 22);
            this.toolStripMenuItem14.Text = "Inspections...";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1336, 843);
            this.Controls.Add(this.scMain);
            this.Controls.Add(this.axLicenseControl1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.splitter1);
            this.Controls.Add(this.axToolbarControl1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CompWorks Buried Piping [Select Line Sections]";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.axMapControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.axToolbarControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.axTOCControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.axLicenseControl1)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.scMain.Panel1.ResumeLayout(false);
            this.scMain.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.scMain)).EndInit();
            this.scMain.ResumeLayout(false);
            this.tcTOC.ResumeLayout(false);
            this.tabTOC.ResumeLayout(false);
            this.tabFilter.ResumeLayout(false);
            this.tabFilter.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.tabFind.ResumeLayout(false);
            this.pnlFindBanner.ResumeLayout(false);
            this.pnlFindBG.ResumeLayout(false);
            this.pnlFindDetails.ResumeLayout(false);
            this.tabDisplay.ResumeLayout(false);
            this.tabDisplay.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.axSymbologyControl1)).EndInit();
            this.scMap.Panel1.ResumeLayout(false);
            this.scMap.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.scMap)).EndInit();
            this.scMap.ResumeLayout(false);
            this.scMapLegend.Panel1.ResumeLayout(false);
            this.scMapLegend.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.scMapLegend)).EndInit();
            this.scMapLegend.ResumeLayout(false);
            this.tcMapLegend.ResumeLayout(false);
            this.tabMapLegendRiskRanking.ResumeLayout(false);
            this.tabMapLegendRiskRanking.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridRiskRankingLegend)).EndInit();
            this.scSelections.Panel1.ResumeLayout(false);
            this.scSelections.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.scSelections)).EndInit();
            this.scSelections.ResumeLayout(false);
            this.cmsTreeViewGeneral.ResumeLayout(false);
            this.pnlCommands.ResumeLayout(false);
            this.tcCommandModes.ResumeLayout(false);
            this.tabBE_LineSection_Association.ResumeLayout(false);
            this.tabBE_LineSection_Association.PerformLayout();
            this.cmsBPNodeLineSection.ResumeLayout(false);
            this.cmsBPRootLineSections.ResumeLayout(false);
            this.cmsBPRootLines.ResumeLayout(false);
            this.cmsBPNodeLines.ResumeLayout(false);
            this.cmsBPNodeBoringLocations.ResumeLayout(false);
            this.cmsBPNodeCorrosionProbes.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).EndInit();
            this.cmsSWRootLines.ResumeLayout(false);
            this.cmsSWNodeLines.ResumeLayout(false);
            this.cmsSWRootFlowSegments.ResumeLayout(false);
            this.cmsSWNodeFlowSegments.ResumeLayout(false);
            this.cmsSWRootComponents.ResumeLayout(false);
            this.cmsSWNodeComponents.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem menuFile;
        private System.Windows.Forms.ToolStripMenuItem menuNewDoc;
        private System.Windows.Forms.ToolStripMenuItem menuOpenDoc;
        private System.Windows.Forms.ToolStripMenuItem menuSaveDoc;
        private System.Windows.Forms.ToolStripMenuItem menuSaveAs;
        private System.Windows.Forms.ToolStripMenuItem menuExitApp;
        private System.Windows.Forms.ToolStripSeparator menuSeparator;
        private ESRI.ArcGIS.Controls.AxMapControl axMapControl1;
        private ESRI.ArcGIS.Controls.AxToolbarControl axToolbarControl1;
        private ESRI.ArcGIS.Controls.AxTOCControl axTOCControl1;
        private ESRI.ArcGIS.Controls.AxLicenseControl axLicenseControl1;
        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel statusBarXY;
        private System.Windows.Forms.SplitContainer scMain;
        private System.Windows.Forms.SplitContainer scMap;
        private System.Windows.Forms.SplitContainer scSelections;
        private System.Windows.Forms.Label lblSelectedFeaturesTitle;
        private System.Windows.Forms.TreeView treeViewMain;
        private System.Windows.Forms.ImageList ilTreeViewMain;
        private System.Windows.Forms.Panel pnlCommands;
        private System.Windows.Forms.TabControl tcCommandModes;
        private System.Windows.Forms.TabPage tabBE_LineSection_Association;
        private System.Windows.Forms.TextBox tbRadius;
        private System.Windows.Forms.Button Reset;
        private System.Windows.Forms.Button Cancel;
        private System.Windows.Forms.Button btnLineSectionsQuery;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ContextMenuStrip cmsBPRootLineSections;
        private System.Windows.Forms.ToolStripSeparator mnuCustomFilterSeparatorLineSections;
        private System.Windows.Forms.ContextMenuStrip cmsBPRootLines;
        private System.Windows.Forms.ToolStripSeparator mnuCustomFilterSeparatorLines;
        private System.Windows.Forms.ContextMenuStrip cmsBPRootBoringLocations;
        private System.Windows.Forms.ContextMenuStrip cmsBPRootCorrosionProbes;
        private System.Windows.Forms.ContextMenuStrip cmsBPNodeLineSection;
        private System.Windows.Forms.ToolStripMenuItem mnuLineSectionProperties;
        private System.Windows.Forms.ContextMenuStrip cmsBPNodeLines;
        private System.Windows.Forms.ToolStripMenuItem mnuLineProperties;
        private System.Windows.Forms.ContextMenuStrip cmsBPNodeBoringLocations;
        private System.Windows.Forms.ToolStripMenuItem mnuBoringLocationProperties;
        private System.Windows.Forms.ContextMenuStrip cmsBPNodeCorrosionProbes;
        private System.Windows.Forms.ToolStripMenuItem mnuCPProperties;
        private System.Windows.Forms.ContextMenuStrip cmsTreeViewGeneral;
        private System.Windows.Forms.ToolStripMenuItem mnuExpandAll;
        private System.Windows.Forms.ToolStripMenuItem mnuCollapseAll;
        private System.Windows.Forms.Button btnCollapseAll;
        private System.Windows.Forms.Button btnExpandAll;
        private System.Windows.Forms.ToolStripMenuItem mnuFilterLines;
        private System.Windows.Forms.ToolStripMenuItem mnuFilterLineSections;
        private System.Windows.Forms.ToolStripMenuItem selectLinesInASystemToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem lineSectionsAssociatedWithLCSsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem lineSectionsThatHaveBeenInspectedToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem lineSectionsWithAHistoryOfLeaksToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem spatialQueryToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem linesNearbyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem lineSectionsNearbyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem spatialQueryToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem linesNearbyToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem lineSectionsNearbyToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem inspectedLinesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem linesContainingLicensedMaterialradioactiveFluidsToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem utilitiesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem synchronizeFeatureTablesToolStripMenuItem;
        private System.Windows.Forms.ToolTip tooltipMain;
        private System.Windows.Forms.ToolStripStatusLabel tsVersion;
        private System.Windows.Forms.ToolStripStatusLabel tsFileName;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem findToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem linesBPToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem lineSectionsBPToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem boringBPExcavationLocationsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem corrosionBPProbesToolStripMenuItem;
        private System.Windows.Forms.TabPage tabPrebuiltSearch;
        private System.Windows.Forms.ToolStripMenuItem associatedLineSectionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem inspectedLineSectionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem linesInASystemToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem lineSectionsAssociatedWithLCOsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem lineSectionsWithAHistoryOfLeaksToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem2;
        private System.Windows.Forms.TabControl tcTOC;
        private System.Windows.Forms.TabPage tabTOC;
        private System.Windows.Forms.ToolStripMenuItem inspectionsToolStripMenuItem;
        private System.Windows.Forms.Button btnExportTreeViewToExcel;
        private System.Windows.Forms.SaveFileDialog saveExcelDialog;
        private System.Windows.Forms.ToolStripMenuItem lineSectionsWithGuidedWaveInspectionsByLevelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem4;
        private System.Windows.Forms.TabPage tabFilter;
        private System.Windows.Forms.Button btnRemoveFilter;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbFilterLayerSelect;
        private System.Windows.Forms.Button btnResetFilter;
        private System.Windows.Forms.Button btnCreateFilter;
        private System.Windows.Forms.ComboBox cbActiveFilter;
        private System.Windows.Forms.Button btnFilterEdit;
        private System.Windows.Forms.Button btnFilterSave;
        private System.IO.FileSystemWatcher fileSystemWatcher1;
        private System.Windows.Forms.CheckBox cbApplyToSelection;
        private System.Windows.Forms.TabPage tabFind;
        private System.Windows.Forms.ToolStripMenuItem lineSectionsByRiskRankingValuesToolStripMenuItem;
        private System.Windows.Forms.Panel pnlFindBG;
        private System.Windows.Forms.Panel pnlFindBanner;
        private System.Windows.Forms.Label lblQueryDesc;
        private System.Windows.Forms.Panel pnlFindDetails;
        private System.Windows.Forms.Label lblQueryParamNote;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ListBox lbFiltersApplied;
        private System.Windows.Forms.TabPage tabDisplay;
        private System.Windows.Forms.RadioButton rbByRiskRankingColor;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbRiskRankingChoices;
        private ESRI.ArcGIS.Controls.AxSymbologyControl axSymbologyControl1;
        private System.Windows.Forms.RadioButton rbByRiskRankingColorRamp;
        private System.Windows.Forms.RadioButton rbByLayer;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.SplitContainer scMapLegend;
        private System.Windows.Forms.ToolStripMenuItem viewToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mapLegendToolStripMenuItem;
        private System.Windows.Forms.TabControl tcMapLegend;
        private System.Windows.Forms.TabPage tabMapLegendRiskRanking;
        private System.Windows.Forms.DataGridView gridRiskRankingLegend;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColLow;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColLowMed;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColMed;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColMedHigh;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColHigh;
        private System.Windows.Forms.ComboBox cbMapLegendRiskFieldToUse;
        private System.Windows.Forms.RadioButton rbMapLegendDisplayLSByRisk;
        private System.Windows.Forms.RadioButton rbMapLegendDisplayLSByLayer;
        private System.Windows.Forms.Label lblMapLegendRiskRankingBanner;
        private System.Windows.Forms.Button btnMapLegendRiskRankingClose;
        private System.Windows.Forms.ToolStripMenuItem lineSectionRiskRankingToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem runDatabaseUpdatesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem layerConfiguration;
        private System.Windows.Forms.ContextMenuStrip cmsSWRootLines;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem5;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem6;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem7;
        private System.Windows.Forms.ContextMenuStrip cmsSWNodeLines;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem8;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ContextMenuStrip cmsSWRootFlowSegments;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem9;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ContextMenuStrip cmsSWNodeFlowSegments;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem10;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem11;
        private System.Windows.Forms.ContextMenuStrip cmsSWRootComponents;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem12;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ContextMenuStrip cmsSWNodeComponents;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem13;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem14;
    }
}

