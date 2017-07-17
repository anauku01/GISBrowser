namespace COMPGIS1
{
    partial class BEDetails
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
            this.tcMain = new System.Windows.Forms.TabControl();
            this.tabGeneral = new System.Windows.Forms.TabPage();
            this.label6 = new System.Windows.Forms.Label();
            this.tbSampleDate = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tbNotes = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tbGroundwaterLevel = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.rbBoring = new System.Windows.Forms.RadioButton();
            this.rbExcavation = new System.Windows.Forms.RadioButton();
            this.label8 = new System.Windows.Forms.Label();
            this.tbPlantNorthing = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.tbPlantEasting = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tbDescription = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tbBEID = new System.Windows.Forms.TextBox();
            this.tabFieldLogs = new System.Windows.Forms.TabPage();
            this.gridFieldLogs = new System.Windows.Forms.DataGridView();
            this.tabLineSections = new System.Windows.Forms.TabPage();
            this.gridLineSections = new System.Windows.Forms.DataGridView();
            this.tabChemicalAnalysis = new System.Windows.Forms.TabPage();
            this.gridChemResults = new System.Windows.Forms.DataGridView();
            this.tabSieveAnalysis = new System.Windows.Forms.TabPage();
            this.gridSieve = new System.Windows.Forms.DataGridView();
            this.tabCP = new System.Windows.Forms.TabPage();
            this.gridCP = new System.Windows.Forms.DataGridView();
            this.tabDCP = new System.Windows.Forms.TabPage();
            this.gridDCP = new System.Windows.Forms.DataGridView();
            this.btnClose = new System.Windows.Forms.Button();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.dataSet1 = new System.Data.DataSet();
            this.soilFieldLogBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.soilFieldLogBindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.tcMain.SuspendLayout();
            this.tabGeneral.SuspendLayout();
            this.tabFieldLogs.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridFieldLogs)).BeginInit();
            this.tabLineSections.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridLineSections)).BeginInit();
            this.tabChemicalAnalysis.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridChemResults)).BeginInit();
            this.tabSieveAnalysis.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridSieve)).BeginInit();
            this.tabCP.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridCP)).BeginInit();
            this.tabDCP.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridDCP)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.soilFieldLogBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.soilFieldLogBindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // tcMain
            // 
            this.tcMain.Controls.Add(this.tabGeneral);
            this.tcMain.Controls.Add(this.tabFieldLogs);
            this.tcMain.Controls.Add(this.tabLineSections);
            this.tcMain.Controls.Add(this.tabChemicalAnalysis);
            this.tcMain.Controls.Add(this.tabSieveAnalysis);
            this.tcMain.Controls.Add(this.tabCP);
            this.tcMain.Controls.Add(this.tabDCP);
            this.tcMain.Location = new System.Drawing.Point(12, 12);
            this.tcMain.Multiline = true;
            this.tcMain.Name = "tcMain";
            this.tcMain.SelectedIndex = 0;
            this.tcMain.Size = new System.Drawing.Size(427, 404);
            this.tcMain.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tcMain.TabIndex = 1;
            // 
            // tabGeneral
            // 
            this.tabGeneral.Controls.Add(this.label6);
            this.tabGeneral.Controls.Add(this.tbSampleDate);
            this.tabGeneral.Controls.Add(this.label5);
            this.tabGeneral.Controls.Add(this.tbNotes);
            this.tabGeneral.Controls.Add(this.label4);
            this.tabGeneral.Controls.Add(this.tbGroundwaterLevel);
            this.tabGeneral.Controls.Add(this.label3);
            this.tabGeneral.Controls.Add(this.rbBoring);
            this.tabGeneral.Controls.Add(this.rbExcavation);
            this.tabGeneral.Controls.Add(this.label8);
            this.tabGeneral.Controls.Add(this.tbPlantNorthing);
            this.tabGeneral.Controls.Add(this.label7);
            this.tabGeneral.Controls.Add(this.tbPlantEasting);
            this.tabGeneral.Controls.Add(this.label2);
            this.tabGeneral.Controls.Add(this.tbDescription);
            this.tabGeneral.Controls.Add(this.label1);
            this.tabGeneral.Controls.Add(this.tbBEID);
            this.tabGeneral.Location = new System.Drawing.Point(4, 40);
            this.tabGeneral.Name = "tabGeneral";
            this.tabGeneral.Padding = new System.Windows.Forms.Padding(3);
            this.tabGeneral.Size = new System.Drawing.Size(419, 360);
            this.tabGeneral.TabIndex = 0;
            this.tabGeneral.Text = "General";
            this.tabGeneral.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(20, 77);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(68, 13);
            this.label6.TabIndex = 22;
            this.label6.Text = "Sample Date";
            // 
            // tbSampleDate
            // 
            this.tbSampleDate.Location = new System.Drawing.Point(139, 74);
            this.tbSampleDate.Name = "tbSampleDate";
            this.tbSampleDate.ReadOnly = true;
            this.tbSampleDate.Size = new System.Drawing.Size(244, 20);
            this.tbSampleDate.TabIndex = 21;
            this.tbSampleDate.TabStop = false;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(20, 232);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(35, 13);
            this.label5.TabIndex = 20;
            this.label5.Text = "Notes";
            // 
            // tbNotes
            // 
            this.tbNotes.Location = new System.Drawing.Point(139, 229);
            this.tbNotes.Multiline = true;
            this.tbNotes.Name = "tbNotes";
            this.tbNotes.ReadOnly = true;
            this.tbNotes.Size = new System.Drawing.Size(244, 71);
            this.tbNotes.TabIndex = 19;
            this.tbNotes.TabStop = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(20, 129);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(97, 13);
            this.label4.TabIndex = 18;
            this.label4.Text = "Groundwater Level";
            // 
            // tbGroundwaterLevel
            // 
            this.tbGroundwaterLevel.Location = new System.Drawing.Point(139, 126);
            this.tbGroundwaterLevel.Name = "tbGroundwaterLevel";
            this.tbGroundwaterLevel.ReadOnly = true;
            this.tbGroundwaterLevel.Size = new System.Drawing.Size(71, 20);
            this.tbGroundwaterLevel.TabIndex = 17;
            this.tbGroundwaterLevel.TabStop = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(20, 53);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(31, 13);
            this.label3.TabIndex = 16;
            this.label3.Text = "Type";
            // 
            // rbBoring
            // 
            this.rbBoring.AutoSize = true;
            this.rbBoring.Checked = true;
            this.rbBoring.Enabled = false;
            this.rbBoring.Location = new System.Drawing.Point(139, 51);
            this.rbBoring.Name = "rbBoring";
            this.rbBoring.Size = new System.Drawing.Size(55, 17);
            this.rbBoring.TabIndex = 0;
            this.rbBoring.TabStop = true;
            this.rbBoring.Text = "Boring";
            this.rbBoring.UseVisualStyleBackColor = true;
            // 
            // rbExcavation
            // 
            this.rbExcavation.AutoSize = true;
            this.rbExcavation.Enabled = false;
            this.rbExcavation.Location = new System.Drawing.Point(200, 51);
            this.rbExcavation.Name = "rbExcavation";
            this.rbExcavation.Size = new System.Drawing.Size(78, 17);
            this.rbExcavation.TabIndex = 1;
            this.rbExcavation.Text = "Excavation";
            this.rbExcavation.UseVisualStyleBackColor = true;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(216, 103);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(74, 13);
            this.label8.TabIndex = 15;
            this.label8.Text = "Plant Northing";
            // 
            // tbPlantNorthing
            // 
            this.tbPlantNorthing.Location = new System.Drawing.Point(312, 100);
            this.tbPlantNorthing.Name = "tbPlantNorthing";
            this.tbPlantNorthing.ReadOnly = true;
            this.tbPlantNorthing.Size = new System.Drawing.Size(71, 20);
            this.tbPlantNorthing.TabIndex = 14;
            this.tbPlantNorthing.TabStop = false;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(20, 103);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(69, 13);
            this.label7.TabIndex = 13;
            this.label7.Text = "Plant Easting";
            // 
            // tbPlantEasting
            // 
            this.tbPlantEasting.Location = new System.Drawing.Point(139, 100);
            this.tbPlantEasting.Name = "tbPlantEasting";
            this.tbPlantEasting.ReadOnly = true;
            this.tbPlantEasting.Size = new System.Drawing.Size(71, 20);
            this.tbPlantEasting.TabIndex = 12;
            this.tbPlantEasting.TabStop = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 155);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(60, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Description";
            // 
            // tbDescription
            // 
            this.tbDescription.Location = new System.Drawing.Point(139, 152);
            this.tbDescription.Multiline = true;
            this.tbDescription.Name = "tbDescription";
            this.tbDescription.ReadOnly = true;
            this.tbDescription.Size = new System.Drawing.Size(244, 71);
            this.tbDescription.TabIndex = 2;
            this.tbDescription.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(109, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Boring/Excavation ID";
            // 
            // tbBEID
            // 
            this.tbBEID.Location = new System.Drawing.Point(139, 24);
            this.tbBEID.Name = "tbBEID";
            this.tbBEID.ReadOnly = true;
            this.tbBEID.Size = new System.Drawing.Size(244, 20);
            this.tbBEID.TabIndex = 2;
            this.tbBEID.TabStop = false;
            // 
            // tabFieldLogs
            // 
            this.tabFieldLogs.Controls.Add(this.gridFieldLogs);
            this.tabFieldLogs.Location = new System.Drawing.Point(4, 40);
            this.tabFieldLogs.Name = "tabFieldLogs";
            this.tabFieldLogs.Padding = new System.Windows.Forms.Padding(3);
            this.tabFieldLogs.Size = new System.Drawing.Size(419, 360);
            this.tabFieldLogs.TabIndex = 1;
            this.tabFieldLogs.Text = "Field Logs";
            this.tabFieldLogs.UseVisualStyleBackColor = true;
            // 
            // gridFieldLogs
            // 
            this.gridFieldLogs.AllowUserToAddRows = false;
            this.gridFieldLogs.AllowUserToDeleteRows = false;
            this.gridFieldLogs.AllowUserToOrderColumns = true;
            this.gridFieldLogs.AllowUserToResizeRows = false;
            this.gridFieldLogs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridFieldLogs.Location = new System.Drawing.Point(15, 15);
            this.gridFieldLogs.Name = "gridFieldLogs";
            this.gridFieldLogs.ReadOnly = true;
            this.gridFieldLogs.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridFieldLogs.Size = new System.Drawing.Size(387, 328);
            this.gridFieldLogs.TabIndex = 0;
            this.gridFieldLogs.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridFieldLogs_CellDoubleClick);
            // 
            // tabLineSections
            // 
            this.tabLineSections.Controls.Add(this.gridLineSections);
            this.tabLineSections.Location = new System.Drawing.Point(4, 40);
            this.tabLineSections.Name = "tabLineSections";
            this.tabLineSections.Size = new System.Drawing.Size(419, 360);
            this.tabLineSections.TabIndex = 4;
            this.tabLineSections.Text = "Line Sections";
            this.tabLineSections.UseVisualStyleBackColor = true;
            // 
            // gridLineSections
            // 
            this.gridLineSections.AllowUserToAddRows = false;
            this.gridLineSections.AllowUserToDeleteRows = false;
            this.gridLineSections.AllowUserToResizeRows = false;
            this.gridLineSections.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridLineSections.Location = new System.Drawing.Point(16, 16);
            this.gridLineSections.Name = "gridLineSections";
            this.gridLineSections.ReadOnly = true;
            this.gridLineSections.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridLineSections.Size = new System.Drawing.Size(387, 328);
            this.gridLineSections.TabIndex = 1;
            this.gridLineSections.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridLineSections_CellDoubleClick);
            // 
            // tabChemicalAnalysis
            // 
            this.tabChemicalAnalysis.Controls.Add(this.gridChemResults);
            this.tabChemicalAnalysis.Location = new System.Drawing.Point(4, 40);
            this.tabChemicalAnalysis.Name = "tabChemicalAnalysis";
            this.tabChemicalAnalysis.Size = new System.Drawing.Size(419, 360);
            this.tabChemicalAnalysis.TabIndex = 2;
            this.tabChemicalAnalysis.Text = "Chemical Analysis";
            this.tabChemicalAnalysis.UseVisualStyleBackColor = true;
            // 
            // gridChemResults
            // 
            this.gridChemResults.AllowUserToAddRows = false;
            this.gridChemResults.AllowUserToDeleteRows = false;
            this.gridChemResults.AllowUserToResizeColumns = false;
            this.gridChemResults.AllowUserToResizeRows = false;
            this.gridChemResults.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridChemResults.Location = new System.Drawing.Point(16, 16);
            this.gridChemResults.Name = "gridChemResults";
            this.gridChemResults.ReadOnly = true;
            this.gridChemResults.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridChemResults.Size = new System.Drawing.Size(387, 328);
            this.gridChemResults.TabIndex = 2;
            this.gridChemResults.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridChemResults_CellDoubleClick);
            // 
            // tabSieveAnalysis
            // 
            this.tabSieveAnalysis.Controls.Add(this.gridSieve);
            this.tabSieveAnalysis.Location = new System.Drawing.Point(4, 40);
            this.tabSieveAnalysis.Name = "tabSieveAnalysis";
            this.tabSieveAnalysis.Size = new System.Drawing.Size(419, 360);
            this.tabSieveAnalysis.TabIndex = 3;
            this.tabSieveAnalysis.Text = "Sieve Analysis";
            this.tabSieveAnalysis.UseVisualStyleBackColor = true;
            // 
            // gridSieve
            // 
            this.gridSieve.AllowUserToAddRows = false;
            this.gridSieve.AllowUserToDeleteRows = false;
            this.gridSieve.AllowUserToResizeColumns = false;
            this.gridSieve.AllowUserToResizeRows = false;
            this.gridSieve.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridSieve.Location = new System.Drawing.Point(16, 16);
            this.gridSieve.Name = "gridSieve";
            this.gridSieve.ReadOnly = true;
            this.gridSieve.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridSieve.Size = new System.Drawing.Size(387, 328);
            this.gridSieve.TabIndex = 2;
            this.gridSieve.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridSieve_CellDoubleClick);
            // 
            // tabCP
            // 
            this.tabCP.Controls.Add(this.gridCP);
            this.tabCP.Location = new System.Drawing.Point(4, 40);
            this.tabCP.Name = "tabCP";
            this.tabCP.Size = new System.Drawing.Size(419, 360);
            this.tabCP.TabIndex = 5;
            this.tabCP.Text = "Corrosion Probes";
            this.tabCP.UseVisualStyleBackColor = true;
            // 
            // gridCP
            // 
            this.gridCP.AllowUserToAddRows = false;
            this.gridCP.AllowUserToDeleteRows = false;
            this.gridCP.AllowUserToResizeColumns = false;
            this.gridCP.AllowUserToResizeRows = false;
            this.gridCP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridCP.Location = new System.Drawing.Point(16, 16);
            this.gridCP.Name = "gridCP";
            this.gridCP.ReadOnly = true;
            this.gridCP.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridCP.Size = new System.Drawing.Size(387, 328);
            this.gridCP.TabIndex = 1;
            this.gridCP.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridCP_CellDoubleClick);
            // 
            // tabDCP
            // 
            this.tabDCP.Controls.Add(this.gridDCP);
            this.tabDCP.Location = new System.Drawing.Point(4, 40);
            this.tabDCP.Name = "tabDCP";
            this.tabDCP.Size = new System.Drawing.Size(419, 360);
            this.tabDCP.TabIndex = 6;
            this.tabDCP.Text = "DCP Results";
            this.tabDCP.UseVisualStyleBackColor = true;
            // 
            // gridDCP
            // 
            this.gridDCP.AllowUserToAddRows = false;
            this.gridDCP.AllowUserToDeleteRows = false;
            this.gridDCP.AllowUserToResizeColumns = false;
            this.gridDCP.AllowUserToResizeRows = false;
            this.gridDCP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridDCP.Location = new System.Drawing.Point(16, 16);
            this.gridDCP.Name = "gridDCP";
            this.gridDCP.ReadOnly = true;
            this.gridDCP.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridDCP.Size = new System.Drawing.Size(387, 328);
            this.gridDCP.TabIndex = 1;
            // 
            // btnClose
            // 
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Location = new System.Drawing.Point(360, 422);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "NewDataSet";
            // 
            // BEDetails
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(453, 460);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.tcMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "BEDetails";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "BEDetails";
            this.tcMain.ResumeLayout(false);
            this.tabGeneral.ResumeLayout(false);
            this.tabGeneral.PerformLayout();
            this.tabFieldLogs.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridFieldLogs)).EndInit();
            this.tabLineSections.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridLineSections)).EndInit();
            this.tabChemicalAnalysis.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridChemResults)).EndInit();
            this.tabSieveAnalysis.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridSieve)).EndInit();
            this.tabCP.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridCP)).EndInit();
            this.tabDCP.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridDCP)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.soilFieldLogBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.soilFieldLogBindingSource1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tcMain;
        private System.Windows.Forms.TabPage tabGeneral;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbDescription;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbBEID;
        private System.Windows.Forms.TabPage tabFieldLogs;
        private System.Windows.Forms.TabPage tabChemicalAnalysis;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox tbPlantNorthing;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox tbPlantEasting;
        private System.Windows.Forms.RadioButton rbExcavation;
        private System.Windows.Forms.RadioButton rbBoring;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TabPage tabSieveAnalysis;
        private System.Windows.Forms.TabPage tabLineSections;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbGroundwaterLevel;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbNotes;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox tbSampleDate;
        private System.Windows.Forms.DataGridView gridFieldLogs;
        private System.Windows.Forms.TabPage tabCP;
        private System.Windows.Forms.DataGridView gridCP;
        private System.Windows.Forms.TabPage tabDCP;
        private System.Windows.Forms.DataGridView gridDCP;
        private System.Windows.Forms.DataGridView gridChemResults;
        private System.Windows.Forms.DataGridView gridLineSections;
        private System.Windows.Forms.DataGridView gridSieve;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Data.DataSet dataSet1;
        private System.Windows.Forms.BindingSource soilFieldLogBindingSource;
        private System.Windows.Forms.BindingSource soilFieldLogBindingSource1;
    }
}