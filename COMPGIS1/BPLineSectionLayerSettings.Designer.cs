namespace COMPGIS1
{
    partial class BPLineSectionLayerSettings
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BPLineSectionLayerSettings));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabGeneral = new System.Windows.Forms.TabPage();
            this.lvLengths = new System.Windows.Forms.ListView();
            this.UnitCol = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.LengthCol = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.label6 = new System.Windows.Forms.Label();
            this.lblLayerDesc = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblLayerName = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tabDisplay = new System.Windows.Forms.TabPage();
            this.rbByRiskLevel = new System.Windows.Forms.RadioButton();
            this.label2 = new System.Windows.Forms.Label();
            this.lblColorRampName = new System.Windows.Forms.Label();
            this.cbRiskRankingChoices = new System.Windows.Forms.ComboBox();
            this.axSymbologyControl1 = new ESRI.ArcGIS.Controls.AxSymbologyControl();
            this.rbColorRamp = new System.Windows.Forms.RadioButton();
            this.rbByLayer = new System.Windows.Forms.RadioButton();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabGeneral.SuspendLayout();
            this.tabDisplay.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.axSymbologyControl1)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabGeneral);
            this.tabControl1.Controls.Add(this.tabDisplay);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(503, 322);
            this.tabControl1.TabIndex = 0;
            // 
            // tabGeneral
            // 
            this.tabGeneral.Controls.Add(this.lvLengths);
            this.tabGeneral.Controls.Add(this.label6);
            this.tabGeneral.Controls.Add(this.lblLayerDesc);
            this.tabGeneral.Controls.Add(this.label3);
            this.tabGeneral.Controls.Add(this.lblLayerName);
            this.tabGeneral.Controls.Add(this.label1);
            this.tabGeneral.Location = new System.Drawing.Point(4, 22);
            this.tabGeneral.Name = "tabGeneral";
            this.tabGeneral.Size = new System.Drawing.Size(495, 296);
            this.tabGeneral.TabIndex = 1;
            this.tabGeneral.Text = "General";
            this.tabGeneral.UseVisualStyleBackColor = true;
            // 
            // lvLengths
            // 
            this.lvLengths.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.lvLengths.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.UnitCol,
            this.LengthCol});
            this.lvLengths.FullRowSelect = true;
            this.lvLengths.Location = new System.Drawing.Point(108, 91);
            this.lvLengths.MultiSelect = false;
            this.lvLengths.Name = "lvLengths";
            this.lvLengths.Size = new System.Drawing.Size(248, 152);
            this.lvLengths.TabIndex = 8;
            this.lvLengths.UseCompatibleStateImageBehavior = false;
            // 
            // UnitCol
            // 
            this.UnitCol.Text = "Unit";
            this.UnitCol.Width = 50;
            // 
            // LengthCol
            // 
            this.LengthCol.Text = "Length";
            this.LengthCol.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.LengthCol.Width = 110;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.label6.Location = new System.Drawing.Point(21, 93);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(60, 13);
            this.label6.TabIndex = 6;
            this.label6.Text = "Lengths (ft)";
            // 
            // lblLayerDesc
            // 
            this.lblLayerDesc.AutoEllipsis = true;
            this.lblLayerDesc.BackColor = System.Drawing.Color.Transparent;
            this.lblLayerDesc.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lblLayerDesc.Location = new System.Drawing.Point(108, 61);
            this.lblLayerDesc.Name = "lblLayerDesc";
            this.lblLayerDesc.Size = new System.Drawing.Size(359, 13);
            this.lblLayerDesc.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.label3.Location = new System.Drawing.Point(21, 61);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(60, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Description";
            // 
            // lblLayerName
            // 
            this.lblLayerName.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lblLayerName.Location = new System.Drawing.Point(108, 30);
            this.lblLayerName.Name = "lblLayerName";
            this.lblLayerName.Size = new System.Drawing.Size(359, 13);
            this.lblLayerName.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.label1.Location = new System.Drawing.Point(21, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Name";
            // 
            // tabDisplay
            // 
            this.tabDisplay.Controls.Add(this.rbByRiskLevel);
            this.tabDisplay.Controls.Add(this.label2);
            this.tabDisplay.Controls.Add(this.lblColorRampName);
            this.tabDisplay.Controls.Add(this.cbRiskRankingChoices);
            this.tabDisplay.Controls.Add(this.axSymbologyControl1);
            this.tabDisplay.Controls.Add(this.rbColorRamp);
            this.tabDisplay.Controls.Add(this.rbByLayer);
            this.tabDisplay.Location = new System.Drawing.Point(4, 22);
            this.tabDisplay.Name = "tabDisplay";
            this.tabDisplay.Padding = new System.Windows.Forms.Padding(3);
            this.tabDisplay.Size = new System.Drawing.Size(495, 296);
            this.tabDisplay.TabIndex = 0;
            this.tabDisplay.Text = "Display";
            this.tabDisplay.UseVisualStyleBackColor = true;
            // 
            // rbByRiskLevel
            // 
            this.rbByRiskLevel.AutoSize = true;
            this.rbByRiskLevel.Location = new System.Drawing.Point(21, 47);
            this.rbByRiskLevel.Name = "rbByRiskLevel";
            this.rbByRiskLevel.Size = new System.Drawing.Size(159, 17);
            this.rbByRiskLevel.TabIndex = 8;
            this.rbByRiskLevel.Text = "Ranking by Risk Level Color";
            this.rbByRiskLevel.UseVisualStyleBackColor = true;
            this.rbByRiskLevel.CheckedChanged += new System.EventHandler(this.rbByLayer_CheckedChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(247, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Risk to Use";
            // 
            // lblColorRampName
            // 
            this.lblColorRampName.AutoEllipsis = true;
            this.lblColorRampName.Location = new System.Drawing.Point(250, 268);
            this.lblColorRampName.Name = "lblColorRampName";
            this.lblColorRampName.Size = new System.Drawing.Size(234, 23);
            this.lblColorRampName.TabIndex = 6;
            this.lblColorRampName.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cbRiskRankingChoices
            // 
            this.cbRiskRankingChoices.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbRiskRankingChoices.FormattingEnabled = true;
            this.cbRiskRankingChoices.Items.AddRange(new object[] {
            "Overall Ranking",
            "Susceptibility Ranking",
            "Consequence Ranking"});
            this.cbRiskRankingChoices.Location = new System.Drawing.Point(319, 10);
            this.cbRiskRankingChoices.Name = "cbRiskRankingChoices";
            this.cbRiskRankingChoices.Size = new System.Drawing.Size(165, 21);
            this.cbRiskRankingChoices.TabIndex = 5;
            this.cbRiskRankingChoices.SelectedIndexChanged += new System.EventHandler(this.cbRiskRankingChoices_SelectedIndexChanged);
            // 
            // axSymbologyControl1
            // 
            this.axSymbologyControl1.Location = new System.Drawing.Point(250, 37);
            this.axSymbologyControl1.Name = "axSymbologyControl1";
            this.axSymbologyControl1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axSymbologyControl1.OcxState")));
            this.axSymbologyControl1.Size = new System.Drawing.Size(234, 228);
            this.axSymbologyControl1.TabIndex = 4;
            this.axSymbologyControl1.OnItemSelected += new ESRI.ArcGIS.Controls.ISymbologyControlEvents_Ax_OnItemSelectedEventHandler(this.axSymbologyControl1_OnItemSelected);
            // 
            // rbColorRamp
            // 
            this.rbColorRamp.AutoSize = true;
            this.rbColorRamp.Location = new System.Drawing.Point(21, 70);
            this.rbColorRamp.Name = "rbColorRamp";
            this.rbColorRamp.Size = new System.Drawing.Size(137, 17);
            this.rbColorRamp.TabIndex = 1;
            this.rbColorRamp.Text = "Ranking by Color Ramp";
            this.rbColorRamp.UseVisualStyleBackColor = true;
            this.rbColorRamp.CheckedChanged += new System.EventHandler(this.rbByLayer_CheckedChanged);
            // 
            // rbByLayer
            // 
            this.rbByLayer.AutoSize = true;
            this.rbByLayer.Checked = true;
            this.rbByLayer.Location = new System.Drawing.Point(21, 24);
            this.rbByLayer.Name = "rbByLayer";
            this.rbByLayer.Size = new System.Drawing.Size(129, 17);
            this.rbByLayer.TabIndex = 0;
            this.rbByLayer.TabStop = true;
            this.rbByLayer.Text = "Default Layer Settings";
            this.rbByLayer.UseVisualStyleBackColor = true;
            this.rbByLayer.CheckedChanged += new System.EventHandler(this.rbByLayer_CheckedChanged);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(436, 340);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(355, 340);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // BPLineSectionLayerSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(527, 375);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.tabControl1);
            this.Name = "BPLineSectionLayerSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Layer Settings : Line Sections";
            this.Load += new System.EventHandler(this.LineSectionLayerSettings_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabGeneral.ResumeLayout(false);
            this.tabGeneral.PerformLayout();
            this.tabDisplay.ResumeLayout(false);
            this.tabDisplay.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.axSymbologyControl1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabDisplay;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.RadioButton rbColorRamp;
        private System.Windows.Forms.RadioButton rbByLayer;
        private ESRI.ArcGIS.Controls.AxSymbologyControl axSymbologyControl1;
        private System.Windows.Forms.TabPage tabGeneral;
        private System.Windows.Forms.Label lblLayerName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblLayerDesc;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbRiskRankingChoices;
        private System.Windows.Forms.Label lblColorRampName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RadioButton rbByRiskLevel;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ListView lvLengths;
        private System.Windows.Forms.ColumnHeader UnitCol;
        private System.Windows.Forms.ColumnHeader LengthCol;
    }
}