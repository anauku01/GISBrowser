namespace COMPGIS1
{
    partial class BPSettingsForm
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
            this.btnClose = new System.Windows.Forms.Button();
            this.btnReIndex = new System.Windows.Forms.Button();
            this.cbLines = new System.Windows.Forms.CheckBox();
            this.cbLineSections = new System.Windows.Forms.CheckBox();
            this.cbBoringLocations = new System.Windows.Forms.CheckBox();
            this.cbCorrosionProbes = new System.Windows.Forms.CheckBox();
            this.lblStatus = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tabControlMain = new System.Windows.Forms.TabControl();
            this.tabSynchronize = new System.Windows.Forms.TabPage();
            this.cbTestStations = new System.Windows.Forms.CheckBox();
            this.cbRectifiers = new System.Windows.Forms.CheckBox();
            this.lblStatus2 = new System.Windows.Forms.Label();
            this.tabControlMain.SuspendLayout();
            this.tabSynchronize.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(345, 329);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnReIndex
            // 
            this.btnReIndex.Location = new System.Drawing.Point(265, 202);
            this.btnReIndex.Name = "btnReIndex";
            this.btnReIndex.Size = new System.Drawing.Size(112, 23);
            this.btnReIndex.TabIndex = 6;
            this.btnReIndex.Text = "Index Tables";
            this.btnReIndex.UseVisualStyleBackColor = true;
            this.btnReIndex.Click += new System.EventHandler(this.btnReIndex_Click);
            // 
            // cbLines
            // 
            this.cbLines.AutoSize = true;
            this.cbLines.Location = new System.Drawing.Point(48, 87);
            this.cbLines.Name = "cbLines";
            this.cbLines.Size = new System.Drawing.Size(51, 17);
            this.cbLines.TabIndex = 0;
            this.cbLines.Text = "Lines";
            this.cbLines.UseVisualStyleBackColor = true;
            this.cbLines.Click += new System.EventHandler(this.cbLines_Click);
            // 
            // cbLineSections
            // 
            this.cbLineSections.AutoSize = true;
            this.cbLineSections.Location = new System.Drawing.Point(48, 110);
            this.cbLineSections.Name = "cbLineSections";
            this.cbLineSections.Size = new System.Drawing.Size(90, 17);
            this.cbLineSections.TabIndex = 1;
            this.cbLineSections.Text = "Line Sections";
            this.cbLineSections.UseVisualStyleBackColor = true;
            this.cbLineSections.Click += new System.EventHandler(this.cbLines_Click);
            // 
            // cbBoringLocations
            // 
            this.cbBoringLocations.AutoSize = true;
            this.cbBoringLocations.Location = new System.Drawing.Point(48, 133);
            this.cbBoringLocations.Name = "cbBoringLocations";
            this.cbBoringLocations.Size = new System.Drawing.Size(163, 17);
            this.cbBoringLocations.TabIndex = 2;
            this.cbBoringLocations.Text = "Boring/Excavation Locations";
            this.cbBoringLocations.UseVisualStyleBackColor = true;
            this.cbBoringLocations.Click += new System.EventHandler(this.cbLines_Click);
            // 
            // cbCorrosionProbes
            // 
            this.cbCorrosionProbes.AutoSize = true;
            this.cbCorrosionProbes.Location = new System.Drawing.Point(48, 156);
            this.cbCorrosionProbes.Name = "cbCorrosionProbes";
            this.cbCorrosionProbes.Size = new System.Drawing.Size(106, 17);
            this.cbCorrosionProbes.TabIndex = 3;
            this.cbCorrosionProbes.Text = "Corrosion Probes";
            this.cbCorrosionProbes.UseVisualStyleBackColor = true;
            this.cbCorrosionProbes.Click += new System.EventHandler(this.cbLines_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.Location = new System.Drawing.Point(25, 245);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(369, 16);
            this.lblStatus.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(20, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(346, 49);
            this.label1.TabIndex = 7;
            this.label1.Text = "This process synchronzes the GIS features with the CompWorks Database records.  T" +
    "his function should only be run when a new set of features is imported to the ma" +
    "p or to initialize a new database.";
            // 
            // tabControlMain
            // 
            this.tabControlMain.Controls.Add(this.tabSynchronize);
            this.tabControlMain.Location = new System.Drawing.Point(12, 12);
            this.tabControlMain.Name = "tabControlMain";
            this.tabControlMain.SelectedIndex = 0;
            this.tabControlMain.Size = new System.Drawing.Size(408, 311);
            this.tabControlMain.TabIndex = 0;
            // 
            // tabSynchronize
            // 
            this.tabSynchronize.Controls.Add(this.cbTestStations);
            this.tabSynchronize.Controls.Add(this.cbRectifiers);
            this.tabSynchronize.Controls.Add(this.lblStatus2);
            this.tabSynchronize.Controls.Add(this.label1);
            this.tabSynchronize.Controls.Add(this.lblStatus);
            this.tabSynchronize.Controls.Add(this.cbBoringLocations);
            this.tabSynchronize.Controls.Add(this.cbLineSections);
            this.tabSynchronize.Controls.Add(this.cbLines);
            this.tabSynchronize.Controls.Add(this.btnReIndex);
            this.tabSynchronize.Controls.Add(this.cbCorrosionProbes);
            this.tabSynchronize.Location = new System.Drawing.Point(4, 22);
            this.tabSynchronize.Name = "tabSynchronize";
            this.tabSynchronize.Padding = new System.Windows.Forms.Padding(3);
            this.tabSynchronize.Size = new System.Drawing.Size(400, 285);
            this.tabSynchronize.TabIndex = 0;
            this.tabSynchronize.Text = "Synchronize";
            this.tabSynchronize.UseVisualStyleBackColor = true;
            // 
            // cbTestStations
            // 
            this.cbTestStations.AutoSize = true;
            this.cbTestStations.Location = new System.Drawing.Point(48, 202);
            this.cbTestStations.Name = "cbTestStations";
            this.cbTestStations.Size = new System.Drawing.Size(88, 17);
            this.cbTestStations.TabIndex = 5;
            this.cbTestStations.Text = "Test Stations";
            this.cbTestStations.UseVisualStyleBackColor = true;
            // 
            // cbRectifiers
            // 
            this.cbRectifiers.AutoSize = true;
            this.cbRectifiers.Location = new System.Drawing.Point(48, 179);
            this.cbRectifiers.Name = "cbRectifiers";
            this.cbRectifiers.Size = new System.Drawing.Size(70, 17);
            this.cbRectifiers.TabIndex = 4;
            this.cbRectifiers.Text = "Rectifiers";
            this.cbRectifiers.UseVisualStyleBackColor = true;
            this.cbRectifiers.Click += new System.EventHandler(this.cbLines_Click);
            // 
            // lblStatus2
            // 
            this.lblStatus2.Location = new System.Drawing.Point(25, 261);
            this.lblStatus2.Name = "lblStatus2";
            this.lblStatus2.Size = new System.Drawing.Size(369, 13);
            this.lblStatus2.TabIndex = 8;
            // 
            // BPSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(428, 360);
            this.Controls.Add(this.tabControlMain);
            this.Controls.Add(this.btnClose);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "BPSettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Feature Configuration";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DBSynchForm_FormClosing);
            this.Load += new System.EventHandler(this.DBSynchForm_Load);
            this.tabControlMain.ResumeLayout(false);
            this.tabSynchronize.ResumeLayout(false);
            this.tabSynchronize.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnReIndex;
        private System.Windows.Forms.CheckBox cbLines;
        private System.Windows.Forms.CheckBox cbLineSections;
        private System.Windows.Forms.CheckBox cbBoringLocations;
        private System.Windows.Forms.CheckBox cbCorrosionProbes;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TabControl tabControlMain;
        private System.Windows.Forms.TabPage tabSynchronize;
        private System.Windows.Forms.Label lblStatus2;
        private System.Windows.Forms.CheckBox cbRectifiers;
        private System.Windows.Forms.CheckBox cbTestStations;
    }
}