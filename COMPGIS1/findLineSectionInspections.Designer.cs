namespace COMPGIS1
{
    partial class findLineSectionInspections
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
            this.cbMethodGeneral = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.tcMain = new System.Windows.Forms.TabControl();
            this.tabInspections = new System.Windows.Forms.TabPage();
            this.tabEvaluations = new System.Windows.Forms.TabPage();
            this.lblResults = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cbWorkScope = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cbMethodEvals = new System.Windows.Forms.ComboBox();
            this.cbApplyToSelection = new System.Windows.Forms.CheckBox();
            this.tcMain.SuspendLayout();
            this.tabInspections.SuspendLayout();
            this.tabEvaluations.SuspendLayout();
            this.SuspendLayout();
            // 
            // cbMethodGeneral
            // 
            this.cbMethodGeneral.DisplayMember = "ObjectName";
            this.cbMethodGeneral.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbMethodGeneral.FormattingEnabled = true;
            this.cbMethodGeneral.Location = new System.Drawing.Point(143, 30);
            this.cbMethodGeneral.Name = "cbMethodGeneral";
            this.cbMethodGeneral.Size = new System.Drawing.Size(213, 21);
            this.cbMethodGeneral.TabIndex = 1;
            this.cbMethodGeneral.ValueMember = "ObjectValue";
            this.cbMethodGeneral.SelectedIndexChanged += new System.EventHandler(this.cbMethod_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(95, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Inspection Method";
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(331, 171);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Enabled = false;
            this.btnOK.Location = new System.Drawing.Point(250, 171);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 0;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // tcMain
            // 
            this.tcMain.Controls.Add(this.tabInspections);
            this.tcMain.Controls.Add(this.tabEvaluations);
            this.tcMain.Location = new System.Drawing.Point(12, 12);
            this.tcMain.Name = "tcMain";
            this.tcMain.SelectedIndex = 0;
            this.tcMain.Size = new System.Drawing.Size(398, 153);
            this.tcMain.TabIndex = 0;
            this.tcMain.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabInspections
            // 
            this.tabInspections.Controls.Add(this.label1);
            this.tabInspections.Controls.Add(this.cbMethodGeneral);
            this.tabInspections.Location = new System.Drawing.Point(4, 22);
            this.tabInspections.Name = "tabInspections";
            this.tabInspections.Padding = new System.Windows.Forms.Padding(3);
            this.tabInspections.Size = new System.Drawing.Size(390, 127);
            this.tabInspections.TabIndex = 0;
            this.tabInspections.Text = "General";
            this.tabInspections.UseVisualStyleBackColor = true;
            // 
            // tabEvaluations
            // 
            this.tabEvaluations.Controls.Add(this.lblResults);
            this.tabEvaluations.Controls.Add(this.label2);
            this.tabEvaluations.Controls.Add(this.label4);
            this.tabEvaluations.Controls.Add(this.cbWorkScope);
            this.tabEvaluations.Controls.Add(this.label3);
            this.tabEvaluations.Controls.Add(this.cbMethodEvals);
            this.tabEvaluations.Location = new System.Drawing.Point(4, 22);
            this.tabEvaluations.Name = "tabEvaluations";
            this.tabEvaluations.Padding = new System.Windows.Forms.Padding(3);
            this.tabEvaluations.Size = new System.Drawing.Size(390, 127);
            this.tabEvaluations.TabIndex = 1;
            this.tabEvaluations.Text = "Evaluations";
            this.tabEvaluations.UseVisualStyleBackColor = true;
            // 
            // lblResults
            // 
            this.lblResults.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblResults.Location = new System.Drawing.Point(140, 102);
            this.lblResults.Name = "lblResults";
            this.lblResults.Size = new System.Drawing.Size(54, 13);
            this.lblResults.TabIndex = 10;
            this.lblResults.Text = "-";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(27, 102);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(104, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Line Sections Found";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(27, 66);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(67, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Work Scope";
            // 
            // cbWorkScope
            // 
            this.cbWorkScope.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbWorkScope.FormattingEnabled = true;
            this.cbWorkScope.Location = new System.Drawing.Point(143, 63);
            this.cbWorkScope.Name = "cbWorkScope";
            this.cbWorkScope.Size = new System.Drawing.Size(213, 21);
            this.cbWorkScope.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(27, 33);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(95, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Inspection Method";
            // 
            // cbMethodEvals
            // 
            this.cbMethodEvals.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbMethodEvals.FormattingEnabled = true;
            this.cbMethodEvals.Items.AddRange(new object[] {
            "All",
            "BE",
            "RT",
            "UT",
            "VT"});
            this.cbMethodEvals.Location = new System.Drawing.Point(143, 30);
            this.cbMethodEvals.Name = "cbMethodEvals";
            this.cbMethodEvals.Size = new System.Drawing.Size(213, 21);
            this.cbMethodEvals.TabIndex = 0;
            this.cbMethodEvals.SelectedIndexChanged += new System.EventHandler(this.cbMethodEvals_SelectedIndexChanged);
            this.cbMethodEvals.SelectionChangeCommitted += new System.EventHandler(this.cbMethodEvals_SelectionChangeCommitted);
            // 
            // cbApplyToSelection
            // 
            this.cbApplyToSelection.AutoSize = true;
            this.cbApplyToSelection.Location = new System.Drawing.Point(12, 175);
            this.cbApplyToSelection.Name = "cbApplyToSelection";
            this.cbApplyToSelection.Size = new System.Drawing.Size(148, 17);
            this.cbApplyToSelection.TabIndex = 26;
            this.cbApplyToSelection.Text = "Apply to Current Selection";
            this.cbApplyToSelection.UseVisualStyleBackColor = true;
            this.cbApplyToSelection.CheckedChanged += new System.EventHandler(this.cbApplyToSelection_CheckedChanged);
            // 
            // findLineSectionInspections
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(422, 206);
            this.Controls.Add(this.cbApplyToSelection);
            this.Controls.Add(this.tcMain);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "findLineSectionInspections";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Find Line Section Inspections";
            this.tcMain.ResumeLayout(false);
            this.tabInspections.ResumeLayout(false);
            this.tabInspections.PerformLayout();
            this.tabEvaluations.ResumeLayout(false);
            this.tabEvaluations.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbMethodGeneral;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.TabControl tcMain;
        private System.Windows.Forms.TabPage tabInspections;
        private System.Windows.Forms.TabPage tabEvaluations;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbMethodEvals;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cbWorkScope;
        private System.Windows.Forms.Label lblResults;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox cbApplyToSelection;
    }
}