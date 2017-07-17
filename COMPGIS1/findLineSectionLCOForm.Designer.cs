namespace COMPGIS1
{
    partial class findLineSectionLCOForm
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
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cbApplyToSelection = new System.Windows.Forms.CheckBox();
            this.cbLCO = new System.Windows.Forms.ComboBox();
            this.cbInclusive = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(218, 103);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 0;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(137, 103);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 13);
            this.label1.TabIndex = 18;
            this.label1.Text = "Show with LCO:";
            // 
            // cbApplyToSelection
            // 
            this.cbApplyToSelection.AutoSize = true;
            this.cbApplyToSelection.Location = new System.Drawing.Point(101, 80);
            this.cbApplyToSelection.Name = "cbApplyToSelection";
            this.cbApplyToSelection.Size = new System.Drawing.Size(148, 17);
            this.cbApplyToSelection.TabIndex = 25;
            this.cbApplyToSelection.Text = "Apply to Current Selection";
            this.cbApplyToSelection.UseVisualStyleBackColor = true;
            this.cbApplyToSelection.CheckedChanged += new System.EventHandler(this.cbApplyToSelection_CheckedChanged);
            // 
            // cbLCO
            // 
            this.cbLCO.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbLCO.FormattingEnabled = true;
            this.cbLCO.Items.AddRange(new object[] {
            "None",
            "< 72 Hours",
            "72 Hours",
            "7 Days",
            "14 Days",
            "30 Days",
            "Any LCO"});
            this.cbLCO.Location = new System.Drawing.Point(101, 21);
            this.cbLCO.Name = "cbLCO";
            this.cbLCO.Size = new System.Drawing.Size(192, 21);
            this.cbLCO.TabIndex = 17;
            this.cbLCO.SelectedIndexChanged += new System.EventHandler(this.cbLCO_SelectedIndexChanged);
            // 
            // cbInclusive
            // 
            this.cbInclusive.AutoSize = true;
            this.cbInclusive.Location = new System.Drawing.Point(101, 57);
            this.cbInclusive.Name = "cbInclusive";
            this.cbInclusive.Size = new System.Drawing.Size(173, 17);
            this.cbInclusive.TabIndex = 26;
            this.cbInclusive.Text = "Include LCO less than selected";
            this.cbInclusive.UseVisualStyleBackColor = true;
            this.cbInclusive.CheckedChanged += new System.EventHandler(this.cbInclusive_CheckedChanged);
            // 
            // findLineSectionLCOForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(309, 142);
            this.Controls.Add(this.cbInclusive);
            this.Controls.Add(this.cbApplyToSelection);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbLCO);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "findLineSectionLCOForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Find - Line Section LCO";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox cbApplyToSelection;
        private System.Windows.Forms.ComboBox cbLCO;
        private System.Windows.Forms.CheckBox cbInclusive;
    }
}