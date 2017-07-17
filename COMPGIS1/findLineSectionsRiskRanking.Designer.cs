namespace COMPGIS1
{
    partial class findLineSectionsRiskRanking
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
            this.clbSusc = new System.Windows.Forms.CheckedListBox();
            this.clbCons = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.lblSuscSelected = new System.Windows.Forms.Label();
            this.lblConsSelected = new System.Windows.Forms.Label();
            this.gbCondition = new System.Windows.Forms.GroupBox();
            this.rbAND = new System.Windows.Forms.RadioButton();
            this.rbOR = new System.Windows.Forms.RadioButton();
            this.cbApplyToSelection = new System.Windows.Forms.CheckBox();
            this.gbCondition.SuspendLayout();
            this.SuspendLayout();
            // 
            // clbSusc
            // 
            this.clbSusc.CheckOnClick = true;
            this.clbSusc.FormattingEnabled = true;
            this.clbSusc.Location = new System.Drawing.Point(17, 32);
            this.clbSusc.Name = "clbSusc";
            this.clbSusc.Size = new System.Drawing.Size(310, 409);
            this.clbSusc.TabIndex = 0;
            // 
            // clbCons
            // 
            this.clbCons.CheckOnClick = true;
            this.clbCons.FormattingEnabled = true;
            this.clbCons.Location = new System.Drawing.Point(342, 32);
            this.clbCons.Name = "clbCons";
            this.clbCons.Size = new System.Drawing.Size(310, 409);
            this.clbCons.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Susceptibility";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(339, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Consequence";
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(577, 470);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(496, 470);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 5;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            // 
            // lblSuscSelected
            // 
            this.lblSuscSelected.Location = new System.Drawing.Point(12, 444);
            this.lblSuscSelected.Name = "lblSuscSelected";
            this.lblSuscSelected.Size = new System.Drawing.Size(310, 23);
            this.lblSuscSelected.TabIndex = 8;
            // 
            // lblConsSelected
            // 
            this.lblConsSelected.Location = new System.Drawing.Point(337, 444);
            this.lblConsSelected.Name = "lblConsSelected";
            this.lblConsSelected.Size = new System.Drawing.Size(310, 23);
            this.lblConsSelected.TabIndex = 9;
            // 
            // gbCondition
            // 
            this.gbCondition.Controls.Add(this.rbOR);
            this.gbCondition.Controls.Add(this.rbAND);
            this.gbCondition.Location = new System.Drawing.Point(20, 447);
            this.gbCondition.Name = "gbCondition";
            this.gbCondition.Size = new System.Drawing.Size(196, 50);
            this.gbCondition.TabIndex = 10;
            this.gbCondition.TabStop = false;
            this.gbCondition.Text = " Condition ";
            // 
            // rbAND
            // 
            this.rbAND.AutoSize = true;
            this.rbAND.Location = new System.Drawing.Point(142, 20);
            this.rbAND.Name = "rbAND";
            this.rbAND.Size = new System.Drawing.Size(48, 17);
            this.rbAND.TabIndex = 0;
            this.rbAND.Text = "AND";
            this.rbAND.UseVisualStyleBackColor = true;
            // 
            // rbOR
            // 
            this.rbOR.AutoSize = true;
            this.rbOR.Checked = true;
            this.rbOR.Location = new System.Drawing.Point(77, 20);
            this.rbOR.Name = "rbOR";
            this.rbOR.Size = new System.Drawing.Size(41, 17);
            this.rbOR.TabIndex = 1;
            this.rbOR.TabStop = true;
            this.rbOR.Text = "OR";
            this.rbOR.UseVisualStyleBackColor = true;
            // 
            // cbApplyToSelection
            // 
            this.cbApplyToSelection.AutoSize = true;
            this.cbApplyToSelection.Location = new System.Drawing.Point(264, 467);
            this.cbApplyToSelection.Name = "cbApplyToSelection";
            this.cbApplyToSelection.Size = new System.Drawing.Size(148, 17);
            this.cbApplyToSelection.TabIndex = 26;
            this.cbApplyToSelection.Text = "Apply to Current Selection";
            this.cbApplyToSelection.UseVisualStyleBackColor = true;
            this.cbApplyToSelection.CheckedChanged += new System.EventHandler(this.cbApplyToSelection_CheckedChanged);
            // 
            // findLineSectionsRiskRanking
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(669, 509);
            this.Controls.Add(this.cbApplyToSelection);
            this.Controls.Add(this.gbCondition);
            this.Controls.Add(this.lblConsSelected);
            this.Controls.Add(this.lblSuscSelected);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.clbCons);
            this.Controls.Add(this.clbSusc);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "findLineSectionsRiskRanking";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Find - Line Sections by Risk Ranking";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.findLineSectionsRiskRanking_FormClosing);
            this.gbCondition.ResumeLayout(false);
            this.gbCondition.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox clbSusc;
        private System.Windows.Forms.CheckedListBox clbCons;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label lblSuscSelected;
        private System.Windows.Forms.Label lblConsSelected;
        private System.Windows.Forms.GroupBox gbCondition;
        private System.Windows.Forms.RadioButton rbOR;
        private System.Windows.Forms.RadioButton rbAND;
        private System.Windows.Forms.CheckBox cbApplyToSelection;
    }
}