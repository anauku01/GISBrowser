namespace COMPGIS1
{
    partial class BPTestStationDetails
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
            this.tcMain = new System.Windows.Forms.TabControl();
            this.tabTestStation = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.tbInServiceDate = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tbTestStation = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.tbTestStationType = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tbRemarks = new System.Windows.Forms.TextBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.tbLocation = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.tcMain.SuspendLayout();
            this.tabTestStation.SuspendLayout();
            this.SuspendLayout();
            // 
            // tcMain
            // 
            this.tcMain.Controls.Add(this.tabTestStation);
            this.tcMain.Location = new System.Drawing.Point(12, 12);
            this.tcMain.Name = "tcMain";
            this.tcMain.SelectedIndex = 0;
            this.tcMain.Size = new System.Drawing.Size(359, 305);
            this.tcMain.TabIndex = 1;
            // 
            // tabTestStation
            // 
            this.tabTestStation.Controls.Add(this.tbLocation);
            this.tabTestStation.Controls.Add(this.textBox1);
            this.tabTestStation.Controls.Add(this.label4);
            this.tabTestStation.Controls.Add(this.tbRemarks);
            this.tabTestStation.Controls.Add(this.label3);
            this.tabTestStation.Controls.Add(this.tbTestStationType);
            this.tabTestStation.Controls.Add(this.label2);
            this.tabTestStation.Controls.Add(this.tbInServiceDate);
            this.tabTestStation.Controls.Add(this.label1);
            this.tabTestStation.Controls.Add(this.tbTestStation);
            this.tabTestStation.Location = new System.Drawing.Point(4, 22);
            this.tabTestStation.Name = "tabTestStation";
            this.tabTestStation.Padding = new System.Windows.Forms.Padding(3);
            this.tabTestStation.Size = new System.Drawing.Size(351, 279);
            this.tabTestStation.TabIndex = 0;
            this.tabTestStation.Text = "Test Station";
            this.tabTestStation.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 105);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(81, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "In Service Date";
            // 
            // tbInServiceDate
            // 
            this.tbInServiceDate.Location = new System.Drawing.Point(113, 102);
            this.tbInServiceDate.Name = "tbInServiceDate";
            this.tbInServiceDate.ReadOnly = true;
            this.tbInServiceDate.Size = new System.Drawing.Size(99, 20);
            this.tbInServiceDate.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Test Station";
            // 
            // tbTestStation
            // 
            this.tbTestStation.Location = new System.Drawing.Point(113, 24);
            this.tbTestStation.Name = "tbTestStation";
            this.tbTestStation.ReadOnly = true;
            this.tbTestStation.Size = new System.Drawing.Size(205, 20);
            this.tbTestStation.TabIndex = 0;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(20, 53);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(91, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Test Station Type";
            // 
            // tbTestStationType
            // 
            this.tbTestStationType.Location = new System.Drawing.Point(113, 50);
            this.tbTestStationType.Name = "tbTestStationType";
            this.tbTestStationType.ReadOnly = true;
            this.tbTestStationType.Size = new System.Drawing.Size(205, 20);
            this.tbTestStationType.TabIndex = 4;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(20, 131);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(49, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Remarks";
            // 
            // tbRemarks
            // 
            this.tbRemarks.Location = new System.Drawing.Point(113, 128);
            this.tbRemarks.Multiline = true;
            this.tbRemarks.Name = "tbRemarks";
            this.tbRemarks.ReadOnly = true;
            this.tbRemarks.Size = new System.Drawing.Size(205, 67);
            this.tbRemarks.TabIndex = 6;
            // 
            // btnClose
            // 
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Location = new System.Drawing.Point(292, 323);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            // 
            // tbLocation
            // 
            this.tbLocation.AutoSize = true;
            this.tbLocation.Location = new System.Drawing.Point(20, 79);
            this.tbLocation.Name = "tbLocation";
            this.tbLocation.Size = new System.Drawing.Size(48, 13);
            this.tbLocation.TabIndex = 9;
            this.tbLocation.Text = "Location";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(113, 76);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(205, 20);
            this.textBox1.TabIndex = 8;
            // 
            // BPTestStationDetails
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 356);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.tcMain);
            this.Name = "BPTestStationDetails";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "BPTestStationDetails";
            this.tcMain.ResumeLayout(false);
            this.tabTestStation.ResumeLayout(false);
            this.tabTestStation.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tcMain;
        private System.Windows.Forms.TabPage tabTestStation;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbInServiceDate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbTestStation;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbTestStationType;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbRemarks;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label tbLocation;
        private System.Windows.Forms.TextBox textBox1;
    }
}