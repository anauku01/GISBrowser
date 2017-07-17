namespace COMPGIS1
{
    partial class BPDBUpdates
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
            this.clbStatus = new System.Windows.Forms.CheckedListBox();
            this.lblBanner = new System.Windows.Forms.Label();
            this.pbMain = new System.Windows.Forms.ProgressBar();
            this.btnClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // clbStatus
            // 
            this.clbStatus.Enabled = false;
            this.clbStatus.FormattingEnabled = true;
            this.clbStatus.Location = new System.Drawing.Point(12, 34);
            this.clbStatus.Name = "clbStatus";
            this.clbStatus.Size = new System.Drawing.Size(316, 94);
            this.clbStatus.TabIndex = 0;
            // 
            // lblBanner
            // 
            this.lblBanner.AutoEllipsis = true;
            this.lblBanner.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.lblBanner.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblBanner.Location = new System.Drawing.Point(13, 8);
            this.lblBanner.Name = "lblBanner";
            this.lblBanner.Size = new System.Drawing.Size(316, 23);
            this.lblBanner.TabIndex = 1;
            this.lblBanner.Text = "Database Updates in Progress";
            this.lblBanner.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pbMain
            // 
            this.pbMain.Location = new System.Drawing.Point(16, 141);
            this.pbMain.Name = "pbMain";
            this.pbMain.Size = new System.Drawing.Size(313, 12);
            this.pbMain.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.pbMain.TabIndex = 2;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(133, 161);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 3;
            this.btnClose.Text = "Cancel";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // BPDBUpdates
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(341, 196);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.pbMain);
            this.Controls.Add(this.lblBanner);
            this.Controls.Add(this.clbStatus);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "BPDBUpdates";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Database Updates";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.DBUpdates_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.CheckedListBox clbStatus;
        private System.Windows.Forms.Label lblBanner;
        private System.Windows.Forms.ProgressBar pbMain;
        private System.Windows.Forms.Button btnClose;
    }
}