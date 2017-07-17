namespace COMPGIS1
{
    partial class ConfigurationForm
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
            this.gridSetup = new System.Windows.Forms.DataGridView();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.ColLayerName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColVisible = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.ColCWIndex = new System.Windows.Forms.DataGridViewComboBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.gridSetup)).BeginInit();
            this.SuspendLayout();
            // 
            // gridSetup
            // 
            this.gridSetup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridSetup.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColLayerName,
            this.ColVisible,
            this.ColCWIndex});
            this.gridSetup.Location = new System.Drawing.Point(12, 12);
            this.gridSetup.Name = "gridSetup";
            this.gridSetup.Size = new System.Drawing.Size(436, 310);
            this.gridSetup.TabIndex = 0;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(373, 334);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(292, 334);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            // 
            // ColLayerName
            // 
            this.ColLayerName.HeaderText = "Layer Name";
            this.ColLayerName.Name = "ColLayerName";
            this.ColLayerName.Width = 125;
            // 
            // ColVisible
            // 
            this.ColVisible.HeaderText = "Visible";
            this.ColVisible.Name = "ColVisible";
            this.ColVisible.Width = 50;
            // 
            // ColCWIndex
            // 
            this.ColCWIndex.HeaderText = "CW Index Field";
            this.ColCWIndex.Name = "ColCWIndex";
            this.ColCWIndex.Width = 200;
            // 
            // ConfigurationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(460, 365);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.gridSetup);
            this.Name = "ConfigurationForm";
            this.Text = "Layer Field Configuration";
            ((System.ComponentModel.ISupportInitialize)(this.gridSetup)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView gridSetup;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColLayerName;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColVisible;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColCWIndex;
    }
}