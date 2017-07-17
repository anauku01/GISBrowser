namespace COMPGIS1
{
    partial class LayerConfigForm
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
            this.tabControlMain = new System.Windows.Forms.TabControl();
            this.tabLayerMap = new System.Windows.Forms.TabPage();
            this.gridLayerMap = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColMappedLayer = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.tabLayerKeyFieldMap = new System.Windows.Forms.TabPage();
            this.btnSelect = new System.Windows.Forms.Button();
            this.gridKeyFields = new System.Windows.Forms.DataGridView();
            this.ColLayerName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColCWIndex = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.toolTipMain = new System.Windows.Forms.ToolTip(this.components);
            this.tabControlMain.SuspendLayout();
            this.tabLayerMap.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridLayerMap)).BeginInit();
            this.tabLayerKeyFieldMap.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridKeyFields)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControlMain
            // 
            this.tabControlMain.Controls.Add(this.tabLayerMap);
            this.tabControlMain.Controls.Add(this.tabLayerKeyFieldMap);
            this.tabControlMain.Location = new System.Drawing.Point(12, 12);
            this.tabControlMain.Name = "tabControlMain";
            this.tabControlMain.SelectedIndex = 0;
            this.tabControlMain.Size = new System.Drawing.Size(410, 311);
            this.tabControlMain.TabIndex = 1;
            // 
            // tabLayerMap
            // 
            this.tabLayerMap.Controls.Add(this.gridLayerMap);
            this.tabLayerMap.Location = new System.Drawing.Point(4, 22);
            this.tabLayerMap.Name = "tabLayerMap";
            this.tabLayerMap.Padding = new System.Windows.Forms.Padding(3);
            this.tabLayerMap.Size = new System.Drawing.Size(402, 285);
            this.tabLayerMap.TabIndex = 1;
            this.tabLayerMap.Text = "Layers Mapping";
            this.tabLayerMap.UseVisualStyleBackColor = true;
            // 
            // gridLayerMap
            // 
            this.gridLayerMap.AllowUserToAddRows = false;
            this.gridLayerMap.AllowUserToDeleteRows = false;
            this.gridLayerMap.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridLayerMap.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.ColMappedLayer});
            this.gridLayerMap.Location = new System.Drawing.Point(6, 6);
            this.gridLayerMap.MultiSelect = false;
            this.gridLayerMap.Name = "gridLayerMap";
            this.gridLayerMap.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridLayerMap.Size = new System.Drawing.Size(390, 273);
            this.gridLayerMap.TabIndex = 3;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.HeaderText = "System Layer Name";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            this.dataGridViewTextBoxColumn1.Width = 150;
            // 
            // ColMappedLayer
            // 
            this.ColMappedLayer.HeaderText = "Mapped Layer";
            this.ColMappedLayer.Name = "ColMappedLayer";
            this.ColMappedLayer.Width = 175;
            // 
            // tabLayerKeyFieldMap
            // 
            this.tabLayerKeyFieldMap.Controls.Add(this.btnSelect);
            this.tabLayerKeyFieldMap.Controls.Add(this.gridKeyFields);
            this.tabLayerKeyFieldMap.Location = new System.Drawing.Point(4, 22);
            this.tabLayerKeyFieldMap.Name = "tabLayerKeyFieldMap";
            this.tabLayerKeyFieldMap.Size = new System.Drawing.Size(402, 285);
            this.tabLayerKeyFieldMap.TabIndex = 3;
            this.tabLayerKeyFieldMap.Text = "Layer Key Field Map";
            this.tabLayerKeyFieldMap.UseVisualStyleBackColor = true;
            // 
            // btnSelect
            // 
            this.btnSelect.Location = new System.Drawing.Point(4, 259);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(75, 23);
            this.btnSelect.TabIndex = 3;
            this.btnSelect.Text = "Select...";
            this.btnSelect.UseVisualStyleBackColor = true;
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // gridKeyFields
            // 
            this.gridKeyFields.AllowUserToAddRows = false;
            this.gridKeyFields.AllowUserToDeleteRows = false;
            this.gridKeyFields.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridKeyFields.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColLayerName,
            this.ColCWIndex});
            this.gridKeyFields.Location = new System.Drawing.Point(6, 6);
            this.gridKeyFields.MultiSelect = false;
            this.gridKeyFields.Name = "gridKeyFields";
            this.gridKeyFields.ReadOnly = true;
            this.gridKeyFields.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridKeyFields.Size = new System.Drawing.Size(394, 250);
            this.gridKeyFields.TabIndex = 2;
            this.gridKeyFields.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridKeyFields_CellDoubleClick);
            // 
            // ColLayerName
            // 
            this.ColLayerName.HeaderText = "Layer Name";
            this.ColLayerName.Name = "ColLayerName";
            this.ColLayerName.ReadOnly = true;
            this.ColLayerName.Width = 150;
            // 
            // ColCWIndex
            // 
            this.ColCWIndex.HeaderText = "CW Index Field";
            this.ColCWIndex.Name = "ColCWIndex";
            this.ColCWIndex.ReadOnly = true;
            this.ColCWIndex.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColCWIndex.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.ColCWIndex.Width = 175;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(343, 329);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(262, 329);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 3;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // toolTipMain
            // 
            this.toolTipMain.IsBalloon = true;
            // 
            // LayerConfigForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(433, 365);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.tabControlMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "LayerConfigForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Layer Configuration";
            this.Load += new System.EventHandler(this.LayerConfigForm_Load);
            this.tabControlMain.ResumeLayout(false);
            this.tabLayerMap.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridLayerMap)).EndInit();
            this.tabLayerKeyFieldMap.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridKeyFields)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControlMain;
        private System.Windows.Forms.TabPage tabLayerKeyFieldMap;
        private System.Windows.Forms.TabPage tabLayerMap;
        private System.Windows.Forms.DataGridView gridLayerMap;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColMappedLayer;
        private System.Windows.Forms.ToolTip toolTipMain;
        private System.Windows.Forms.Button btnSelect;
        public System.Windows.Forms.DataGridView gridKeyFields;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColLayerName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColCWIndex;
    }
}