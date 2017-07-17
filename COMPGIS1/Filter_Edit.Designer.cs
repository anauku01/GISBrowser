namespace COMPGIS1
{
    partial class Filter_Edit
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
            this.lbFieldName = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cbCondition = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cbValue = new System.Windows.Forms.ComboBox();
            this.rbAND = new System.Windows.Forms.RadioButton();
            this.rbOR = new System.Windows.Forms.RadioButton();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.lvSelections = new System.Windows.Forms.ListView();
            this.colFieldName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colCondition = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colValue = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colOperator = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.dsWorking = new System.Data.DataSet();
            this.dataTable1 = new System.Data.DataTable();
            this.gbOperator = new System.Windows.Forms.GroupBox();
            this.tbFilterName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.tbFilterDescription = new System.Windows.Forms.TextBox();
            this.scMain = new System.Windows.Forms.SplitContainer();
            ((System.ComponentModel.ISupportInitialize)(this.dsWorking)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataTable1)).BeginInit();
            this.gbOperator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.scMain)).BeginInit();
            this.scMain.Panel1.SuspendLayout();
            this.scMain.Panel2.SuspendLayout();
            this.scMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbFieldName
            // 
            this.lbFieldName.FormattingEnabled = true;
            this.lbFieldName.Location = new System.Drawing.Point(3, 25);
            this.lbFieldName.Name = "lbFieldName";
            this.lbFieldName.Size = new System.Drawing.Size(185, 290);
            this.lbFieldName.TabIndex = 0;
            this.lbFieldName.SelectedIndexChanged += new System.EventHandler(this.lbFieldName_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Field Name";
            // 
            // cbCondition
            // 
            this.cbCondition.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbCondition.FormattingEnabled = true;
            this.cbCondition.Items.AddRange(new object[] {
            "=",
            ">",
            ">=",
            "<",
            "<=",
            "<>",
            "LIKE"});
            this.cbCondition.Location = new System.Drawing.Point(257, 25);
            this.cbCondition.Name = "cbCondition";
            this.cbCondition.Size = new System.Drawing.Size(237, 21);
            this.cbCondition.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(200, 28);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Condition";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(200, 55);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(34, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Value";
            // 
            // cbValue
            // 
            this.cbValue.DisplayMember = "ObjectName";
            this.cbValue.FormattingEnabled = true;
            this.cbValue.Location = new System.Drawing.Point(257, 52);
            this.cbValue.Name = "cbValue";
            this.cbValue.Size = new System.Drawing.Size(237, 21);
            this.cbValue.TabIndex = 5;
            this.cbValue.ValueMember = "ObjectValue";
            this.cbValue.TextChanged += new System.EventHandler(this.cbValue_TextChanged);
            // 
            // rbAND
            // 
            this.rbAND.AutoSize = true;
            this.rbAND.Checked = true;
            this.rbAND.Location = new System.Drawing.Point(13, 14);
            this.rbAND.Name = "rbAND";
            this.rbAND.Size = new System.Drawing.Size(48, 17);
            this.rbAND.TabIndex = 7;
            this.rbAND.TabStop = true;
            this.rbAND.Text = "AND";
            this.rbAND.UseVisualStyleBackColor = true;
            // 
            // rbOR
            // 
            this.rbOR.AutoSize = true;
            this.rbOR.Location = new System.Drawing.Point(76, 14);
            this.rbOR.Name = "rbOR";
            this.rbOR.Size = new System.Drawing.Size(41, 17);
            this.rbOR.TabIndex = 8;
            this.rbOR.Text = "OR";
            this.rbOR.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(437, 442);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 9;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(356, 442);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 10;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(419, 96);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 11;
            this.btnAdd.Text = "Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // lvSelections
            // 
            this.lvSelections.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colFieldName,
            this.colCondition,
            this.colValue,
            this.colOperator});
            this.lvSelections.FullRowSelect = true;
            this.lvSelections.GridLines = true;
            this.lvSelections.Location = new System.Drawing.Point(203, 125);
            this.lvSelections.Name = "lvSelections";
            this.lvSelections.Size = new System.Drawing.Size(291, 190);
            this.lvSelections.TabIndex = 12;
            this.lvSelections.UseCompatibleStateImageBehavior = false;
            this.lvSelections.View = System.Windows.Forms.View.Details;
            // 
            // colFieldName
            // 
            this.colFieldName.Text = "Field";
            this.colFieldName.Width = 105;
            // 
            // colCondition
            // 
            this.colCondition.Text = "";
            this.colCondition.Width = 30;
            // 
            // colValue
            // 
            this.colValue.Text = "Value";
            this.colValue.Width = 105;
            // 
            // colOperator
            // 
            this.colOperator.Text = "";
            this.colOperator.Width = 30;
            // 
            // dsWorking
            // 
            this.dsWorking.DataSetName = "NewDataSet";
            this.dsWorking.Tables.AddRange(new System.Data.DataTable[] {
            this.dataTable1});
            // 
            // dataTable1
            // 
            this.dataTable1.TableName = "Table1";
            // 
            // gbOperator
            // 
            this.gbOperator.Controls.Add(this.rbAND);
            this.gbOperator.Controls.Add(this.rbOR);
            this.gbOperator.Enabled = false;
            this.gbOperator.Location = new System.Drawing.Point(257, 79);
            this.gbOperator.Name = "gbOperator";
            this.gbOperator.Size = new System.Drawing.Size(127, 40);
            this.gbOperator.TabIndex = 13;
            this.gbOperator.TabStop = false;
            // 
            // tbFilterName
            // 
            this.tbFilterName.Location = new System.Drawing.Point(69, 6);
            this.tbFilterName.MaxLength = 150;
            this.tbFilterName.Name = "tbFilterName";
            this.tbFilterName.Size = new System.Drawing.Size(425, 20);
            this.tbFilterName.TabIndex = 14;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(35, 13);
            this.label4.TabIndex = 15;
            this.label4.Text = "Name";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(3, 35);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(60, 13);
            this.label5.TabIndex = 17;
            this.label5.Text = "Description";
            // 
            // tbFilterDescription
            // 
            this.tbFilterDescription.Location = new System.Drawing.Point(69, 32);
            this.tbFilterDescription.MaxLength = 255;
            this.tbFilterDescription.Multiline = true;
            this.tbFilterDescription.Name = "tbFilterDescription";
            this.tbFilterDescription.Size = new System.Drawing.Size(425, 63);
            this.tbFilterDescription.TabIndex = 16;
            // 
            // scMain
            // 
            this.scMain.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.scMain.IsSplitterFixed = true;
            this.scMain.Location = new System.Drawing.Point(12, 12);
            this.scMain.Name = "scMain";
            this.scMain.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // scMain.Panel1
            // 
            this.scMain.Panel1.Controls.Add(this.label4);
            this.scMain.Panel1.Controls.Add(this.label5);
            this.scMain.Panel1.Controls.Add(this.tbFilterName);
            this.scMain.Panel1.Controls.Add(this.tbFilterDescription);
            // 
            // scMain.Panel2
            // 
            this.scMain.Panel2.BackColor = System.Drawing.SystemColors.Control;
            this.scMain.Panel2.Controls.Add(this.lbFieldName);
            this.scMain.Panel2.Controls.Add(this.label1);
            this.scMain.Panel2.Controls.Add(this.cbCondition);
            this.scMain.Panel2.Controls.Add(this.label2);
            this.scMain.Panel2.Controls.Add(this.cbValue);
            this.scMain.Panel2.Controls.Add(this.gbOperator);
            this.scMain.Panel2.Controls.Add(this.label3);
            this.scMain.Panel2.Controls.Add(this.lvSelections);
            this.scMain.Panel2.Controls.Add(this.btnAdd);
            this.scMain.Size = new System.Drawing.Size(500, 430);
            this.scMain.SplitterDistance = 100;
            this.scMain.TabIndex = 18;
            // 
            // Filter_Edit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(531, 477);
            this.Controls.Add(this.scMain);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "Filter_Edit";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Filter Editor";
            this.Load += new System.EventHandler(this.Filter_Edit_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dsWorking)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataTable1)).EndInit();
            this.gbOperator.ResumeLayout(false);
            this.gbOperator.PerformLayout();
            this.scMain.Panel1.ResumeLayout(false);
            this.scMain.Panel1.PerformLayout();
            this.scMain.Panel2.ResumeLayout(false);
            this.scMain.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.scMain)).EndInit();
            this.scMain.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox lbFieldName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbCondition;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbValue;
        private System.Windows.Forms.RadioButton rbAND;
        private System.Windows.Forms.RadioButton rbOR;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.ListView lvSelections;
        private System.Windows.Forms.ColumnHeader colFieldName;
        private System.Windows.Forms.ColumnHeader colCondition;
        private System.Windows.Forms.ColumnHeader colValue;
        private System.Data.DataSet dsWorking;
        private System.Data.DataTable dataTable1;
        private System.Windows.Forms.ColumnHeader colOperator;
        private System.Windows.Forms.GroupBox gbOperator;
        private System.Windows.Forms.TextBox tbFilterName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbFilterDescription;
        private System.Windows.Forms.SplitContainer scMain;
    }
}