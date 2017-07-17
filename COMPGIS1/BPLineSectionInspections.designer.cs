namespace COMPGIS1
{
    partial class BPLineSectionInspections
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
            this.label3 = new System.Windows.Forms.Label();
            this.tbEndPoint = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tbBeginPoint = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tbLineNumber = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tbLineSection = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.tbUnit = new System.Windows.Forms.TextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabGeneral = new System.Windows.Forms.TabPage();
            this.gridGeneralInspections = new System.Windows.Forms.DataGridView();
            this.tabEvaluations = new System.Windows.Forms.TabPage();
            this.gridEvaluations = new System.Windows.Forms.DataGridView();
            this.btnClose = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabGeneral.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridGeneralInspections)).BeginInit();
            this.tabEvaluations.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridEvaluations)).BeginInit();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(213, 46);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 15;
            this.label3.Text = "End Point";
            // 
            // tbEndPoint
            // 
            this.tbEndPoint.Location = new System.Drawing.Point(272, 43);
            this.tbEndPoint.Name = "tbEndPoint";
            this.tbEndPoint.ReadOnly = true;
            this.tbEndPoint.Size = new System.Drawing.Size(72, 20);
            this.tbEndPoint.TabIndex = 14;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(20, 46);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(61, 13);
            this.label4.TabIndex = 13;
            this.label4.Text = "Begin Point";
            // 
            // tbBeginPoint
            // 
            this.tbBeginPoint.Location = new System.Drawing.Point(139, 43);
            this.tbBeginPoint.Name = "tbBeginPoint";
            this.tbBeginPoint.ReadOnly = true;
            this.tbBeginPoint.Size = new System.Drawing.Size(71, 20);
            this.tbBeginPoint.TabIndex = 12;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(366, 24);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "Line Number";
            // 
            // tbLineNumber
            // 
            this.tbLineNumber.Location = new System.Drawing.Point(485, 21);
            this.tbLineNumber.Name = "tbLineNumber";
            this.tbLineNumber.ReadOnly = true;
            this.tbLineNumber.Size = new System.Drawing.Size(205, 20);
            this.tbLineNumber.TabIndex = 10;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Line Section";
            // 
            // tbLineSection
            // 
            this.tbLineSection.Location = new System.Drawing.Point(139, 17);
            this.tbLineSection.Name = "tbLineSection";
            this.tbLineSection.ReadOnly = true;
            this.tbLineSection.Size = new System.Drawing.Size(205, 20);
            this.tbLineSection.TabIndex = 8;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(366, 50);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(26, 13);
            this.label10.TabIndex = 21;
            this.label10.Text = "Unit";
            // 
            // tbUnit
            // 
            this.tbUnit.Location = new System.Drawing.Point(485, 47);
            this.tbUnit.Name = "tbUnit";
            this.tbUnit.ReadOnly = true;
            this.tbUnit.Size = new System.Drawing.Size(71, 20);
            this.tbUnit.TabIndex = 20;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabGeneral);
            this.tabControl1.Controls.Add(this.tabEvaluations);
            this.tabControl1.Location = new System.Drawing.Point(23, 83);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(667, 415);
            this.tabControl1.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tabControl1.TabIndex = 22;
            // 
            // tabGeneral
            // 
            this.tabGeneral.Controls.Add(this.gridGeneralInspections);
            this.tabGeneral.Location = new System.Drawing.Point(4, 22);
            this.tabGeneral.Name = "tabGeneral";
            this.tabGeneral.Padding = new System.Windows.Forms.Padding(3);
            this.tabGeneral.Size = new System.Drawing.Size(659, 389);
            this.tabGeneral.TabIndex = 0;
            this.tabGeneral.Text = "General";
            this.tabGeneral.UseVisualStyleBackColor = true;
            // 
            // gridGeneralInspections
            // 
            this.gridGeneralInspections.AllowUserToAddRows = false;
            this.gridGeneralInspections.AllowUserToDeleteRows = false;
            this.gridGeneralInspections.AllowUserToResizeRows = false;
            this.gridGeneralInspections.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridGeneralInspections.Location = new System.Drawing.Point(6, 6);
            this.gridGeneralInspections.Name = "gridGeneralInspections";
            this.gridGeneralInspections.ReadOnly = true;
            this.gridGeneralInspections.Size = new System.Drawing.Size(647, 377);
            this.gridGeneralInspections.TabIndex = 0;
            this.gridGeneralInspections.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridGeneralInspections_CellDoubleClick);
            // 
            // tabEvaluations
            // 
            this.tabEvaluations.Controls.Add(this.gridEvaluations);
            this.tabEvaluations.Location = new System.Drawing.Point(4, 22);
            this.tabEvaluations.Name = "tabEvaluations";
            this.tabEvaluations.Padding = new System.Windows.Forms.Padding(3);
            this.tabEvaluations.Size = new System.Drawing.Size(659, 389);
            this.tabEvaluations.TabIndex = 1;
            this.tabEvaluations.Text = "Evaluations";
            this.tabEvaluations.UseVisualStyleBackColor = true;
            // 
            // gridEvaluations
            // 
            this.gridEvaluations.AllowUserToAddRows = false;
            this.gridEvaluations.AllowUserToDeleteRows = false;
            this.gridEvaluations.AllowUserToResizeRows = false;
            this.gridEvaluations.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridEvaluations.Location = new System.Drawing.Point(6, 6);
            this.gridEvaluations.Name = "gridEvaluations";
            this.gridEvaluations.ReadOnly = true;
            this.gridEvaluations.Size = new System.Drawing.Size(647, 377);
            this.gridEvaluations.TabIndex = 1;
            this.gridEvaluations.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridEvaluations_CellDoubleClick);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(611, 504);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 23;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // LineSectionInspections
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(707, 534);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.tbUnit);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tbEndPoint);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.tbBeginPoint);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tbLineNumber);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbLineSection);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "LineSectionInspections";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "LineSectionInspections";
            this.tabControl1.ResumeLayout(false);
            this.tabGeneral.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridGeneralInspections)).EndInit();
            this.tabEvaluations.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridEvaluations)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbEndPoint;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbBeginPoint;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbLineNumber;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbLineSection;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox tbUnit;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabGeneral;
        private System.Windows.Forms.TabPage tabEvaluations;
        private System.Windows.Forms.DataGridView gridGeneralInspections;
        private System.Windows.Forms.DataGridView gridEvaluations;
        private System.Windows.Forms.Button btnClose;

    }
}