namespace APOffice
{
    partial class frmGesTabEcr
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnLiv3 = new System.Windows.Forms.Button();
            this.btnLiv2 = new System.Windows.Forms.Button();
            this.btnLiv1 = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.dgv3 = new APOffice.APDataGridView();
            this.dgv1 = new APOffice.APDataGridView();
            this.dgv2 = new APOffice.APDataGridView();
            this.menuStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.excelToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(872, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // esciToolStripMenuItem
            // 
            this.esciToolStripMenuItem.Name = "esciToolStripMenuItem";
            this.esciToolStripMenuItem.Size = new System.Drawing.Size(39, 20);
            this.esciToolStripMenuItem.Text = "Esci";
            this.esciToolStripMenuItem.Click += new System.EventHandler(this.esciToolStripMenuItem_Click);
            // 
            // excelToolStripMenuItem
            // 
            this.excelToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.excelToolStripMenuItem.Name = "excelToolStripMenuItem";
            this.excelToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.excelToolStripMenuItem.Text = "Excel";
            this.excelToolStripMenuItem.Click += new System.EventHandler(this.excelToolStripMenuItem_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Gainsboro;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.btnLiv3);
            this.panel1.Controls.Add(this.btnLiv2);
            this.panel1.Controls.Add(this.btnLiv1);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.dgv3);
            this.panel1.Controls.Add(this.dgv1);
            this.panel1.Controls.Add(this.dgv2);
            this.panel1.Location = new System.Drawing.Point(4, 32);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(861, 483);
            this.panel1.TabIndex = 2;
            // 
            // btnLiv3
            // 
            this.btnLiv3.Location = new System.Drawing.Point(778, 217);
            this.btnLiv3.Name = "btnLiv3";
            this.btnLiv3.Size = new System.Drawing.Size(75, 19);
            this.btnLiv3.TabIndex = 27;
            this.btnLiv3.Text = "Nuovo";
            this.btnLiv3.UseVisualStyleBackColor = true;
            this.btnLiv3.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnLiv2
            // 
            this.btnLiv2.Location = new System.Drawing.Point(777, 4);
            this.btnLiv2.Name = "btnLiv2";
            this.btnLiv2.Size = new System.Drawing.Size(75, 19);
            this.btnLiv2.TabIndex = 26;
            this.btnLiv2.Text = "Nuovo";
            this.btnLiv2.UseVisualStyleBackColor = true;
            this.btnLiv2.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnLiv1
            // 
            this.btnLiv1.Location = new System.Drawing.Point(352, 4);
            this.btnLiv1.Name = "btnLiv1";
            this.btnLiv1.Size = new System.Drawing.Size(75, 19);
            this.btnLiv1.TabIndex = 25;
            this.btnLiv1.Text = "Nuovo";
            this.btnLiv1.UseVisualStyleBackColor = true;
            this.btnLiv1.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 8);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(46, 13);
            this.label5.TabIndex = 23;
            this.label5.Text = "Livello 1";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(432, 8);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(46, 13);
            this.label4.TabIndex = 22;
            this.label4.Text = "Livello 2";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 220);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(46, 13);
            this.label3.TabIndex = 21;
            this.label3.Text = "Livello 3";
            // 
            // dgv3
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Lavender;
            this.dgv3.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv3.Location = new System.Drawing.Point(6, 236);
            this.dgv3.Name = "dgv3";
            this.dgv3.Size = new System.Drawing.Size(846, 236);
            this.dgv3.TabIndex = 20;
            this.dgv3.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv3_CellDoubleClick);
            this.dgv3.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellEndEdit);
            this.dgv3.RowEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv3_RowEnter);
            // 
            // dgv1
            // 
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.Lavender;
            this.dgv1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(6, 24);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(420, 185);
            this.dgv1.TabIndex = 18;
            this.dgv1.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellEndEdit);
            this.dgv1.RowEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_RowEnter);
            // 
            // dgv2
            // 
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.Lavender;
            this.dgv2.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle3;
            this.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv2.Location = new System.Drawing.Point(435, 24);
            this.dgv2.Name = "dgv2";
            this.dgv2.Size = new System.Drawing.Size(418, 185);
            this.dgv2.TabIndex = 17;
            this.dgv2.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellEndEdit);
            this.dgv2.RowEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv2_RowEnter);
            // 
            // frmGesTabEcr
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLight;
            this.ClientSize = new System.Drawing.Size(872, 519);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesTabEcr";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Gestione tabelle merceologiche";
            this.Load += new System.EventHandler(this.frmGesTabEcr_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesTabEcr_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnLiv3;
        private System.Windows.Forms.Button btnLiv2;
        private System.Windows.Forms.Button btnLiv1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private APOffice.APDataGridView dgv3;
        private APOffice.APDataGridView dgv1;
        private APOffice.APDataGridView dgv2;
        private System.Windows.Forms.ToolStripMenuItem excelToolStripMenuItem;
    }
}
