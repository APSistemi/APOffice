namespace APOffice
{
    partial class frmUtyImpianto
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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnExp = new System.Windows.Forms.Button();
            this.btnImp = new System.Windows.Forms.Button();
            this.dgv1 = new APOffice.APDataGridView();
            this.lblPath = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.cmbMdf = new System.Windows.Forms.ComboBox();
            this.chkTab = new System.Windows.Forms.CheckBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(631, 24);
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
            // btnExp
            // 
            this.btnExp.Location = new System.Drawing.Point(31, 49);
            this.btnExp.Name = "btnExp";
            this.btnExp.Size = new System.Drawing.Size(179, 26);
            this.btnExp.TabIndex = 1;
            this.btnExp.Text = "Esporta tabelle";
            this.btnExp.UseVisualStyleBackColor = true;
            this.btnExp.Click += new System.EventHandler(this.btnExp_Click);
            // 
            // btnImp
            // 
            this.btnImp.Location = new System.Drawing.Point(476, 34);
            this.btnImp.Name = "btnImp";
            this.btnImp.Size = new System.Drawing.Size(124, 41);
            this.btnImp.TabIndex = 2;
            this.btnImp.Text = "Importa tabelle";
            this.btnImp.UseVisualStyleBackColor = true;
            this.btnImp.Click += new System.EventHandler(this.btnImp_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(31, 82);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(569, 380);
            this.dgv1.TabIndex = 3;
            // 
            // lblPath
            // 
            this.lblPath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblPath.Location = new System.Drawing.Point(211, 30);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(253, 18);
            this.lblPath.TabIndex = 4;
            this.lblPath.Text = "C:\\APproject\\Temp\\SQL\\TabImport.txt";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(2, 478);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(626, 13);
            this.progressBar1.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(216, 62);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 16;
            this.label1.Text = "DataBase Mdf";
            // 
            // cmbMdf
            // 
            this.cmbMdf.FormattingEnabled = true;
            this.cmbMdf.Location = new System.Drawing.Point(297, 55);
            this.cmbMdf.Name = "cmbMdf";
            this.cmbMdf.Size = new System.Drawing.Size(167, 21);
            this.cmbMdf.TabIndex = 15;
            this.cmbMdf.SelectionChangeCommitted += new System.EventHandler(this.cmbMdf_SelectionChangeCommitted);
            this.cmbMdf.Leave += new System.EventHandler(this.cmbMdf_Leave);
            // 
            // chkTab
            // 
            this.chkTab.AutoSize = true;
            this.chkTab.Checked = true;
            this.chkTab.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkTab.Location = new System.Drawing.Point(31, 27);
            this.chkTab.Name = "chkTab";
            this.chkTab.Size = new System.Drawing.Size(81, 17);
            this.chkTab.TabIndex = 17;
            this.chkTab.Text = "Solo tabelle";
            this.chkTab.UseVisualStyleBackColor = true;
            // 
            // frmUtyImpianto
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(631, 493);
            this.Controls.Add(this.chkTab);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbMdf);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.lblPath);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.btnImp);
            this.Controls.Add(this.btnExp);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyImpianto";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Impianto";
            this.Load += new System.EventHandler(this.frmUtyImpianto_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Button btnExp;
        private System.Windows.Forms.Button btnImp;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbMdf;
        private System.Windows.Forms.CheckBox chkTab;
    }
}
