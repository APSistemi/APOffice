namespace APOffice
{
    partial class frmGesImpTerm
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
            this.recuperoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ftpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.cmbTer = new System.Windows.Forms.ComboBox();
            this.dgv1 = new APOffice.APDataGridView();
            this.btnImp = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.grpInv = new System.Windows.Forms.GroupBox();
            this.lblNrv = new System.Windows.Forms.Label();
            this.cmbNte = new System.Windows.Forms.ComboBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.grpInv.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.recuperoToolStripMenuItem,
            this.ftpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(284, 24);
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
            // recuperoToolStripMenuItem
            // 
            this.recuperoToolStripMenuItem.Name = "recuperoToolStripMenuItem";
            this.recuperoToolStripMenuItem.Size = new System.Drawing.Size(69, 20);
            this.recuperoToolStripMenuItem.Text = "Recupero";
            this.recuperoToolStripMenuItem.Click += new System.EventHandler(this.recuperoToolStripMenuItem_Click);
            // 
            // ftpToolStripMenuItem
            // 
            this.ftpToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.ftpToolStripMenuItem.Name = "ftpToolStripMenuItem";
            this.ftpToolStripMenuItem.Size = new System.Drawing.Size(36, 20);
            this.ftpToolStripMenuItem.Text = "Ftp";
            this.ftpToolStripMenuItem.Click += new System.EventHandler(this.ftpToolStripMenuItem_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 43);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Terminale";
            // 
            // cmbTer
            // 
            this.cmbTer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTer.FormattingEnabled = true;
            this.cmbTer.Location = new System.Drawing.Point(64, 35);
            this.cmbTer.Name = "cmbTer";
            this.cmbTer.Size = new System.Drawing.Size(214, 21);
            this.cmbTer.TabIndex = 2;
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(6, 147);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(272, 114);
            this.dgv1.TabIndex = 3;
            // 
            // btnImp
            // 
            this.btnImp.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnImp.Location = new System.Drawing.Point(6, 61);
            this.btnImp.Name = "btnImp";
            this.btnImp.Size = new System.Drawing.Size(272, 41);
            this.btnImp.TabIndex = 4;
            this.btnImp.Text = "Importa";
            this.btnImp.UseVisualStyleBackColor = true;
            this.btnImp.Click += new System.EventHandler(this.btnImp_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(143, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(97, 13);
            this.label3.TabIndex = 16;
            this.label3.Text = "Numero rilevazione";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(1, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(97, 13);
            this.label2.TabIndex = 15;
            this.label2.Text = "Numero terminalino";
            // 
            // grpInv
            // 
            this.grpInv.Controls.Add(this.lblNrv);
            this.grpInv.Controls.Add(this.cmbNte);
            this.grpInv.Controls.Add(this.label3);
            this.grpInv.Controls.Add(this.label2);
            this.grpInv.Location = new System.Drawing.Point(6, 97);
            this.grpInv.Name = "grpInv";
            this.grpInv.Size = new System.Drawing.Size(272, 44);
            this.grpInv.TabIndex = 17;
            this.grpInv.TabStop = false;
            // 
            // lblNrv
            // 
            this.lblNrv.BackColor = System.Drawing.Color.White;
            this.lblNrv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblNrv.Location = new System.Drawing.Point(146, 23);
            this.lblNrv.Name = "lblNrv";
            this.lblNrv.Size = new System.Drawing.Size(100, 19);
            this.lblNrv.TabIndex = 17;
            this.lblNrv.Text = "01";
            this.lblNrv.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cmbNte
            // 
            this.cmbNte.FormattingEnabled = true;
            this.cmbNte.Location = new System.Drawing.Point(3, 22);
            this.cmbNte.Name = "cmbNte";
            this.cmbNte.Size = new System.Drawing.Size(126, 21);
            this.cmbNte.TabIndex = 15;
            this.cmbNte.SelectedIndexChanged += new System.EventHandler(this.cmbNte_SelectedIndexChanged);
            // 
            // frmGesImpTerm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 266);
            this.ControlBox = false;
            this.Controls.Add(this.btnImp);
            this.Controls.Add(this.grpInv);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.cmbTer);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesImpTerm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Importazione da terminalino";
            this.Load += new System.EventHandler(this.frmGesImpTerm_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.grpInv.ResumeLayout(false);
            this.grpInv.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbTer;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Button btnImp;
        private System.Windows.Forms.ToolStripMenuItem recuperoToolStripMenuItem;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox grpInv;
        private System.Windows.Forms.ComboBox cmbNte;
        private System.Windows.Forms.Label lblNrv;
        private System.Windows.Forms.ToolStripMenuItem ftpToolStripMenuItem;
    }
}
