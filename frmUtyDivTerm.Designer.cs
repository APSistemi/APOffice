namespace APOffice
{
    partial class frmUtyDivTerm
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
            this.panel2 = new System.Windows.Forms.Panel();
            this.rdbAll = new System.Windows.Forms.RadioButton();
            this.rdbAtt = new System.Windows.Forms.RadioButton();
            this.chkAnaArt = new System.Windows.Forms.CheckBox();
            this.chkTabUsr = new System.Windows.Forms.CheckBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.dgv1 = new APOffice.APDataGridView();
            this.dgv2 = new APOffice.APDataGridView();
            this.btnFtp = new System.Windows.Forms.Button();
            this.lblPath = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(492, 24);
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
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.rdbAll);
            this.panel2.Controls.Add(this.rdbAtt);
            this.panel2.Controls.Add(this.chkAnaArt);
            this.panel2.Location = new System.Drawing.Point(278, 79);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(111, 56);
            this.panel2.TabIndex = 19;
            // 
            // rdbAll
            // 
            this.rdbAll.AutoSize = true;
            this.rdbAll.Location = new System.Drawing.Point(3, 16);
            this.rdbAll.Name = "rdbAll";
            this.rdbAll.Size = new System.Drawing.Size(46, 17);
            this.rdbAll.TabIndex = 1;
            this.rdbAll.Text = "Tutti";
            this.rdbAll.UseVisualStyleBackColor = true;
            // 
            // rdbAtt
            // 
            this.rdbAtt.AutoSize = true;
            this.rdbAtt.Checked = true;
            this.rdbAtt.Location = new System.Drawing.Point(3, 1);
            this.rdbAtt.Name = "rdbAtt";
            this.rdbAtt.Size = new System.Drawing.Size(104, 17);
            this.rdbAtt.TabIndex = 0;
            this.rdbAtt.TabStop = true;
            this.rdbAtt.Text = "Solo articoli attivi";
            this.rdbAtt.UseVisualStyleBackColor = true;
            // 
            // chkAnaArt
            // 
            this.chkAnaArt.AutoSize = true;
            this.chkAnaArt.Location = new System.Drawing.Point(3, 34);
            this.chkAnaArt.Name = "chkAnaArt";
            this.chkAnaArt.Size = new System.Drawing.Size(99, 17);
            this.chkAnaArt.TabIndex = 18;
            this.chkAnaArt.Text = "Impianto articoli";
            this.chkAnaArt.UseVisualStyleBackColor = true;
            // 
            // chkTabUsr
            // 
            this.chkTabUsr.AutoSize = true;
            this.chkTabUsr.Location = new System.Drawing.Point(279, 59);
            this.chkTabUsr.Name = "chkTabUsr";
            this.chkTabUsr.Size = new System.Drawing.Size(90, 17);
            this.chkTabUsr.TabIndex = 20;
            this.chkTabUsr.Text = "Tabella utenti";
            this.chkTabUsr.UseVisualStyleBackColor = true;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(398, 58);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(88, 79);
            this.btnOk.TabIndex = 21;
            this.btnOk.Text = "Genera";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(7, 266);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(479, 23);
            this.progressBar1.TabIndex = 22;
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(9, 59);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(253, 153);
            this.dgv1.TabIndex = 23;
            // 
            // dgv2
            // 
            this.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv2.Location = new System.Drawing.Point(7, 218);
            this.dgv2.Name = "dgv2";
            this.dgv2.Size = new System.Drawing.Size(479, 44);
            this.dgv2.TabIndex = 24;
            // 
            // btnFtp
            // 
            this.btnFtp.Location = new System.Drawing.Point(398, 139);
            this.btnFtp.Name = "btnFtp";
            this.btnFtp.Size = new System.Drawing.Size(88, 73);
            this.btnFtp.TabIndex = 25;
            this.btnFtp.Text = "Invio FTP";
            this.btnFtp.UseVisualStyleBackColor = true;
            this.btnFtp.Click += new System.EventHandler(this.btnFtp_Click);
            // 
            // lblPath
            // 
            this.lblPath.AutoSize = true;
            this.lblPath.Location = new System.Drawing.Point(10, 39);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(163, 13);
            this.lblPath.TabIndex = 26;
            this.lblPath.Text = "C:\\ApProject\\Temp\\DivPhone7\\";
            // 
            // frmUtyDivTerm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(492, 297);
            this.ControlBox = false;
            this.Controls.Add(this.lblPath);
            this.Controls.Add(this.btnFtp);
            this.Controls.Add(this.dgv2);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.chkTabUsr);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyDivTerm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Divulgazioni a terminalini";
            this.Load += new System.EventHandler(this.frmUtyDivTerm_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyDivTerm_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.RadioButton rdbAll;
        private System.Windows.Forms.RadioButton rdbAtt;
        private System.Windows.Forms.CheckBox chkAnaArt;
        private System.Windows.Forms.CheckBox chkTabUsr;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.ProgressBar progressBar1;
        private APOffice.APDataGridView dgv1;
        private APOffice.APDataGridView dgv2;
        private System.Windows.Forms.Button btnFtp;
        private System.Windows.Forms.Label lblPath;
    }
}
