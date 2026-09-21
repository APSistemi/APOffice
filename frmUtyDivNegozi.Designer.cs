namespace APOffice
{
    partial class frmUtyDivNegozi
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
            this.chkTabGrt = new System.Windows.Forms.CheckBox();
            this.chkTabUsr = new System.Windows.Forms.CheckBox();
            this.chkTabCam = new System.Windows.Forms.CheckBox();
            this.chkTabPpo = new System.Windows.Forms.CheckBox();
            this.chkTabUmi = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dgv2 = new APOffice.APDataGridView();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btnOk = new System.Windows.Forms.Button();
            this.dgv1 = new APOffice.APDataGridView();
            this.chkTabRep = new System.Windows.Forms.CheckBox();
            this.chkAnaArt = new System.Windows.Forms.CheckBox();
            this.chkTabIva = new System.Windows.Forms.CheckBox();
            this.chkArtVar = new System.Windows.Forms.CheckBox();
            this.dtpArtVar = new System.Windows.Forms.DateTimePicker();
            this.chkTabEcr = new System.Windows.Forms.CheckBox();
            this.menuStrip1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).BeginInit();
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
            this.menuStrip1.Size = new System.Drawing.Size(684, 24);
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
            this.panel2.Location = new System.Drawing.Point(126, 37);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(111, 35);
            this.panel2.TabIndex = 33;
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
            // chkTabGrt
            // 
            this.chkTabGrt.AutoSize = true;
            this.chkTabGrt.Location = new System.Drawing.Point(31, 252);
            this.chkTabGrt.Name = "chkTabGrt";
            this.chkTabGrt.Size = new System.Drawing.Size(130, 17);
            this.chkTabGrt.TabIndex = 32;
            this.chkTabGrt.Text = "Tabella gruppi tessere";
            this.chkTabGrt.UseVisualStyleBackColor = true;
            // 
            // chkTabUsr
            // 
            this.chkTabUsr.AutoSize = true;
            this.chkTabUsr.Location = new System.Drawing.Point(31, 222);
            this.chkTabUsr.Name = "chkTabUsr";
            this.chkTabUsr.Size = new System.Drawing.Size(90, 17);
            this.chkTabUsr.TabIndex = 31;
            this.chkTabUsr.Text = "Tabella utenti";
            this.chkTabUsr.UseVisualStyleBackColor = true;
            // 
            // chkTabCam
            // 
            this.chkTabCam.AutoSize = true;
            this.chkTabCam.Location = new System.Drawing.Point(31, 284);
            this.chkTabCam.Name = "chkTabCam";
            this.chkTabCam.Size = new System.Drawing.Size(146, 17);
            this.chkTabCam.TabIndex = 30;
            this.chkTabCam.Text = "Tabella campagne fidelity";
            this.chkTabCam.UseVisualStyleBackColor = true;
            // 
            // chkTabPpo
            // 
            this.chkTabPpo.AutoSize = true;
            this.chkTabPpo.Location = new System.Drawing.Point(31, 191);
            this.chkTabPpo.Name = "chkTabPpo";
            this.chkTabPpo.Size = new System.Drawing.Size(128, 17);
            this.chkTabPpo.TabIndex = 29;
            this.chkTabPpo.Text = "Tabella causali cassa";
            this.chkTabPpo.UseVisualStyleBackColor = true;
            // 
            // chkTabUmi
            // 
            this.chkTabUmi.AutoSize = true;
            this.chkTabUmi.Location = new System.Drawing.Point(31, 160);
            this.chkTabUmi.Name = "chkTabUmi";
            this.chkTabUmi.Size = new System.Drawing.Size(133, 17);
            this.chkTabUmi.TabIndex = 27;
            this.chkTabUmi.Text = "Tabella umità di misura";
            this.chkTabUmi.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(26, 319);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 13);
            this.label1.TabIndex = 26;
            this.label1.Text = "Scartati";
            // 
            // dgv2
            // 
            this.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv2.Location = new System.Drawing.Point(26, 338);
            this.dgv2.Name = "dgv2";
            this.dgv2.Size = new System.Drawing.Size(617, 160);
            this.dgv2.TabIndex = 25;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(26, 517);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(623, 13);
            this.progressBar1.TabIndex = 24;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(519, 27);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(124, 150);
            this.btnOk.TabIndex = 23;
            this.btnOk.Text = "Genera";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(256, 27);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(234, 150);
            this.dgv1.TabIndex = 22;
            // 
            // chkTabRep
            // 
            this.chkTabRep.AutoSize = true;
            this.chkTabRep.Location = new System.Drawing.Point(31, 106);
            this.chkTabRep.Name = "chkTabRep";
            this.chkTabRep.Size = new System.Drawing.Size(124, 17);
            this.chkTabRep.TabIndex = 21;
            this.chkTabRep.Text = "Tabella reparti cassa";
            this.chkTabRep.UseVisualStyleBackColor = true;
            // 
            // chkAnaArt
            // 
            this.chkAnaArt.AutoSize = true;
            this.chkAnaArt.Location = new System.Drawing.Point(31, 44);
            this.chkAnaArt.Name = "chkAnaArt";
            this.chkAnaArt.Size = new System.Drawing.Size(99, 17);
            this.chkAnaArt.TabIndex = 20;
            this.chkAnaArt.Text = "Impianto articoli";
            this.chkAnaArt.UseVisualStyleBackColor = true;
            // 
            // chkTabIva
            // 
            this.chkTabIva.AutoSize = true;
            this.chkTabIva.Location = new System.Drawing.Point(31, 79);
            this.chkTabIva.Name = "chkTabIva";
            this.chkTabIva.Size = new System.Drawing.Size(81, 17);
            this.chkTabIva.TabIndex = 19;
            this.chkTabIva.Text = "Tabella IVA";
            this.chkTabIva.UseVisualStyleBackColor = true;
            // 
            // chkArtVar
            // 
            this.chkArtVar.AutoSize = true;
            this.chkArtVar.Location = new System.Drawing.Point(256, 232);
            this.chkArtVar.Name = "chkArtVar";
            this.chkArtVar.Size = new System.Drawing.Size(99, 17);
            this.chkArtVar.TabIndex = 34;
            this.chkArtVar.Text = "Articoli variati al";
            this.chkArtVar.UseVisualStyleBackColor = true;
            // 
            // dtpArtVar
            // 
            this.dtpArtVar.Location = new System.Drawing.Point(361, 227);
            this.dtpArtVar.Name = "dtpArtVar";
            this.dtpArtVar.Size = new System.Drawing.Size(200, 20);
            this.dtpArtVar.TabIndex = 35;
            // 
            // chkTabEcr
            // 
            this.chkTabEcr.AutoSize = true;
            this.chkTabEcr.Location = new System.Drawing.Point(31, 131);
            this.chkTabEcr.Name = "chkTabEcr";
            this.chkTabEcr.Size = new System.Drawing.Size(121, 17);
            this.chkTabEcr.TabIndex = 36;
            this.chkTabEcr.Text = "Tabella merceologie";
            this.chkTabEcr.UseVisualStyleBackColor = true;
            // 
            // frmUtyDivNegozi
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(684, 545);
            this.ControlBox = false;
            this.Controls.Add(this.chkTabEcr);
            this.Controls.Add(this.dtpArtVar);
            this.Controls.Add(this.chkArtVar);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.chkTabGrt);
            this.Controls.Add(this.chkTabUsr);
            this.Controls.Add(this.chkTabCam);
            this.Controls.Add(this.chkTabPpo);
            this.Controls.Add(this.chkTabUmi);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgv2);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.chkTabRep);
            this.Controls.Add(this.chkAnaArt);
            this.Controls.Add(this.chkTabIva);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyDivNegozi";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Divulgazioni ai negozi";
            this.Load += new System.EventHandler(this.frmUtyDivNegozi_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyDivNegozi_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.RadioButton rdbAll;
        private System.Windows.Forms.RadioButton rdbAtt;
        private System.Windows.Forms.CheckBox chkTabGrt;
        private System.Windows.Forms.CheckBox chkTabUsr;
        private System.Windows.Forms.CheckBox chkTabCam;
        private System.Windows.Forms.CheckBox chkTabPpo;
        private System.Windows.Forms.CheckBox chkTabUmi;
        private System.Windows.Forms.Label label1;
        private APOffice.APDataGridView dgv2;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button btnOk;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.CheckBox chkTabRep;
        private System.Windows.Forms.CheckBox chkAnaArt;
        private System.Windows.Forms.CheckBox chkTabIva;
        private System.Windows.Forms.CheckBox chkArtVar;
        private System.Windows.Forms.DateTimePicker dtpArtVar;
        private System.Windows.Forms.CheckBox chkTabEcr;
    }
}
