namespace APOffice
{
    partial class frmUtyDivPos
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
            this.chkTabIva = new System.Windows.Forms.CheckBox();
            this.chkAnaArt = new System.Windows.Forms.CheckBox();
            this.chkTabRep = new System.Windows.Forms.CheckBox();
            this.dgv1 = new APOffice.APDataGridView();
            this.btnOk = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.dgv2 = new APOffice.APDataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.chkTabUmi = new System.Windows.Forms.CheckBox();
            this.btnRep = new System.Windows.Forms.Button();
            this.chkTabPpo = new System.Windows.Forms.CheckBox();
            this.chkTabCam = new System.Windows.Forms.CheckBox();
            this.chkTabUsr = new System.Windows.Forms.CheckBox();
            this.chkTabGrt = new System.Windows.Forms.CheckBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.rdbAll = new System.Windows.Forms.RadioButton();
            this.rdbAtt = new System.Windows.Forms.RadioButton();
            this.button1 = new System.Windows.Forms.Button();
            this.lblMsg = new System.Windows.Forms.Label();
            this.chkTabMdo = new System.Windows.Forms.CheckBox();
            this.chkTabSta = new System.Windows.Forms.CheckBox();
            this.chkTabPag = new System.Windows.Forms.CheckBox();
            this.chkTabNot = new System.Windows.Forms.CheckBox();
            this.chkTabPos = new System.Windows.Forms.CheckBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(636, 24);
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
            // chkTabIva
            // 
            this.chkTabIva.AutoSize = true;
            this.chkTabIva.Location = new System.Drawing.Point(12, 63);
            this.chkTabIva.Name = "chkTabIva";
            this.chkTabIva.Size = new System.Drawing.Size(81, 17);
            this.chkTabIva.TabIndex = 1;
            this.chkTabIva.Text = "Tabella IVA";
            this.chkTabIva.UseVisualStyleBackColor = true;
            // 
            // chkAnaArt
            // 
            this.chkAnaArt.AutoSize = true;
            this.chkAnaArt.Location = new System.Drawing.Point(12, 35);
            this.chkAnaArt.Name = "chkAnaArt";
            this.chkAnaArt.Size = new System.Drawing.Size(99, 17);
            this.chkAnaArt.TabIndex = 2;
            this.chkAnaArt.Text = "Impianto articoli";
            this.chkAnaArt.UseVisualStyleBackColor = true;
            // 
            // chkTabRep
            // 
            this.chkTabRep.AutoSize = true;
            this.chkTabRep.Location = new System.Drawing.Point(12, 86);
            this.chkTabRep.Name = "chkTabRep";
            this.chkTabRep.Size = new System.Drawing.Size(124, 17);
            this.chkTabRep.TabIndex = 3;
            this.chkTabRep.Text = "Tabella reparti cassa";
            this.chkTabRep.UseVisualStyleBackColor = true;
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(237, 38);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(234, 150);
            this.dgv1.TabIndex = 4;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(500, 38);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(124, 150);
            this.btnOk.TabIndex = 5;
            this.btnOk.Text = "Genera";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(7, 531);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(623, 13);
            this.progressBar1.TabIndex = 6;
            // 
            // dgv2
            // 
            this.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv2.Location = new System.Drawing.Point(7, 352);
            this.dgv2.Name = "dgv2";
            this.dgv2.Size = new System.Drawing.Size(617, 160);
            this.dgv2.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 337);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Scartati";
            // 
            // chkTabUmi
            // 
            this.chkTabUmi.AutoSize = true;
            this.chkTabUmi.Location = new System.Drawing.Point(12, 109);
            this.chkTabUmi.Name = "chkTabUmi";
            this.chkTabUmi.Size = new System.Drawing.Size(133, 17);
            this.chkTabUmi.TabIndex = 10;
            this.chkTabUmi.Text = "Tabella umità di misura";
            this.chkTabUmi.UseVisualStyleBackColor = true;
            // 
            // btnRep
            // 
            this.btnRep.Location = new System.Drawing.Point(500, 193);
            this.btnRep.Name = "btnRep";
            this.btnRep.Size = new System.Drawing.Size(124, 150);
            this.btnRep.TabIndex = 11;
            this.btnRep.Text = "Generazione articoli reparto";
            this.btnRep.UseVisualStyleBackColor = true;
            this.btnRep.Click += new System.EventHandler(this.btnRep_Click);
            // 
            // chkTabPpo
            // 
            this.chkTabPpo.AutoSize = true;
            this.chkTabPpo.Location = new System.Drawing.Point(12, 132);
            this.chkTabPpo.Name = "chkTabPpo";
            this.chkTabPpo.Size = new System.Drawing.Size(128, 17);
            this.chkTabPpo.TabIndex = 12;
            this.chkTabPpo.Text = "Tabella causali cassa";
            this.chkTabPpo.UseVisualStyleBackColor = true;
            // 
            // chkTabCam
            // 
            this.chkTabCam.AutoSize = true;
            this.chkTabCam.Location = new System.Drawing.Point(12, 201);
            this.chkTabCam.Name = "chkTabCam";
            this.chkTabCam.Size = new System.Drawing.Size(146, 17);
            this.chkTabCam.TabIndex = 13;
            this.chkTabCam.Text = "Tabella campagne fidelity";
            this.chkTabCam.UseVisualStyleBackColor = true;
            // 
            // chkTabUsr
            // 
            this.chkTabUsr.AutoSize = true;
            this.chkTabUsr.Location = new System.Drawing.Point(12, 155);
            this.chkTabUsr.Name = "chkTabUsr";
            this.chkTabUsr.Size = new System.Drawing.Size(90, 17);
            this.chkTabUsr.TabIndex = 14;
            this.chkTabUsr.Text = "Tabella utenti";
            this.chkTabUsr.UseVisualStyleBackColor = true;
            // 
            // chkTabGrt
            // 
            this.chkTabGrt.AutoSize = true;
            this.chkTabGrt.Location = new System.Drawing.Point(12, 178);
            this.chkTabGrt.Name = "chkTabGrt";
            this.chkTabGrt.Size = new System.Drawing.Size(130, 17);
            this.chkTabGrt.TabIndex = 15;
            this.chkTabGrt.Text = "Tabella gruppi tessere";
            this.chkTabGrt.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.rdbAll);
            this.panel2.Controls.Add(this.rdbAtt);
            this.panel2.Location = new System.Drawing.Point(107, 28);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(111, 35);
            this.panel2.TabIndex = 17;
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
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(285, 245);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 18;
            this.button1.Text = "Prova beep";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // lblMsg
            // 
            this.lblMsg.AutoSize = true;
            this.lblMsg.Location = new System.Drawing.Point(7, 514);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(16, 13);
            this.lblMsg.TabIndex = 19;
            this.lblMsg.Text = "...";
            // 
            // chkTabMdo
            // 
            this.chkTabMdo.AutoSize = true;
            this.chkTabMdo.Location = new System.Drawing.Point(12, 224);
            this.chkTabMdo.Name = "chkTabMdo";
            this.chkTabMdo.Size = new System.Drawing.Size(139, 17);
            this.chkTabMdo.TabIndex = 20;
            this.chkTabMdo.Text = "Tabella Mov/documenti";
            this.chkTabMdo.UseVisualStyleBackColor = true;
            this.chkTabMdo.CheckedChanged += new System.EventHandler(this.chkTabMdo_CheckedChanged);
            // 
            // chkTabSta
            // 
            this.chkTabSta.AutoSize = true;
            this.chkTabSta.Location = new System.Drawing.Point(12, 269);
            this.chkTabSta.Name = "chkTabSta";
            this.chkTabSta.Size = new System.Drawing.Size(89, 17);
            this.chkTabSta.TabIndex = 21;
            this.chkTabSta.Text = "Tabella Stato";
            this.chkTabSta.UseVisualStyleBackColor = true;
            // 
            // chkTabPag
            // 
            this.chkTabPag.AutoSize = true;
            this.chkTabPag.Location = new System.Drawing.Point(12, 247);
            this.chkTabPag.Name = "chkTabPag";
            this.chkTabPag.Size = new System.Drawing.Size(114, 17);
            this.chkTabPag.TabIndex = 22;
            this.chkTabPag.Text = "Tabella Pagamenti";
            this.chkTabPag.UseVisualStyleBackColor = true;
            // 
            // chkTabNot
            // 
            this.chkTabNot.AutoSize = true;
            this.chkTabNot.Location = new System.Drawing.Point(12, 292);
            this.chkTabNot.Name = "chkTabNot";
            this.chkTabNot.Size = new System.Drawing.Size(87, 17);
            this.chkTabNot.TabIndex = 23;
            this.chkTabNot.Text = "Tabella Note";
            this.chkTabNot.UseVisualStyleBackColor = true;
            // 
            // chkTabPos
            // 
            this.chkTabPos.AutoSize = true;
            this.chkTabPos.Location = new System.Drawing.Point(12, 315);
            this.chkTabPos.Name = "chkTabPos";
            this.chkTabPos.Size = new System.Drawing.Size(86, 17);
            this.chkTabPos.TabIndex = 24;
            this.chkTabPos.Text = "Tabella POS";
            this.chkTabPos.UseVisualStyleBackColor = true;
            // 
            // frmUtyDivPos
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(636, 548);
            this.ControlBox = false;
            this.Controls.Add(this.chkTabPos);
            this.Controls.Add(this.chkTabNot);
            this.Controls.Add(this.chkTabPag);
            this.Controls.Add(this.chkTabSta);
            this.Controls.Add(this.chkTabMdo);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.chkTabGrt);
            this.Controls.Add(this.chkTabUsr);
            this.Controls.Add(this.chkTabCam);
            this.Controls.Add(this.chkTabPpo);
            this.Controls.Add(this.btnRep);
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
            this.Name = "frmUtyDivPos";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Divulgazioni ApShop";
            this.Load += new System.EventHandler(this.frmGesDivPos_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyDivPos_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.CheckBox chkTabIva;
        private System.Windows.Forms.CheckBox chkAnaArt;
        private System.Windows.Forms.CheckBox chkTabRep;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.ProgressBar progressBar1;
        private APOffice.APDataGridView dgv2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkTabUmi;
        private System.Windows.Forms.Button btnRep;
        private System.Windows.Forms.CheckBox chkTabPpo;
        private System.Windows.Forms.CheckBox chkTabCam;
        private System.Windows.Forms.CheckBox chkTabUsr;
        private System.Windows.Forms.CheckBox chkTabGrt;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.RadioButton rdbAll;
        private System.Windows.Forms.RadioButton rdbAtt;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label lblMsg;
        private System.Windows.Forms.CheckBox chkTabMdo;
        private System.Windows.Forms.CheckBox chkTabSta;
        private System.Windows.Forms.CheckBox chkTabPag;
        private System.Windows.Forms.CheckBox chkTabNot;
        private System.Windows.Forms.CheckBox chkTabPos;
    }
}
