namespace APOffice
{
    partial class frmSeekArt
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
            this.nuovoArticoloToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtArtTas = new System.Windows.Forms.TextBox();
            this.pnlOpz1 = new System.Windows.Forms.Panel();
            this.rdbArtCel = new System.Windows.Forms.RadioButton();
            this.rdbArtCan = new System.Windows.Forms.RadioButton();
            this.rdbArtWeb = new System.Windows.Forms.RadioButton();
            this.cmbArtEti = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.chkVal = new System.Windows.Forms.CheckBox();
            this.cmbArtReb = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtArtPlu = new System.Windows.Forms.TextBox();
            this.txtArfCod = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.cmbArtRep = new System.Windows.Forms.ComboBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.cmbArtFor = new System.Windows.Forms.ComboBox();
            this.txtEanCod = new System.Windows.Forms.TextBox();
            this.txtArtDes = new System.Windows.Forms.TextBox();
            this.txtArtCod = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.lblTrovati = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.panel1.SuspendLayout();
            this.pnlOpz1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.nuovoArticoloToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(689, 24);
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
            // nuovoArticoloToolStripMenuItem
            // 
            this.nuovoArticoloToolStripMenuItem.Name = "nuovoArticoloToolStripMenuItem";
            this.nuovoArticoloToolStripMenuItem.Size = new System.Drawing.Size(98, 20);
            this.nuovoArticoloToolStripMenuItem.Text = "Nuovo articolo";
            this.nuovoArticoloToolStripMenuItem.Click += new System.EventHandler(this.nuovoArticoloToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 191);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(665, 221);
            this.dgv1.TabIndex = 17;
            this.dgv1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellContentClick);
            this.dgv1.DoubleClick += new System.EventHandler(this.dgv1_DoubleClick);
            this.dgv1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dgv1_KeyDown);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.txtArtTas);
            this.panel1.Controls.Add(this.pnlOpz1);
            this.panel1.Controls.Add(this.cmbArtEti);
            this.panel1.Controls.Add(this.label11);
            this.panel1.Controls.Add(this.chkVal);
            this.panel1.Controls.Add(this.cmbArtReb);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.txtArtPlu);
            this.panel1.Controls.Add(this.txtArfCod);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.cmbArtRep);
            this.panel1.Controls.Add(this.textBox4);
            this.panel1.Controls.Add(this.cmbArtFor);
            this.panel1.Controls.Add(this.txtEanCod);
            this.panel1.Controls.Add(this.txtArtDes);
            this.panel1.Controls.Add(this.txtArtCod);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(12, 28);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(665, 155);
            this.panel1.TabIndex = 18;
            // 
            // txtArtTas
            // 
            this.txtArtTas.Location = new System.Drawing.Point(620, 39);
            this.txtArtTas.MaxLength = 4;
            this.txtArtTas.Name = "txtArtTas";
            this.txtArtTas.Size = new System.Drawing.Size(30, 20);
            this.txtArtTas.TabIndex = 39;
            this.txtArtTas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArtTas_KeyDown);
            // 
            // pnlOpz1
            // 
            this.pnlOpz1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlOpz1.Controls.Add(this.rdbArtCel);
            this.pnlOpz1.Controls.Add(this.rdbArtCan);
            this.pnlOpz1.Controls.Add(this.rdbArtWeb);
            this.pnlOpz1.Location = new System.Drawing.Point(0, 125);
            this.pnlOpz1.Name = "pnlOpz1";
            this.pnlOpz1.Size = new System.Drawing.Size(661, 25);
            this.pnlOpz1.TabIndex = 38;
            // 
            // rdbArtCel
            // 
            this.rdbArtCel.AutoSize = true;
            this.rdbArtCel.Location = new System.Drawing.Point(84, 3);
            this.rdbArtCel.Name = "rdbArtCel";
            this.rdbArtCel.Size = new System.Drawing.Size(68, 17);
            this.rdbArtCel.TabIndex = 2;
            this.rdbArtCel.TabStop = true;
            this.rdbArtCel.Text = "Celiachia";
            this.rdbArtCel.UseVisualStyleBackColor = true;
            this.rdbArtCel.CheckedChanged += new System.EventHandler(this.rdbArtCel_CheckedChanged);
            // 
            // rdbArtCan
            // 
            this.rdbArtCan.AutoSize = true;
            this.rdbArtCan.Location = new System.Drawing.Point(196, 3);
            this.rdbArtCan.Name = "rdbArtCan";
            this.rdbArtCan.Size = new System.Drawing.Size(71, 17);
            this.rdbArtCan.TabIndex = 1;
            this.rdbArtCan.TabStop = true;
            this.rdbArtCan.Text = "Cancellati";
            this.rdbArtCan.UseVisualStyleBackColor = true;
            this.rdbArtCan.CheckedChanged += new System.EventHandler(this.rdbArtCan_CheckedChanged);
            // 
            // rdbArtWeb
            // 
            this.rdbArtWeb.AutoSize = true;
            this.rdbArtWeb.Location = new System.Drawing.Point(3, 4);
            this.rdbArtWeb.Name = "rdbArtWeb";
            this.rdbArtWeb.Size = new System.Drawing.Size(48, 17);
            this.rdbArtWeb.TabIndex = 0;
            this.rdbArtWeb.TabStop = true;
            this.rdbArtWeb.Text = "Web";
            this.rdbArtWeb.UseVisualStyleBackColor = true;
            this.rdbArtWeb.CheckedChanged += new System.EventHandler(this.rdbArtWeb_CheckedChanged);
            // 
            // cmbArtEti
            // 
            this.cmbArtEti.FormattingEnabled = true;
            this.cmbArtEti.Location = new System.Drawing.Point(330, 96);
            this.cmbArtEti.Name = "cmbArtEti";
            this.cmbArtEti.Size = new System.Drawing.Size(100, 21);
            this.cmbArtEti.TabIndex = 37;
            this.cmbArtEti.SelectionChangeCommitted += new System.EventHandler(this.cmbArtEti_SelectionChangeCommitted);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(254, 103);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(72, 13);
            this.label11.TabIndex = 36;
            this.label11.Text = "Tipo etichetta";
            // 
            // chkVal
            // 
            this.chkVal.AutoSize = true;
            this.chkVal.Location = new System.Drawing.Point(13, 99);
            this.chkVal.Name = "chkVal";
            this.chkVal.Size = new System.Drawing.Size(168, 17);
            this.chkVal.TabIndex = 35;
            this.chkVal.Text = "Valori per risultati > 500 articoli";
            this.chkVal.UseVisualStyleBackColor = true;
            // 
            // cmbArtReb
            // 
            this.cmbArtReb.FormattingEnabled = true;
            this.cmbArtReb.Location = new System.Drawing.Point(551, 68);
            this.cmbArtReb.Name = "cmbArtReb";
            this.cmbArtReb.Size = new System.Drawing.Size(100, 21);
            this.cmbArtReb.TabIndex = 34;
            this.cmbArtReb.SelectionChangeCommitted += new System.EventHandler(this.cmbArtReb_SelectionChangeCommitted);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(462, 75);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(84, 13);
            this.label9.TabIndex = 33;
            this.label9.Text = "Reparto bilancia";
            // 
            // txtArtPlu
            // 
            this.txtArtPlu.Location = new System.Drawing.Point(551, 39);
            this.txtArtPlu.MaxLength = 6;
            this.txtArtPlu.Name = "txtArtPlu";
            this.txtArtPlu.Size = new System.Drawing.Size(66, 20);
            this.txtArtPlu.TabIndex = 32;
            this.txtArtPlu.TextChanged += new System.EventHandler(this.txtArtPlu_TextChanged);
            this.txtArtPlu.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArtPlu_KeyDown);
            // 
            // txtArfCod
            // 
            this.txtArfCod.Location = new System.Drawing.Point(330, 39);
            this.txtArfCod.Name = "txtArfCod";
            this.txtArfCod.Size = new System.Drawing.Size(100, 20);
            this.txtArfCod.TabIndex = 31;
            this.txtArfCod.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArfCod_KeyDown);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(453, 46);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(100, 13);
            this.label8.TabIndex = 30;
            this.label8.Text = "Bilancia PLU/Tasto";
            // 
            // cmbArtRep
            // 
            this.cmbArtRep.FormattingEnabled = true;
            this.cmbArtRep.Location = new System.Drawing.Point(551, 97);
            this.cmbArtRep.Name = "cmbArtRep";
            this.cmbArtRep.Size = new System.Drawing.Size(100, 21);
            this.cmbArtRep.TabIndex = 29;
            this.cmbArtRep.SelectionChangeCommitted += new System.EventHandler(this.cmbArtRep_SelectionChangeCommitted);
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(330, 68);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(100, 20);
            this.textBox4.TabIndex = 28;
            // 
            // cmbArtFor
            // 
            this.cmbArtFor.FormattingEnabled = true;
            this.cmbArtFor.Location = new System.Drawing.Point(67, 67);
            this.cmbArtFor.Name = "cmbArtFor";
            this.cmbArtFor.Size = new System.Drawing.Size(167, 21);
            this.cmbArtFor.TabIndex = 27;
            this.cmbArtFor.SelectionChangeCommitted += new System.EventHandler(this.cmbArtFor_SelectionChangeCommitted);
            // 
            // txtEanCod
            // 
            this.txtEanCod.Location = new System.Drawing.Point(67, 35);
            this.txtEanCod.MaxLength = 13;
            this.txtEanCod.Name = "txtEanCod";
            this.txtEanCod.Size = new System.Drawing.Size(100, 20);
            this.txtEanCod.TabIndex = 26;
            this.txtEanCod.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtEanCod_KeyDown);
            // 
            // txtArtDes
            // 
            this.txtArtDes.Location = new System.Drawing.Point(330, 11);
            this.txtArtDes.Name = "txtArtDes";
            this.txtArtDes.Size = new System.Drawing.Size(321, 20);
            this.txtArtDes.TabIndex = 25;
            this.txtArtDes.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArtDes_KeyDown);
            // 
            // txtArtCod
            // 
            this.txtArtCod.Location = new System.Drawing.Point(67, 10);
            this.txtArtCod.MaxLength = 7;
            this.txtArtCod.Name = "txtArtCod";
            this.txtArtCod.Size = new System.Drawing.Size(100, 20);
            this.txtArtCod.TabIndex = 24;
            this.txtArtCod.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArtCod_KeyDown);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(500, 104);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(45, 13);
            this.label7.TabIndex = 23;
            this.label7.Text = "Reparto";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(10, 73);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(48, 13);
            this.label6.TabIndex = 22;
            this.label6.Text = "Fornitore";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(222, 42);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(100, 13);
            this.label5.TabIndex = 21;
            this.label5.Text = "Articolo del fornitore";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(206, 14);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(126, 13);
            this.label4.TabIndex = 20;
            this.label4.Text = "Descrizione (+ più parole)";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(259, 73);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 13);
            this.label3.TabIndex = 19;
            this.label3.Text = "Merceologia";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 13);
            this.label2.TabIndex = 18;
            this.label2.Text = "Barcode";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 13);
            this.label1.TabIndex = 17;
            this.label1.Text = "Codice";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(565, 417);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(40, 13);
            this.label10.TabIndex = 19;
            this.label10.Text = "Trovati";
            // 
            // lblTrovati
            // 
            this.lblTrovati.Location = new System.Drawing.Point(621, 412);
            this.lblTrovati.Name = "lblTrovati";
            this.lblTrovati.Size = new System.Drawing.Size(56, 23);
            this.lblTrovati.TabIndex = 20;
            this.lblTrovati.Text = "0";
            this.lblTrovati.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 433);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(689, 10);
            this.progressBar1.TabIndex = 21;
            // 
            // frmSeekArt
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(689, 445);
            this.ControlBox = false;
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.lblTrovati);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmSeekArt";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ricerca articolo";
            this.Load += new System.EventHandler(this.frmSeekArt_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSeekArt_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.pnlOpz1.ResumeLayout(false);
            this.pnlOpz1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox txtArtPlu;
        private System.Windows.Forms.TextBox txtArfCod;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cmbArtRep;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.ComboBox cmbArtFor;
        private System.Windows.Forms.TextBox txtEanCod;
        private System.Windows.Forms.TextBox txtArtDes;
        private System.Windows.Forms.TextBox txtArtCod;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolStripMenuItem nuovoArticoloToolStripMenuItem;
        private System.Windows.Forms.ComboBox cmbArtReb;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label lblTrovati;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.CheckBox chkVal;
        private System.Windows.Forms.ComboBox cmbArtEti;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Panel pnlOpz1;
        private System.Windows.Forms.RadioButton rdbArtWeb;
        private System.Windows.Forms.TextBox txtArtTas;
        private System.Windows.Forms.RadioButton rdbArtCan;
        private System.Windows.Forms.RadioButton rdbArtCel;
    }
}
