namespace APOffice
{
    partial class frmGesStatMerceologia
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
            this.stampaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.btnEstrai = new System.Windows.Forms.Button();
            this.dtpDtf = new System.Windows.Forms.DateTimePicker();
            this.dtpDti = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lblSco = new System.Windows.Forms.Label();
            this.lblQta = new System.Windows.Forms.Label();
            this.lblVen = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.cmbNeg = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.lblCdv = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.lblNiv = new System.Windows.Forms.Label();
            this.lblMar = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.lblVls = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.stampaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(914, 24);
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
            // stampaToolStripMenuItem
            // 
            this.stampaToolStripMenuItem.Name = "stampaToolStripMenuItem";
            this.stampaToolStripMenuItem.Size = new System.Drawing.Size(59, 20);
            this.stampaToolStripMenuItem.Text = "Stampa";
            this.stampaToolStripMenuItem.Click += new System.EventHandler(this.stampaToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.Location = new System.Drawing.Point(7, 71);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(901, 410);
            this.dgv1.TabIndex = 1;
            // 
            // btnEstrai
            // 
            this.btnEstrai.Location = new System.Drawing.Point(744, 37);
            this.btnEstrai.Name = "btnEstrai";
            this.btnEstrai.Size = new System.Drawing.Size(164, 23);
            this.btnEstrai.TabIndex = 12;
            this.btnEstrai.Text = "Estrai";
            this.btnEstrai.UseVisualStyleBackColor = true;
            this.btnEstrai.Click += new System.EventHandler(this.btnEstrai_Click);
            // 
            // dtpDtf
            // 
            this.dtpDtf.Location = new System.Drawing.Point(265, 36);
            this.dtpDtf.Name = "dtpDtf";
            this.dtpDtf.Size = new System.Drawing.Size(200, 20);
            this.dtpDtf.TabIndex = 11;
            // 
            // dtpDti
            // 
            this.dtpDti.Location = new System.Drawing.Point(10, 36);
            this.dtpDti.Name = "dtpDti";
            this.dtpDti.Size = new System.Drawing.Size(200, 20);
            this.dtpDti.TabIndex = 10;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(234, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(15, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "al";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(-24, 274);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(23, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Dal";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(133, 497);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(32, 13);
            this.label3.TabIndex = 15;
            this.label3.Text = "Pezzi";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(254, 498);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(47, 13);
            this.label4.TabIndex = 16;
            this.label4.Text = "Venduto";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 496);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(48, 13);
            this.label5.TabIndex = 17;
            this.label5.Text = "Scontrini";
            // 
            // lblSco
            // 
            this.lblSco.BackColor = System.Drawing.Color.White;
            this.lblSco.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblSco.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSco.Location = new System.Drawing.Point(55, 491);
            this.lblSco.Name = "lblSco";
            this.lblSco.Size = new System.Drawing.Size(72, 18);
            this.lblSco.TabIndex = 18;
            this.lblSco.Text = "0";
            this.lblSco.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblQta
            // 
            this.lblQta.BackColor = System.Drawing.Color.White;
            this.lblQta.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblQta.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQta.Location = new System.Drawing.Point(172, 492);
            this.lblQta.Name = "lblQta";
            this.lblQta.Size = new System.Drawing.Size(72, 18);
            this.lblQta.TabIndex = 19;
            this.lblQta.Text = "0";
            this.lblQta.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblVen
            // 
            this.lblVen.BackColor = System.Drawing.Color.White;
            this.lblVen.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblVen.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVen.Location = new System.Drawing.Point(303, 493);
            this.lblVen.Name = "lblVen";
            this.lblVen.Size = new System.Drawing.Size(100, 18);
            this.lblVen.TabIndex = 20;
            this.lblVen.Text = "0";
            this.lblVen.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(1, 541);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(912, 10);
            this.progressBar1.TabIndex = 21;
            // 
            // cmbNeg
            // 
            this.cmbNeg.FormattingEnabled = true;
            this.cmbNeg.Location = new System.Drawing.Point(553, 37);
            this.cmbNeg.Name = "cmbNeg";
            this.cmbNeg.Size = new System.Drawing.Size(164, 21);
            this.cmbNeg.TabIndex = 22;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(487, 42);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(46, 13);
            this.label6.TabIndex = 23;
            this.label6.Text = "Negozio";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(607, 497);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(93, 13);
            this.label7.TabIndex = 24;
            this.label7.Text = "Costo del venduto";
            // 
            // lblCdv
            // 
            this.lblCdv.BackColor = System.Drawing.Color.White;
            this.lblCdv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCdv.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCdv.Location = new System.Drawing.Point(703, 491);
            this.lblCdv.Name = "lblCdv";
            this.lblCdv.Size = new System.Drawing.Size(95, 18);
            this.lblCdv.TabIndex = 25;
            this.lblCdv.Text = "0";
            this.lblCdv.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(408, 496);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(94, 13);
            this.label8.TabIndex = 26;
            this.label8.Text = "Venduto netto IVA";
            // 
            // lblNiv
            // 
            this.lblNiv.BackColor = System.Drawing.Color.White;
            this.lblNiv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblNiv.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNiv.Location = new System.Drawing.Point(504, 490);
            this.lblNiv.Name = "lblNiv";
            this.lblNiv.Size = new System.Drawing.Size(97, 18);
            this.lblNiv.TabIndex = 27;
            this.lblNiv.Text = "0";
            this.lblNiv.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblMar
            // 
            this.lblMar.BackColor = System.Drawing.Color.White;
            this.lblMar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblMar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMar.Location = new System.Drawing.Point(843, 491);
            this.lblMar.Name = "lblMar";
            this.lblMar.Size = new System.Drawing.Size(62, 18);
            this.lblMar.TabIndex = 28;
            this.lblMar.Text = "0";
            this.lblMar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(803, 495);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(34, 13);
            this.label9.TabIndex = 29;
            this.label9.Text = "Marg.";
            // 
            // lblVls
            // 
            this.lblVls.BackColor = System.Drawing.Color.White;
            this.lblVls.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblVls.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVls.Location = new System.Drawing.Point(303, 518);
            this.lblVls.Name = "lblVls";
            this.lblVls.Size = new System.Drawing.Size(100, 18);
            this.lblVls.TabIndex = 30;
            this.lblVls.Text = "0";
            this.lblVls.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(198, 523);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(104, 13);
            this.label11.TabIndex = 31;
            this.label11.Text = "Venduto lordo sconti";
            // 
            // frmGesStatMerceologia
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(914, 552);
            this.ControlBox = false;
            this.Controls.Add(this.label11);
            this.Controls.Add(this.lblVls);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.lblMar);
            this.Controls.Add(this.lblNiv);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.lblCdv);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.cmbNeg);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.lblVen);
            this.Controls.Add(this.lblQta);
            this.Controls.Add(this.lblSco);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnEstrai);
            this.Controls.Add(this.dtpDtf);
            this.Controls.Add(this.dtpDti);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesStatMerceologia";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Stampa merceologie";
            this.Load += new System.EventHandler(this.frmGesStaReparto_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesStaReparto_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Button btnEstrai;
        private System.Windows.Forms.DateTimePicker dtpDtf;
        private System.Windows.Forms.DateTimePicker dtpDti;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolStripMenuItem stampaToolStripMenuItem;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblSco;
        private System.Windows.Forms.Label lblQta;
        private System.Windows.Forms.Label lblVen;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ComboBox cmbNeg;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label lblCdv;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label lblNiv;
        private System.Windows.Forms.Label lblMar;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label lblVls;
        private System.Windows.Forms.Label label11;
    }
}
