namespace APOffice
{
    partial class frmGesStat
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
            this.repartoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.iVAToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.stampaXArticoloToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.articoliToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.utilityToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.verificaPagamentiToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.verificaTotaliToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.verificaTotaliDaDettaglioToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpIni = new System.Windows.Forms.DateTimePicker();
            this.dtpFin = new System.Windows.Forms.DateTimePicker();
            this.btnEstrai = new System.Windows.Forms.Button();
            this.dgv2 = new APOffice.APDataGridView();
            this.lblMsg = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lblTotAcq = new System.Windows.Forms.Label();
            this.lblTotVen = new System.Windows.Forms.Label();
            this.lblTotSct = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.lblTotPro = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.repartoToolStripMenuItem,
            this.iVAToolStripMenuItem,
            this.stampaXArticoloToolStripMenuItem,
            this.utilityToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1009, 24);
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
            // repartoToolStripMenuItem
            // 
            this.repartoToolStripMenuItem.Name = "repartoToolStripMenuItem";
            this.repartoToolStripMenuItem.Size = new System.Drawing.Size(111, 20);
            this.repartoToolStripMenuItem.Text = "Stampa x Reparto";
            this.repartoToolStripMenuItem.Click += new System.EventHandler(this.repartoToolStripMenuItem_Click);
            // 
            // iVAToolStripMenuItem
            // 
            this.iVAToolStripMenuItem.Name = "iVAToolStripMenuItem";
            this.iVAToolStripMenuItem.Size = new System.Drawing.Size(87, 20);
            this.iVAToolStripMenuItem.Text = "Stampa x IVA";
            this.iVAToolStripMenuItem.Click += new System.EventHandler(this.iVAToolStripMenuItem_Click);
            // 
            // stampaXArticoloToolStripMenuItem
            // 
            this.stampaXArticoloToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.articoliToolStripMenuItem});
            this.stampaXArticoloToolStripMenuItem.Name = "stampaXArticoloToolStripMenuItem";
            this.stampaXArticoloToolStripMenuItem.Size = new System.Drawing.Size(112, 20);
            this.stampaXArticoloToolStripMenuItem.Text = "Stampe x Articolo";
            // 
            // articoliToolStripMenuItem
            // 
            this.articoliToolStripMenuItem.Name = "articoliToolStripMenuItem";
            this.articoliToolStripMenuItem.Size = new System.Drawing.Size(155, 22);
            this.articoliToolStripMenuItem.Text = "Articoli venduti";
            this.articoliToolStripMenuItem.Click += new System.EventHandler(this.articoliToolStripMenuItem_Click);
            // 
            // utilityToolStripMenuItem
            // 
            this.utilityToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.verificaPagamentiToolStripMenuItem,
            this.verificaTotaliToolStripMenuItem,
            this.verificaTotaliDaDettaglioToolStripMenuItem});
            this.utilityToolStripMenuItem.Name = "utilityToolStripMenuItem";
            this.utilityToolStripMenuItem.Size = new System.Drawing.Size(50, 20);
            this.utilityToolStripMenuItem.Text = "Utility";
            // 
            // verificaPagamentiToolStripMenuItem
            // 
            this.verificaPagamentiToolStripMenuItem.Name = "verificaPagamentiToolStripMenuItem";
            this.verificaPagamentiToolStripMenuItem.Size = new System.Drawing.Size(208, 22);
            this.verificaPagamentiToolStripMenuItem.Text = "Verifica pagamenti";
            this.verificaPagamentiToolStripMenuItem.Click += new System.EventHandler(this.verificaPagamentiToolStripMenuItem_Click);
            // 
            // verificaTotaliToolStripMenuItem
            // 
            this.verificaTotaliToolStripMenuItem.Name = "verificaTotaliToolStripMenuItem";
            this.verificaTotaliToolStripMenuItem.Size = new System.Drawing.Size(208, 22);
            this.verificaTotaliToolStripMenuItem.Text = "Verifica totali";
            this.verificaTotaliToolStripMenuItem.Click += new System.EventHandler(this.verificaTotaliToolStripMenuItem_Click);
            // 
            // verificaTotaliDaDettaglioToolStripMenuItem
            // 
            this.verificaTotaliDaDettaglioToolStripMenuItem.Name = "verificaTotaliDaDettaglioToolStripMenuItem";
            this.verificaTotaliDaDettaglioToolStripMenuItem.Size = new System.Drawing.Size(208, 22);
            this.verificaTotaliDaDettaglioToolStripMenuItem.Text = "Verifica totali da dettaglio";
            this.verificaTotaliDaDettaglioToolStripMenuItem.Click += new System.EventHandler(this.verificaTotaliDaDettaglioToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 74);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(985, 264);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(23, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Dal";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(255, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(15, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "al";
            // 
            // dtpIni
            // 
            this.dtpIni.Location = new System.Drawing.Point(42, 31);
            this.dtpIni.Name = "dtpIni";
            this.dtpIni.Size = new System.Drawing.Size(200, 20);
            this.dtpIni.TabIndex = 4;
            // 
            // dtpFin
            // 
            this.dtpFin.Location = new System.Drawing.Point(279, 32);
            this.dtpFin.Name = "dtpFin";
            this.dtpFin.Size = new System.Drawing.Size(200, 20);
            this.dtpFin.TabIndex = 5;
            // 
            // btnEstrai
            // 
            this.btnEstrai.Location = new System.Drawing.Point(485, 31);
            this.btnEstrai.Name = "btnEstrai";
            this.btnEstrai.Size = new System.Drawing.Size(75, 23);
            this.btnEstrai.TabIndex = 6;
            this.btnEstrai.Text = "Estrai";
            this.btnEstrai.UseVisualStyleBackColor = true;
            this.btnEstrai.Click += new System.EventHandler(this.btnEstrai_Click);
            // 
            // dgv2
            // 
            this.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv2.Location = new System.Drawing.Point(12, 358);
            this.dgv2.Name = "dgv2";
            this.dgv2.RowHeadersWidth = 20;
            this.dgv2.Size = new System.Drawing.Size(985, 272);
            this.dgv2.TabIndex = 9;
            this.dgv2.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv2_CellDoubleClick);
            // 
            // lblMsg
            // 
            this.lblMsg.AutoSize = true;
            this.lblMsg.Location = new System.Drawing.Point(13, 638);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(16, 13);
            this.lblMsg.TabIndex = 10;
            this.lblMsg.Text = "...";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(2, 659);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(1005, 5);
            this.progressBar1.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(812, 643);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(79, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Totale venduto";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(226, 645);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(76, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Totale acquisti";
            // 
            // lblTotAcq
            // 
            this.lblTotAcq.BackColor = System.Drawing.Color.White;
            this.lblTotAcq.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotAcq.Location = new System.Drawing.Point(303, 636);
            this.lblTotAcq.Name = "lblTotAcq";
            this.lblTotAcq.Size = new System.Drawing.Size(100, 23);
            this.lblTotAcq.TabIndex = 15;
            this.lblTotAcq.Text = "0";
            this.lblTotAcq.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblTotVen
            // 
            this.lblTotVen.BackColor = System.Drawing.Color.White;
            this.lblTotVen.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotVen.Location = new System.Drawing.Point(897, 634);
            this.lblTotVen.Name = "lblTotVen";
            this.lblTotVen.Size = new System.Drawing.Size(100, 23);
            this.lblTotVen.TabIndex = 8;
            this.lblTotVen.Text = "0";
            this.lblTotVen.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblTotSct
            // 
            this.lblTotSct.BackColor = System.Drawing.Color.White;
            this.lblTotSct.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotSct.Location = new System.Drawing.Point(693, 635);
            this.lblTotSct.Name = "lblTotSct";
            this.lblTotSct.Size = new System.Drawing.Size(100, 23);
            this.lblTotSct.TabIndex = 17;
            this.lblTotSct.Text = "0";
            this.lblTotSct.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(579, 644);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(111, 13);
            this.label6.TabIndex = 16;
            this.label6.Text = "Totale sconti su totale";
            // 
            // lblTotPro
            // 
            this.lblTotPro.BackColor = System.Drawing.Color.White;
            this.lblTotPro.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotPro.Location = new System.Drawing.Point(473, 637);
            this.lblTotPro.Name = "lblTotPro";
            this.lblTotPro.Size = new System.Drawing.Size(100, 23);
            this.lblTotPro.TabIndex = 18;
            this.lblTotPro.Text = "0";
            this.lblTotPro.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(406, 646);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(61, 13);
            this.label3.TabIndex = 19;
            this.label3.Text = "Totale prof.";
            // 
            // frmGesStat
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1009, 666);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblTotPro);
            this.Controls.Add(this.lblTotSct);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.lblTotAcq);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.dgv2);
            this.Controls.Add(this.lblTotVen);
            this.Controls.Add(this.btnEstrai);
            this.Controls.Add(this.dtpFin);
            this.Controls.Add(this.dtpIni);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesStat";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Estrazione dati";
            this.Load += new System.EventHandler(this.frmGesStat_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesStat_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpIni;
        private System.Windows.Forms.DateTimePicker dtpFin;
        private System.Windows.Forms.Button btnEstrai;
        private System.Windows.Forms.ToolStripMenuItem repartoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem iVAToolStripMenuItem;
        private APOffice.APDataGridView dgv2;
        private System.Windows.Forms.Label lblMsg;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ToolStripMenuItem stampaXArticoloToolStripMenuItem;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblTotAcq;
        private System.Windows.Forms.Label lblTotVen;
        private System.Windows.Forms.Label lblTotSct;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ToolStripMenuItem utilityToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem verificaPagamentiToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem verificaTotaliToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem verificaTotaliDaDettaglioToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem articoliToolStripMenuItem;
        private System.Windows.Forms.Label lblTotPro;
        private System.Windows.Forms.Label label3;
    }
}
