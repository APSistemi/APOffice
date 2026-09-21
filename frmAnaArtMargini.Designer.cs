namespace APOffice
{
    partial class frmAnaArtMargini
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
            this.pnlTop = new System.Windows.Forms.Panel();
            this.lblTitoloListino = new System.Windows.Forms.Label();
            this.cmbLisVen = new System.Windows.Forms.ComboBox();
            this.pnlCard = new System.Windows.Forms.Panel();
            this.lblHdrParam = new System.Windows.Forms.Label();
            this.lblHdrAttuale = new System.Windows.Forms.Label();
            this.lblHdrNuovo = new System.Windows.Forms.Label();
            this.pnlSepHdr = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.lblPrc = new System.Windows.Forms.Label();
            this.txtPrc = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.lblPrv = new System.Windows.Forms.Label();
            this.txtPrv = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.lblMav = new System.Windows.Forms.Label();
            this.txtMav = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.lblRic = new System.Windows.Forms.Label();
            this.txtRic = new System.Windows.Forms.TextBox();
            this.lblHint = new System.Windows.Forms.Label();
            this.pnlBottom = new System.Windows.Forms.Panel();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.btnOk = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.pnlTop.SuspendLayout();
            this.pnlCard.SuspendLayout();
            this.pnlBottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(624, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // esciToolStripMenuItem
            // 
            this.esciToolStripMenuItem.Name = "esciToolStripMenuItem";
            this.esciToolStripMenuItem.Size = new System.Drawing.Size(81, 20);
            this.esciToolStripMenuItem.Text = "Chiudi (Esc)";
            this.esciToolStripMenuItem.Click += new System.EventHandler(this.esciToolStripMenuItem_Click);
            // 
            // pnlTop
            // 
            this.pnlTop.BackColor = System.Drawing.Color.White;
            this.pnlTop.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlTop.Controls.Add(this.lblTitoloListino);
            this.pnlTop.Controls.Add(this.cmbLisVen);
            this.pnlTop.Location = new System.Drawing.Point(16, 32);
            this.pnlTop.Name = "pnlTop";
            this.pnlTop.Size = new System.Drawing.Size(592, 52);
            this.pnlTop.TabIndex = 1;
            // 
            // lblTitoloListino
            // 
            this.lblTitoloListino.AutoSize = true;
            this.lblTitoloListino.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitoloListino.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(41)))), ((int)(((byte)(59)))));
            this.lblTitoloListino.Location = new System.Drawing.Point(14, 15);
            this.lblTitoloListino.Name = "lblTitoloListino";
            this.lblTitoloListino.Size = new System.Drawing.Size(130, 19);
            this.lblTitoloListino.TabIndex = 0;
            this.lblTitoloListino.Text = "Listino di Vendita:";
            // 
            // cmbLisVen
            // 
            this.cmbLisVen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLisVen.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbLisVen.FormattingEnabled = true;
            this.cmbLisVen.Location = new System.Drawing.Point(152, 11);
            this.cmbLisVen.Name = "cmbLisVen";
            this.cmbLisVen.Size = new System.Drawing.Size(424, 28);
            this.cmbLisVen.TabIndex = 1;
            // 
            // pnlCard
            // 
            this.pnlCard.BackColor = System.Drawing.Color.White;
            this.pnlCard.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlCard.Controls.Add(this.lblHdrParam);
            this.pnlCard.Controls.Add(this.lblHdrAttuale);
            this.pnlCard.Controls.Add(this.lblHdrNuovo);
            this.pnlCard.Controls.Add(this.pnlSepHdr);
            this.pnlCard.Controls.Add(this.label1);
            this.pnlCard.Controls.Add(this.lblPrc);
            this.pnlCard.Controls.Add(this.txtPrc);
            this.pnlCard.Controls.Add(this.label2);
            this.pnlCard.Controls.Add(this.lblPrv);
            this.pnlCard.Controls.Add(this.txtPrv);
            this.pnlCard.Controls.Add(this.label3);
            this.pnlCard.Controls.Add(this.lblMav);
            this.pnlCard.Controls.Add(this.txtMav);
            this.pnlCard.Controls.Add(this.label4);
            this.pnlCard.Controls.Add(this.lblRic);
            this.pnlCard.Controls.Add(this.txtRic);
            this.pnlCard.Controls.Add(this.lblHint);
            this.pnlCard.Location = new System.Drawing.Point(16, 92);
            this.pnlCard.Name = "pnlCard";
            this.pnlCard.Size = new System.Drawing.Size(592, 290);
            this.pnlCard.TabIndex = 2;
            // 
            // lblHdrParam
            // 
            this.lblHdrParam.AutoSize = true;
            this.lblHdrParam.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHdrParam.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(116)))), ((int)(((byte)(139)))));
            this.lblHdrParam.Location = new System.Drawing.Point(16, 12);
            this.lblHdrParam.Name = "lblHdrParam";
            this.lblHdrParam.Size = new System.Drawing.Size(81, 15);
            this.lblHdrParam.TabIndex = 0;
            this.lblHdrParam.Text = "PARAMETRO";
            // 
            // lblHdrAttuale
            // 
            this.lblHdrAttuale.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHdrAttuale.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(116)))), ((int)(((byte)(139)))));
            this.lblHdrAttuale.Location = new System.Drawing.Point(234, 12);
            this.lblHdrAttuale.Name = "lblHdrAttuale";
            this.lblHdrAttuale.Size = new System.Drawing.Size(150, 15);
            this.lblHdrAttuale.TabIndex = 1;
            this.lblHdrAttuale.Text = "VALORE ATTUALE";
            this.lblHdrAttuale.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblHdrNuovo
            // 
            this.lblHdrNuovo.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHdrNuovo.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(99)))), ((int)(((byte)(235)))));
            this.lblHdrNuovo.Location = new System.Drawing.Point(408, 12);
            this.lblHdrNuovo.Name = "lblHdrNuovo";
            this.lblHdrNuovo.Size = new System.Drawing.Size(168, 15);
            this.lblHdrNuovo.TabIndex = 2;
            this.lblHdrNuovo.Text = "NUOVO VALORE";
            this.lblHdrNuovo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // pnlSepHdr
            // 
            this.pnlSepHdr.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(226)))), ((int)(((byte)(232)))), ((int)(((byte)(240)))));
            this.pnlSepHdr.Location = new System.Drawing.Point(14, 32);
            this.pnlSepHdr.Name = "pnlSepHdr";
            this.pnlSepHdr.Size = new System.Drawing.Size(562, 1);
            this.pnlSepHdr.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(41)))), ((int)(((byte)(59)))));
            this.label1.Location = new System.Drawing.Point(14, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(171, 19);
            this.label1.TabIndex = 4;
            this.label1.Text = "🏷  Costo Base Acquisto";
            // 
            // lblPrc
            // 
            this.lblPrc.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.lblPrc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblPrc.Font = new System.Drawing.Font("Segoe UI", 13.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPrc.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.lblPrc.Location = new System.Drawing.Point(234, 41);
            this.lblPrc.Name = "lblPrc";
            this.lblPrc.Size = new System.Drawing.Size(150, 32);
            this.lblPrc.TabIndex = 5;
            this.lblPrc.Text = "0,000";
            this.lblPrc.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtPrc
            // 
            this.txtPrc.BackColor = System.Drawing.Color.White;
            this.txtPrc.Font = new System.Drawing.Font("Segoe UI", 13.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPrc.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(15)))), ((int)(((byte)(23)))), ((int)(((byte)(42)))));
            this.txtPrc.Location = new System.Drawing.Point(408, 41);
            this.txtPrc.Name = "txtPrc";
            this.txtPrc.Size = new System.Drawing.Size(168, 31);
            this.txtPrc.TabIndex = 6;
            this.txtPrc.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtPrc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtVal_KeyDown);
            this.txtPrc.Leave += new System.EventHandler(this.txtVal_Leave);
            this.txtPrc.Enter += new System.EventHandler(this.txtVal_Enter);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(41)))), ((int)(((byte)(59)))));
            this.label2.Location = new System.Drawing.Point(14, 96);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(187, 19);
            this.label2.TabIndex = 7;
            this.label2.Text = "💰  Prezzo Vendita (Ivato)";
            // 
            // lblPrv
            // 
            this.lblPrv.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(239)))), ((int)(((byte)(246)))), ((int)(((byte)(255)))));
            this.lblPrv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblPrv.Font = new System.Drawing.Font("Segoe UI", 13.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPrv.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(64)))), ((int)(((byte)(175)))));
            this.lblPrv.Location = new System.Drawing.Point(234, 89);
            this.lblPrv.Name = "lblPrv";
            this.lblPrv.Size = new System.Drawing.Size(150, 32);
            this.lblPrv.TabIndex = 8;
            this.lblPrv.Text = "0,00";
            this.lblPrv.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtPrv
            // 
            this.txtPrv.BackColor = System.Drawing.Color.White;
            this.txtPrv.Font = new System.Drawing.Font("Segoe UI", 13.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPrv.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(64)))), ((int)(((byte)(175)))));
            this.txtPrv.Location = new System.Drawing.Point(408, 89);
            this.txtPrv.Name = "txtPrv";
            this.txtPrv.Size = new System.Drawing.Size(168, 31);
            this.txtPrv.TabIndex = 9;
            this.txtPrv.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtPrv.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtVal_KeyDown);
            this.txtPrv.Leave += new System.EventHandler(this.txtVal_Leave);
            this.txtPrv.Enter += new System.EventHandler(this.txtVal_Enter);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(41)))), ((int)(((byte)(59)))));
            this.label3.Location = new System.Drawing.Point(14, 144);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(109, 19);
            this.label3.TabIndex = 10;
            this.label3.Text = "📈  Margine %";
            // 
            // lblMav
            // 
            this.lblMav.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(236)))), ((int)(((byte)(253)))), ((int)(((byte)(245)))));
            this.lblMav.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblMav.Font = new System.Drawing.Font("Segoe UI", 13.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMav.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(6)))), ((int)(((byte)(95)))), ((int)(((byte)(70)))));
            this.lblMav.Location = new System.Drawing.Point(234, 137);
            this.lblMav.Name = "lblMav";
            this.lblMav.Size = new System.Drawing.Size(150, 32);
            this.lblMav.TabIndex = 11;
            this.lblMav.Text = "0,00";
            this.lblMav.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtMav
            // 
            this.txtMav.BackColor = System.Drawing.Color.White;
            this.txtMav.Font = new System.Drawing.Font("Segoe UI", 13.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMav.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(6)))), ((int)(((byte)(95)))), ((int)(((byte)(70)))));
            this.txtMav.Location = new System.Drawing.Point(408, 137);
            this.txtMav.Name = "txtMav";
            this.txtMav.Size = new System.Drawing.Size(168, 31);
            this.txtMav.TabIndex = 12;
            this.txtMav.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtMav.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtVal_KeyDown);
            this.txtMav.Leave += new System.EventHandler(this.txtVal_Leave);
            this.txtMav.Enter += new System.EventHandler(this.txtVal_Enter);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(41)))), ((int)(((byte)(59)))));
            this.label4.Location = new System.Drawing.Point(14, 192);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(110, 19);
            this.label4.TabIndex = 13;
            this.label4.Text = "📊  Ricarico %";
            // 
            // lblRic
            // 
            this.lblRic.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(242)))), ((int)(((byte)(255)))));
            this.lblRic.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblRic.Font = new System.Drawing.Font("Segoe UI", 13.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRic.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(55)))), ((int)(((byte)(48)))), ((int)(((byte)(163)))));
            this.lblRic.Location = new System.Drawing.Point(234, 185);
            this.lblRic.Name = "lblRic";
            this.lblRic.Size = new System.Drawing.Size(150, 32);
            this.lblRic.TabIndex = 14;
            this.lblRic.Text = "0,00";
            this.lblRic.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtRic
            // 
            this.txtRic.BackColor = System.Drawing.Color.White;
            this.txtRic.Font = new System.Drawing.Font("Segoe UI", 13.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRic.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(55)))), ((int)(((byte)(48)))), ((int)(((byte)(163)))));
            this.txtRic.Location = new System.Drawing.Point(408, 185);
            this.txtRic.Name = "txtRic";
            this.txtRic.Size = new System.Drawing.Size(168, 31);
            this.txtRic.TabIndex = 15;
            this.txtRic.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtRic.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtVal_KeyDown);
            this.txtRic.Leave += new System.EventHandler(this.txtVal_Leave);
            this.txtRic.Enter += new System.EventHandler(this.txtVal_Enter);
            // 
            // lblHint
            // 
            this.lblHint.Font = new System.Drawing.Font("Segoe UI", 8.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHint.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(116)))), ((int)(((byte)(139)))));
            this.lblHint.Location = new System.Drawing.Point(14, 238);
            this.lblHint.Name = "lblHint";
            this.lblHint.Size = new System.Drawing.Size(562, 36);
            this.lblHint.TabIndex = 16;
            this.lblHint.Text = "💡 Digita un nuovo valore e premi INVIO (o passa a un altro campo) per ricalcolar" +
    "e automaticamente gli altri parametri.";
            this.lblHint.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pnlBottom
            // 
            this.pnlBottom.Controls.Add(this.btnAnnulla);
            this.pnlBottom.Controls.Add(this.btnOk);
            this.pnlBottom.Location = new System.Drawing.Point(16, 392);
            this.pnlBottom.Name = "pnlBottom";
            this.pnlBottom.Size = new System.Drawing.Size(592, 52);
            this.pnlBottom.TabIndex = 3;
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.Font = new System.Drawing.Font("Segoe UI", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(0, 4);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(180, 44);
            this.btnAnnulla.TabIndex = 1;
            this.btnAnnulla.Text = "Annulla (Esc)";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            this.btnAnnulla.Click += new System.EventHandler(this.btnAnnulla_Click);
            // 
            // btnOk
            // 
            this.btnOk.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOk.Location = new System.Drawing.Point(302, 4);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(290, 44);
            this.btnOk.TabIndex = 0;
            this.btnOk.Text = "CONFERMA (F5 / Invio)";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // frmAnaArtMargini
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(244)))), ((int)(((byte)(246)))));
            this.ClientSize = new System.Drawing.Size(624, 456);
            this.Controls.Add(this.pnlBottom);
            this.Controls.Add(this.pnlCard);
            this.Controls.Add(this.pnlTop);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmAnaArtMargini";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Margini Articolo";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmAnaArtMargini_FormClosing);
            this.Load += new System.EventHandler(this.frmAnaArtMargini_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAnaArtMargini_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.pnlTop.ResumeLayout(false);
            this.pnlTop.PerformLayout();
            this.pnlCard.ResumeLayout(false);
            this.pnlCard.PerformLayout();
            this.pnlBottom.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Panel pnlTop;
        private System.Windows.Forms.Label lblTitoloListino;
        private System.Windows.Forms.ComboBox cmbLisVen;
        private System.Windows.Forms.Panel pnlCard;
        private System.Windows.Forms.Label lblHdrParam;
        private System.Windows.Forms.Label lblHdrAttuale;
        private System.Windows.Forms.Label lblHdrNuovo;
        private System.Windows.Forms.Panel pnlSepHdr;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblPrc;
        private System.Windows.Forms.TextBox txtPrc;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblPrv;
        private System.Windows.Forms.TextBox txtPrv;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblMav;
        private System.Windows.Forms.TextBox txtMav;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblRic;
        private System.Windows.Forms.TextBox txtRic;
        private System.Windows.Forms.Label lblHint;
        private System.Windows.Forms.Panel pnlBottom;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.Button btnOk;
    }
}