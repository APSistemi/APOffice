namespace APOffice
{
    partial class frmGesFattElettroniche
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
            this.dgv1 = new APOffice.APDataGridView();
            this.btnOk = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblCnfBan = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.lblCnfCfi = new System.Windows.Forms.Label();
            this.lblCnfAgp = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.lblCnfNcv = new System.Windows.Forms.Label();
            this.lblCnfIba = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.lblCnfInd = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lblCnfRag = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblCnfMai = new System.Windows.Forms.Label();
            this.lblCnfPiv = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.chkAll = new System.Windows.Forms.CheckBox();
            this.chkEle = new System.Windows.Forms.CheckBox();
            this.cmbYea = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.lblCnt = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(891, 24);
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
            // dgv1
            // 
            this.dgv1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(5, 297);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(881, 276);
            this.dgv1.TabIndex = 2;
            this.dgv1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellClick);
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            // 
            // btnOk
            // 
            this.btnOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOk.Location = new System.Drawing.Point(593, 219);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(293, 55);
            this.btnOk.TabIndex = 3;
            this.btnOk.Text = "Generazione file XML ed INVIO";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.lblCnfBan);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.lblCnfCfi);
            this.panel1.Controls.Add(this.lblCnfAgp);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.lblCnfNcv);
            this.panel1.Controls.Add(this.lblCnfIba);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.lblCnfInd);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.lblCnfRag);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.lblCnfMai);
            this.panel1.Controls.Add(this.lblCnfPiv);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(6, 27);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(880, 186);
            this.panel1.TabIndex = 4;
            // 
            // lblCnfBan
            // 
            this.lblCnfBan.BackColor = System.Drawing.Color.White;
            this.lblCnfBan.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCnfBan.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCnfBan.Location = new System.Drawing.Point(285, 110);
            this.lblCnfBan.Name = "lblCnfBan";
            this.lblCnfBan.Size = new System.Drawing.Size(586, 26);
            this.lblCnfBan.TabIndex = 23;
            this.lblCnfBan.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(495, 158);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(147, 13);
            this.label9.TabIndex = 22;
            this.label9.Text = "PEC dell\'agenzia delle entrate";
            // 
            // lblCnfCfi
            // 
            this.lblCnfCfi.BackColor = System.Drawing.Color.White;
            this.lblCnfCfi.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCnfCfi.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCnfCfi.Location = new System.Drawing.Point(717, 40);
            this.lblCnfCfi.Name = "lblCnfCfi";
            this.lblCnfCfi.Size = new System.Drawing.Size(153, 26);
            this.lblCnfCfi.TabIndex = 11;
            this.lblCnfCfi.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblCnfAgp
            // 
            this.lblCnfAgp.BackColor = System.Drawing.Color.White;
            this.lblCnfAgp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCnfAgp.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCnfAgp.Location = new System.Drawing.Point(52, 145);
            this.lblCnfAgp.Name = "lblCnfAgp";
            this.lblCnfAgp.Size = new System.Drawing.Size(424, 26);
            this.lblCnfAgp.TabIndex = 21;
            this.lblCnfAgp.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(8, 155);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(32, 13);
            this.label8.TabIndex = 20;
            this.label8.Text = "Dest.";
            // 
            // lblCnfNcv
            // 
            this.lblCnfNcv.BackColor = System.Drawing.Color.White;
            this.lblCnfNcv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCnfNcv.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCnfNcv.Location = new System.Drawing.Point(433, 39);
            this.lblCnfNcv.Name = "lblCnfNcv";
            this.lblCnfNcv.Size = new System.Drawing.Size(43, 26);
            this.lblCnfNcv.TabIndex = 19;
            this.lblCnfNcv.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblCnfIba
            // 
            this.lblCnfIba.BackColor = System.Drawing.Color.White;
            this.lblCnfIba.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCnfIba.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCnfIba.Location = new System.Drawing.Point(52, 109);
            this.lblCnfIba.Name = "lblCnfIba";
            this.lblCnfIba.Size = new System.Drawing.Size(224, 26);
            this.lblCnfIba.TabIndex = 18;
            this.lblCnfIba.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(8, 120);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(32, 13);
            this.label6.TabIndex = 17;
            this.label6.Text = "IBAN";
            // 
            // lblCnfInd
            // 
            this.lblCnfInd.BackColor = System.Drawing.Color.White;
            this.lblCnfInd.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCnfInd.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCnfInd.Location = new System.Drawing.Point(52, 39);
            this.lblCnfInd.Name = "lblCnfInd";
            this.lblCnfInd.Size = new System.Drawing.Size(375, 26);
            this.lblCnfInd.TabIndex = 16;
            this.lblCnfInd.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(5, 55);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(45, 13);
            this.label5.TabIndex = 15;
            this.label5.Text = "Indirizzo";
            // 
            // lblCnfRag
            // 
            this.lblCnfRag.BackColor = System.Drawing.Color.White;
            this.lblCnfRag.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCnfRag.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCnfRag.Location = new System.Drawing.Point(52, 7);
            this.lblCnfRag.Name = "lblCnfRag";
            this.lblCnfRag.Size = new System.Drawing.Size(424, 26);
            this.lblCnfRag.TabIndex = 14;
            this.lblCnfRag.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(5, 20);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "Nome";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(659, 49);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 16);
            this.label4.TabIndex = 12;
            this.label4.Text = "C.fiscale";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 87);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(28, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "PEC";
            // 
            // lblCnfMai
            // 
            this.lblCnfMai.BackColor = System.Drawing.Color.White;
            this.lblCnfMai.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCnfMai.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCnfMai.Location = new System.Drawing.Point(52, 73);
            this.lblCnfMai.Name = "lblCnfMai";
            this.lblCnfMai.Size = new System.Drawing.Size(424, 26);
            this.lblCnfMai.TabIndex = 9;
            this.lblCnfMai.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblCnfPiv
            // 
            this.lblCnfPiv.BackColor = System.Drawing.Color.White;
            this.lblCnfPiv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCnfPiv.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCnfPiv.Location = new System.Drawing.Point(717, 7);
            this.lblCnfPiv.Name = "lblCnfPiv";
            this.lblCnfPiv.Size = new System.Drawing.Size(153, 26);
            this.lblCnfPiv.TabIndex = 8;
            this.lblCnfPiv.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(664, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 16);
            this.label1.TabIndex = 7;
            this.label1.Text = "P.IVA";
            // 
            // chkAll
            // 
            this.chkAll.AutoSize = true;
            this.chkAll.Location = new System.Drawing.Point(7, 257);
            this.chkAll.Name = "chkAll";
            this.chkAll.Size = new System.Drawing.Size(92, 17);
            this.chkAll.TabIndex = 61;
            this.chkAll.Text = "Seleziona tutti";
            this.chkAll.UseVisualStyleBackColor = true;
            this.chkAll.Click += new System.EventHandler(this.chkAll_Click);
            // 
            // chkEle
            // 
            this.chkEle.AutoSize = true;
            this.chkEle.Location = new System.Drawing.Point(288, 228);
            this.chkEle.Name = "chkEle";
            this.chkEle.Size = new System.Drawing.Size(211, 17);
            this.chkEle.TabIndex = 62;
            this.chkEle.Text = "Visualizza anche documenti già laborati";
            this.chkEle.UseVisualStyleBackColor = true;
            this.chkEle.CheckedChanged += new System.EventHandler(this.chkEle_CheckedChanged);
            // 
            // cmbYea
            // 
            this.cmbYea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbYea.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbYea.FormattingEnabled = true;
            this.cmbYea.Location = new System.Drawing.Point(60, 218);
            this.cmbYea.Name = "cmbYea";
            this.cmbYea.Size = new System.Drawing.Size(116, 24);
            this.cmbYea.TabIndex = 64;
            this.cmbYea.SelectionChangeCommitted += new System.EventHandler(this.cmbYea_SelectionChangeCommitted);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(6, 228);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(39, 16);
            this.label7.TabIndex = 63;
            this.label7.Text = "Anno";
            // 
            // lblCnt
            // 
            this.lblCnt.AutoSize = true;
            this.lblCnt.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblCnt.Location = new System.Drawing.Point(859, 281);
            this.lblCnt.Name = "lblCnt";
            this.lblCnt.Size = new System.Drawing.Size(13, 13);
            this.lblCnt.TabIndex = 65;
            this.lblCnt.Text = "0";
            this.lblCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // frmGesFattElettroniche
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(891, 585);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            this.ControlBox = true;
            this.MinimumSize = new System.Drawing.Size(910, 620);
            this.Controls.Add(this.lblCnt);
            this.Controls.Add(this.cmbYea);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.chkEle);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.chkAll);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesFattElettroniche";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Invio fatture Elettroniche";
            this.Load += new System.EventHandler(this.frmGesFattElettroniche_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesFattElettroniche_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblCnfMai;
        private System.Windows.Forms.Label lblCnfPiv;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblCnfCfi;
        private System.Windows.Forms.Label lblCnfRag;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblCnfInd;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblCnfIba;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label lblCnfNcv;
        private System.Windows.Forms.CheckBox chkAll;
        private System.Windows.Forms.CheckBox chkEle;
        private System.Windows.Forms.ComboBox cmbYea;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label lblCnfAgp;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label lblCnt;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label lblCnfBan;
    }
}
