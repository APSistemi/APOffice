namespace APOffice
{
    partial class frmAnaFornitore
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
            this.label2 = new System.Windows.Forms.Label();
            this.chkCliAnn = new System.Windows.Forms.CheckBox();
            this.label11 = new System.Windows.Forms.Label();
            this.txtForCel = new System.Windows.Forms.TextBox();
            this.txtForFax = new System.Windows.Forms.TextBox();
            this.txtForTel = new System.Windows.Forms.TextBox();
            this.txtForCfi = new System.Windows.Forms.TextBox();
            this.txtForPiv = new System.Windows.Forms.TextBox();
            this.txtForMai = new System.Windows.Forms.TextBox();
            this.txtForPrv = new System.Windows.Forms.TextBox();
            this.txtForLoc = new System.Windows.Forms.TextBox();
            this.txtForCap = new System.Windows.Forms.TextBox();
            this.txtForInd = new System.Windows.Forms.TextBox();
            this.txtForDes = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtForCod = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.txtForNot = new System.Windows.Forms.TextBox();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(698, 24);
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
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(16, 44);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 20);
            this.label2.TabIndex = 81;
            this.label2.Text = "Codice";
            // 
            // chkCliAnn
            // 
            this.chkCliAnn.AutoSize = true;
            this.chkCliAnn.Location = new System.Drawing.Point(441, 55);
            this.chkCliAnn.Name = "chkCliAnn";
            this.chkCliAnn.Size = new System.Drawing.Size(15, 14);
            this.chkCliAnn.TabIndex = 80;
            this.chkCliAnn.UseVisualStyleBackColor = true;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(292, 49);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(69, 20);
            this.label11.TabIndex = 79;
            this.label11.Text = "Annullato";
            // 
            // txtForCel
            // 
            this.txtForCel.Location = new System.Drawing.Point(353, 282);
            this.txtForCel.MaxLength = 20;
            this.txtForCel.Name = "txtForCel";
            this.txtForCel.Size = new System.Drawing.Size(102, 20);
            this.txtForCel.TabIndex = 70;
            // 
            // txtForFax
            // 
            this.txtForFax.Location = new System.Drawing.Point(128, 307);
            this.txtForFax.MaxLength = 20;
            this.txtForFax.Name = "txtForFax";
            this.txtForFax.Size = new System.Drawing.Size(102, 20);
            this.txtForFax.TabIndex = 72;
            // 
            // txtForTel
            // 
            this.txtForTel.Location = new System.Drawing.Point(128, 281);
            this.txtForTel.MaxLength = 20;
            this.txtForTel.Name = "txtForTel";
            this.txtForTel.Size = new System.Drawing.Size(102, 20);
            this.txtForTel.TabIndex = 69;
            // 
            // txtForCfi
            // 
            this.txtForCfi.Location = new System.Drawing.Point(353, 252);
            this.txtForCfi.MaxLength = 16;
            this.txtForCfi.Name = "txtForCfi";
            this.txtForCfi.Size = new System.Drawing.Size(102, 20);
            this.txtForCfi.TabIndex = 67;
            this.txtForCfi.Validated += new System.EventHandler(this.txtForCfi_Validated);
            // 
            // txtForPiv
            // 
            this.txtForPiv.Location = new System.Drawing.Point(128, 250);
            this.txtForPiv.MaxLength = 11;
            this.txtForPiv.Name = "txtForPiv";
            this.txtForPiv.Size = new System.Drawing.Size(102, 20);
            this.txtForPiv.TabIndex = 64;
            this.txtForPiv.Validated += new System.EventHandler(this.txtForPiv_Validated);
            // 
            // txtForMai
            // 
            this.txtForMai.Location = new System.Drawing.Point(353, 309);
            this.txtForMai.MaxLength = 50;
            this.txtForMai.Name = "txtForMai";
            this.txtForMai.Size = new System.Drawing.Size(327, 20);
            this.txtForMai.TabIndex = 74;
            // 
            // txtForPrv
            // 
            this.txtForPrv.Location = new System.Drawing.Point(353, 207);
            this.txtForPrv.MaxLength = 2;
            this.txtForPrv.Name = "txtForPrv";
            this.txtForPrv.Size = new System.Drawing.Size(102, 20);
            this.txtForPrv.TabIndex = 62;
            // 
            // txtForLoc
            // 
            this.txtForLoc.Location = new System.Drawing.Point(128, 171);
            this.txtForLoc.MaxLength = 50;
            this.txtForLoc.Name = "txtForLoc";
            this.txtForLoc.Size = new System.Drawing.Size(327, 20);
            this.txtForLoc.TabIndex = 59;
            // 
            // txtForCap
            // 
            this.txtForCap.Location = new System.Drawing.Point(128, 207);
            this.txtForCap.MaxLength = 5;
            this.txtForCap.Name = "txtForCap";
            this.txtForCap.Size = new System.Drawing.Size(102, 20);
            this.txtForCap.TabIndex = 61;
            // 
            // txtForInd
            // 
            this.txtForInd.Location = new System.Drawing.Point(128, 136);
            this.txtForInd.MaxLength = 50;
            this.txtForInd.Name = "txtForInd";
            this.txtForInd.Size = new System.Drawing.Size(327, 20);
            this.txtForInd.TabIndex = 58;
            // 
            // txtForDes
            // 
            this.txtForDes.Location = new System.Drawing.Point(128, 102);
            this.txtForDes.MaxLength = 70;
            this.txtForDes.Name = "txtForDes";
            this.txtForDes.Size = new System.Drawing.Size(552, 20);
            this.txtForDes.TabIndex = 57;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(16, 208);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(48, 20);
            this.label12.TabIndex = 78;
            this.label12.Text = "C.A.P.";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.Location = new System.Drawing.Point(243, 250);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(100, 20);
            this.label16.TabIndex = 77;
            this.label16.Text = "Codice Fiscale";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(16, 100);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(110, 20);
            this.label10.TabIndex = 76;
            this.label10.Text = "Ragione Sociale";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(16, 136);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(62, 20);
            this.label9.TabIndex = 75;
            this.label9.Text = "Indirizzo";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(16, 172);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(37, 20);
            this.label8.TabIndex = 73;
            this.label8.Text = "Città";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(243, 213);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(68, 20);
            this.label7.TabIndex = 71;
            this.label7.Text = "Provincia";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(16, 247);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(40, 20);
            this.label6.TabIndex = 68;
            this.label6.Text = "P.Iva";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(16, 280);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(63, 20);
            this.label5.TabIndex = 66;
            this.label5.Text = "Telefono";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(16, 316);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(30, 20);
            this.label4.TabIndex = 65;
            this.label4.Text = "Fax";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(243, 282);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 20);
            this.label3.TabIndex = 63;
            this.label3.Text = "Cellulare";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(243, 307);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(49, 20);
            this.label1.TabIndex = 60;
            this.label1.Text = "E-mail";
            // 
            // txtForCod
            // 
            this.txtForCod.Enabled = false;
            this.txtForCod.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtForCod.Location = new System.Drawing.Point(128, 44);
            this.txtForCod.Name = "txtForCod";
            this.txtForCod.Size = new System.Drawing.Size(100, 22);
            this.txtForCod.TabIndex = 82;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(16, 341);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(37, 20);
            this.label13.TabIndex = 83;
            this.label13.Text = "Note";
            // 
            // txtForNot
            // 
            this.txtForNot.Location = new System.Drawing.Point(128, 337);
            this.txtForNot.MaxLength = 100;
            this.txtForNot.Name = "txtForNot";
            this.txtForNot.Size = new System.Drawing.Size(552, 20);
            this.txtForNot.TabIndex = 84;
            // 
            // frmAnaFornitore
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(698, 402);
            this.Controls.Add(this.txtForNot);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.txtForCod);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.chkCliAnn);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.txtForCel);
            this.Controls.Add(this.txtForFax);
            this.Controls.Add(this.txtForTel);
            this.Controls.Add(this.txtForCfi);
            this.Controls.Add(this.txtForPiv);
            this.Controls.Add(this.txtForMai);
            this.Controls.Add(this.txtForPrv);
            this.Controls.Add(this.txtForLoc);
            this.Controls.Add(this.txtForCap);
            this.Controls.Add(this.txtForInd);
            this.Controls.Add(this.txtForDes);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmAnaFornitore";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Anagrafica fornitore";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmAnaFornitore_FormClosing);
            this.Load += new System.EventHandler(this.frmAnaFornitore_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAnaFornitore_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox chkCliAnn;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox txtForCel;
        private System.Windows.Forms.TextBox txtForFax;
        private System.Windows.Forms.TextBox txtForTel;
        private System.Windows.Forms.TextBox txtForCfi;
        private System.Windows.Forms.TextBox txtForPiv;
        private System.Windows.Forms.TextBox txtForMai;
        private System.Windows.Forms.TextBox txtForPrv;
        private System.Windows.Forms.TextBox txtForLoc;
        private System.Windows.Forms.TextBox txtForCap;
        private System.Windows.Forms.TextBox txtForInd;
        private System.Windows.Forms.TextBox txtForDes;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtForCod;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtForNot;
    }
}