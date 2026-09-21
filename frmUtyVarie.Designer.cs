namespace APOffice
{
    partial class frmUtyVarie
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
            this.button1 = new System.Windows.Forms.Button();
            this.btnBkpOff = new System.Windows.Forms.Button();
            this.btnOptOff = new System.Windows.Forms.Button();
            this.btnShrOff = new System.Windows.Forms.Button();
            this.btnBkpLog = new System.Windows.Forms.Button();
            this.btnOptLog = new System.Windows.Forms.Button();
            this.btnShrLog = new System.Windows.Forms.Button();
            this.btnBkpSta = new System.Windows.Forms.Button();
            this.btnOptSta = new System.Windows.Forms.Button();
            this.btnShrSta = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblStatOff = new System.Windows.Forms.Label();
            this.lblStatLog = new System.Windows.Forms.Label();
            this.lblStatSta = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnDelNoEan = new System.Windows.Forms.Button();
            this.btnDelNoPrices = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.btnAlignApShop = new System.Windows.Forms.Button();
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
            this.menuStrip1.Size = new System.Drawing.Size(640, 24);
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
            // button1
            // 
            this.button1.Enabled = false;
            this.button1.Location = new System.Drawing.Point(29, 47);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(198, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Controllo cartelle";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnBkpOff
            // 
            this.btnBkpOff.Location = new System.Drawing.Point(29, 120);
            this.btnBkpOff.Name = "btnBkpOff";
            this.btnBkpOff.Size = new System.Drawing.Size(180, 40);
            this.btnBkpOff.TabIndex = 2;
            this.btnBkpOff.Text = "Backup APOffice";
            this.btnBkpOff.UseVisualStyleBackColor = true;
            this.btnBkpOff.Click += new System.EventHandler(this.btnDb_Click);
            // 
            // btnOptOff
            // 
            this.btnOptOff.Location = new System.Drawing.Point(215, 120);
            this.btnOptOff.Name = "btnOptOff";
            this.btnOptOff.Size = new System.Drawing.Size(180, 40);
            this.btnOptOff.TabIndex = 3;
            this.btnOptOff.Text = "Ottimizza APOffice";
            this.btnOptOff.UseVisualStyleBackColor = true;
            this.btnOptOff.Click += new System.EventHandler(this.btnDb_Click);
            // 
            // btnShrOff
            // 
            this.btnShrOff.Location = new System.Drawing.Point(401, 120);
            this.btnShrOff.Name = "btnShrOff";
            this.btnShrOff.Size = new System.Drawing.Size(180, 40);
            this.btnShrOff.TabIndex = 4;
            this.btnShrOff.Text = "Compatta APOffice";
            this.btnShrOff.UseVisualStyleBackColor = true;
            this.btnShrOff.Click += new System.EventHandler(this.btnDb_Click);
            // 
            // btnBkpLog
            // 
            this.btnBkpLog.Location = new System.Drawing.Point(29, 205);
            this.btnBkpLog.Name = "btnBkpLog";
            this.btnBkpLog.Size = new System.Drawing.Size(180, 40);
            this.btnBkpLog.TabIndex = 5;
            this.btnBkpLog.Text = "Backup APLog";
            this.btnBkpLog.UseVisualStyleBackColor = true;
            this.btnBkpLog.Click += new System.EventHandler(this.btnDb_Click);
            // 
            // btnOptLog
            // 
            this.btnOptLog.Location = new System.Drawing.Point(215, 205);
            this.btnOptLog.Name = "btnOptLog";
            this.btnOptLog.Size = new System.Drawing.Size(180, 40);
            this.btnOptLog.TabIndex = 6;
            this.btnOptLog.Text = "Ottimizza APLog";
            this.btnOptLog.UseVisualStyleBackColor = true;
            this.btnOptLog.Click += new System.EventHandler(this.btnDb_Click);
            // 
            // btnShrLog
            // 
            this.btnShrLog.Location = new System.Drawing.Point(401, 205);
            this.btnShrLog.Name = "btnShrLog";
            this.btnShrLog.Size = new System.Drawing.Size(180, 40);
            this.btnShrLog.TabIndex = 7;
            this.btnShrLog.Text = "Compatta APLog";
            this.btnShrLog.UseVisualStyleBackColor = true;
            this.btnShrLog.Click += new System.EventHandler(this.btnDb_Click);
            // 
            // btnBkpSta
            // 
            this.btnBkpSta.Location = new System.Drawing.Point(29, 290);
            this.btnBkpSta.Name = "btnBkpSta";
            this.btnBkpSta.Size = new System.Drawing.Size(180, 40);
            this.btnBkpSta.TabIndex = 8;
            this.btnBkpSta.Text = "Backup APStat";
            this.btnBkpSta.UseVisualStyleBackColor = true;
            this.btnBkpSta.Click += new System.EventHandler(this.btnDb_Click);
            // 
            // btnOptSta
            // 
            this.btnOptSta.Location = new System.Drawing.Point(215, 290);
            this.btnOptSta.Name = "btnOptSta";
            this.btnOptSta.Size = new System.Drawing.Size(180, 40);
            this.btnOptSta.TabIndex = 9;
            this.btnOptSta.Text = "Ottimizza APStat";
            this.btnOptSta.UseVisualStyleBackColor = true;
            this.btnOptSta.Click += new System.EventHandler(this.btnDb_Click);
            // 
            // btnShrSta
            // 
            this.btnShrSta.Location = new System.Drawing.Point(401, 290);
            this.btnShrSta.Name = "btnShrSta";
            this.btnShrSta.Size = new System.Drawing.Size(180, 40);
            this.btnShrSta.TabIndex = 10;
            this.btnShrSta.Text = "Compatta APStat";
            this.btnShrSta.UseVisualStyleBackColor = true;
            this.btnShrSta.Click += new System.EventHandler(this.btnDb_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(26, 100);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(130, 16);
            this.label1.TabIndex = 11;
            this.label1.Text = "Database APOffice";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(26, 185);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(111, 16);
            this.label2.TabIndex = 12;
            this.label2.Text = "Database APLog";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(26, 270);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(117, 16);
            this.label3.TabIndex = 13;
            this.label3.Text = "Database APStat";
            // 
            // lblStatOff
            // 
            this.lblStatOff.AutoSize = true;
            this.lblStatOff.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStatOff.ForeColor = System.Drawing.Color.DimGray;
            this.lblStatOff.Location = new System.Drawing.Point(30, 165);
            this.lblStatOff.Name = "lblStatOff";
            this.lblStatOff.Size = new System.Drawing.Size(217, 14);
            this.lblStatOff.TabIndex = 14;
            this.lblStatOff.Text = "Dimensione: - MB | Frag.: - %";
            // 
            // lblStatLog
            // 
            this.lblStatLog.AutoSize = true;
            this.lblStatLog.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStatLog.ForeColor = System.Drawing.Color.DimGray;
            this.lblStatLog.Location = new System.Drawing.Point(30, 250);
            this.lblStatLog.Name = "lblStatLog";
            this.lblStatLog.Size = new System.Drawing.Size(217, 14);
            this.lblStatLog.TabIndex = 15;
            this.lblStatLog.Text = "Dimensione: - MB | Frag.: - %";
            // 
            // lblStatSta
            // 
            this.lblStatSta.AutoSize = true;
            this.lblStatSta.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStatSta.ForeColor = System.Drawing.Color.DimGray;
            this.lblStatSta.Location = new System.Drawing.Point(30, 335);
            this.lblStatSta.Name = "lblStatSta";
            this.lblStatSta.Size = new System.Drawing.Size(217, 14);
            this.lblStatSta.TabIndex = 16;
            this.lblStatSta.Text = "Dimensione: - MB | Frag.: - %";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(26, 360);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(200, 16);
            this.label4.TabIndex = 17;
            this.label4.Text = "Pulizia Anagrafica Articoli";
            // 
            // btnDelNoEan
            // 
            this.btnDelNoEan.Location = new System.Drawing.Point(29, 385);
            this.btnDelNoEan.Name = "btnDelNoEan";
            this.btnDelNoEan.Size = new System.Drawing.Size(265, 42);
            this.btnDelNoEan.TabIndex = 18;
            this.btnDelNoEan.Text = "Elimina Articoli Senza Barcode";
            this.btnDelNoEan.UseVisualStyleBackColor = true;
            this.btnDelNoEan.Click += new System.EventHandler(this.btnDelNoEan_Click);
            // 
            // btnDelNoPrices
            // 
            this.btnDelNoPrices.Location = new System.Drawing.Point(316, 385);
            this.btnDelNoPrices.Name = "btnDelNoPrices";
            this.btnDelNoPrices.Size = new System.Drawing.Size(265, 42);
            this.btnDelNoPrices.TabIndex = 19;
            this.btnDelNoPrices.Text = "Elimina Articoli Senza Prezzi/Costi";
            this.btnDelNoPrices.UseVisualStyleBackColor = true;
            this.btnDelNoPrices.Click += new System.EventHandler(this.btnDelNoPrices_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(26, 440);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(306, 16);
            this.label5.TabIndex = 20;
            this.label5.Text = "Strumenti Allineamento Frontend APShop";
            // 
            // btnAlignApShop
            // 
            this.btnAlignApShop.Location = new System.Drawing.Point(29, 465);
            this.btnAlignApShop.Name = "btnAlignApShop";
            this.btnAlignApShop.Size = new System.Drawing.Size(552, 42);
            this.btnAlignApShop.TabIndex = 21;
            this.btnAlignApShop.Text = "Confronta e Allinea Anagrafica con APShop POS";
            this.btnAlignApShop.UseVisualStyleBackColor = true;
            this.btnAlignApShop.Click += new System.EventHandler(this.btnAlignApShop_Click);
            // 
            // frmUtyVarie
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(610, 530);
            this.Controls.Add(this.btnAlignApShop);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnDelNoPrices);
            this.Controls.Add(this.btnDelNoEan);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lblStatSta);
            this.Controls.Add(this.lblStatLog);
            this.Controls.Add(this.lblStatOff);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnShrSta);
            this.Controls.Add(this.btnOptSta);
            this.Controls.Add(this.btnBkpSta);
            this.Controls.Add(this.btnShrLog);
            this.Controls.Add(this.btnOptLog);
            this.Controls.Add(this.btnBkpLog);
            this.Controls.Add(this.btnShrOff);
            this.Controls.Add(this.btnOptOff);
            this.Controls.Add(this.btnBkpOff);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyVarie";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Utilità Varie";
            this.Load += new System.EventHandler(this.frmUtyVarie_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnBkpOff;
        private System.Windows.Forms.Button btnOptOff;
        private System.Windows.Forms.Button btnShrOff;
        private System.Windows.Forms.Button btnBkpLog;
        private System.Windows.Forms.Button btnOptLog;
        private System.Windows.Forms.Button btnShrLog;
        private System.Windows.Forms.Button btnBkpSta;
        private System.Windows.Forms.Button btnOptSta;
        private System.Windows.Forms.Button btnShrSta;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblStatOff;
        private System.Windows.Forms.Label lblStatLog;
        private System.Windows.Forms.Label lblStatSta;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnDelNoEan;
        private System.Windows.Forms.Button btnDelNoPrices;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnAlignApShop;
    }
}