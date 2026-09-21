namespace APOffice
{
    partial class frmAnaArtIngredienti
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
            this.stampaEtichetteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.nuovaRigaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.dtpDsc = new System.Windows.Forms.DateTimePicker();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.lblEan = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.lblImp = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lblPrv = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lblTar = new System.Windows.Forms.Label();
            this.lblPes = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblDes = new System.Windows.Forms.Label();
            this.txtQta = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.stampaEtichetteToolStripMenuItem,
            this.nuovaRigaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(788, 24);
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
            // stampaEtichetteToolStripMenuItem
            // 
            this.stampaEtichetteToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.stampaEtichetteToolStripMenuItem.Name = "stampaEtichetteToolStripMenuItem";
            this.stampaEtichetteToolStripMenuItem.Size = new System.Drawing.Size(108, 20);
            this.stampaEtichetteToolStripMenuItem.Text = "Stampa etichette";
            this.stampaEtichetteToolStripMenuItem.Click += new System.EventHandler(this.stampaEtichetteToolStripMenuItem_Click);
            // 
            // nuovaRigaToolStripMenuItem
            // 
            this.nuovaRigaToolStripMenuItem.Name = "nuovaRigaToolStripMenuItem";
            this.nuovaRigaToolStripMenuItem.Size = new System.Drawing.Size(77, 20);
            this.nuovaRigaToolStripMenuItem.Text = "Nuova riga";
            this.nuovaRigaToolStripMenuItem.Click += new System.EventHandler(this.nuovaRigaToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(3, 121);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(783, 411);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellClick);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.dtpDsc);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.lblEan);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.lblImp);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.lblPrv);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.lblTar);
            this.panel1.Controls.Add(this.lblPes);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.lblDes);
            this.panel1.Controls.Add(this.txtQta);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(5, 28);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(781, 87);
            this.panel1.TabIndex = 2;
            // 
            // dtpDsc
            // 
            this.dtpDsc.CustomFormat = "";
            this.dtpDsc.Location = new System.Drawing.Point(524, 52);
            this.dtpDsc.Name = "dtpDsc";
            this.dtpDsc.Size = new System.Drawing.Size(157, 20);
            this.dtpDsc.TabIndex = 32;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(464, 55);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(55, 13);
            this.label7.TabIndex = 31;
            this.label7.Text = "Scadenza";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(574, 21);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(26, 13);
            this.label8.TabIndex = 30;
            this.label8.Text = "Ean";
            // 
            // lblEan
            // 
            this.lblEan.BackColor = System.Drawing.Color.White;
            this.lblEan.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblEan.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblEan.Location = new System.Drawing.Point(606, 10);
            this.lblEan.Name = "lblEan";
            this.lblEan.Size = new System.Drawing.Size(170, 23);
            this.lblEan.TabIndex = 29;
            this.lblEan.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(350, 55);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(42, 13);
            this.label6.TabIndex = 28;
            this.label6.Text = "Importo";
            // 
            // lblImp
            // 
            this.lblImp.BackColor = System.Drawing.Color.White;
            this.lblImp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblImp.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblImp.Location = new System.Drawing.Point(400, 45);
            this.lblImp.Name = "lblImp";
            this.lblImp.Size = new System.Drawing.Size(55, 23);
            this.lblImp.TabIndex = 27;
            this.lblImp.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(232, 55);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(39, 13);
            this.label5.TabIndex = 26;
            this.label5.Text = "Prezzo";
            // 
            // lblPrv
            // 
            this.lblPrv.BackColor = System.Drawing.Color.White;
            this.lblPrv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblPrv.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPrv.Location = new System.Drawing.Point(282, 45);
            this.lblPrv.Name = "lblPrv";
            this.lblPrv.Size = new System.Drawing.Size(55, 23);
            this.lblPrv.TabIndex = 25;
            this.lblPrv.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(128, 55);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 13);
            this.label4.TabIndex = 24;
            this.label4.Text = "Tara";
            // 
            // lblTar
            // 
            this.lblTar.BackColor = System.Drawing.Color.White;
            this.lblTar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTar.Location = new System.Drawing.Point(165, 45);
            this.lblTar.Name = "lblTar";
            this.lblTar.Size = new System.Drawing.Size(55, 23);
            this.lblTar.TabIndex = 23;
            this.lblTar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblPes
            // 
            this.lblPes.BackColor = System.Drawing.Color.White;
            this.lblPes.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblPes.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPes.Location = new System.Drawing.Point(69, 45);
            this.lblPes.Name = "lblPes";
            this.lblPes.Size = new System.Drawing.Size(55, 23);
            this.lblPes.TabIndex = 22;
            this.lblPes.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(0, 55);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(31, 13);
            this.label3.TabIndex = 21;
            this.label3.Text = "Peso";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(0, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 13);
            this.label2.TabIndex = 20;
            this.label2.Text = "Descrizione";
            // 
            // lblDes
            // 
            this.lblDes.BackColor = System.Drawing.Color.White;
            this.lblDes.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblDes.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDes.Location = new System.Drawing.Point(69, 10);
            this.lblDes.Name = "lblDes";
            this.lblDes.Size = new System.Drawing.Size(491, 23);
            this.lblDes.TabIndex = 19;
            this.lblDes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtQta
            // 
            this.txtQta.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtQta.Location = new System.Drawing.Point(706, 57);
            this.txtQta.MaxLength = 3;
            this.txtQta.Name = "txtQta";
            this.txtQta.Size = new System.Drawing.Size(68, 26);
            this.txtQta.TabIndex = 18;
            this.txtQta.Text = "1";
            this.txtQta.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(693, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 13);
            this.label1.TabIndex = 17;
            this.label1.Text = "Numero etichette";
            // 
            // frmAnaArtIngredienti
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(788, 535);
            this.ControlBox = false;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmAnaArtIngredienti";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ingredienti";
            this.Load += new System.EventHandler(this.frmAnaArtIngredienti_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAnaArtIngredienti_KeyDown);
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
        private System.Windows.Forms.ToolStripMenuItem stampaEtichetteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem nuovaRigaToolStripMenuItem;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label lblEan;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label lblImp;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblPrv;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblTar;
        private System.Windows.Forms.Label lblPes;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblDes;
        private System.Windows.Forms.TextBox txtQta;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dtpDsc;
        private System.Windows.Forms.Label label7;
    }
}
