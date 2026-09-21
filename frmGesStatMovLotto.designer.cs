namespace APOffice
{
    partial class frmGesStatMovLotto
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
            this.letturaBilanceToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.lblArt = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblQkg = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.lblQkgSez = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblQkgTot = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lblQkgCar = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblQkgSca = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.pDFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.letturaBilanceToolStripMenuItem,
            this.excelToolStripMenuItem,
            this.pDFToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(922, 29);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // esciToolStripMenuItem
            // 
            this.esciToolStripMenuItem.Name = "esciToolStripMenuItem";
            this.esciToolStripMenuItem.Size = new System.Drawing.Size(48, 25);
            this.esciToolStripMenuItem.Text = "Esci";
            this.esciToolStripMenuItem.Click += new System.EventHandler(this.esciToolStripMenuItem_Click);
            // 
            // letturaBilanceToolStripMenuItem
            // 
            this.letturaBilanceToolStripMenuItem.Name = "letturaBilanceToolStripMenuItem";
            this.letturaBilanceToolStripMenuItem.Size = new System.Drawing.Size(12, 25);
            // 
            // excelToolStripMenuItem
            // 
            this.excelToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.excelToolStripMenuItem.Name = "excelToolStripMenuItem";
            this.excelToolStripMenuItem.Size = new System.Drawing.Size(56, 25);
            this.excelToolStripMenuItem.Text = "Excel";
            this.excelToolStripMenuItem.Click += new System.EventHandler(this.excelToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(7, 84);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(910, 446);
            this.dgv1.TabIndex = 2;
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            // 
            // lblArt
            // 
            this.lblArt.BackColor = System.Drawing.Color.Beige;
            this.lblArt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblArt.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblArt.Location = new System.Drawing.Point(0, 31);
            this.lblArt.Name = "lblArt";
            this.lblArt.Size = new System.Drawing.Size(636, 33);
            this.lblArt.TabIndex = 41;
            this.lblArt.Text = "label1";
            this.lblArt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(1, 656);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(921, 7);
            this.progressBar1.TabIndex = 48;
            // 
            // lblQkg
            // 
            this.lblQkg.BackColor = System.Drawing.Color.White;
            this.lblQkg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblQkg.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQkg.Location = new System.Drawing.Point(814, 36);
            this.lblQkg.Name = "lblQkg";
            this.lblQkg.Size = new System.Drawing.Size(100, 38);
            this.lblQkg.TabIndex = 50;
            this.lblQkg.Text = "0.00";
            this.lblQkg.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(749, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(44, 16);
            this.label1.TabIndex = 49;
            this.label1.Text = "Peso";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(12, 552);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(205, 16);
            this.label6.TabIndex = 55;
            this.label6.Text = "SEZIONATO: peso in entrata";
            // 
            // lblQkgSez
            // 
            this.lblQkgSez.BackColor = System.Drawing.Color.White;
            this.lblQkgSez.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblQkgSez.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQkgSez.Location = new System.Drawing.Point(227, 544);
            this.lblQkgSez.Name = "lblQkgSez";
            this.lblQkgSez.Size = new System.Drawing.Size(100, 29);
            this.lblQkgSez.TabIndex = 56;
            this.lblQkgSez.Text = "0.00";
            this.lblQkgSez.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.lblQkgTot);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.lblQkgCar);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.lblQkgSca);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Location = new System.Drawing.Point(7, 586);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(907, 67);
            this.panel1.TabIndex = 57;
            // 
            // lblQkgTot
            // 
            this.lblQkgTot.BackColor = System.Drawing.Color.White;
            this.lblQkgTot.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblQkgTot.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQkgTot.Location = new System.Drawing.Point(793, 13);
            this.lblQkgTot.Name = "lblQkgTot";
            this.lblQkgTot.Size = new System.Drawing.Size(100, 29);
            this.lblQkgTot.TabIndex = 58;
            this.lblQkgTot.Text = "0.00";
            this.lblQkgTot.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(661, 26);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(116, 16);
            this.label5.TabIndex = 57;
            this.label5.Text = "Peso rimanente";
            // 
            // lblQkgCar
            // 
            this.lblQkgCar.BackColor = System.Drawing.Color.White;
            this.lblQkgCar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblQkgCar.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQkgCar.Location = new System.Drawing.Point(221, 13);
            this.lblQkgCar.Name = "lblQkgCar";
            this.lblQkgCar.Size = new System.Drawing.Size(100, 29);
            this.lblQkgCar.TabIndex = 56;
            this.lblQkgCar.Text = "0.00";
            this.lblQkgCar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(1, 26);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(218, 16);
            this.label2.TabIndex = 55;
            this.label2.Text = "PORZIONATO: Peso in entrata";
            // 
            // lblQkgSca
            // 
            this.lblQkgSca.BackColor = System.Drawing.Color.White;
            this.lblQkgSca.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblQkgSca.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQkgSca.Location = new System.Drawing.Point(494, 18);
            this.lblQkgSca.Name = "lblQkgSca";
            this.lblQkgSca.Size = new System.Drawing.Size(100, 29);
            this.lblQkgSca.TabIndex = 54;
            this.lblQkgSca.Text = "0.00";
            this.lblQkgSca.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(373, 26);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(105, 16);
            this.label3.TabIndex = 53;
            this.label3.Text = "Peso in uscita";
            // 
            // pDFToolStripMenuItem
            // 
            this.pDFToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.pDFToolStripMenuItem.Name = "pDFToolStripMenuItem";
            this.pDFToolStripMenuItem.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.pDFToolStripMenuItem.Size = new System.Drawing.Size(50, 25);
            this.pDFToolStripMenuItem.Text = "PDF";
            this.pDFToolStripMenuItem.Click += new System.EventHandler(this.pDFToolStripMenuItem_Click);
            // 
            // frmGesStatMovLotto
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(922, 665);
            this.ControlBox = false;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.lblQkgSez);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.lblQkg);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.lblArt);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesStatMovLotto";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Movimenti per lotto";
            this.Load += new System.EventHandler(this.frmGesStaMovLegami_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesStaMovLegami_KeyDown);
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
        private System.Windows.Forms.Label lblArt;
        private System.Windows.Forms.ToolStripMenuItem letturaBilanceToolStripMenuItem;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblQkg;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolStripMenuItem excelToolStripMenuItem;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label lblQkgSez;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblQkgTot;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblQkgCar;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblQkgSca;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ToolStripMenuItem pDFToolStripMenuItem;
    }
}
