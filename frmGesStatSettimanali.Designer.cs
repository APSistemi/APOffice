namespace APOffice
{
    partial class frmGesStatSettimanali
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
            this.pDFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.panel1 = new System.Windows.Forms.Panel();
            this.rdbFas = new System.Windows.Forms.RadioButton();
            this.rdbRep = new System.Windows.Forms.RadioButton();
            this.btnEstrai = new System.Windows.Forms.Button();
            this.dtpDay = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.cmbNeg = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.lblDti = new System.Windows.Forms.Label();
            this.lblDtf = new System.Windows.Forms.Label();
            this.dtpDti = new System.Windows.Forms.DateTimePicker();
            this.dtpDtf = new System.Windows.Forms.DateTimePicker();
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
            this.pDFToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(926, 24);
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
            // pDFToolStripMenuItem
            // 
            this.pDFToolStripMenuItem.Name = "pDFToolStripMenuItem";
            this.pDFToolStripMenuItem.Size = new System.Drawing.Size(40, 20);
            this.pDFToolStripMenuItem.Text = "PDF";
            this.pDFToolStripMenuItem.Click += new System.EventHandler(this.pDFToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 119);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(902, 488);
            this.dgv1.TabIndex = 1;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 630);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(926, 11);
            this.progressBar1.TabIndex = 2;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.rdbFas);
            this.panel1.Controls.Add(this.rdbRep);
            this.panel1.Location = new System.Drawing.Point(12, 37);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(323, 33);
            this.panel1.TabIndex = 3;
            // 
            // rdbFas
            // 
            this.rdbFas.AutoSize = true;
            this.rdbFas.Location = new System.Drawing.Point(154, 9);
            this.rdbFas.Name = "rdbFas";
            this.rdbFas.Size = new System.Drawing.Size(101, 17);
            this.rdbFas.TabIndex = 1;
            this.rdbFas.Text = "Per fascia oraria";
            this.rdbFas.UseVisualStyleBackColor = true;
            // 
            // rdbRep
            // 
            this.rdbRep.AutoSize = true;
            this.rdbRep.Checked = true;
            this.rdbRep.Location = new System.Drawing.Point(6, 9);
            this.rdbRep.Name = "rdbRep";
            this.rdbRep.Size = new System.Drawing.Size(77, 17);
            this.rdbRep.TabIndex = 0;
            this.rdbRep.TabStop = true;
            this.rdbRep.Text = "Per reparto";
            this.rdbRep.UseVisualStyleBackColor = true;
            // 
            // btnEstrai
            // 
            this.btnEstrai.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEstrai.Location = new System.Drawing.Point(776, 27);
            this.btnEstrai.Name = "btnEstrai";
            this.btnEstrai.Size = new System.Drawing.Size(138, 47);
            this.btnEstrai.TabIndex = 4;
            this.btnEstrai.Text = "Estrazione";
            this.btnEstrai.UseVisualStyleBackColor = true;
            this.btnEstrai.Click += new System.EventHandler(this.btnEstrai_Click);
            // 
            // dtpDay
            // 
            this.dtpDay.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDay.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDay.Location = new System.Drawing.Point(455, 32);
            this.dtpDay.Name = "dtpDay";
            this.dtpDay.Size = new System.Drawing.Size(261, 22);
            this.dtpDay.TabIndex = 5;
            this.dtpDay.ValueChanged += new System.EventHandler(this.dtpDay_ValueChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(354, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(97, 18);
            this.label1.TabIndex = 6;
            this.label1.Text = "Settimana del";
            // 
            // cmbNeg
            // 
            this.cmbNeg.FormattingEnabled = true;
            this.cmbNeg.Location = new System.Drawing.Point(168, 76);
            this.cmbNeg.Name = "cmbNeg";
            this.cmbNeg.Size = new System.Drawing.Size(164, 21);
            this.cmbNeg.TabIndex = 23;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(42, 83);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 13);
            this.label2.TabIndex = 24;
            this.label2.Text = "Negozio";
            // 
            // lblDti
            // 
            this.lblDti.AutoSize = true;
            this.lblDti.Location = new System.Drawing.Point(365, 68);
            this.lblDti.Name = "lblDti";
            this.lblDti.Size = new System.Drawing.Size(56, 13);
            this.lblDti.TabIndex = 25;
            this.lblDti.Text = "Data inizio";
            // 
            // lblDtf
            // 
            this.lblDtf.AutoSize = true;
            this.lblDtf.Location = new System.Drawing.Point(365, 93);
            this.lblDtf.Name = "lblDtf";
            this.lblDtf.Size = new System.Drawing.Size(50, 13);
            this.lblDtf.TabIndex = 26;
            this.lblDtf.Text = "Data fine";
            // 
            // dtpDti
            // 
            this.dtpDti.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDti.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDti.Location = new System.Drawing.Point(455, 60);
            this.dtpDti.Name = "dtpDti";
            this.dtpDti.Size = new System.Drawing.Size(261, 22);
            this.dtpDti.TabIndex = 27;
            // 
            // dtpDtf
            // 
            this.dtpDtf.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDtf.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDtf.Location = new System.Drawing.Point(455, 89);
            this.dtpDtf.Name = "dtpDtf";
            this.dtpDtf.Size = new System.Drawing.Size(261, 22);
            this.dtpDtf.TabIndex = 28;
            // 
            // frmGesStatSettimanali
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(926, 641);
            this.ControlBox = false;
            this.Controls.Add(this.dtpDtf);
            this.Controls.Add(this.dtpDti);
            this.Controls.Add(this.lblDtf);
            this.Controls.Add(this.lblDti);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cmbNeg);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dtpDay);
            this.Controls.Add(this.btnEstrai);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesStatSettimanali";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Estrazione statistiche settimanali";
            this.Load += new System.EventHandler(this.frmGesStatSettimanali_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesStatSettimanali_KeyDown);
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
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RadioButton rdbFas;
        private System.Windows.Forms.RadioButton rdbRep;
        private System.Windows.Forms.Button btnEstrai;
        private System.Windows.Forms.DateTimePicker dtpDay;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbNeg;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblDti;
        private System.Windows.Forms.Label lblDtf;
        private System.Windows.Forms.ToolStripMenuItem pDFToolStripMenuItem;
        private System.Windows.Forms.DateTimePicker dtpDti;
        private System.Windows.Forms.DateTimePicker dtpDtf;
    }
}
