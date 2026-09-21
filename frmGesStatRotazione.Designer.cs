namespace APOffice
{
    partial class frmGesStatRotazione
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdbArtMov = new System.Windows.Forms.RadioButton();
            this.rdbArtAll = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpIni = new System.Windows.Forms.DateTimePicker();
            this.dtpFin = new System.Windows.Forms.DateTimePicker();
            this.btnEstrai = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rdbSelMon = new System.Windows.Forms.RadioButton();
            this.rdbSelSet = new System.Windows.Forms.RadioButton();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.rdbTipVal = new System.Windows.Forms.RadioButton();
            this.rdbTipQta = new System.Windows.Forms.RadioButton();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.excelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.excelToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(863, 24);
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
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 140);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(839, 523);
            this.dgv1.TabIndex = 17;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdbArtMov);
            this.groupBox1.Controls.Add(this.rdbArtAll);
            this.groupBox1.Location = new System.Drawing.Point(575, 64);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(275, 70);
            this.groupBox1.TabIndex = 19;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Articoli";
            this.groupBox1.Visible = false;
            // 
            // rdbArtMov
            // 
            this.rdbArtMov.AutoSize = true;
            this.rdbArtMov.Location = new System.Drawing.Point(13, 42);
            this.rdbArtMov.Name = "rdbArtMov";
            this.rdbArtMov.Size = new System.Drawing.Size(105, 17);
            this.rdbArtMov.TabIndex = 23;
            this.rdbArtMov.Text = "Solo movimentati";
            this.rdbArtMov.UseVisualStyleBackColor = true;
            // 
            // rdbArtAll
            // 
            this.rdbArtAll.AutoSize = true;
            this.rdbArtAll.Checked = true;
            this.rdbArtAll.Location = new System.Drawing.Point(13, 19);
            this.rdbArtAll.Name = "rdbArtAll";
            this.rdbArtAll.Size = new System.Drawing.Size(46, 17);
            this.rdbArtAll.TabIndex = 22;
            this.rdbArtAll.TabStop = true;
            this.rdbArtAll.Text = "Tutti";
            this.rdbArtAll.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(14, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(23, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "Dal";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(256, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(15, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "al";
            // 
            // dtpIni
            // 
            this.dtpIni.Location = new System.Drawing.Point(43, 32);
            this.dtpIni.Name = "dtpIni";
            this.dtpIni.Size = new System.Drawing.Size(200, 20);
            this.dtpIni.TabIndex = 14;
            // 
            // dtpFin
            // 
            this.dtpFin.Location = new System.Drawing.Point(280, 33);
            this.dtpFin.Name = "dtpFin";
            this.dtpFin.Size = new System.Drawing.Size(200, 20);
            this.dtpFin.TabIndex = 15;
            // 
            // btnEstrai
            // 
            this.btnEstrai.BackColor = System.Drawing.SystemColors.Control;
            this.btnEstrai.Location = new System.Drawing.Point(486, 32);
            this.btnEstrai.Name = "btnEstrai";
            this.btnEstrai.Size = new System.Drawing.Size(365, 23);
            this.btnEstrai.TabIndex = 16;
            this.btnEstrai.Text = "Estrai";
            this.btnEstrai.UseVisualStyleBackColor = false;
            this.btnEstrai.Click += new System.EventHandler(this.btnEstrai_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rdbSelMon);
            this.groupBox2.Controls.Add(this.rdbSelSet);
            this.groupBox2.Location = new System.Drawing.Point(14, 65);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(275, 70);
            this.groupBox2.TabIndex = 24;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Estrazione";
            // 
            // rdbSelMon
            // 
            this.rdbSelMon.AutoSize = true;
            this.rdbSelMon.Checked = true;
            this.rdbSelMon.Location = new System.Drawing.Point(6, 19);
            this.rdbSelMon.Name = "rdbSelMon";
            this.rdbSelMon.Size = new System.Drawing.Size(55, 17);
            this.rdbSelMon.TabIndex = 23;
            this.rdbSelMon.TabStop = true;
            this.rdbSelMon.Text = "Mesile";
            this.rdbSelMon.UseVisualStyleBackColor = true;
            // 
            // rdbSelSet
            // 
            this.rdbSelSet.AutoSize = true;
            this.rdbSelSet.Location = new System.Drawing.Point(6, 40);
            this.rdbSelSet.Name = "rdbSelSet";
            this.rdbSelSet.Size = new System.Drawing.Size(80, 17);
            this.rdbSelSet.TabIndex = 22;
            this.rdbSelSet.Text = "Settimanale";
            this.rdbSelSet.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.rdbTipVal);
            this.groupBox3.Controls.Add(this.rdbTipQta);
            this.groupBox3.Location = new System.Drawing.Point(294, 65);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(275, 70);
            this.groupBox3.TabIndex = 25;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Tipo";
            // 
            // rdbTipVal
            // 
            this.rdbTipVal.AutoSize = true;
            this.rdbTipVal.Location = new System.Drawing.Point(13, 42);
            this.rdbTipVal.Name = "rdbTipVal";
            this.rdbTipVal.Size = new System.Drawing.Size(55, 17);
            this.rdbTipVal.TabIndex = 23;
            this.rdbTipVal.Text = "Valore";
            this.rdbTipVal.UseVisualStyleBackColor = true;
            // 
            // rdbTipQta
            // 
            this.rdbTipQta.AutoSize = true;
            this.rdbTipQta.Checked = true;
            this.rdbTipQta.Location = new System.Drawing.Point(13, 19);
            this.rdbTipQta.Name = "rdbTipQta";
            this.rdbTipQta.Size = new System.Drawing.Size(65, 17);
            this.rdbTipQta.TabIndex = 22;
            this.rdbTipQta.TabStop = true;
            this.rdbTipQta.Text = "Quantità";
            this.rdbTipQta.UseVisualStyleBackColor = true;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 675);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(863, 14);
            this.progressBar1.TabIndex = 26;
            // 
            // excelToolStripMenuItem
            // 
            this.excelToolStripMenuItem.Name = "excelToolStripMenuItem";
            this.excelToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.excelToolStripMenuItem.Text = "Excel";
            this.excelToolStripMenuItem.Click += new System.EventHandler(this.excelToolStripMenuItem_Click);
            // 
            // frmGesStatRotazione
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(863, 687);
            this.ControlBox = false;
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.btnEstrai);
            this.Controls.Add(this.dtpFin);
            this.Controls.Add(this.dtpIni);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesStatRotazione";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Rotazione articoli";
            this.Load += new System.EventHandler(this.frmGesStatRotazione_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesStatRotazione_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdbArtMov;
        private System.Windows.Forms.RadioButton rdbArtAll;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpIni;
        private System.Windows.Forms.DateTimePicker dtpFin;
        private System.Windows.Forms.Button btnEstrai;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton rdbSelMon;
        private System.Windows.Forms.RadioButton rdbSelSet;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.RadioButton rdbTipVal;
        private System.Windows.Forms.RadioButton rdbTipQta;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ToolStripMenuItem excelToolStripMenuItem;
    }
}
