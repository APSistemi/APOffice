namespace APOffice
{
    partial class frmGesStatCtrlTotali
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
            this.procediToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dtpDay = new System.Windows.Forms.DateTimePicker();
            this.dgv1 = new APOffice.APDataGridView();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblTot = new System.Windows.Forms.Label();
            this.lblPag = new System.Windows.Forms.Label();
            this.lblVen = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.procediToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(911, 24);
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
            // procediToolStripMenuItem
            // 
            this.procediToolStripMenuItem.Name = "procediToolStripMenuItem";
            this.procediToolStripMenuItem.Size = new System.Drawing.Size(59, 20);
            this.procediToolStripMenuItem.Text = "Procedi";
            this.procediToolStripMenuItem.Click += new System.EventHandler(this.procediToolStripMenuItem_Click);
            // 
            // dtpDay
            // 
            this.dtpDay.Location = new System.Drawing.Point(12, 36);
            this.dtpDay.Name = "dtpDay";
            this.dtpDay.Size = new System.Drawing.Size(200, 20);
            this.dtpDay.TabIndex = 3;
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 81);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(888, 422);
            this.dgv1.TabIndex = 4;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(7, 554);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(898, 10);
            this.progressBar1.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(293, 530);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(33, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Totali";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(702, 530);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Vendite";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(486, 530);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(57, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Pagamenti";
            // 
            // lblTot
            // 
            this.lblTot.BackColor = System.Drawing.Color.White;
            this.lblTot.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTot.Location = new System.Drawing.Point(365, 520);
            this.lblTot.Name = "lblTot";
            this.lblTot.Size = new System.Drawing.Size(100, 23);
            this.lblTot.TabIndex = 9;
            this.lblTot.Text = "0";
            this.lblTot.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblPag
            // 
            this.lblPag.BackColor = System.Drawing.Color.White;
            this.lblPag.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblPag.Location = new System.Drawing.Point(562, 520);
            this.lblPag.Name = "lblPag";
            this.lblPag.Size = new System.Drawing.Size(100, 23);
            this.lblPag.TabIndex = 10;
            this.lblPag.Text = "0";
            this.lblPag.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblVen
            // 
            this.lblVen.BackColor = System.Drawing.Color.White;
            this.lblVen.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblVen.Location = new System.Drawing.Point(783, 520);
            this.lblVen.Name = "lblVen";
            this.lblVen.Size = new System.Drawing.Size(100, 23);
            this.lblVen.TabIndex = 11;
            this.lblVen.Text = "0";
            this.lblVen.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // frmGesStatCtrlTotali
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(911, 569);
            this.Controls.Add(this.lblVen);
            this.Controls.Add(this.lblPag);
            this.Controls.Add(this.lblTot);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.dtpDay);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesStatCtrlTotali";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmGesStatCtrlTotali";
            this.Load += new System.EventHandler(this.frmGesStatCtrlTotali_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.DateTimePicker dtpDay;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.ToolStripMenuItem procediToolStripMenuItem;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblTot;
        private System.Windows.Forms.Label lblPag;
        private System.Windows.Forms.Label lblVen;
    }
}
