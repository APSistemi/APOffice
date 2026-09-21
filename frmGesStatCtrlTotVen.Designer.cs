namespace APOffice
{
    partial class frmGesStatCtrlTotVen
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
            this.lblVen = new System.Windows.Forms.Label();
            this.lblPag = new System.Windows.Forms.Label();
            this.lblTot = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.dgv1 = new APOffice.APDataGridView();
            this.dtpDay = new System.Windows.Forms.DateTimePicker();
            this.lblRes = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
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
            this.menuStrip1.Size = new System.Drawing.Size(910, 24);
            this.menuStrip1.TabIndex = 12;
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
            // lblVen
            // 
            this.lblVen.BackColor = System.Drawing.Color.White;
            this.lblVen.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblVen.Location = new System.Drawing.Point(783, 521);
            this.lblVen.Name = "lblVen";
            this.lblVen.Size = new System.Drawing.Size(100, 23);
            this.lblVen.TabIndex = 21;
            this.lblVen.Text = "0";
            this.lblVen.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblPag
            // 
            this.lblPag.BackColor = System.Drawing.Color.White;
            this.lblPag.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblPag.Location = new System.Drawing.Point(562, 521);
            this.lblPag.Name = "lblPag";
            this.lblPag.Size = new System.Drawing.Size(100, 23);
            this.lblPag.TabIndex = 20;
            this.lblPag.Text = "0";
            this.lblPag.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblTot
            // 
            this.lblTot.BackColor = System.Drawing.Color.White;
            this.lblTot.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTot.Location = new System.Drawing.Point(365, 521);
            this.lblTot.Name = "lblTot";
            this.lblTot.Size = new System.Drawing.Size(100, 23);
            this.lblTot.TabIndex = 19;
            this.lblTot.Text = "0";
            this.lblTot.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(486, 531);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(57, 13);
            this.label3.TabIndex = 18;
            this.label3.Text = "Pagamenti";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(702, 531);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 17;
            this.label2.Text = "Vendite";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(293, 531);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(33, 13);
            this.label1.TabIndex = 16;
            this.label1.Text = "Totali";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(7, 555);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(898, 10);
            this.progressBar1.TabIndex = 15;
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 82);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(888, 422);
            this.dgv1.TabIndex = 14;
            // 
            // dtpDay
            // 
            this.dtpDay.Location = new System.Drawing.Point(12, 37);
            this.dtpDay.Name = "dtpDay";
            this.dtpDay.Size = new System.Drawing.Size(200, 20);
            this.dtpDay.TabIndex = 13;
            // 
            // lblRes
            // 
            this.lblRes.BackColor = System.Drawing.Color.White;
            this.lblRes.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblRes.Location = new System.Drawing.Point(165, 521);
            this.lblRes.Name = "lblRes";
            this.lblRes.Size = new System.Drawing.Size(100, 23);
            this.lblRes.TabIndex = 22;
            this.lblRes.Text = "0";
            this.lblRes.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(112, 533);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(28, 13);
            this.label4.TabIndex = 23;
            this.label4.Text = "Resi";
            // 
            // frmGesStaCtrlTotVen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(910, 567);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lblRes);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.lblVen);
            this.Controls.Add(this.lblPag);
            this.Controls.Add(this.lblTot);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.dtpDay);
            this.Name = "frmGesStaCtrlTotVen";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmGesStaCtrlTotVen";
            this.Load += new System.EventHandler(this.frmGesStaCtrlTotVen_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem procediToolStripMenuItem;
        private System.Windows.Forms.Label lblVen;
        private System.Windows.Forms.Label lblPag;
        private System.Windows.Forms.Label lblTot;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.DateTimePicker dtpDay;
        private System.Windows.Forms.Label lblRes;
        private System.Windows.Forms.Label label4;
    }
}
