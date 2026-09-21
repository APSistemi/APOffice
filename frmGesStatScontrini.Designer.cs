namespace APOffice
{
    partial class frmGesStatScontrini
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
            this.label1 = new System.Windows.Forms.Label();
            this.lblTotVen = new System.Windows.Forms.Label();
            this.lblTotPro = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.btnTip = new System.Windows.Forms.Button();
            this.pDFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
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
            this.menuStrip1.Size = new System.Drawing.Size(859, 24);
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
            this.dgv1.Location = new System.Drawing.Point(12, 37);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(751, 478);
            this.dgv1.TabIndex = 1;
            this.dgv1.DoubleClick += new System.EventHandler(this.dgv1_DoubleClick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(568, 540);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(37, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Totale";
            // 
            // lblTotVen
            // 
            this.lblTotVen.BackColor = System.Drawing.Color.White;
            this.lblTotVen.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotVen.Location = new System.Drawing.Point(625, 528);
            this.lblTotVen.Name = "lblTotVen";
            this.lblTotVen.Size = new System.Drawing.Size(100, 23);
            this.lblTotVen.TabIndex = 19;
            this.lblTotVen.Text = "0";
            this.lblTotVen.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblTotPro
            // 
            this.lblTotPro.BackColor = System.Drawing.Color.White;
            this.lblTotPro.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotPro.Location = new System.Drawing.Point(462, 530);
            this.lblTotPro.Name = "lblTotPro";
            this.lblTotPro.Size = new System.Drawing.Size(100, 23);
            this.lblTotPro.TabIndex = 23;
            this.lblTotPro.Text = "0";
            this.lblTotPro.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(385, 539);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(61, 13);
            this.label5.TabIndex = 22;
            this.label5.Text = "Totale prof.";
            // 
            // btnTip
            // 
            this.btnTip.Location = new System.Drawing.Point(772, 37);
            this.btnTip.Name = "btnTip";
            this.btnTip.Size = new System.Drawing.Size(75, 75);
            this.btnTip.TabIndex = 24;
            this.btnTip.Text = "Tipi pagamento";
            this.btnTip.UseVisualStyleBackColor = true;
            this.btnTip.Click += new System.EventHandler(this.btnTip_Click);
            // 
            // pDFToolStripMenuItem
            // 
            this.pDFToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.pDFToolStripMenuItem.Name = "pDFToolStripMenuItem";
            this.pDFToolStripMenuItem.Size = new System.Drawing.Size(40, 20);
            this.pDFToolStripMenuItem.Text = "PDF";
            this.pDFToolStripMenuItem.Click += new System.EventHandler(this.pDFToolStripMenuItem_Click);
            // 
            // frmGesStatScontrini
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(859, 562);
            this.ControlBox = false;
            this.Controls.Add(this.btnTip);
            this.Controls.Add(this.lblTotPro);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.lblTotVen);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesStatScontrini";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Statistiche per scontrino";
            this.Load += new System.EventHandler(this.frmStatScontrini_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmStatScontrini_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblTotVen;
        private System.Windows.Forms.Label lblTotPro;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnTip;
        private System.Windows.Forms.ToolStripMenuItem pDFToolStripMenuItem;
    }
}
