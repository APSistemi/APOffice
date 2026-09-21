namespace APOffice
{
    partial class frmUtyCtrlArtPos
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
            this.excelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.btnEstrai = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnMdb = new System.Windows.Forms.Button();
            this.lblPath = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btnAnn = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
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
            this.menuStrip1.Size = new System.Drawing.Size(650, 24);
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
            // excelToolStripMenuItem
            // 
            this.excelToolStripMenuItem.Name = "excelToolStripMenuItem";
            this.excelToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.excelToolStripMenuItem.Text = "Excel";
            this.excelToolStripMenuItem.Click += new System.EventHandler(this.excelToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(7, 78);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(638, 477);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellContentClick);
            // 
            // btnEstrai
            // 
            this.btnEstrai.Location = new System.Drawing.Point(297, 31);
            this.btnEstrai.Name = "btnEstrai";
            this.btnEstrai.Size = new System.Drawing.Size(75, 23);
            this.btnEstrai.TabIndex = 2;
            this.btnEstrai.Text = "Estrai";
            this.btnEstrai.UseVisualStyleBackColor = true;
            this.btnEstrai.Click += new System.EventHandler(this.btnEstrai_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "DB cassa";
            // 
            // btnMdb
            // 
            this.btnMdb.Location = new System.Drawing.Point(81, 31);
            this.btnMdb.Name = "btnMdb";
            this.btnMdb.Size = new System.Drawing.Size(42, 23);
            this.btnMdb.TabIndex = 4;
            this.btnMdb.Text = "...";
            this.btnMdb.UseVisualStyleBackColor = true;
            this.btnMdb.Click += new System.EventHandler(this.btnMdb_Click);
            // 
            // lblPath
            // 
            this.lblPath.AutoSize = true;
            this.lblPath.Location = new System.Drawing.Point(29, 60);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(16, 13);
            this.lblPath.TabIndex = 5;
            this.lblPath.Text = "...";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 568);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(650, 12);
            this.progressBar1.TabIndex = 6;
            // 
            // btnAnn
            // 
            this.btnAnn.Location = new System.Drawing.Point(464, 31);
            this.btnAnn.Name = "btnAnn";
            this.btnAnn.Size = new System.Drawing.Size(174, 41);
            this.btnAnn.TabIndex = 7;
            this.btnAnn.Text = "Annullamento barcode non corriscondenti";
            this.btnAnn.UseVisualStyleBackColor = true;
            this.btnAnn.Click += new System.EventHandler(this.btnAnn_Click);
            // 
            // frmUtyCtrlArtPos
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(650, 582);
            this.ControlBox = false;
            this.Controls.Add(this.btnAnn);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.lblPath);
            this.Controls.Add(this.btnMdb);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnEstrai);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyCtrlArtPos";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Controllo articoli cassa";
            this.Load += new System.EventHandler(this.frmUtyCtrlArtPos_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyCtrlArtPos_KeyDown);
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
        private System.Windows.Forms.Button btnEstrai;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnMdb;
        private System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ToolStripMenuItem excelToolStripMenuItem;
        private System.Windows.Forms.Button btnAnn;
    }
}
