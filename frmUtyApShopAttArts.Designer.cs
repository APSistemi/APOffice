namespace APOffice
{
    partial class frmUtyApShopAttArts
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
            this.lblPath = new System.Windows.Forms.Label();
            this.btnMdb = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.estrazioneToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblCnt = new System.Windows.Forms.Label();
            this.btnOk = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.estrazioneToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(792, 24);
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
            this.dgv1.Location = new System.Drawing.Point(21, 76);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(746, 450);
            this.dgv1.TabIndex = 1;
            // 
            // lblPath
            // 
            this.lblPath.AutoSize = true;
            this.lblPath.Location = new System.Drawing.Point(128, 46);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(16, 13);
            this.lblPath.TabIndex = 8;
            this.lblPath.Text = "...";
            // 
            // btnMdb
            // 
            this.btnMdb.Location = new System.Drawing.Point(81, 39);
            this.btnMdb.Name = "btnMdb";
            this.btnMdb.Size = new System.Drawing.Size(42, 23);
            this.btnMdb.TabIndex = 7;
            this.btnMdb.Text = "...";
            this.btnMdb.UseVisualStyleBackColor = true;
            this.btnMdb.Click += new System.EventHandler(this.btnMdb_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 49);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "DB cassa";
            // 
            // estrazioneToolStripMenuItem
            // 
            this.estrazioneToolStripMenuItem.Name = "estrazioneToolStripMenuItem";
            this.estrazioneToolStripMenuItem.Size = new System.Drawing.Size(72, 20);
            this.estrazioneToolStripMenuItem.Text = "Estrazione";
            this.estrazioneToolStripMenuItem.Click += new System.EventHandler(this.estrazioneToolStripMenuItem_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(2, 547);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(788, 11);
            this.progressBar1.TabIndex = 9;
            // 
            // lblCnt
            // 
            this.lblCnt.AutoSize = true;
            this.lblCnt.Location = new System.Drawing.Point(723, 530);
            this.lblCnt.Name = "lblCnt";
            this.lblCnt.Size = new System.Drawing.Size(13, 13);
            this.lblCnt.TabIndex = 10;
            this.lblCnt.Text = "0";
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(576, 39);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(191, 23);
            this.btnOk.TabIndex = 11;
            this.btnOk.Text = "Attivazione";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // frmUtyApShopAttArts
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(792, 559);
            this.ControlBox = false;
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.lblCnt);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.lblPath);
            this.Controls.Add(this.btnMdb);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyApShopAttArts";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Attivazione articoli se presenti in ApShop mediante ricerca barcode";
            this.Load += new System.EventHandler(this.frmUtyApShopAttArts_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyApShopAttArts_KeyDown);
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
        private System.Windows.Forms.ToolStripMenuItem estrazioneToolStripMenuItem;
        private System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.Button btnMdb;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblCnt;
        private System.Windows.Forms.Button btnOk;
    }
}
