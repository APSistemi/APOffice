namespace APOffice
{
    partial class frmUtyCtrlArtStat
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
            this.utilitàToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.recuperoStatoDaApShopToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.btnEstrai = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnAgg = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblCnt = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.utilitàToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(831, 24);
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
            // utilitàToolStripMenuItem
            // 
            this.utilitàToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.recuperoStatoDaApShopToolStripMenuItem});
            this.utilitàToolStripMenuItem.Name = "utilitàToolStripMenuItem";
            this.utilitàToolStripMenuItem.Size = new System.Drawing.Size(50, 20);
            this.utilitàToolStripMenuItem.Text = "Utilità";
            // 
            // recuperoStatoDaApShopToolStripMenuItem
            // 
            this.recuperoStatoDaApShopToolStripMenuItem.Name = "recuperoStatoDaApShopToolStripMenuItem";
            this.recuperoStatoDaApShopToolStripMenuItem.Size = new System.Drawing.Size(214, 22);
            this.recuperoStatoDaApShopToolStripMenuItem.Text = "Recupero stato da ApShop";
            this.recuperoStatoDaApShopToolStripMenuItem.Click += new System.EventHandler(this.recuperoStatoDaApShopToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 103);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(807, 478);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellClick);
            // 
            // btnEstrai
            // 
            this.btnEstrai.Location = new System.Drawing.Point(164, 43);
            this.btnEstrai.Name = "btnEstrai";
            this.btnEstrai.Size = new System.Drawing.Size(154, 23);
            this.btnEstrai.TabIndex = 2;
            this.btnEstrai.Text = "Estrai";
            this.btnEstrai.UseVisualStyleBackColor = true;
            this.btnEstrai.Click += new System.EventHandler(this.btnEstrai_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 53);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(141, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Estrazione articoli NON attivi";
            // 
            // btnAgg
            // 
            this.btnAgg.Location = new System.Drawing.Point(665, 43);
            this.btnAgg.Name = "btnAgg";
            this.btnAgg.Size = new System.Drawing.Size(154, 23);
            this.btnAgg.TabIndex = 4;
            this.btnAgg.Text = "Aggiorna STATO";
            this.btnAgg.UseVisualStyleBackColor = true;
            this.btnAgg.Click += new System.EventHandler(this.btnAgg_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 610);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(831, 10);
            this.progressBar1.TabIndex = 5;
            // 
            // lblCnt
            // 
            this.lblCnt.AutoSize = true;
            this.lblCnt.Location = new System.Drawing.Point(760, 587);
            this.lblCnt.Name = "lblCnt";
            this.lblCnt.Size = new System.Drawing.Size(13, 13);
            this.lblCnt.TabIndex = 6;
            this.lblCnt.Text = "0";
            // 
            // frmUtyCtrlArtStat
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(831, 619);
            this.ControlBox = false;
            this.Controls.Add(this.lblCnt);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnAgg);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnEstrai);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyCtrlArtStat";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Estrazione articoli da statistiche con stato NO ATTIVO";
            this.Load += new System.EventHandler(this.frmUtyCtrlArts_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyCtrlArts_KeyDown);
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
        private System.Windows.Forms.Button btnAgg;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblCnt;
        private System.Windows.Forms.ToolStripMenuItem utilitàToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem recuperoStatoDaApShopToolStripMenuItem;
    }
}
