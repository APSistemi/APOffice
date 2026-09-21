namespace APOffice
{
    partial class frmUtyArtEstrazioni
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
            this.btnEstrai = new System.Windows.Forms.Button();
            this.btnAgg = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblCnt = new System.Windows.Forms.Label();
            this.pnlRdb1 = new System.Windows.Forms.Panel();
            this.rdbArtIng = new System.Windows.Forms.RadioButton();
            this.rdbArtWeb = new System.Windows.Forms.RadioButton();
            this.chkSelAll = new System.Windows.Forms.CheckBox();
            this.chkPrnPrv = new System.Windows.Forms.CheckBox();
            this.rdbArtCel = new System.Windows.Forms.RadioButton();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.pnlRdb1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
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
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(8, 98);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(816, 479);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            // 
            // btnEstrai
            // 
            this.btnEstrai.Location = new System.Drawing.Point(414, 35);
            this.btnEstrai.Name = "btnEstrai";
            this.btnEstrai.Size = new System.Drawing.Size(194, 38);
            this.btnEstrai.TabIndex = 2;
            this.btnEstrai.Text = "Estrai";
            this.btnEstrai.UseVisualStyleBackColor = true;
            this.btnEstrai.Click += new System.EventHandler(this.btnEstrai_Click);
            // 
            // btnAgg
            // 
            this.btnAgg.Location = new System.Drawing.Point(615, 35);
            this.btnAgg.Name = "btnAgg";
            this.btnAgg.Size = new System.Drawing.Size(209, 38);
            this.btnAgg.TabIndex = 4;
            this.btnAgg.Text = "Generazione ";
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
            this.lblCnt.Location = new System.Drawing.Point(760, 589);
            this.lblCnt.Name = "lblCnt";
            this.lblCnt.Size = new System.Drawing.Size(13, 13);
            this.lblCnt.TabIndex = 6;
            this.lblCnt.Text = "0";
            // 
            // pnlRdb1
            // 
            this.pnlRdb1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnlRdb1.Controls.Add(this.rdbArtCel);
            this.pnlRdb1.Controls.Add(this.rdbArtIng);
            this.pnlRdb1.Controls.Add(this.rdbArtWeb);
            this.pnlRdb1.Location = new System.Drawing.Point(8, 35);
            this.pnlRdb1.Name = "pnlRdb1";
            this.pnlRdb1.Size = new System.Drawing.Size(386, 38);
            this.pnlRdb1.TabIndex = 7;
            // 
            // rdbArtIng
            // 
            this.rdbArtIng.AutoSize = true;
            this.rdbArtIng.Location = new System.Drawing.Point(104, 8);
            this.rdbArtIng.Name = "rdbArtIng";
            this.rdbArtIng.Size = new System.Drawing.Size(99, 17);
            this.rdbArtIng.TabIndex = 2;
            this.rdbArtIng.TabStop = true;
            this.rdbArtIng.Text = "Libro ingredienti";
            this.rdbArtIng.UseVisualStyleBackColor = true;
            // 
            // rdbArtWeb
            // 
            this.rdbArtWeb.AutoSize = true;
            this.rdbArtWeb.Location = new System.Drawing.Point(8, 7);
            this.rdbArtWeb.Name = "rdbArtWeb";
            this.rdbArtWeb.Size = new System.Drawing.Size(48, 17);
            this.rdbArtWeb.TabIndex = 1;
            this.rdbArtWeb.TabStop = true;
            this.rdbArtWeb.Text = "Web";
            this.rdbArtWeb.UseVisualStyleBackColor = true;
            // 
            // chkSelAll
            // 
            this.chkSelAll.AutoSize = true;
            this.chkSelAll.Checked = true;
            this.chkSelAll.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSelAll.Location = new System.Drawing.Point(8, 79);
            this.chkSelAll.Name = "chkSelAll";
            this.chkSelAll.Size = new System.Drawing.Size(92, 17);
            this.chkSelAll.TabIndex = 8;
            this.chkSelAll.Text = "Seleziona tutti";
            this.chkSelAll.UseVisualStyleBackColor = true;
            this.chkSelAll.CheckedChanged += new System.EventHandler(this.chkAll_Click);
            // 
            // chkPrnPrv
            // 
            this.chkPrnPrv.AutoSize = true;
            this.chkPrnPrv.Location = new System.Drawing.Point(114, 79);
            this.chkPrnPrv.Name = "chkPrnPrv";
            this.chkPrnPrv.Size = new System.Drawing.Size(79, 17);
            this.chkPrnPrv.TabIndex = 9;
            this.chkPrnPrv.Text = "Con prezzo";
            this.chkPrnPrv.UseVisualStyleBackColor = true;
            // 
            // rdbArtCel
            // 
            this.rdbArtCel.AutoSize = true;
            this.rdbArtCel.Location = new System.Drawing.Point(224, 9);
            this.rdbArtCel.Name = "rdbArtCel";
            this.rdbArtCel.Size = new System.Drawing.Size(68, 17);
            this.rdbArtCel.TabIndex = 3;
            this.rdbArtCel.TabStop = true;
            this.rdbArtCel.Text = "Celiachia";
            this.rdbArtCel.UseVisualStyleBackColor = true;
            // 
            // frmUtyArtEstrazioni
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(831, 619);
            this.ControlBox = false;
            this.Controls.Add(this.chkPrnPrv);
            this.Controls.Add(this.chkSelAll);
            this.Controls.Add(this.pnlRdb1);
            this.Controls.Add(this.lblCnt);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnAgg);
            this.Controls.Add(this.btnEstrai);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyArtEstrazioni";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Estrazioni articoli";
            this.Load += new System.EventHandler(this.frmUtyCtrlArts_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyCtrlArts_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.pnlRdb1.ResumeLayout(false);
            this.pnlRdb1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Button btnEstrai;
        private System.Windows.Forms.Button btnAgg;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblCnt;
        private System.Windows.Forms.Panel pnlRdb1;
        private System.Windows.Forms.RadioButton rdbArtWeb;
        private System.Windows.Forms.RadioButton rdbArtIng;
        private System.Windows.Forms.CheckBox chkSelAll;
        private System.Windows.Forms.CheckBox chkPrnPrv;
        private System.Windows.Forms.RadioButton rdbArtCel;
    }
}
