namespace APOffice
{
    partial class frmPrnForArticoli
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
            this.cmbArtFor = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnAtt = new System.Windows.Forms.Button();
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
            this.excelToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(828, 24);
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
            this.excelToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.excelToolStripMenuItem.Name = "excelToolStripMenuItem";
            this.excelToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.excelToolStripMenuItem.Text = "Excel";
            this.excelToolStripMenuItem.Click += new System.EventHandler(this.excelToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(6, 63);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(817, 472);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            // 
            // cmbArtFor
            // 
            this.cmbArtFor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbArtFor.FormattingEnabled = true;
            this.cmbArtFor.Location = new System.Drawing.Point(76, 32);
            this.cmbArtFor.Name = "cmbArtFor";
            this.cmbArtFor.Size = new System.Drawing.Size(227, 21);
            this.cmbArtFor.TabIndex = 28;
            this.cmbArtFor.SelectionChangeCommitted += new System.EventHandler(this.cmbArtFor_SelectionChangeCommitted);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 13);
            this.label1.TabIndex = 29;
            this.label1.Text = "Fornitore";
            // 
            // btnAtt
            // 
            this.btnAtt.Location = new System.Drawing.Point(668, 31);
            this.btnAtt.Name = "btnAtt";
            this.btnAtt.Size = new System.Drawing.Size(155, 23);
            this.btnAtt.TabIndex = 30;
            this.btnAtt.Text = "Attiva alle casse";
            this.btnAtt.UseVisualStyleBackColor = true;
            this.btnAtt.Click += new System.EventHandler(this.btnAtt_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(1, 541);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(770, 10);
            this.progressBar1.TabIndex = 31;
            // 
            // lblCnt
            // 
            this.lblCnt.AutoSize = true;
            this.lblCnt.Location = new System.Drawing.Point(772, 538);
            this.lblCnt.Name = "lblCnt";
            this.lblCnt.Size = new System.Drawing.Size(13, 13);
            this.lblCnt.TabIndex = 32;
            this.lblCnt.Text = "0";
            // 
            // frmPrnForArticoli
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(828, 553);
            this.ControlBox = false;
            this.Controls.Add(this.lblCnt);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnAtt);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbArtFor);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmPrnForArticoli";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Articoli per fornitore";
            this.Load += new System.EventHandler(this.frmPrnForArticoli_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPrnForArticoli_KeyDown);
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
        private System.Windows.Forms.ComboBox cmbArtFor;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolStripMenuItem excelToolStripMenuItem;
        private System.Windows.Forms.Button btnAtt;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblCnt;
    }
}
