namespace APOffice
{
    partial class frmGesDivCtrlTab
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.articoliToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.abbinamentoAutomaticoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.dgv2 = new APOffice.APDataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.txtSeek = new System.Windows.Forms.TextBox();
            this.btnEcr = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.articoliToolStripMenuItem,
            this.abbinamentoAutomaticoToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(884, 24);
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
            // articoliToolStripMenuItem
            // 
            this.articoliToolStripMenuItem.Name = "articoliToolStripMenuItem";
            this.articoliToolStripMenuItem.Size = new System.Drawing.Size(57, 20);
            this.articoliToolStripMenuItem.Text = "Articoli";
            this.articoliToolStripMenuItem.Click += new System.EventHandler(this.articoliToolStripMenuItem_Click);
            // 
            // abbinamentoAutomaticoToolStripMenuItem
            // 
            this.abbinamentoAutomaticoToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.abbinamentoAutomaticoToolStripMenuItem.Name = "abbinamentoAutomaticoToolStripMenuItem";
            this.abbinamentoAutomaticoToolStripMenuItem.Size = new System.Drawing.Size(156, 20);
            this.abbinamentoAutomaticoToolStripMenuItem.Text = "Abbinamento automatico";
            this.abbinamentoAutomaticoToolStripMenuItem.Click += new System.EventHandler(this.abbinamentoAutomaticoToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Lavender;
            this.dgv1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 39);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(535, 425);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            this.dgv1.CurrentCellChanged += new System.EventHandler(this.dgv1_CurrentCellChanged);
            // 
            // dgv2
            // 
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.Lavender;
            this.dgv2.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv2.Location = new System.Drawing.Point(553, 89);
            this.dgv2.Name = "dgv2";
            this.dgv2.RowHeadersWidth = 20;
            this.dgv2.Size = new System.Drawing.Size(324, 374);
            this.dgv2.TabIndex = 2;
            this.dgv2.DoubleClick += new System.EventHandler(this.dgv2_DoubleClick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(554, 59);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(44, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Ricerca";
            // 
            // txtSeek
            // 
            this.txtSeek.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSeek.Location = new System.Drawing.Point(621, 56);
            this.txtSeek.Name = "txtSeek";
            this.txtSeek.Size = new System.Drawing.Size(255, 24);
            this.txtSeek.TabIndex = 4;
            this.txtSeek.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSeek_KeyDown);
            // 
            // btnEcr
            // 
            this.btnEcr.Location = new System.Drawing.Point(802, 28);
            this.btnEcr.Name = "btnEcr";
            this.btnEcr.Size = new System.Drawing.Size(75, 23);
            this.btnEcr.TabIndex = 5;
            this.btnEcr.Text = "ECR";
            this.btnEcr.UseVisualStyleBackColor = true;
            this.btnEcr.Click += new System.EventHandler(this.btnEcr_Click);
            // 
            // frmGesDivCtrlTab
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 480);
            this.ControlBox = false;
            this.Controls.Add(this.btnEcr);
            this.Controls.Add(this.txtSeek);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgv2);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesDivCtrlTab";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Abbinamento tabelle anagrafiche del fornitore";
            this.Load += new System.EventHandler(this.frmGesDivCtrlTab_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private APOffice.APDataGridView dgv2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtSeek;
        private System.Windows.Forms.ToolStripMenuItem articoliToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem abbinamentoAutomaticoToolStripMenuItem;
        private System.Windows.Forms.Button btnEcr;
    }
}
