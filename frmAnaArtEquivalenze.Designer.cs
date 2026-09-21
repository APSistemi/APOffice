namespace APOffice
{
    partial class frmAnaArtEquivalenze
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.lblDes = new System.Windows.Forms.Label();
            this.txtPrv = new System.Windows.Forms.TextBox();
            this.btnAgg = new System.Windows.Forms.Button();
            this.chkAll = new System.Windows.Forms.CheckBox();
            this.lblAll = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(675, 24);
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
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv1.DefaultCellStyle = dataGridViewCellStyle4;
            this.dgv1.Location = new System.Drawing.Point(3, 80);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(669, 346);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            // 
            // lblDes
            // 
            this.lblDes.AutoSize = true;
            this.lblDes.Location = new System.Drawing.Point(108, 55);
            this.lblDes.Name = "lblDes";
            this.lblDes.Size = new System.Drawing.Size(88, 13);
            this.lblDes.TabIndex = 2;
            this.lblDes.Text = "Prezzo di vendita";
            // 
            // txtPrv
            // 
            this.txtPrv.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPrv.Location = new System.Drawing.Point(202, 42);
            this.txtPrv.Name = "txtPrv";
            this.txtPrv.Size = new System.Drawing.Size(100, 26);
            this.txtPrv.TabIndex = 3;
            this.txtPrv.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // btnAgg
            // 
            this.btnAgg.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAgg.Location = new System.Drawing.Point(542, 40);
            this.btnAgg.Name = "btnAgg";
            this.btnAgg.Size = new System.Drawing.Size(130, 28);
            this.btnAgg.TabIndex = 4;
            this.btnAgg.Text = "Aggiorna";
            this.btnAgg.UseVisualStyleBackColor = true;
            this.btnAgg.Click += new System.EventHandler(this.btnAgg_Click);
            // 
            // chkAll
            // 
            this.chkAll.AutoSize = true;
            this.chkAll.Checked = true;
            this.chkAll.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkAll.Location = new System.Drawing.Point(34, 51);
            this.chkAll.Name = "chkAll";
            this.chkAll.Size = new System.Drawing.Size(15, 14);
            this.chkAll.TabIndex = 61;
            this.chkAll.UseVisualStyleBackColor = true;
            this.chkAll.Click += new System.EventHandler(this.chkAll_Click);
            // 
            // lblAll
            // 
            this.lblAll.AutoSize = true;
            this.lblAll.Location = new System.Drawing.Point(3, 54);
            this.lblAll.Name = "lblAll";
            this.lblAll.Size = new System.Drawing.Size(28, 13);
            this.lblAll.TabIndex = 62;
            this.lblAll.Text = "Tutti";
            // 
            // frmAnaArtEquivalenze
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(675, 453);
            this.ControlBox = false;
            this.Controls.Add(this.lblAll);
            this.Controls.Add(this.chkAll);
            this.Controls.Add(this.btnAgg);
            this.Controls.Add(this.txtPrv);
            this.Controls.Add(this.lblDes);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmAnaArtEquivalenze";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Equivalenze articoli";
            this.Load += new System.EventHandler(this.frmAnaArtEquivalenze_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAnaArtEquivalenze_KeyDown);
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
        private System.Windows.Forms.Label lblDes;
        private System.Windows.Forms.TextBox txtPrv;
        private System.Windows.Forms.Button btnAgg;
        private System.Windows.Forms.CheckBox chkAll;
        private System.Windows.Forms.Label lblAll;
    }
}
