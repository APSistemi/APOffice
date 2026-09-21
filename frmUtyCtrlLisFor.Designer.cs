namespace APOffice
{
    partial class frmUtyCtrlLisFor
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
            this.txtLis = new System.Windows.Forms.TextBox();
            this.btnFil = new System.Windows.Forms.Button();
            this.dgv2 = new APOffice.APDataGridView();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(873, 24);
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
            this.dgv1.Location = new System.Drawing.Point(28, 89);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(811, 200);
            this.dgv1.TabIndex = 1;
            // 
            // txtLis
            // 
            this.txtLis.Location = new System.Drawing.Point(28, 45);
            this.txtLis.Name = "txtLis";
            this.txtLis.Size = new System.Drawing.Size(455, 20);
            this.txtLis.TabIndex = 2;
            this.txtLis.Text = "C:\\ApProject\\Temp\\DivFornitori\\Venfri\\LISTINO2017M1.xls";
            // 
            // btnFil
            // 
            this.btnFil.Location = new System.Drawing.Point(764, 45);
            this.btnFil.Name = "btnFil";
            this.btnFil.Size = new System.Drawing.Size(75, 23);
            this.btnFil.TabIndex = 3;
            this.btnFil.Text = "Carica";
            this.btnFil.UseVisualStyleBackColor = true;
            this.btnFil.Click += new System.EventHandler(this.btnFil_Click);
            // 
            // dgv2
            // 
            this.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv2.Location = new System.Drawing.Point(28, 323);
            this.dgv2.Name = "dgv2";
            this.dgv2.Size = new System.Drawing.Size(811, 290);
            this.dgv2.TabIndex = 4;
            this.dgv2.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv2_CellDoubleClick);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 624);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(873, 10);
            this.progressBar1.TabIndex = 5;
            // 
            // frmUtyCtrlLisFor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(873, 633);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.dgv2);
            this.Controls.Add(this.btnFil);
            this.Controls.Add(this.txtLis);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyCtrlLisFor";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmUtyCtrlLisFor";
            this.Load += new System.EventHandler(this.frmUtyCtrlLisFor_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyCtrlLisFor_KeyDown);
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
        private System.Windows.Forms.TextBox txtLis;
        private System.Windows.Forms.Button btnFil;
        private APOffice.APDataGridView dgv2;
        private System.Windows.Forms.ProgressBar progressBar1;
    }
}
