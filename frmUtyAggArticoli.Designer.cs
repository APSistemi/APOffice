namespace APOffice
{
    partial class frmUtyAggArticoli
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
            this.btnSeek = new System.Windows.Forms.Button();
            this.txtSeek = new System.Windows.Forms.TextBox();
            this.label47 = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(935, 24);
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
            this.dgv1.Location = new System.Drawing.Point(12, 66);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(911, 513);
            this.dgv1.TabIndex = 1;
            // 
            // btnSeek
            // 
            this.btnSeek.Location = new System.Drawing.Point(814, 37);
            this.btnSeek.Name = "btnSeek";
            this.btnSeek.Size = new System.Drawing.Size(55, 23);
            this.btnSeek.TabIndex = 104;
            this.btnSeek.Text = "...";
            this.btnSeek.UseVisualStyleBackColor = true;
            this.btnSeek.Click += new System.EventHandler(this.btnSeek_Click);
            // 
            // txtSeek
            // 
            this.txtSeek.Location = new System.Drawing.Point(151, 39);
            this.txtSeek.Name = "txtSeek";
            this.txtSeek.Size = new System.Drawing.Size(657, 20);
            this.txtSeek.TabIndex = 103;
            this.txtSeek.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSeek_KeyDown);
            // 
            // label47
            // 
            this.label47.AutoSize = true;
            this.label47.ForeColor = System.Drawing.Color.Blue;
            this.label47.Location = new System.Drawing.Point(66, 44);
            this.label47.Name = "label47";
            this.label47.Size = new System.Drawing.Size(79, 13);
            this.label47.TabIndex = 102;
            this.label47.Text = "Ricerca veloce";
            // 
            // frmUtyAggArticoli
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(935, 605);
            this.ControlBox = false;
            this.Controls.Add(this.btnSeek);
            this.Controls.Add(this.txtSeek);
            this.Controls.Add(this.label47);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyAggArticoli";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Variazioni articoli";
            this.Load += new System.EventHandler(this.frmUtyAggArticoli_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyAggArticoli_KeyDown);
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
        private System.Windows.Forms.Button btnSeek;
        private System.Windows.Forms.TextBox txtSeek;
        private System.Windows.Forms.Label label47;
    }
}
