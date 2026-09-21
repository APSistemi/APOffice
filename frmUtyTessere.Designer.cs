namespace APOffice
{
    partial class frmUtyTessere
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
            this.txtIni = new System.Windows.Forms.TextBox();
            this.txtQta = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnFill = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.cmbTesGru = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
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
            this.menuStrip1.Size = new System.Drawing.Size(464, 24);
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
            this.dgv1.Location = new System.Drawing.Point(13, 90);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(437, 303);
            this.dgv1.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 43);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Codice inizio";
            // 
            // txtIni
            // 
            this.txtIni.Location = new System.Drawing.Point(75, 36);
            this.txtIni.MaxLength = 13;
            this.txtIni.Name = "txtIni";
            this.txtIni.Size = new System.Drawing.Size(100, 20);
            this.txtIni.TabIndex = 3;
            // 
            // txtQta
            // 
            this.txtQta.Location = new System.Drawing.Point(213, 36);
            this.txtQta.Name = "txtQta";
            this.txtQta.Size = new System.Drawing.Size(52, 20);
            this.txtQta.TabIndex = 5;
            this.txtQta.Text = "0";
            this.txtQta.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(182, 43);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(27, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Q.tà";
            // 
            // btnFill
            // 
            this.btnFill.Location = new System.Drawing.Point(274, 36);
            this.btnFill.Name = "btnFill";
            this.btnFill.Size = new System.Drawing.Size(83, 48);
            this.btnFill.TabIndex = 6;
            this.btnFill.Text = "Carica";
            this.btnFill.UseVisualStyleBackColor = true;
            this.btnFill.Click += new System.EventHandler(this.btnFill_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(363, 38);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(87, 46);
            this.btnSave.TabIndex = 7;
            this.btnSave.Text = "Salva";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // cmbTesGru
            // 
            this.cmbTesGru.FormattingEnabled = true;
            this.cmbTesGru.Location = new System.Drawing.Point(75, 63);
            this.cmbTesGru.Name = "cmbTesGru";
            this.cmbTesGru.Size = new System.Drawing.Size(190, 21);
            this.cmbTesGru.TabIndex = 137;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 66);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(42, 13);
            this.label3.TabIndex = 138;
            this.label3.Text = "Gruppo";
            // 
            // frmUtyTessere
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(464, 405);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cmbTesGru);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnFill);
            this.Controls.Add(this.txtQta);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtIni);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyTessere";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Generazione tessere";
            this.Load += new System.EventHandler(this.frmUtyTessere_Load);
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
        private System.Windows.Forms.TextBox txtIni;
        private System.Windows.Forms.TextBox txtQta;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnFill;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.ComboBox cmbTesGru;
        private System.Windows.Forms.Label label3;
    }
}
