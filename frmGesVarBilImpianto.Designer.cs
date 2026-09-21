namespace APOffice
{
    partial class frmGesVarBilImpianto
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.rdbAll = new System.Windows.Forms.RadioButton();
            this.rdbAtt = new System.Windows.Forms.RadioButton();
            this.btnFill = new System.Windows.Forms.Button();
            this.cmbReb = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(407, 24);
            this.menuStrip1.TabIndex = 13;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // esciToolStripMenuItem
            // 
            this.esciToolStripMenuItem.Name = "esciToolStripMenuItem";
            this.esciToolStripMenuItem.Size = new System.Drawing.Size(39, 20);
            this.esciToolStripMenuItem.Text = "Esci";
            this.esciToolStripMenuItem.Click += new System.EventHandler(this.esciToolStripMenuItem_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.rdbAll);
            this.panel1.Controls.Add(this.rdbAtt);
            this.panel1.Location = new System.Drawing.Point(2, 39);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(402, 31);
            this.panel1.TabIndex = 15;
            // 
            // rdbAll
            // 
            this.rdbAll.AutoSize = true;
            this.rdbAll.Location = new System.Drawing.Point(285, 4);
            this.rdbAll.Name = "rdbAll";
            this.rdbAll.Size = new System.Drawing.Size(46, 17);
            this.rdbAll.TabIndex = 1;
            this.rdbAll.Text = "Tutti";
            this.rdbAll.UseVisualStyleBackColor = true;
            // 
            // rdbAtt
            // 
            this.rdbAtt.AutoSize = true;
            this.rdbAtt.Checked = true;
            this.rdbAtt.Location = new System.Drawing.Point(12, 4);
            this.rdbAtt.Name = "rdbAtt";
            this.rdbAtt.Size = new System.Drawing.Size(104, 17);
            this.rdbAtt.TabIndex = 0;
            this.rdbAtt.TabStop = true;
            this.rdbAtt.Text = "Solo articoli attivi";
            this.rdbAtt.UseVisualStyleBackColor = true;
            // 
            // btnFill
            // 
            this.btnFill.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFill.Location = new System.Drawing.Point(96, 149);
            this.btnFill.Name = "btnFill";
            this.btnFill.Size = new System.Drawing.Size(210, 76);
            this.btnFill.TabIndex = 14;
            this.btnFill.Text = "Genera impianto per le bilance";
            this.btnFill.UseVisualStyleBackColor = true;
            this.btnFill.Click += new System.EventHandler(this.btnFill_Click);
            // 
            // cmbReb
            // 
            this.cmbReb.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbReb.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbReb.FormattingEnabled = true;
            this.cmbReb.Location = new System.Drawing.Point(51, 99);
            this.cmbReb.Name = "cmbReb";
            this.cmbReb.Size = new System.Drawing.Size(304, 32);
            this.cmbReb.TabIndex = 17;
            this.cmbReb.SelectionChangeCommitted += new System.EventHandler(this.cmbReb_SelectionChangeCommitted);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 117);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 13);
            this.label1.TabIndex = 18;
            this.label1.Text = "Banco";
            // 
            // frmGesVarBilImpianto
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(407, 236);
            this.ControlBox = false;
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbReb);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnFill);
            this.KeyPreview = true;
            this.Name = "frmGesVarBilImpianto";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Impianto a bilance";
            this.Load += new System.EventHandler(this.frmGesVarBilImpianto_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesVarBilImpianto_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RadioButton rdbAll;
        private System.Windows.Forms.RadioButton rdbAtt;
        private System.Windows.Forms.Button btnFill;
        private System.Windows.Forms.ComboBox cmbReb;
        private System.Windows.Forms.Label label1;
    }
}