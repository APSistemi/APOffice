namespace APOffice
{
    partial class frmUtyPrnGiacenza
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
            this.csvToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excel2ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.btnEstrai = new System.Windows.Forms.Button();
            this.lblArts = new System.Windows.Forms.Label();
            this.lblCos = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.cmbNeg = new System.Windows.Forms.ComboBox();
            this.chkArtAtt = new System.Windows.Forms.CheckBox();
            this.lblMsg = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.excelToolStripMenuItem,
            this.csvToolStripMenuItem,
            this.excel2ToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(882, 24);
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
            this.excelToolStripMenuItem.Name = "excelToolStripMenuItem";
            this.excelToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.excelToolStripMenuItem.Text = "Excel";
            this.excelToolStripMenuItem.Click += new System.EventHandler(this.excelToolStripMenuItem_Click);
            // 
            // csvToolStripMenuItem
            // 
            this.csvToolStripMenuItem.Name = "csvToolStripMenuItem";
            this.csvToolStripMenuItem.Size = new System.Drawing.Size(38, 20);
            this.csvToolStripMenuItem.Text = "Csv";
            this.csvToolStripMenuItem.Click += new System.EventHandler(this.csvToolStripMenuItem_Click);
            // 
            // excel2ToolStripMenuItem
            // 
            this.excel2ToolStripMenuItem.Name = "excel2ToolStripMenuItem";
            this.excel2ToolStripMenuItem.Size = new System.Drawing.Size(54, 20);
            this.excel2ToolStripMenuItem.Text = "Excel 2";
            this.excel2ToolStripMenuItem.Click += new System.EventHandler(this.excel2ToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(4, 58);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(873, 462);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            // 
            // btnEstrai
            // 
            this.btnEstrai.Location = new System.Drawing.Point(763, 28);
            this.btnEstrai.Name = "btnEstrai";
            this.btnEstrai.Size = new System.Drawing.Size(114, 23);
            this.btnEstrai.TabIndex = 10;
            this.btnEstrai.Text = "Estrazione";
            this.btnEstrai.UseVisualStyleBackColor = true;
            this.btnEstrai.Click += new System.EventHandler(this.btnEstrai_Click);
            // 
            // lblArts
            // 
            this.lblArts.AutoSize = true;
            this.lblArts.Location = new System.Drawing.Point(539, 527);
            this.lblArts.Name = "lblArts";
            this.lblArts.Size = new System.Drawing.Size(13, 13);
            this.lblArts.TabIndex = 11;
            this.lblArts.Text = "0";
            // 
            // lblCos
            // 
            this.lblCos.AutoSize = true;
            this.lblCos.Location = new System.Drawing.Point(767, 527);
            this.lblCos.Name = "lblCos";
            this.lblCos.Size = new System.Drawing.Size(13, 13);
            this.lblCos.TabIndex = 15;
            this.lblCos.Text = "0";
            this.lblCos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(486, 527);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 17;
            this.label2.Text = "Articoli";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(698, 527);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(37, 13);
            this.label4.TabIndex = 18;
            this.label4.Text = "Valore";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 547);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(882, 5);
            this.progressBar1.TabIndex = 19;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(369, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 13);
            this.label1.TabIndex = 13;
            this.label1.Text = "Negozio";
            // 
            // cmbNeg
            // 
            this.cmbNeg.FormattingEnabled = true;
            this.cmbNeg.Location = new System.Drawing.Point(421, 31);
            this.cmbNeg.Name = "cmbNeg";
            this.cmbNeg.Size = new System.Drawing.Size(227, 21);
            this.cmbNeg.TabIndex = 14;
            // 
            // chkArtAtt
            // 
            this.chkArtAtt.AutoSize = true;
            this.chkArtAtt.Checked = true;
            this.chkArtAtt.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkArtAtt.Location = new System.Drawing.Point(15, 38);
            this.chkArtAtt.Name = "chkArtAtt";
            this.chkArtAtt.Size = new System.Drawing.Size(105, 17);
            this.chkArtAtt.TabIndex = 20;
            this.chkArtAtt.Text = "Solo articoli attivi";
            this.chkArtAtt.UseVisualStyleBackColor = true;
            // 
            // lblMsg
            // 
            this.lblMsg.AutoSize = true;
            this.lblMsg.Location = new System.Drawing.Point(4, 527);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(16, 13);
            this.lblMsg.TabIndex = 21;
            this.lblMsg.Text = "...";
            // 
            // frmUtyPrnGiacenza
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(882, 552);
            this.ControlBox = false;
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.chkArtAtt);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblCos);
            this.Controls.Add(this.cmbNeg);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblArts);
            this.Controls.Add(this.btnEstrai);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyPrnGiacenza";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Stampa giacenza";
            this.Load += new System.EventHandler(this.frmUtyPrnGiacenza_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyPrnGiacenza_KeyDown);
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
        private System.Windows.Forms.Label lblArts;
        private System.Windows.Forms.ToolStripMenuItem excelToolStripMenuItem;
        private System.Windows.Forms.Label lblCos;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbNeg;
        private System.Windows.Forms.ToolStripMenuItem csvToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excel2ToolStripMenuItem;
        private System.Windows.Forms.CheckBox chkArtAtt;
        private System.Windows.Forms.Label lblMsg;
    }
}
