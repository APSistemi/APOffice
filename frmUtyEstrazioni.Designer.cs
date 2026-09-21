namespace APOffice
{
    partial class frmUtyEstrazioni
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
            this.estrazioneToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.generazioneToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.chkAtt = new System.Windows.Forms.CheckBox();
            this.dgv1 = new APOffice.APDataGridView();
            this.chkKonzOpen = new System.Windows.Forms.CheckBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblCnt = new System.Windows.Forms.Label();
            this.chkRep = new System.Windows.Forms.CheckBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.estrazioneToolStripMenuItem,
            this.generazioneToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(703, 24);
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
            // estrazioneToolStripMenuItem
            // 
            this.estrazioneToolStripMenuItem.Name = "estrazioneToolStripMenuItem";
            this.estrazioneToolStripMenuItem.Size = new System.Drawing.Size(72, 20);
            this.estrazioneToolStripMenuItem.Text = "Estrazione";
            this.estrazioneToolStripMenuItem.Click += new System.EventHandler(this.estrazioneToolStripMenuItem_Click);
            // 
            // generazioneToolStripMenuItem
            // 
            this.generazioneToolStripMenuItem.Name = "generazioneToolStripMenuItem";
            this.generazioneToolStripMenuItem.Size = new System.Drawing.Size(84, 20);
            this.generazioneToolStripMenuItem.Text = "Generazione";
            this.generazioneToolStripMenuItem.Click += new System.EventHandler(this.generazioneToolStripMenuItem_Click);
            // 
            // chkAtt
            // 
            this.chkAtt.AutoSize = true;
            this.chkAtt.Checked = true;
            this.chkAtt.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkAtt.Location = new System.Drawing.Point(528, 29);
            this.chkAtt.Name = "chkAtt";
            this.chkAtt.Size = new System.Drawing.Size(72, 17);
            this.chkAtt.TabIndex = 1;
            this.chkAtt.Text = "Solo attivi";
            this.chkAtt.UseVisualStyleBackColor = true;
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(6, 52);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(691, 475);
            this.dgv1.TabIndex = 2;
            // 
            // chkKonzOpen
            // 
            this.chkKonzOpen.AutoSize = true;
            this.chkKonzOpen.Checked = true;
            this.chkKonzOpen.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkKonzOpen.Location = new System.Drawing.Point(11, 29);
            this.chkKonzOpen.Name = "chkKonzOpen";
            this.chkKonzOpen.Size = new System.Drawing.Size(76, 17);
            this.chkKonzOpen.TabIndex = 3;
            this.chkKonzOpen.Text = "KonzOpen";
            this.chkKonzOpen.UseVisualStyleBackColor = true;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 539);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(703, 8);
            this.progressBar1.TabIndex = 4;
            // 
            // lblCnt
            // 
            this.lblCnt.AutoSize = true;
            this.lblCnt.Location = new System.Drawing.Point(642, 31);
            this.lblCnt.Name = "lblCnt";
            this.lblCnt.Size = new System.Drawing.Size(13, 13);
            this.lblCnt.TabIndex = 5;
            this.lblCnt.Text = "0";
            this.lblCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // chkRep
            // 
            this.chkRep.AutoSize = true;
            this.chkRep.Location = new System.Drawing.Point(209, 27);
            this.chkRep.Name = "chkRep";
            this.chkRep.Size = new System.Drawing.Size(64, 17);
            this.chkRep.TabIndex = 6;
            this.chkRep.Text = "Reparto";
            this.chkRep.UseVisualStyleBackColor = true;
            this.chkRep.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // frmUtyEstrazioni
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(703, 547);
            this.ControlBox = false;
            this.Controls.Add(this.chkRep);
            this.Controls.Add(this.lblCnt);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.chkKonzOpen);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.chkAtt);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyEstrazioni";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Estrazioni anagrafiche";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.CheckBox chkAtt;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.ToolStripMenuItem estrazioneToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem generazioneToolStripMenuItem;
        private System.Windows.Forms.CheckBox chkKonzOpen;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblCnt;
        private System.Windows.Forms.CheckBox chkRep;
    }
}
