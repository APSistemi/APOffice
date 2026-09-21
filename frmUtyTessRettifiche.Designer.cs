namespace APOffice
{
    partial class frmUtyTessRettifiche
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
            this.chkSgn = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lblDes = new System.Windows.Forms.Label();
            this.txtPun = new System.Windows.Forms.TextBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.lblTesCod = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(591, 24);
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
            // chkSgn
            // 
            this.chkSgn.Appearance = System.Windows.Forms.Appearance.Button;
            this.chkSgn.Checked = true;
            this.chkSgn.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSgn.Font = new System.Drawing.Font("Microsoft Sans Serif", 48F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkSgn.Location = new System.Drawing.Point(246, 93);
            this.chkSgn.Name = "chkSgn";
            this.chkSgn.Size = new System.Drawing.Size(104, 87);
            this.chkSgn.TabIndex = 1;
            this.chkSgn.Text = "+";
            this.chkSgn.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.chkSgn.UseVisualStyleBackColor = true;
            this.chkSgn.Click += new System.EventHandler(this.chkSgn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(138, 124);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 24);
            this.label1.TabIndex = 2;
            this.label1.Text = "Punti in ";
            // 
            // lblDes
            // 
            this.lblDes.AutoSize = true;
            this.lblDes.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDes.Location = new System.Drawing.Point(25, 214);
            this.lblDes.Name = "lblDes";
            this.lblDes.Size = new System.Drawing.Size(170, 24);
            this.lblDes.TabIndex = 3;
            this.lblDes.Text = "Punti aggiungere";
            // 
            // txtPun
            // 
            this.txtPun.Font = new System.Drawing.Font("Microsoft Sans Serif", 26.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPun.Location = new System.Drawing.Point(246, 205);
            this.txtPun.Name = "txtPun";
            this.txtPun.Size = new System.Drawing.Size(150, 47);
            this.txtPun.TabIndex = 4;
            this.txtPun.Text = "0";
            this.txtPun.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // btnOk
            // 
            this.btnOk.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOk.Location = new System.Drawing.Point(354, 282);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(198, 137);
            this.btnOk.TabIndex = 5;
            this.btnOk.Text = "Conferma";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(25, 45);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(91, 24);
            this.label3.TabIndex = 6;
            this.label3.Text = "Tessera:";
            // 
            // lblTesCod
            // 
            this.lblTesCod.AutoSize = true;
            this.lblTesCod.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold);
            this.lblTesCod.Location = new System.Drawing.Point(157, 45);
            this.lblTesCod.Name = "lblTesCod";
            this.lblTesCod.Size = new System.Drawing.Size(28, 24);
            this.lblTesCod.TabIndex = 7;
            this.lblTesCod.Text = "...";
            // 
            // frmUtyTessRettifiche
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(591, 449);
            this.ControlBox = false;
            this.Controls.Add(this.lblTesCod);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.txtPun);
            this.Controls.Add(this.lblDes);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.chkSgn);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyTessRettifiche";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Rettifiche punti";
            this.Load += new System.EventHandler(this.frmUtyTessRettifiche_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyTessRettifiche_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.CheckBox chkSgn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblDes;
        private System.Windows.Forms.TextBox txtPun;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblTesCod;
    }
}