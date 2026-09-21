namespace APOffice
{
    partial class frmGesStatMovLegami
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
            this.letturaBilanceToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.lblArt = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lblQkg = new System.Windows.Forms.Label();
            this.lblLegQkg = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblBilQkg = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.letturaBilanceToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(820, 29);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // esciToolStripMenuItem
            // 
            this.esciToolStripMenuItem.Name = "esciToolStripMenuItem";
            this.esciToolStripMenuItem.Size = new System.Drawing.Size(48, 25);
            this.esciToolStripMenuItem.Text = "Esci";
            this.esciToolStripMenuItem.Click += new System.EventHandler(this.esciToolStripMenuItem_Click);
            // 
            // letturaBilanceToolStripMenuItem
            // 
            this.letturaBilanceToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.letturaBilanceToolStripMenuItem.Name = "letturaBilanceToolStripMenuItem";
            this.letturaBilanceToolStripMenuItem.Size = new System.Drawing.Size(124, 25);
            this.letturaBilanceToolStripMenuItem.Text = "Lettura Bilance";
            this.letturaBilanceToolStripMenuItem.Click += new System.EventHandler(this.letturaBilanceToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 84);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(795, 457);
            this.dgv1.TabIndex = 2;
            // 
            // lblArt
            // 
            this.lblArt.BackColor = System.Drawing.Color.Beige;
            this.lblArt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblArt.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblArt.Location = new System.Drawing.Point(0, 31);
            this.lblArt.Name = "lblArt";
            this.lblArt.Size = new System.Drawing.Size(636, 33);
            this.lblArt.TabIndex = 41;
            this.lblArt.Text = "label1";
            this.lblArt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(642, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(44, 16);
            this.label1.TabIndex = 42;
            this.label1.Text = "Peso";
            // 
            // lblQkg
            // 
            this.lblQkg.BackColor = System.Drawing.Color.White;
            this.lblQkg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblQkg.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQkg.Location = new System.Drawing.Point(707, 34);
            this.lblQkg.Name = "lblQkg";
            this.lblQkg.Size = new System.Drawing.Size(100, 38);
            this.lblQkg.TabIndex = 43;
            this.lblQkg.Text = "0.00";
            this.lblQkg.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblLegQkg
            // 
            this.lblLegQkg.BackColor = System.Drawing.Color.White;
            this.lblLegQkg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblLegQkg.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLegQkg.Location = new System.Drawing.Point(707, 544);
            this.lblLegQkg.Name = "lblLegQkg";
            this.lblLegQkg.Size = new System.Drawing.Size(100, 38);
            this.lblLegQkg.TabIndex = 45;
            this.lblLegQkg.Text = "0.00";
            this.lblLegQkg.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(573, 558);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(118, 16);
            this.label3.TabIndex = 44;
            this.label3.Text = "Peso movimenti";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(264, 557);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(99, 16);
            this.label2.TabIndex = 46;
            this.label2.Text = "Peso bilance";
            // 
            // lblBilQkg
            // 
            this.lblBilQkg.BackColor = System.Drawing.Color.White;
            this.lblBilQkg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblBilQkg.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblBilQkg.Location = new System.Drawing.Point(378, 544);
            this.lblBilQkg.Name = "lblBilQkg";
            this.lblBilQkg.Size = new System.Drawing.Size(100, 38);
            this.lblBilQkg.TabIndex = 47;
            this.lblBilQkg.Text = "0.00";
            this.lblBilQkg.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(1, 587);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(819, 7);
            this.progressBar1.TabIndex = 48;
            // 
            // frmGesStatMovLegami
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(820, 594);
            this.ControlBox = false;
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.lblBilQkg);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblLegQkg);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblQkg);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblArt);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesStatMovLegami";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Movimenti per lotto";
            this.Load += new System.EventHandler(this.frmGesStaMovLegami_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesStaMovLegami_KeyDown);
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
        private System.Windows.Forms.Label lblArt;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblQkg;
        private System.Windows.Forms.Label lblLegQkg;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ToolStripMenuItem letturaBilanceToolStripMenuItem;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblBilQkg;
        private System.Windows.Forms.ProgressBar progressBar1;
    }
}
