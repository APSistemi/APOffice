namespace APOffice
{
    partial class frmUtyTessSostituzione
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
            this.label1 = new System.Windows.Forms.Label();
            this.lblOldTes = new System.Windows.Forms.Label();
            this.lblOldDes = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtNewCod = new System.Windows.Forms.TextBox();
            this.btnSeek = new System.Windows.Forms.Button();
            this.lblNewDes = new System.Windows.Forms.Label();
            this.lblOldPun = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnOk = new System.Windows.Forms.Button();
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
            this.menuStrip1.Size = new System.Drawing.Size(587, 24);
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
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(11, 66);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(138, 20);
            this.label1.TabIndex = 1;
            this.label1.Text = "Tessera vecchia";
            // 
            // lblOldTes
            // 
            this.lblOldTes.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblOldTes.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold);
            this.lblOldTes.Location = new System.Drawing.Point(161, 61);
            this.lblOldTes.Name = "lblOldTes";
            this.lblOldTes.Size = new System.Drawing.Size(183, 23);
            this.lblOldTes.TabIndex = 2;
            this.lblOldTes.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblOldDes
            // 
            this.lblOldDes.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblOldDes.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold);
            this.lblOldDes.Location = new System.Drawing.Point(15, 101);
            this.lblOldDes.Name = "lblOldDes";
            this.lblOldDes.Size = new System.Drawing.Size(556, 23);
            this.lblOldDes.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 180);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(126, 20);
            this.label2.TabIndex = 4;
            this.label2.Text = "Tessera nuova";
            // 
            // txtNewCod
            // 
            this.txtNewCod.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold);
            this.txtNewCod.Location = new System.Drawing.Point(161, 173);
            this.txtNewCod.Name = "txtNewCod";
            this.txtNewCod.ReadOnly = true;
            this.txtNewCod.Size = new System.Drawing.Size(183, 26);
            this.txtNewCod.TabIndex = 5;
            this.txtNewCod.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtNewCod.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNewCod_KeyDown);
            // 
            // btnSeek
            // 
            this.btnSeek.Location = new System.Drawing.Point(367, 175);
            this.btnSeek.Name = "btnSeek";
            this.btnSeek.Size = new System.Drawing.Size(54, 23);
            this.btnSeek.TabIndex = 6;
            this.btnSeek.Text = "...";
            this.btnSeek.UseVisualStyleBackColor = true;
            this.btnSeek.Click += new System.EventHandler(this.btnSeek_Click);
            // 
            // lblNewDes
            // 
            this.lblNewDes.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblNewDes.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold);
            this.lblNewDes.Location = new System.Drawing.Point(25, 209);
            this.lblNewDes.Name = "lblNewDes";
            this.lblNewDes.Size = new System.Drawing.Size(546, 23);
            this.lblNewDes.TabIndex = 7;
            // 
            // lblOldPun
            // 
            this.lblOldPun.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblOldPun.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold);
            this.lblOldPun.Location = new System.Drawing.Point(495, 61);
            this.lblOldPun.Name = "lblOldPun";
            this.lblOldPun.Size = new System.Drawing.Size(76, 23);
            this.lblOldPun.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(420, 66);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(50, 20);
            this.label3.TabIndex = 9;
            this.label3.Text = "Punti";
            // 
            // btnOk
            // 
            this.btnOk.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOk.Location = new System.Drawing.Point(367, 280);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(206, 113);
            this.btnOk.TabIndex = 10;
            this.btnOk.Text = "Conferma";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // frmUtyTessSostituzione
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(587, 440);
            this.ControlBox = false;
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblOldPun);
            this.Controls.Add(this.lblNewDes);
            this.Controls.Add(this.btnSeek);
            this.Controls.Add(this.txtNewCod);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblOldDes);
            this.Controls.Add(this.lblOldTes);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyTessSostituzione";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Sostituzione tessera";
            this.Load += new System.EventHandler(this.frmUtyTessSostituzione_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyTessSostituzione_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblOldTes;
        private System.Windows.Forms.Label lblOldDes;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtNewCod;
        private System.Windows.Forms.Button btnSeek;
        private System.Windows.Forms.Label lblNewDes;
        private System.Windows.Forms.Label lblOldPun;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnOk;
    }
}