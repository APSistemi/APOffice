namespace APOffice
{
    partial class frmAnaArtLegami
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
            this.utilityToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cancellaAnnullatiToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lblArt = new System.Windows.Forms.Label();
            this.dgv1 = new APOffice.APDataGridView();
            this.btnSeek = new System.Windows.Forms.Button();
            this.txtSeek = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblInc = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
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
            this.utilityToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(724, 29);
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
            // utilityToolStripMenuItem
            // 
            this.utilityToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.utilityToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cancellaAnnullatiToolStripMenuItem});
            this.utilityToolStripMenuItem.Name = "utilityToolStripMenuItem";
            this.utilityToolStripMenuItem.Size = new System.Drawing.Size(63, 25);
            this.utilityToolStripMenuItem.Text = "Utility";
            // 
            // cancellaAnnullatiToolStripMenuItem
            // 
            this.cancellaAnnullatiToolStripMenuItem.Name = "cancellaAnnullatiToolStripMenuItem";
            this.cancellaAnnullatiToolStripMenuItem.Size = new System.Drawing.Size(202, 26);
            this.cancellaAnnullatiToolStripMenuItem.Text = "Cancella annullati";
            this.cancellaAnnullatiToolStripMenuItem.Click += new System.EventHandler(this.cancellaAnnullatiToolStripMenuItem_Click);
            // 
            // lblArt
            // 
            this.lblArt.BackColor = System.Drawing.Color.Beige;
            this.lblArt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblArt.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblArt.Location = new System.Drawing.Point(0, 30);
            this.lblArt.Name = "lblArt";
            this.lblArt.Size = new System.Drawing.Size(724, 33);
            this.lblArt.TabIndex = 1;
            this.lblArt.Text = "label1";
            this.lblArt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 137);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(700, 413);
            this.dgv1.TabIndex = 2;
            this.dgv1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellContentClick);
            this.dgv1.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellContentDoubleClick);
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            this.dgv1.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellEndEdit);
            this.dgv1.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgv1_CellFormatting);
            this.dgv1.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgv1_EditingControlShowing);
            // 
            // btnSeek
            // 
            this.btnSeek.Location = new System.Drawing.Point(513, 85);
            this.btnSeek.Name = "btnSeek";
            this.btnSeek.Size = new System.Drawing.Size(62, 32);
            this.btnSeek.TabIndex = 8;
            this.btnSeek.Text = "...";
            this.btnSeek.UseVisualStyleBackColor = true;
            this.btnSeek.Click += new System.EventHandler(this.btnSeek_Click);
            // 
            // txtSeek
            // 
            this.txtSeek.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSeek.Location = new System.Drawing.Point(126, 88);
            this.txtSeek.Name = "txtSeek";
            this.txtSeek.Size = new System.Drawing.Size(378, 29);
            this.txtSeek.TabIndex = 7;
            this.txtSeek.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSeek_KeyDown);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 84);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(134, 46);
            this.label1.TabIndex = 6;
            this.label1.Text = "Ricerca nuovo articolo";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(522, 566);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Incidenza";
            // 
            // lblInc
            // 
            this.lblInc.BackColor = System.Drawing.Color.White;
            this.lblInc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblInc.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInc.Location = new System.Drawing.Point(603, 556);
            this.lblInc.Name = "lblInc";
            this.lblInc.Size = new System.Drawing.Size(100, 23);
            this.lblInc.TabIndex = 10;
            this.lblInc.Text = "0.00";
            this.lblInc.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnOK
            // 
            this.btnOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOK.Location = new System.Drawing.Point(603, 71);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(109, 60);
            this.btnOK.TabIndex = 11;
            this.btnOK.Text = "generazione barcode";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // frmAnaArtLegami
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(724, 586);
            this.ControlBox = false;
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.lblInc);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnSeek);
            this.Controls.Add(this.txtSeek);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.lblArt);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmAnaArtLegami";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Legami articolo";
            this.Load += new System.EventHandler(this.frmAnaArtDistinta_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAnaArtDistinta_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Label lblArt;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Button btnSeek;
        private System.Windows.Forms.TextBox txtSeek;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblInc;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.ToolStripMenuItem utilityToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cancellaAnnullatiToolStripMenuItem;
    }
}
