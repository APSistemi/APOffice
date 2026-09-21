namespace APOffice
{
    partial class frmGesDocPrint
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
            this.btnPrint = new System.Windows.Forms.Button();
            this.cmbCauTra = new System.Windows.Forms.ComboBox();
            this.cmbAspBen = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtNumCol = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtDocNot = new System.Windows.Forms.TextBox();
            this.lblCopie = new System.Windows.Forms.Label();
            this.chkCopie = new System.Windows.Forms.CheckBox();
            this.chkPrezzi = new System.Windows.Forms.CheckBox();
            this.lblPrezzi = new System.Windows.Forms.Label();
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
            this.menuStrip1.Size = new System.Drawing.Size(708, 24);
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
            // btnPrint
            // 
            this.btnPrint.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPrint.Location = new System.Drawing.Point(430, 41);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(240, 151);
            this.btnPrint.TabIndex = 2;
            this.btnPrint.Text = "Stampa";
            this.btnPrint.UseVisualStyleBackColor = true;
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // cmbCauTra
            // 
            this.cmbCauTra.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbCauTra.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbCauTra.FormattingEnabled = true;
            this.cmbCauTra.Location = new System.Drawing.Point(159, 42);
            this.cmbCauTra.Name = "cmbCauTra";
            this.cmbCauTra.Size = new System.Drawing.Size(240, 32);
            this.cmbCauTra.TabIndex = 3;
            // 
            // cmbAspBen
            // 
            this.cmbAspBen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAspBen.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbAspBen.FormattingEnabled = true;
            this.cmbAspBen.Location = new System.Drawing.Point(159, 83);
            this.cmbAspBen.Name = "cmbAspBen";
            this.cmbAspBen.Size = new System.Drawing.Size(240, 32);
            this.cmbAspBen.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(11, 49);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(141, 18);
            this.label1.TabIndex = 6;
            this.label1.Text = "Causale di trasporto";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(11, 90);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(120, 18);
            this.label2.TabIndex = 7;
            this.label2.Text = "Aspetto esteriore";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(11, 135);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(92, 18);
            this.label3.TabIndex = 8;
            this.label3.Text = "Numero colli";
            // 
            // txtNumCol
            // 
            this.txtNumCol.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtNumCol.Location = new System.Drawing.Point(159, 128);
            this.txtNumCol.Name = "txtNumCol";
            this.txtNumCol.Size = new System.Drawing.Size(100, 26);
            this.txtNumCol.TabIndex = 9;
            this.txtNumCol.Text = "1";
            this.txtNumCol.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(12, 224);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(40, 18);
            this.label4.TabIndex = 10;
            this.label4.Text = "Note";
            // 
            // txtDocNot
            // 
            this.txtDocNot.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDocNot.Location = new System.Drawing.Point(159, 216);
            this.txtDocNot.Name = "txtDocNot";
            this.txtDocNot.Size = new System.Drawing.Size(511, 26);
            this.txtDocNot.TabIndex = 11;
            // 
            // lblCopie
            // 
            this.lblCopie.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F);
            this.lblCopie.Location = new System.Drawing.Point(11, 169);
            this.lblCopie.Name = "lblCopie";
            this.lblCopie.Size = new System.Drawing.Size(100, 46);
            this.lblCopie.TabIndex = 12;
            this.lblCopie.Text = "2 Copie su stesso foglio";
            // 
            // chkCopie
            // 
            this.chkCopie.AutoSize = true;
            this.chkCopie.Location = new System.Drawing.Point(159, 178);
            this.chkCopie.Name = "chkCopie";
            this.chkCopie.Size = new System.Drawing.Size(15, 14);
            this.chkCopie.TabIndex = 13;
            this.chkCopie.UseVisualStyleBackColor = true;
            // 
            // chkPrezzi
            // 
            this.chkPrezzi.AutoSize = true;
            this.chkPrezzi.Checked = true;
            this.chkPrezzi.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkPrezzi.Location = new System.Drawing.Point(379, 178);
            this.chkPrezzi.Name = "chkPrezzi";
            this.chkPrezzi.Size = new System.Drawing.Size(15, 14);
            this.chkPrezzi.TabIndex = 14;
            this.chkPrezzi.UseVisualStyleBackColor = true;
            // 
            // lblPrezzi
            // 
            this.lblPrezzi.AutoSize = true;
            this.lblPrezzi.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPrezzi.Location = new System.Drawing.Point(283, 175);
            this.lblPrezzi.Name = "lblPrezzi";
            this.lblPrezzi.Size = new System.Drawing.Size(82, 18);
            this.lblPrezzi.TabIndex = 15;
            this.lblPrezzi.Text = "Con Prezzi";
            // 
            // frmGesDocPrint
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(708, 261);
            this.ControlBox = false;
            this.Controls.Add(this.lblPrezzi);
            this.Controls.Add(this.chkPrezzi);
            this.Controls.Add(this.chkCopie);
            this.Controls.Add(this.lblCopie);
            this.Controls.Add(this.txtDocNot);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtNumCol);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbAspBen);
            this.Controls.Add(this.cmbCauTra);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesDocPrint";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Stampa documento";
            this.Load += new System.EventHandler(this.frmGesDocPrint_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesDocPrint_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Button btnPrint;
        private System.Windows.Forms.ComboBox cmbCauTra;
        private System.Windows.Forms.ComboBox cmbAspBen;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtNumCol;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtDocNot;
        private System.Windows.Forms.Label lblCopie;
        private System.Windows.Forms.CheckBox chkCopie;
        private System.Windows.Forms.CheckBox chkPrezzi;
        private System.Windows.Forms.Label lblPrezzi;
    }
}