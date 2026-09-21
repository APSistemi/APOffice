namespace APOffice
{
    partial class frmGesStatStoCancella
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
            this.btnCan = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblMsg1 = new System.Windows.Forms.Label();
            this.lblMsg2 = new System.Windows.Forms.Label();
            this.chkSta = new System.Windows.Forms.CheckBox();
            this.chkDoc = new System.Windows.Forms.CheckBox();
            this.cmbYea = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.chkDve = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.chkMov = new System.Windows.Forms.CheckBox();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(741, 24);
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
            // btnCan
            // 
            this.btnCan.Location = new System.Drawing.Point(510, 92);
            this.btnCan.Name = "btnCan";
            this.btnCan.Size = new System.Drawing.Size(171, 135);
            this.btnCan.TabIndex = 68;
            this.btnCan.Text = "Cancella";
            this.btnCan.UseVisualStyleBackColor = true;
            this.btnCan.Click += new System.EventHandler(this.btnCan_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 321);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(741, 23);
            this.progressBar1.TabIndex = 69;
            // 
            // lblMsg1
            // 
            this.lblMsg1.AutoSize = true;
            this.lblMsg1.Location = new System.Drawing.Point(5, 303);
            this.lblMsg1.Name = "lblMsg1";
            this.lblMsg1.Size = new System.Drawing.Size(16, 13);
            this.lblMsg1.TabIndex = 70;
            this.lblMsg1.Text = "...";
            // 
            // lblMsg2
            // 
            this.lblMsg2.AutoSize = true;
            this.lblMsg2.Location = new System.Drawing.Point(713, 303);
            this.lblMsg2.Name = "lblMsg2";
            this.lblMsg2.Size = new System.Drawing.Size(16, 13);
            this.lblMsg2.TabIndex = 71;
            this.lblMsg2.Text = "...";
            // 
            // chkSta
            // 
            this.chkSta.AutoSize = true;
            this.chkSta.Checked = true;
            this.chkSta.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSta.Location = new System.Drawing.Point(261, 104);
            this.chkSta.Name = "chkSta";
            this.chkSta.Size = new System.Drawing.Size(15, 14);
            this.chkSta.TabIndex = 77;
            this.chkSta.UseVisualStyleBackColor = true;
            // 
            // chkDoc
            // 
            this.chkDoc.AutoSize = true;
            this.chkDoc.Checked = true;
            this.chkDoc.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDoc.Location = new System.Drawing.Point(261, 146);
            this.chkDoc.Name = "chkDoc";
            this.chkDoc.Size = new System.Drawing.Size(15, 14);
            this.chkDoc.TabIndex = 76;
            this.chkDoc.UseVisualStyleBackColor = true;
            // 
            // cmbYea
            // 
            this.cmbYea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbYea.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbYea.FormattingEnabled = true;
            this.cmbYea.Location = new System.Drawing.Point(160, 60);
            this.cmbYea.Name = "cmbYea";
            this.cmbYea.Size = new System.Drawing.Size(116, 24);
            this.cmbYea.TabIndex = 75;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(69, 106);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(97, 13);
            this.label3.TabIndex = 74;
            this.label3.Text = "Venduto statistiche";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(69, 147);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(101, 13);
            this.label2.TabIndex = 73;
            this.label2.Text = "Documenti acquisto";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(69, 66);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 13);
            this.label1.TabIndex = 72;
            this.label1.Text = "Fino all\'anno";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(69, 192);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(96, 13);
            this.label4.TabIndex = 78;
            this.label4.Text = "Documenti vendita";
            // 
            // chkDve
            // 
            this.chkDve.AutoSize = true;
            this.chkDve.Checked = true;
            this.chkDve.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDve.Location = new System.Drawing.Point(261, 192);
            this.chkDve.Name = "chkDve";
            this.chkDve.Size = new System.Drawing.Size(15, 14);
            this.chkDve.TabIndex = 79;
            this.chkDve.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(69, 236);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(108, 13);
            this.label5.TabIndex = 80;
            this.label5.Text = "Documenti movimenti";
            // 
            // chkMov
            // 
            this.chkMov.AutoSize = true;
            this.chkMov.Checked = true;
            this.chkMov.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkMov.Location = new System.Drawing.Point(261, 236);
            this.chkMov.Name = "chkMov";
            this.chkMov.Size = new System.Drawing.Size(15, 14);
            this.chkMov.TabIndex = 81;
            this.chkMov.UseVisualStyleBackColor = true;
            // 
            // frmGesStatStoCancella
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(741, 338);
            this.ControlBox = false;
            this.Controls.Add(this.chkMov);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.chkDve);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.chkSta);
            this.Controls.Add(this.chkDoc);
            this.Controls.Add(this.cmbYea);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblMsg2);
            this.Controls.Add(this.lblMsg1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnCan);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesStatStoCancella";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Cancellazione dati";
            this.Load += new System.EventHandler(this.frmGesStaStoCancella_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesStaStoCancella_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Button btnCan;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblMsg1;
        private System.Windows.Forms.Label lblMsg2;
        private System.Windows.Forms.CheckBox chkSta;
        private System.Windows.Forms.CheckBox chkDoc;
        private System.Windows.Forms.ComboBox cmbYea;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox chkDve;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox chkMov;
    }
}