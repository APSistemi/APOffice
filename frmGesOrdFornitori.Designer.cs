namespace APOffice
{
    partial class frmGesOrdFornitori
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.apPhoneFTPToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importazioneDaTerminalinoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.stampaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tuttoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.erroriToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.btnInvio = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblOftYea = new System.Windows.Forms.Label();
            this.lblOftNum = new System.Windows.Forms.Label();
            this.cmbOftSta = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.lblOrfTot = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.cmbOftFor = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.lblOrfRig = new System.Windows.Forms.Label();
            this.lblOrfCol = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtSeek = new System.Windows.Forms.TextBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.apPhoneFTPToolStripMenuItem,
            this.importazioneDaTerminalinoToolStripMenuItem,
            this.stampaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(793, 24);
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
            // apPhoneFTPToolStripMenuItem
            // 
            this.apPhoneFTPToolStripMenuItem.Name = "apPhoneFTPToolStripMenuItem";
            this.apPhoneFTPToolStripMenuItem.Size = new System.Drawing.Size(12, 20);
            // 
            // importazioneDaTerminalinoToolStripMenuItem
            // 
            this.importazioneDaTerminalinoToolStripMenuItem.Name = "importazioneDaTerminalinoToolStripMenuItem";
            this.importazioneDaTerminalinoToolStripMenuItem.Size = new System.Drawing.Size(169, 20);
            this.importazioneDaTerminalinoToolStripMenuItem.Text = "Importazione da terminalino";
            this.importazioneDaTerminalinoToolStripMenuItem.Click += new System.EventHandler(this.importazioneDaTerminalinoToolStripMenuItem_Click);
            // 
            // stampaToolStripMenuItem
            // 
            this.stampaToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tuttoToolStripMenuItem,
            this.erroriToolStripMenuItem});
            this.stampaToolStripMenuItem.Name = "stampaToolStripMenuItem";
            this.stampaToolStripMenuItem.Size = new System.Drawing.Size(59, 20);
            this.stampaToolStripMenuItem.Text = "Stampa";
            // 
            // tuttoToolStripMenuItem
            // 
            this.tuttoToolStripMenuItem.Name = "tuttoToolStripMenuItem";
            this.tuttoToolStripMenuItem.Size = new System.Drawing.Size(103, 22);
            this.tuttoToolStripMenuItem.Text = "Tutto";
            this.tuttoToolStripMenuItem.Click += new System.EventHandler(this.tuttoToolStripMenuItem_Click);
            // 
            // erroriToolStripMenuItem
            // 
            this.erroriToolStripMenuItem.Name = "erroriToolStripMenuItem";
            this.erroriToolStripMenuItem.Size = new System.Drawing.Size(103, 22);
            this.erroriToolStripMenuItem.Text = "Errori";
            this.erroriToolStripMenuItem.Click += new System.EventHandler(this.erroriToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 75);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(796, 390);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            this.dgv1.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellEndEdit);
            this.dgv1.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgv1_CellFormatting);
            this.dgv1.CurrentCellChanged += new System.EventHandler(this.dgv1_CurrentCellChanged);
            this.dgv1.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgv1_EditingControlShowing);
            // 
            // btnInvio
            // 
            this.btnInvio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnInvio.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnInvio.Location = new System.Drawing.Point(677, 28);
            this.btnInvio.Name = "btnInvio";
            this.btnInvio.Size = new System.Drawing.Size(131, 40);
            this.btnInvio.TabIndex = 2;
            this.btnInvio.Text = "Invio";
            this.btnInvio.UseVisualStyleBackColor = true;
            this.btnInvio.Click += new System.EventHandler(this.btnInvio_Click);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(645, 485);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 15);
            this.label1.TabIndex = 99;
            this.label1.Text = "Totale costo";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(282, 43);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(36, 15);
            this.label2.TabIndex = 101;
            this.label2.Text = "Anno";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(383, 43);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(51, 15);
            this.label3.TabIndex = 102;
            this.label3.Text = "Numero";
            // 
            // lblOftYea
            // 
            this.lblOftYea.BackColor = System.Drawing.Color.White;
            this.lblOftYea.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblOftYea.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblOftYea.Location = new System.Drawing.Point(321, 38);
            this.lblOftYea.Name = "lblOftYea";
            this.lblOftYea.Size = new System.Drawing.Size(57, 25);
            this.lblOftYea.TabIndex = 103;
            this.lblOftYea.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblOftNum
            // 
            this.lblOftNum.BackColor = System.Drawing.Color.White;
            this.lblOftNum.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblOftNum.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblOftNum.Location = new System.Drawing.Point(434, 38);
            this.lblOftNum.Name = "lblOftNum";
            this.lblOftNum.Size = new System.Drawing.Size(57, 25);
            this.lblOftNum.TabIndex = 104;
            this.lblOftNum.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cmbOftSta
            // 
            this.cmbOftSta.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbOftSta.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbOftSta.FormattingEnabled = true;
            this.cmbOftSta.Location = new System.Drawing.Point(539, 39);
            this.cmbOftSta.Name = "cmbOftSta";
            this.cmbOftSta.Size = new System.Drawing.Size(121, 23);
            this.cmbOftSta.TabIndex = 105;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(497, 43);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(34, 15);
            this.label4.TabIndex = 106;
            this.label4.Text = "Stato";
            // 
            // lblOrfTot
            // 
            this.lblOrfTot.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblOrfTot.BackColor = System.Drawing.Color.White;
            this.lblOrfTot.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblOrfTot.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblOrfTot.Location = new System.Drawing.Point(720, 480);
            this.lblOrfTot.Name = "lblOrfTot";
            this.lblOrfTot.Size = new System.Drawing.Size(88, 25);
            this.lblOrfTot.TabIndex = 107;
            this.lblOrfTot.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(12, 43);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(55, 15);
            this.label5.TabIndex = 108;
            this.label5.Text = "Fornitore";
            // 
            // cmbOftFor
            // 
            this.cmbOftFor.AllowDrop = true;
            this.cmbOftFor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.Simple;
            this.cmbOftFor.Enabled = false;
            this.cmbOftFor.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbOftFor.FormattingEnabled = true;
            this.cmbOftFor.Location = new System.Drawing.Point(70, 38);
            this.cmbOftFor.Name = "cmbOftFor";
            this.cmbOftFor.Size = new System.Drawing.Size(206, 25);
            this.cmbOftFor.TabIndex = 109;
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(375, 485);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(37, 15);
            this.label6.TabIndex = 110;
            this.label6.Text = "Righe";
            // 
            // lblOrfRig
            // 
            this.lblOrfRig.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblOrfRig.BackColor = System.Drawing.Color.White;
            this.lblOrfRig.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblOrfRig.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblOrfRig.Location = new System.Drawing.Point(415, 480);
            this.lblOrfRig.Name = "lblOrfRig";
            this.lblOrfRig.Size = new System.Drawing.Size(80, 25);
            this.lblOrfRig.TabIndex = 111;
            this.lblOrfRig.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblOrfCol
            // 
            this.lblOrfCol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblOrfCol.BackColor = System.Drawing.Color.White;
            this.lblOrfCol.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblOrfCol.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblOrfCol.Location = new System.Drawing.Point(545, 480);
            this.lblOrfCol.Name = "lblOrfCol";
            this.lblOrfCol.Size = new System.Drawing.Size(80, 25);
            this.lblOrfCol.TabIndex = 112;
            this.lblOrfCol.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label9
            // 
            this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(510, 485);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(30, 15);
            this.label9.TabIndex = 113;
            this.label9.Text = "Colli";
            // 
            // label7
            // 
            this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label7.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(12, 475);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(121, 35);
            this.label7.TabIndex = 114;
            this.label7.Text = "Ricerca per descrizione sull\'ordine";
            // 
            // txtSeek
            // 
            this.txtSeek.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtSeek.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSeek.Location = new System.Drawing.Point(135, 480);
            this.txtSeek.MaxLength = 20;
            this.txtSeek.Name = "txtSeek";
            this.txtSeek.Size = new System.Drawing.Size(220, 25);
            this.txtSeek.TabIndex = 115;
            this.txtSeek.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSeek_KeyDown);
            // 
            // frmGesOrdFornitori
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(820, 520);
            this.Controls.Add(this.txtSeek);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.lblOrfCol);
            this.Controls.Add(this.lblOrfRig);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.cmbOftFor);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.lblOrfTot);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cmbOftSta);
            this.Controls.Add(this.lblOftNum);
            this.Controls.Add(this.lblOftYea);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnInvio);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(780, 460);
            this.Name = "frmGesOrdFornitori";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ordine a fornitori";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmGesOrdFornitori_FormClosing);
            this.Load += new System.EventHandler(this.frmGesOrdFornitori_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesOrdFornitori_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importazioneDaTerminalinoToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Button btnInvio;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblOftYea;
        private System.Windows.Forms.Label lblOftNum;
        private System.Windows.Forms.ComboBox cmbOftSta;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblOrfTot;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cmbOftFor;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label lblOrfRig;
        private System.Windows.Forms.Label lblOrfCol;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ToolStripMenuItem stampaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tuttoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem erroriToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem apPhoneFTPToolStripMenuItem;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtSeek;
    }
}
