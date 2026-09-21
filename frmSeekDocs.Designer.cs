namespace APOffice
{
    partial class frmSeekDocs
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lottiToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.elenchiToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excelDocumentiToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excelElencoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excelElencoCsvToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excelElencoDettaglioToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importFattureDaScontriniToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fattureElettronicheToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cmbMovFat = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lblDocTip = new System.Windows.Forms.Label();
            this.cmbCauTpd = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbYea = new System.Windows.Forms.ComboBox();
            this.btnNew = new System.Windows.Forms.Button();
            this.btnSeek = new System.Windows.Forms.Button();
            this.btnDocRaggruppa = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.lblTotVal = new System.Windows.Forms.Label();
            this.txtSeek = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.dgv1 = new APOffice.APDataGridView();
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
            this.lottiToolStripMenuItem,
            this.elenchiToolStripMenuItem,
            this.importFattureDaScontriniToolStripMenuItem,
            this.fattureElettronicheToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(803, 29);
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
            // lottiToolStripMenuItem
            // 
            this.lottiToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.lottiToolStripMenuItem.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lottiToolStripMenuItem.Name = "lottiToolStripMenuItem";
            this.lottiToolStripMenuItem.Size = new System.Drawing.Size(57, 25);
            this.lottiToolStripMenuItem.Text = "Lotti";
            this.lottiToolStripMenuItem.Click += new System.EventHandler(this.lottiToolStripMenuItem_Click);
            // 
            // elenchiToolStripMenuItem
            // 
            this.elenchiToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.excelDocumentiToolStripMenuItem,
            this.excelElencoToolStripMenuItem,
            this.excelElencoCsvToolStripMenuItem,
            this.excelElencoDettaglioToolStripMenuItem});
            this.elenchiToolStripMenuItem.Name = "elenchiToolStripMenuItem";
            this.elenchiToolStripMenuItem.Size = new System.Drawing.Size(71, 25);
            this.elenchiToolStripMenuItem.Text = "Elenchi";
            // 
            // excelDocumentiToolStripMenuItem
            // 
            this.excelDocumentiToolStripMenuItem.Name = "excelDocumentiToolStripMenuItem";
            this.excelDocumentiToolStripMenuItem.Size = new System.Drawing.Size(228, 26);
            this.excelDocumentiToolStripMenuItem.Text = "Excel tutti documenti";
            this.excelDocumentiToolStripMenuItem.Click += new System.EventHandler(this.excelDocumentiToolStripMenuItem_Click);
            // 
            // excelElencoToolStripMenuItem
            // 
            this.excelElencoToolStripMenuItem.Name = "excelElencoToolStripMenuItem";
            this.excelElencoToolStripMenuItem.Size = new System.Drawing.Size(228, 26);
            this.excelElencoToolStripMenuItem.Text = "Excel elenco";
            this.excelElencoToolStripMenuItem.Click += new System.EventHandler(this.excelElencoToolStripMenuItem_Click);
            // 
            // excelElencoCsvToolStripMenuItem
            // 
            this.excelElencoCsvToolStripMenuItem.Name = "excelElencoCsvToolStripMenuItem";
            this.excelElencoCsvToolStripMenuItem.Size = new System.Drawing.Size(228, 26);
            this.excelElencoCsvToolStripMenuItem.Text = "Excel elenco csv";
            this.excelElencoCsvToolStripMenuItem.Click += new System.EventHandler(this.excelElencoCsvToolStripMenuItem_Click);
            // 
            // excelElencoDettaglioToolStripMenuItem
            // 
            this.excelElencoDettaglioToolStripMenuItem.Name = "excelElencoDettaglioToolStripMenuItem";
            this.excelElencoDettaglioToolStripMenuItem.Size = new System.Drawing.Size(228, 26);
            this.excelElencoDettaglioToolStripMenuItem.Text = "Excel elenco dettaglio";
            this.excelElencoDettaglioToolStripMenuItem.Click += new System.EventHandler(this.excelElencoDettaglioToolStripMenuItem_Click);
            // 
            // importFattureDaScontriniToolStripMenuItem
            // 
            this.importFattureDaScontriniToolStripMenuItem.Name = "importFattureDaScontriniToolStripMenuItem";
            this.importFattureDaScontriniToolStripMenuItem.Size = new System.Drawing.Size(181, 25);
            this.importFattureDaScontriniToolStripMenuItem.Text = "Import fatture da cassa";
            this.importFattureDaScontriniToolStripMenuItem.Click += new System.EventHandler(this.importFattureDaScontriniToolStripMenuItem_Click);
            // 
            // fattureElettronicheToolStripMenuItem
            // 
            this.fattureElettronicheToolStripMenuItem.Name = "fattureElettronicheToolStripMenuItem";
            this.fattureElettronicheToolStripMenuItem.Size = new System.Drawing.Size(156, 25);
            this.fattureElettronicheToolStripMenuItem.Text = "Fatture elettroniche";
            this.fattureElettronicheToolStripMenuItem.Click += new System.EventHandler(this.fattureElettronicheToolStripMenuItem_Click);
            // 
            // cmbMovFat
            // 
            this.cmbMovFat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbMovFat.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbMovFat.FormattingEnabled = true;
            this.cmbMovFat.Location = new System.Drawing.Point(195, 36);
            this.cmbMovFat.Name = "cmbMovFat";
            this.cmbMovFat.Size = new System.Drawing.Size(185, 25);
            this.cmbMovFat.TabIndex = 2;
            this.cmbMovFat.SelectionChangeCommitted += new System.EventHandler(this.cmbDocDoc_SelectionChangeCommitted);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(122, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 15);
            this.label1.TabIndex = 3;
            this.label1.Text = "Documento";
            // 
            // lblDocTip
            // 
            this.lblDocTip.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDocTip.Location = new System.Drawing.Point(105, 71);
            this.lblDocTip.Name = "lblDocTip";
            this.lblDocTip.Size = new System.Drawing.Size(87, 18);
            this.lblDocTip.TabIndex = 4;
            this.lblDocTip.Text = "Causale";
            this.lblDocTip.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cmbCauTpd
            // 
            this.cmbCauTpd.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbCauTpd.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbCauTpd.FormattingEnabled = true;
            this.cmbCauTpd.Location = new System.Drawing.Point(195, 67);
            this.cmbCauTpd.Name = "cmbCauTpd";
            this.cmbCauTpd.Size = new System.Drawing.Size(185, 25);
            this.cmbCauTpd.TabIndex = 5;
            this.cmbCauTpd.SelectionChangeCommitted += new System.EventHandler(this.cmbCauTpd_SelectionChangeCommitted);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(12, 41);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(36, 15);
            this.label3.TabIndex = 6;
            this.label3.Text = "Anno";
            // 
            // cmbYea
            // 
            this.cmbYea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbYea.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbYea.FormattingEnabled = true;
            this.cmbYea.Location = new System.Drawing.Point(52, 36);
            this.cmbYea.Name = "cmbYea";
            this.cmbYea.Size = new System.Drawing.Size(65, 25);
            this.cmbYea.TabIndex = 7;
            this.cmbYea.SelectedIndexChanged += new System.EventHandler(this.cmbYea_SelectedIndexChanged);
            // 
            // btnNew
            // 
            this.btnNew.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnNew.Location = new System.Drawing.Point(595, 34);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(245, 28);
            this.btnNew.TabIndex = 8;
            this.btnNew.Text = "Nuovo";
            this.btnNew.UseVisualStyleBackColor = true;
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnSeek
            // 
            this.btnSeek.Location = new System.Drawing.Point(388, 34);
            this.btnSeek.Name = "btnSeek";
            this.btnSeek.Size = new System.Drawing.Size(86, 58);
            this.btnSeek.TabIndex = 9;
            this.btnSeek.Text = "Estrai";
            this.btnSeek.UseVisualStyleBackColor = true;
            this.btnSeek.Click += new System.EventHandler(this.btnSeek_Click);
            // 
            // btnDocRaggruppa
            // 
            this.btnDocRaggruppa.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDocRaggruppa.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDocRaggruppa.Location = new System.Drawing.Point(12, 100);
            this.btnDocRaggruppa.Name = "btnDocRaggruppa";
            this.btnDocRaggruppa.Size = new System.Drawing.Size(828, 32);
            this.btnDocRaggruppa.TabIndex = 10;
            this.btnDocRaggruppa.Text = "Raggruppamento documenti per fattura differita";
            this.btnDocRaggruppa.UseVisualStyleBackColor = true;
            this.btnDocRaggruppa.Click += new System.EventHandler(this.btnDocRaggruppa_Click);
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(650, 526);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 17);
            this.label2.TabIndex = 11;
            this.label2.Text = "Totale";
            // 
            // lblTotVal
            // 
            this.lblTotVal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblTotVal.BackColor = System.Drawing.Color.White;
            this.lblTotVal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblTotVal.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTotVal.Location = new System.Drawing.Point(702, 523);
            this.lblTotVal.Name = "lblTotVal";
            this.lblTotVal.Size = new System.Drawing.Size(138, 24);
            this.lblTotVal.TabIndex = 12;
            this.lblTotVal.Text = "0,00";
            this.lblTotVal.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtSeek
            // 
            this.txtSeek.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSeek.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSeek.Location = new System.Drawing.Point(595, 67);
            this.txtSeek.MaxLength = 20;
            this.txtSeek.Name = "txtSeek";
            this.txtSeek.Size = new System.Drawing.Size(245, 25);
            this.txtSeek.TabIndex = 13;
            this.txtSeek.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSeek_KeyDown);
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(490, 72);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(100, 15);
            this.label4.TabIndex = 14;
            this.label4.Text = "Ricerca per nome";
            // 
            // dgv1
            // 
            this.dgv1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv1.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgv1.Location = new System.Drawing.Point(12, 138);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(828, 376);
            this.dgv1.TabIndex = 1;
            this.dgv1.DoubleClick += new System.EventHandler(this.dgv1_DoubleClick);
            // 
            // frmSeekDocs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(852, 554);
            this.Controls.Add(this.cmbMovFat);
            this.Controls.Add(this.txtSeek);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lblTotVal);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnDocRaggruppa);
            this.Controls.Add(this.btnSeek);
            this.Controls.Add(this.btnNew);
            this.Controls.Add(this.cmbYea);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cmbCauTpd);
            this.Controls.Add(this.lblDocTip);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(840, 520);
            this.Name = "frmSeekDocs";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ricerca documenti";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmSeekDocs_FormClosing);
            this.Load += new System.EventHandler(this.frmGesSeekDocs_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSeekDocs_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.ComboBox cmbMovFat;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblDocTip;
        private System.Windows.Forms.ComboBox cmbCauTpd;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbYea;
        private System.Windows.Forms.Button btnNew;
        private System.Windows.Forms.Button btnSeek;
        private System.Windows.Forms.Button btnDocRaggruppa;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblTotVal;
        private System.Windows.Forms.TextBox txtSeek;
        private System.Windows.Forms.Label label4;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.ToolStripMenuItem lottiToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importFattureDaScontriniToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fattureElettronicheToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem elenchiToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excelDocumentiToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excelElencoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excelElencoCsvToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excelElencoDettaglioToolStripMenuItem;
    }
}
