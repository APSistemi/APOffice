namespace APOffice
{
    partial class frmUtyGenCosti
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.estraiToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.generaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.panelHeader = new System.Windows.Forms.Panel();
            this.nudPercScorporo = new System.Windows.Forms.NumericUpDown();
            this.lblPercScorporo = new System.Windows.Forms.Label();
            this.btnEsci = new System.Windows.Forms.Button();
            this.btnGenera = new System.Windows.Forms.Button();
            this.btnEstrai = new System.Windows.Forms.Button();
            this.chkAggiornaAna = new System.Windows.Forms.CheckBox();
            this.chkSoloAttivi = new System.Windows.Forms.CheckBox();
            this.cmbFornitore = new System.Windows.Forms.ComboBox();
            this.lblFornitore = new System.Windows.Forms.Label();
            this.dtpData = new System.Windows.Forms.DateTimePicker();
            this.lblData = new System.Windows.Forms.Label();
            this.dgv1 = new APOffice.APDataGridView();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblStatus = new System.Windows.Forms.Label();
            this.lblRiepilogo = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            this.panelHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudPercScorporo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(244)))), ((int)(((byte)(246)))));
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.estraiToolStripMenuItem,
            this.generaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(964, 24);
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
            // estraiToolStripMenuItem
            // 
            this.estraiToolStripMenuItem.Name = "estraiToolStripMenuItem";
            this.estraiToolStripMenuItem.Size = new System.Drawing.Size(102, 20);
            this.estraiToolStripMenuItem.Text = "Estrai Anteprima";
            this.estraiToolStripMenuItem.Click += new System.EventHandler(this.btnEstrai_Click);
            // 
            // generaToolStripMenuItem
            // 
            this.generaToolStripMenuItem.Name = "generaToolStripMenuItem";
            this.generaToolStripMenuItem.Size = new System.Drawing.Size(87, 20);
            this.generaToolStripMenuItem.Text = "Genera Costi";
            this.generaToolStripMenuItem.Click += new System.EventHandler(this.btnGenera_Click);
            // 
            // panelHeader
            // 
            this.panelHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelHeader.BackColor = System.Drawing.Color.White;
            this.panelHeader.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelHeader.Controls.Add(this.nudPercScorporo);
            this.panelHeader.Controls.Add(this.lblPercScorporo);
            this.panelHeader.Controls.Add(this.btnEsci);
            this.panelHeader.Controls.Add(this.btnGenera);
            this.panelHeader.Controls.Add(this.btnEstrai);
            this.panelHeader.Controls.Add(this.chkAggiornaAna);
            this.panelHeader.Controls.Add(this.chkSoloAttivi);
            this.panelHeader.Controls.Add(this.cmbFornitore);
            this.panelHeader.Controls.Add(this.lblFornitore);
            this.panelHeader.Controls.Add(this.dtpData);
            this.panelHeader.Controls.Add(this.lblData);
            this.panelHeader.Location = new System.Drawing.Point(12, 32);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(940, 96);
            this.panelHeader.TabIndex = 1;
            // 
            // nudPercScorporo
            // 
            this.nudPercScorporo.DecimalPlaces = 2;
            this.nudPercScorporo.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold);
            this.nudPercScorporo.Location = new System.Drawing.Point(520, 23);
            this.nudPercScorporo.Maximum = new decimal(new int[] {
            100,
            0,
            0,
            0});
            this.nudPercScorporo.Name = "nudPercScorporo";
            this.nudPercScorporo.Size = new System.Drawing.Size(110, 25);
            this.nudPercScorporo.TabIndex = 5;
            this.nudPercScorporo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.nudPercScorporo.ValueChanged += new System.EventHandler(this.nudPercScorporo_ValueChanged);
            // 
            // lblPercScorporo
            // 
            this.lblPercScorporo.AutoSize = true;
            this.lblPercScorporo.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold);
            this.lblPercScorporo.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.lblPercScorporo.Location = new System.Drawing.Point(517, 7);
            this.lblPercScorporo.Name = "lblPercScorporo";
            this.lblPercScorporo.Size = new System.Drawing.Size(120, 13);
            this.lblPercScorporo.TabIndex = 4;
            this.lblPercScorporo.Text = "Coeff. / % Ricarico:";
            // 
            // btnEsci
            // 
            this.btnEsci.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnEsci.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnEsci.Location = new System.Drawing.Point(825, 50);
            this.btnEsci.Name = "btnEsci";
            this.btnEsci.Size = new System.Drawing.Size(100, 34);
            this.btnEsci.TabIndex = 10;
            this.btnEsci.Text = "Esci";
            this.btnEsci.UseVisualStyleBackColor = true;
            this.btnEsci.Click += new System.EventHandler(this.btnEsci_Click);
            // 
            // btnGenera
            // 
            this.btnGenera.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnGenera.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnGenera.Location = new System.Drawing.Point(665, 50);
            this.btnGenera.Name = "btnGenera";
            this.btnGenera.Size = new System.Drawing.Size(154, 34);
            this.btnGenera.TabIndex = 9;
            this.btnGenera.Text = "Genera Costi";
            this.btnGenera.UseVisualStyleBackColor = true;
            this.btnGenera.Click += new System.EventHandler(this.btnGenera_Click);
            // 
            // btnEstrai
            // 
            this.btnEstrai.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnEstrai.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnEstrai.Location = new System.Drawing.Point(665, 10);
            this.btnEstrai.Name = "btnEstrai";
            this.btnEstrai.Size = new System.Drawing.Size(260, 34);
            this.btnEstrai.TabIndex = 8;
            this.btnEstrai.Text = "Estrai Anteprima";
            this.btnEstrai.UseVisualStyleBackColor = true;
            this.btnEstrai.Click += new System.EventHandler(this.btnEstrai_Click);
            // 
            // chkAggiornaAna
            // 
            this.chkAggiornaAna.AutoSize = true;
            this.chkAggiornaAna.Checked = true;
            this.chkAggiornaAna.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkAggiornaAna.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.chkAggiornaAna.Location = new System.Drawing.Point(315, 60);
            this.chkAggiornaAna.Name = "chkAggiornaAna";
            this.chkAggiornaAna.Size = new System.Drawing.Size(262, 19);
            this.chkAggiornaAna.TabIndex = 7;
            this.chkAggiornaAna.Text = "Aggiorna costo ultimo anagrafica (AnaArticoli)";
            this.chkAggiornaAna.UseVisualStyleBackColor = true;
            // 
            // chkSoloAttivi
            // 
            this.chkSoloAttivi.AutoSize = true;
            this.chkSoloAttivi.Checked = true;
            this.chkSoloAttivi.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSoloAttivi.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.chkSoloAttivi.Location = new System.Drawing.Point(120, 60);
            this.chkSoloAttivi.Name = "chkSoloAttivi";
            this.chkSoloAttivi.Size = new System.Drawing.Size(185, 19);
            this.chkSoloAttivi.TabIndex = 6;
            this.chkSoloAttivi.Text = "Solo articoli attivi (Stato \'A\')";
            this.chkSoloAttivi.UseVisualStyleBackColor = true;
            this.chkSoloAttivi.CheckedChanged += new System.EventHandler(this.chkSoloAttivi_CheckedChanged);
            // 
            // cmbFornitore
            // 
            this.cmbFornitore.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cmbFornitore.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbFornitore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFornitore.Font = new System.Drawing.Font("Segoe UI", 9.75F);
            this.cmbFornitore.FormattingEnabled = true;
            this.cmbFornitore.Location = new System.Drawing.Point(120, 23);
            this.cmbFornitore.Name = "cmbFornitore";
            this.cmbFornitore.Size = new System.Drawing.Size(385, 25);
            this.cmbFornitore.TabIndex = 3;
            // 
            // lblFornitore
            // 
            this.lblFornitore.AutoSize = true;
            this.lblFornitore.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold);
            this.lblFornitore.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.lblFornitore.Location = new System.Drawing.Point(117, 7);
            this.lblFornitore.Name = "lblFornitore";
            this.lblFornitore.Size = new System.Drawing.Size(126, 13);
            this.lblFornitore.TabIndex = 2;
            this.lblFornitore.Text = "Fornitore Destinazione:";
            // 
            // dtpData
            // 
            this.dtpData.CustomFormat = "dd/MM/yyyy";
            this.dtpData.Font = new System.Drawing.Font("Segoe UI", 9.75F);
            this.dtpData.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpData.Location = new System.Drawing.Point(10, 23);
            this.dtpData.Name = "dtpData";
            this.dtpData.Size = new System.Drawing.Size(98, 25);
            this.dtpData.TabIndex = 1;
            this.dtpData.ValueChanged += new System.EventHandler(this.dtpData_ValueChanged);
            // 
            // lblData
            // 
            this.lblData.AutoSize = true;
            this.lblData.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold);
            this.lblData.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.lblData.Location = new System.Drawing.Point(7, 7);
            this.lblData.Name = "lblData";
            this.lblData.Size = new System.Drawing.Size(107, 13);
            this.lblData.TabIndex = 0;
            this.lblData.Text = "Data Registrazione:";
            // 
            // dgv1
            // 
            this.dgv1.AllowUserToAddRows = false;
            this.dgv1.AllowUserToDeleteRows = false;
            this.dgv1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 134);
            this.dgv1.Name = "dgv1";
            this.dgv1.ReadOnly = true;
            this.dgv1.RowHeadersVisible = false;
            this.dgv1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv1.Size = new System.Drawing.Size(940, 440);
            this.dgv1.TabIndex = 2;
            this.dgv1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellClick);
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(12, 582);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(940, 16);
            this.progressBar1.TabIndex = 3;
            // 
            // lblStatus
            // 
            this.lblStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblStatus.AutoSize = true;
            this.lblStatus.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStatus.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(71)))), ((int)(((byte)(85)))), ((int)(((byte)(105)))));
            this.lblStatus.Location = new System.Drawing.Point(12, 604);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(46, 15);
            this.lblStatus.TabIndex = 4;
            this.lblStatus.Text = "Pronto.";
            // 
            // lblRiepilogo
            // 
            this.lblRiepilogo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblRiepilogo.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRiepilogo.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(41)))), ((int)(((byte)(59)))));
            this.lblRiepilogo.Location = new System.Drawing.Point(652, 604);
            this.lblRiepilogo.Name = "lblRiepilogo";
            this.lblRiepilogo.Size = new System.Drawing.Size(300, 15);
            this.lblRiepilogo.TabIndex = 5;
            this.lblRiepilogo.Text = "Articoli estratti: 0";
            this.lblRiepilogo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // frmUtyGenCosti
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(244)))), ((int)(((byte)(246)))));
            this.ClientSize = new System.Drawing.Size(964, 628);
            this.Controls.Add(this.lblRiepilogo);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.panelHeader);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(800, 500);
            this.Name = "frmUtyGenCosti";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Generazione Costi Acquisto da Prezzo di Vendita";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmUtyGenCosti_FormClosing);
            this.Load += new System.EventHandler(this.frmUtyGenCosti_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyGenCosti_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.panelHeader.ResumeLayout(false);
            this.panelHeader.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudPercScorporo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem estraiToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem generaToolStripMenuItem;
        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.Label lblData;
        private System.Windows.Forms.DateTimePicker dtpData;
        private System.Windows.Forms.Label lblFornitore;
        private System.Windows.Forms.ComboBox cmbFornitore;
        private System.Windows.Forms.Label lblPercScorporo;
        private System.Windows.Forms.NumericUpDown nudPercScorporo;
        private System.Windows.Forms.CheckBox chkSoloAttivi;
        private System.Windows.Forms.CheckBox chkAggiornaAna;
        private System.Windows.Forms.Button btnEstrai;
        private System.Windows.Forms.Button btnGenera;
        private System.Windows.Forms.Button btnEsci;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label lblRiepilogo;
    }
}
