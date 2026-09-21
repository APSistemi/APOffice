namespace APOffice
{
    partial class frmGesVariazioni
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.invioVariazioniACasseEBilanceToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.stampaEtichetteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.utilityToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.etichetteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.annullaTutteLeEtichetteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.attivaTutteLeEtichetteToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.recuperoEtichetteToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.designerEtichetteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importTerminalinoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.casseEBilanceToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.annullaTutteLeVariazioniToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.attivaTutteLeVariazioniToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.recuperoVariazioniToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importTerminalinoToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.importTerminalinoADVToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.terminalinoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.impiantoCassaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.impiantoEtichetteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.bilanciaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.svuotaElencoVariazioniToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.promoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.refreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.dgv1 = new APOffice.APDataGridView();
            this.label3 = new System.Windows.Forms.Label();
            this.dgv2 = new APOffice.APDataGridView();
            this.dgv3 = new APOffice.APDataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.lblVarCnt = new System.Windows.Forms.Label();
            this.lblEtiCnt = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.cmbEtiRep = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.nudEtiRow = new System.Windows.Forms.NumericUpDown();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudEtiRow)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.invioVariazioniACasseEBilanceToolStripMenuItem,
            this.stampaEtichetteToolStripMenuItem,
            this.utilityToolStripMenuItem,
            this.promoToolStripMenuItem,
            this.refreshToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(873, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            this.menuStrip1.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.menuStrip1_ItemClicked);
            // 
            // esciToolStripMenuItem
            // 
            this.esciToolStripMenuItem.Name = "esciToolStripMenuItem";
            this.esciToolStripMenuItem.Size = new System.Drawing.Size(39, 20);
            this.esciToolStripMenuItem.Text = "Esci";
            this.esciToolStripMenuItem.Click += new System.EventHandler(this.esciToolStripMenuItem_Click);
            // 
            // invioVariazioniACasseEBilanceToolStripMenuItem
            // 
            this.invioVariazioniACasseEBilanceToolStripMenuItem.Name = "invioVariazioniACasseEBilanceToolStripMenuItem";
            this.invioVariazioniACasseEBilanceToolStripMenuItem.Size = new System.Drawing.Size(188, 20);
            this.invioVariazioniACasseEBilanceToolStripMenuItem.Text = "Invio variazioni a casse e bilance";
            this.invioVariazioniACasseEBilanceToolStripMenuItem.Click += new System.EventHandler(this.invioVariazioniACasseEBilanceToolStripMenuItem_Click);
            // 
            // stampaEtichetteToolStripMenuItem
            // 
            this.stampaEtichetteToolStripMenuItem.Name = "stampaEtichetteToolStripMenuItem";
            this.stampaEtichetteToolStripMenuItem.Size = new System.Drawing.Size(108, 20);
            this.stampaEtichetteToolStripMenuItem.Text = "Stampa etichette";
            this.stampaEtichetteToolStripMenuItem.Click += new System.EventHandler(this.stampaEtichetteToolStripMenuItem_Click);
            // 
            // utilityToolStripMenuItem
            // 
            this.utilityToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.utilityToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.etichetteToolStripMenuItem,
            this.casseEBilanceToolStripMenuItem,
            this.terminalinoToolStripMenuItem,
            this.impiantoCassaToolStripMenuItem,
            this.impiantoEtichetteToolStripMenuItem,
            this.bilanciaToolStripMenuItem,
            this.svuotaElencoVariazioniToolStripMenuItem});
            this.utilityToolStripMenuItem.Name = "utilityToolStripMenuItem";
            this.utilityToolStripMenuItem.Size = new System.Drawing.Size(50, 20);
            this.utilityToolStripMenuItem.Text = "Utility";
            // 
            // etichetteToolStripMenuItem
            // 
            this.etichetteToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.annullaTutteLeEtichetteToolStripMenuItem,
            this.attivaTutteLeEtichetteToolStripMenuItem1,
            this.recuperoEtichetteToolStripMenuItem1,
            this.importTerminalinoToolStripMenuItem,
            this.impiantoEtichetteToolStripMenuItem,
            this.designerEtichetteToolStripMenuItem});
            this.etichetteToolStripMenuItem.Name = "etichetteToolStripMenuItem";
            this.etichetteToolStripMenuItem.Size = new System.Drawing.Size(186, 22);
            this.etichetteToolStripMenuItem.Text = "Etichette";
            // 
            // annullaTutteLeEtichetteToolStripMenuItem
            // 
            this.annullaTutteLeEtichetteToolStripMenuItem.Name = "annullaTutteLeEtichetteToolStripMenuItem";
            this.annullaTutteLeEtichetteToolStripMenuItem.Size = new System.Drawing.Size(204, 22);
            this.annullaTutteLeEtichetteToolStripMenuItem.Text = "Annulla tutte le etichette";
            this.annullaTutteLeEtichetteToolStripMenuItem.Click += new System.EventHandler(this.annullaTutteLeEtichetteToolStripMenuItem_Click);
            // 
            // attivaTutteLeEtichetteToolStripMenuItem1
            // 
            this.attivaTutteLeEtichetteToolStripMenuItem1.Name = "attivaTutteLeEtichetteToolStripMenuItem1";
            this.attivaTutteLeEtichetteToolStripMenuItem1.Size = new System.Drawing.Size(204, 22);
            this.attivaTutteLeEtichetteToolStripMenuItem1.Text = "Attiva tutte le etichette";
            this.attivaTutteLeEtichetteToolStripMenuItem1.Click += new System.EventHandler(this.attivaTutteLeEtichetteToolStripMenuItem1_Click);
            // 
            // recuperoEtichetteToolStripMenuItem1
            // 
            this.recuperoEtichetteToolStripMenuItem1.Name = "recuperoEtichetteToolStripMenuItem1";
            this.recuperoEtichetteToolStripMenuItem1.Size = new System.Drawing.Size(204, 22);
            this.recuperoEtichetteToolStripMenuItem1.Text = "Recupero etichette";
            this.recuperoEtichetteToolStripMenuItem1.Click += new System.EventHandler(this.recuperoEtichetteToolStripMenuItem1_Click);
            // 
            // importTerminalinoToolStripMenuItem
            // 
            this.importTerminalinoToolStripMenuItem.Name = "importTerminalinoToolStripMenuItem";
            this.importTerminalinoToolStripMenuItem.Size = new System.Drawing.Size(204, 22);
            this.importTerminalinoToolStripMenuItem.Text = "Import terminalino";
            this.importTerminalinoToolStripMenuItem.Click += new System.EventHandler(this.importTerminalinoToolStripMenuItem_Click);
            // 
            // designerEtichetteToolStripMenuItem
            // 
            this.designerEtichetteToolStripMenuItem.Name = "designerEtichetteToolStripMenuItem";
            this.designerEtichetteToolStripMenuItem.Size = new System.Drawing.Size(303, 22);
            this.designerEtichetteToolStripMenuItem.Text = "Crea/Modifica Formati Etichette (Designer)";
            this.designerEtichetteToolStripMenuItem.Click += new System.EventHandler(this.designerEtichetteToolStripMenuItem_Click);
            // 
            // casseEBilanceToolStripMenuItem
            // 
            this.casseEBilanceToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.annullaTutteLeVariazioniToolStripMenuItem,
            this.attivaTutteLeVariazioniToolStripMenuItem,
            this.recuperoVariazioniToolStripMenuItem,
            this.importTerminalinoToolStripMenuItem1,
            this.importTerminalinoADVToolStripMenuItem});
            this.casseEBilanceToolStripMenuItem.Name = "casseEBilanceToolStripMenuItem";
            this.casseEBilanceToolStripMenuItem.Size = new System.Drawing.Size(186, 22);
            this.casseEBilanceToolStripMenuItem.Text = "Casse e bilance";
            // 
            // annullaTutteLeVariazioniToolStripMenuItem
            // 
            this.annullaTutteLeVariazioniToolStripMenuItem.Name = "annullaTutteLeVariazioniToolStripMenuItem";
            this.annullaTutteLeVariazioniToolStripMenuItem.Size = new System.Drawing.Size(303, 22);
            this.annullaTutteLeVariazioniToolStripMenuItem.Text = "Annulla tutte le variazioni";
            this.annullaTutteLeVariazioniToolStripMenuItem.Click += new System.EventHandler(this.annullaTutteLeVariazioniToolStripMenuItem_Click);
            // 
            // attivaTutteLeVariazioniToolStripMenuItem
            // 
            this.attivaTutteLeVariazioniToolStripMenuItem.Name = "attivaTutteLeVariazioniToolStripMenuItem";
            this.attivaTutteLeVariazioniToolStripMenuItem.Size = new System.Drawing.Size(303, 22);
            this.attivaTutteLeVariazioniToolStripMenuItem.Text = "Attiva tutte le variazioni";
            this.attivaTutteLeVariazioniToolStripMenuItem.Click += new System.EventHandler(this.attivaTutteLeVariazioniToolStripMenuItem_Click);
            // 
            // recuperoVariazioniToolStripMenuItem
            // 
            this.recuperoVariazioniToolStripMenuItem.Name = "recuperoVariazioniToolStripMenuItem";
            this.recuperoVariazioniToolStripMenuItem.Size = new System.Drawing.Size(303, 22);
            this.recuperoVariazioniToolStripMenuItem.Text = "Recupero variazioni";
            this.recuperoVariazioniToolStripMenuItem.Click += new System.EventHandler(this.recuperoVariazioniToolStripMenuItem_Click);
            // 
            // importTerminalinoToolStripMenuItem1
            // 
            this.importTerminalinoToolStripMenuItem1.Name = "importTerminalinoToolStripMenuItem1";
            this.importTerminalinoToolStripMenuItem1.Size = new System.Drawing.Size(303, 22);
            this.importTerminalinoToolStripMenuItem1.Text = "Import terminalino + attivazione e etichette";
            this.importTerminalinoToolStripMenuItem1.Click += new System.EventHandler(this.importTerminalinoToolStripMenuItem1_Click);
            // 
            // importTerminalinoADVToolStripMenuItem
            // 
            this.importTerminalinoADVToolStripMenuItem.Name = "importTerminalinoADVToolStripMenuItem";
            this.importTerminalinoADVToolStripMenuItem.Size = new System.Drawing.Size(303, 22);
            this.importTerminalinoADVToolStripMenuItem.Text = "Import terminalino Variazioni Prezzi (ADV)";
            this.importTerminalinoADVToolStripMenuItem.Click += new System.EventHandler(this.importTerminalinoADVToolStripMenuItem_Click);
            // 
            // terminalinoToolStripMenuItem
            // 
            this.terminalinoToolStripMenuItem.Name = "terminalinoToolStripMenuItem";
            this.terminalinoToolStripMenuItem.Size = new System.Drawing.Size(186, 22);
            this.terminalinoToolStripMenuItem.Text = "Impianto terminalino";
            this.terminalinoToolStripMenuItem.Click += new System.EventHandler(this.terminalinoToolStripMenuItem_Click);
            // 
            // impiantoCassaToolStripMenuItem
            // 
            this.impiantoCassaToolStripMenuItem.Name = "impiantoCassaToolStripMenuItem";
            this.impiantoCassaToolStripMenuItem.Size = new System.Drawing.Size(186, 22);
            this.impiantoCassaToolStripMenuItem.Text = "Impianto a cassa";
            this.impiantoCassaToolStripMenuItem.Click += new System.EventHandler(this.impiantoCassaToolStripMenuItem_Click);
            // 
            // impiantoEtichetteToolStripMenuItem
            // 
            this.impiantoEtichetteToolStripMenuItem.Name = "impiantoEtichetteToolStripMenuItem";
            this.impiantoEtichetteToolStripMenuItem.Size = new System.Drawing.Size(186, 22);
            this.impiantoEtichetteToolStripMenuItem.Text = "Impianto etichette";
            this.impiantoEtichetteToolStripMenuItem.Click += new System.EventHandler(this.impiantoEtichetteToolStripMenuItem_Click);
            // 
            // bilanciaToolStripMenuItem
            // 
            this.bilanciaToolStripMenuItem.Name = "bilanciaToolStripMenuItem";
            this.bilanciaToolStripMenuItem.Size = new System.Drawing.Size(186, 22);
            this.bilanciaToolStripMenuItem.Text = "Impianto a bilance";
            this.bilanciaToolStripMenuItem.Click += new System.EventHandler(this.bilanciaToolStripMenuItem_Click);
            // 
            // svuotaElencoVariazioniToolStripMenuItem
            // 
            this.svuotaElencoVariazioniToolStripMenuItem.Name = "svuotaElencoVariazioniToolStripMenuItem";
            this.svuotaElencoVariazioniToolStripMenuItem.Size = new System.Drawing.Size(204, 22);
            this.svuotaElencoVariazioniToolStripMenuItem.Text = "Svuota Elenco Variazioni";
            this.svuotaElencoVariazioniToolStripMenuItem.Click += new System.EventHandler(this.svuotaElencoVariazioniToolStripMenuItem_Click);
            // 
            // promoToolStripMenuItem
            // 
            this.promoToolStripMenuItem.Name = "promoToolStripMenuItem";
            this.promoToolStripMenuItem.Size = new System.Drawing.Size(55, 20);
            this.promoToolStripMenuItem.Text = "Promo";
            this.promoToolStripMenuItem.Click += new System.EventHandler(this.promoToolStripMenuItem_Click);
            // 
            // refreshToolStripMenuItem
            // 
            this.refreshToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.refreshToolStripMenuItem.Name = "refreshToolStripMenuItem";
            this.refreshToolStripMenuItem.Size = new System.Drawing.Size(58, 20);
            this.refreshToolStripMenuItem.Text = "Refresh";
            this.refreshToolStripMenuItem.Click += new System.EventHandler(this.refreshToolStripMenuItem_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Casse e bilance";
            // 
            // dgv1
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Lavender;
            this.dgv1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(6, 48);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(861, 273);
            this.dgv1.TabIndex = 5;
            this.dgv1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellClick);
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            this.dgv1.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgv1_CellFormatting);
            this.dgv1.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgv_CurrentCellDirtyStateChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(424, 327);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(49, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Etichette";
            // 
            // dgv2
            // 
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.Lavender;
            this.dgv2.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv2.Location = new System.Drawing.Point(6, 342);
            this.dgv2.Name = "dgv2";
            this.dgv2.RowHeadersWidth = 20;
            this.dgv2.Size = new System.Drawing.Size(416, 276);
            this.dgv2.TabIndex = 8;
            this.dgv2.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv2_CellClick);
            this.dgv2.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv2_CellDoubleClick);
            // 
            // dgv3
            // 
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.Lavender;
            this.dgv3.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle3;
            this.dgv3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv3.Location = new System.Drawing.Point(426, 342);
            this.dgv3.Name = "dgv3";
            this.dgv3.RowHeadersWidth = 20;
            this.dgv3.Size = new System.Drawing.Size(442, 276);
            this.dgv3.TabIndex = 10;
            this.dgv3.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv3_CellClick);
            this.dgv3.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv3_CellDoubleClick);
            this.dgv3.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgv_CurrentCellDirtyStateChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 327);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(39, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "Offerte";
            // 
            // lblVarCnt
            // 
            this.lblVarCnt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblVarCnt.Location = new System.Drawing.Point(811, 26);
            this.lblVarCnt.Name = "lblVarCnt";
            this.lblVarCnt.Size = new System.Drawing.Size(54, 18);
            this.lblVarCnt.TabIndex = 12;
            this.lblVarCnt.Text = "0";
            this.lblVarCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblEtiCnt
            // 
            this.lblEtiCnt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblEtiCnt.Location = new System.Drawing.Point(811, 323);
            this.lblEtiCnt.Name = "lblEtiCnt";
            this.lblEtiCnt.Size = new System.Drawing.Size(53, 18);
            this.lblEtiCnt.TabIndex = 13;
            this.lblEtiCnt.Text = "0";
            this.lblEtiCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 620);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(873, 10);
            this.progressBar1.TabIndex = 14;
            // 
            // cmbEtiRep
            // 
            this.cmbEtiRep.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbEtiRep.FormattingEnabled = true;
            this.cmbEtiRep.Location = new System.Drawing.Point(528, 321);
            this.cmbEtiRep.Name = "cmbEtiRep";
            this.cmbEtiRep.Size = new System.Drawing.Size(121, 21);
            this.cmbEtiRep.TabIndex = 15;
            this.cmbEtiRep.SelectionChangeCommitted += new System.EventHandler(this.cmbEtiRep_SelectionChangeCommitted);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(481, 327);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(45, 13);
            this.label4.TabIndex = 16;
            this.label4.Text = "Reparto";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(669, 326);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(55, 13);
            this.label5.TabIndex = 17;
            this.label5.Text = "Riga inizio";
            // 
            // nudEtiRow
            // 
            this.nudEtiRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nudEtiRow.Location = new System.Drawing.Point(731, 323);
            this.nudEtiRow.Maximum = new decimal(new int[] {
            12,
            0,
            0,
            0});
            this.nudEtiRow.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudEtiRow.Name = "nudEtiRow";
            this.nudEtiRow.Size = new System.Drawing.Size(34, 20);
            this.nudEtiRow.TabIndex = 18;
            this.nudEtiRow.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.nudEtiRow.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // frmGesVariazioni
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(873, 631);
            this.ControlBox = false;
            this.Controls.Add(this.nudEtiRow);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cmbEtiRep);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.lblVarCnt);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dgv3);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dgv2);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.lblEtiCnt);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesVariazioni";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Gestione invio variazioni/stampa etichette";
            this.Load += new System.EventHandler(this.frmGesVariazioni_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesVariazioni_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudEtiRow)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Label label1;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Label label3;
        private APOffice.APDataGridView dgv2;
        private System.Windows.Forms.ToolStripMenuItem stampaEtichetteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem invioVariazioniACasseEBilanceToolStripMenuItem;
        private APOffice.APDataGridView dgv3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblVarCnt;
        private System.Windows.Forms.Label lblEtiCnt;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ToolStripMenuItem utilityToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem etichetteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem annullaTutteLeEtichetteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem attivaTutteLeEtichetteToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem recuperoEtichetteToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem casseEBilanceToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem annullaTutteLeVariazioniToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem attivaTutteLeVariazioniToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem recuperoVariazioniToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem terminalinoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem impiantoCassaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem impiantoEtichetteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem bilanciaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importTerminalinoToolStripMenuItem;
        private System.Windows.Forms.ComboBox cmbEtiRep;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.NumericUpDown nudEtiRow;
        private System.Windows.Forms.ToolStripMenuItem refreshToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importTerminalinoToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem importTerminalinoADVToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem promoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem designerEtichetteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem svuotaElencoVariazioniToolStripMenuItem;
    }
}
