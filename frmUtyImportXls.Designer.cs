namespace APOffice
{
    partial class frmUtyImportXls
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.pnlTop = new System.Windows.Forms.Panel();
            this.lblStepInfo = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.pnlBottom = new System.Windows.Forms.Panel();
            this.btnExit = new System.Windows.Forms.Button();
            this.btnBack = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.tabWizard = new System.Windows.Forms.TabControl();
            this.tabStep1 = new System.Windows.Forms.TabPage();
            this.label3 = new System.Windows.Forms.Label();
            this.lblFileStatus = new System.Windows.Forms.Label();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.tabStep2 = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.cmbMapping = new System.Windows.Forms.ComboBox();
            this.btnNewMapping = new System.Windows.Forms.Button();
            this.tabStep3 = new System.Windows.Forms.TabPage();
            this.dgvImport = new APOffice.APDataGridView();
            this.tabStep4 = new System.Windows.Forms.TabPage();
            this.btnAcquire = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblSummary = new System.Windows.Forms.Label();
            this.chkUpdateStorePrice = new System.Windows.Forms.CheckBox();
            this.chkInsertNew = new System.Windows.Forms.CheckBox();
            this.btnEditMapping = new System.Windows.Forms.Button();
            this.chkSendToCasse = new System.Windows.Forms.CheckBox();
            this.chkSendToStampa = new System.Windows.Forms.CheckBox();
            this.lblFileSelected = new System.Windows.Forms.Label();
            this.chkSetActiveStatus = new System.Windows.Forms.CheckBox();
            this.lblLegend = new System.Windows.Forms.Label();
            this.pnlTop.SuspendLayout();
            this.pnlBottom.SuspendLayout();
            this.tabWizard.SuspendLayout();
            this.tabStep1.SuspendLayout();
            this.tabStep2.SuspendLayout();
            this.tabStep3.SuspendLayout();
            this.tabStep4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvImport)).BeginInit();
            this.SuspendLayout();
            // 
            // pnlTop
            // 
            this.pnlTop.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(52)))), ((int)(((byte)(73)))), ((int)(((byte)(94)))));
            this.pnlTop.Controls.Add(this.lblStepInfo);
            this.pnlTop.Controls.Add(this.lblTitle);
            this.pnlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlTop.Location = new System.Drawing.Point(0, 0);
            this.pnlTop.Name = "pnlTop";
            this.pnlTop.Size = new System.Drawing.Size(900, 60);
            this.pnlTop.TabIndex = 9;
            // 
            // lblStepInfo
            // 
            this.lblStepInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblStepInfo.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStepInfo.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(195)))), ((int)(((byte)(199)))));
            this.lblStepInfo.Location = new System.Drawing.Point(688, 19);
            this.lblStepInfo.Name = "lblStepInfo";
            this.lblStepInfo.Size = new System.Drawing.Size(200, 23);
            this.lblStepInfo.TabIndex = 1;
            this.lblStepInfo.Text = "Passaggio 1 di 4";
            this.lblStepInfo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(44)))), ((int)(((byte)(62)))), ((int)(((byte)(80)))));
            this.lblTitle.Location = new System.Drawing.Point(12, 13);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(262, 30);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "Importazione Dati Excel";
            this.pnlBottom.Controls.Add(this.progressBar1);
            this.pnlBottom.Controls.Add(this.btnExit);
            this.pnlBottom.Controls.Add(this.btnBack);
            this.pnlBottom.Controls.Add(this.btnNext);
            this.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnlBottom.Location = new System.Drawing.Point(0, 515);
            this.pnlBottom.Name = "pnlBottom";
            this.pnlBottom.Size = new System.Drawing.Size(900, 60);
            this.pnlBottom.TabIndex = 10;
            // 
            // btnExit
            // 
            this.btnExit.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExit.Location = new System.Drawing.Point(12, 22);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(100, 30);
            this.btnExit.TabIndex = 2;
            this.btnExit.Text = "ESCI";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.esciToolStripMenuItem_Click);
            // 
            // btnBack
            // 
            this.btnBack.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBack.Enabled = false;
            this.btnBack.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBack.Location = new System.Drawing.Point(682, 22);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(100, 30);
            this.btnBack.TabIndex = 1;
            this.btnBack.Text = "< INDIETRO";
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // btnNext
            // 
            this.btnNext.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnNext.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNext.Location = new System.Drawing.Point(788, 22);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(100, 30);
            this.btnNext.TabIndex = 0;
            this.btnNext.Text = "AVANTI >";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // tabWizard
            // 
            this.tabWizard.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.tabWizard.Controls.Add(this.tabStep1);
            this.tabWizard.Controls.Add(this.tabStep2);
            this.tabWizard.Controls.Add(this.tabStep3);
            this.tabWizard.Controls.Add(this.tabStep4);
            this.tabWizard.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabWizard.ItemSize = new System.Drawing.Size(0, 1);
            this.tabWizard.Location = new System.Drawing.Point(0, 60);
            this.tabWizard.Name = "tabWizard";
            this.tabWizard.SelectedIndex = 0;
            this.tabWizard.Size = new System.Drawing.Size(900, 455);
            this.tabWizard.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tabWizard.TabIndex = 11;
            // 
            // tabStep1
            // 
            this.tabStep1.BackColor = System.Drawing.Color.White;
            this.tabStep1.Controls.Add(this.label3);
            this.tabStep1.Controls.Add(this.lblFileStatus);
            this.tabStep1.Controls.Add(this.btnSelectFile);
            this.tabStep1.Controls.Add(this.txtFilePath);
            this.tabStep1.Location = new System.Drawing.Point(4, 5);
            this.tabStep1.Name = "tabStep1";
            this.tabStep1.Size = new System.Drawing.Size(892, 456);
            this.tabStep1.TabIndex = 0;
            this.tabStep1.Text = "File";
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.DimGray;
            this.label3.Location = new System.Drawing.Point(200, 100);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(500, 40);
            this.label3.TabIndex = 4;
            this.label3.Text = "Per iniziare, seleziona il file Excel (.xls, .xlsx) o CSV che desideri importare.";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblFileStatus
            // 
            this.lblFileStatus.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFileStatus.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(128)))), ((int)(((byte)(185)))));
            this.lblFileStatus.Location = new System.Drawing.Point(200, 240);
            this.lblFileStatus.Name = "lblFileStatus";
            this.lblFileStatus.Size = new System.Drawing.Size(500, 30);
            this.lblFileStatus.TabIndex = 3;
            this.lblFileStatus.Text = "Nessun file selezionato";
            this.lblFileStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(52)))), ((int)(((byte)(152)))), ((int)(((byte)(219)))));
            this.btnSelectFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSelectFile.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSelectFile.ForeColor = System.Drawing.Color.White;
            this.btnSelectFile.Location = new System.Drawing.Point(300, 150);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(300, 60);
            this.btnSelectFile.TabIndex = 1;
            this.btnSelectFile.Text = "CARICA FILE EXCEL / CSV";
            this.btnSelectFile.UseVisualStyleBackColor = false;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // txtFilePath
            // 
            this.txtFilePath.BackColor = System.Drawing.Color.White;
            this.txtFilePath.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtFilePath.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFilePath.Location = new System.Drawing.Point(200, 215);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.Size = new System.Drawing.Size(500, 16);
            this.txtFilePath.TabIndex = 2;
            this.txtFilePath.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // lblFileSelected
            // 
            this.lblFileSelected.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFileSelected.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(128)))), ((int)(((byte)(185)))));
            this.lblFileSelected.Location = new System.Drawing.Point(10, 10);
            this.lblFileSelected.Name = "lblFileSelected";
            this.lblFileSelected.Size = new System.Drawing.Size(872, 23);
            this.lblFileSelected.TabIndex = 11;
            this.lblFileSelected.Text = "File in uso: -";
            this.lblFileSelected.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // tabStep2
            // 
            this.tabStep2.Controls.Add(this.lblFileSelected);
            this.tabStep2.Controls.Add(this.btnEditMapping);
            this.tabStep2.Controls.Add(this.btnNewMapping);
            this.tabStep2.Controls.Add(this.label2);
            this.tabStep2.Controls.Add(this.label1);
            this.tabStep2.Controls.Add(this.cmbMapping);
            this.tabStep2.Location = new System.Drawing.Point(4, 5);
            this.tabStep2.Name = "tabStep2";
            this.tabStep2.Size = new System.Drawing.Size(892, 456);
            this.tabStep2.TabIndex = 1;
            this.tabStep2.Text = "Config";
            // 
            // btnNewMapping
            // 
            this.btnNewMapping.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(128)))), ((int)(((byte)(185)))));
            this.btnNewMapping.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNewMapping.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.btnNewMapping.ForeColor = System.Drawing.Color.White;
            this.btnNewMapping.Location = new System.Drawing.Point(235, 260);
            this.btnNewMapping.Name = "btnNewMapping";
            this.btnNewMapping.Size = new System.Drawing.Size(200, 35);
            this.btnNewMapping.TabIndex = 9;
            this.btnNewMapping.Text = "+ NUOVA MAPPATURA";
            this.btnNewMapping.UseVisualStyleBackColor = false;
            this.btnNewMapping.Click += new System.EventHandler(this.btnNewMapping_Click);
            // 
            // btnEditMapping
            // 
            this.btnEditMapping.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(126)))), ((int)(((byte)(34)))));
            this.btnEditMapping.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEditMapping.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.btnEditMapping.ForeColor = System.Drawing.Color.White;
            this.btnEditMapping.Location = new System.Drawing.Point(445, 260);
            this.btnEditMapping.Name = "btnEditMapping";
            this.btnEditMapping.Size = new System.Drawing.Size(200, 35);
            this.btnEditMapping.TabIndex = 10;
            this.btnEditMapping.Text = "MODIFICA MAPPATURA";
            this.btnEditMapping.UseVisualStyleBackColor = false;
            this.btnEditMapping.Click += new System.EventHandler(this.btnEditMapping_Click);
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.DimGray;
            this.label2.Location = new System.Drawing.Point(200, 100);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(500, 60);
            this.label2.TabIndex = 8;
            this.label2.Text = "Seleziona la mappatura delle colonne desiderata.\r\n(Le mappature definiscono la po" +
    "sizione di EAN, Descrizione e Prezzi)";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI Bold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(375, 180);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(148, 21);
            this.label1.TabIndex = 4;
            this.label1.Text = "Scegli Mappatura:";
            // 
            // cmbMapping
            // 
            this.cmbMapping.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbMapping.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbMapping.FormattingEnabled = true;
            this.cmbMapping.Location = new System.Drawing.Point(250, 210);
            this.cmbMapping.Name = "cmbMapping";
            this.cmbMapping.Size = new System.Drawing.Size(400, 29);
            this.cmbMapping.TabIndex = 3;
            // 
            // tabStep3
            // 
            this.tabStep3.Controls.Add(this.lblLegend);
            this.tabStep3.Controls.Add(this.chkSetActiveStatus);
            this.tabStep3.Controls.Add(this.chkSendToCasse);
            this.tabStep3.Controls.Add(this.chkSendToStampa);
            this.tabStep3.Controls.Add(this.chkUpdateStorePrice);
            this.tabStep3.Controls.Add(this.chkInsertNew);
            this.tabStep3.Controls.Add(this.dgvImport);
            this.tabStep3.Location = new System.Drawing.Point(4, 5);
            this.tabStep3.Name = "tabStep3";
            this.tabStep3.Size = new System.Drawing.Size(892, 456);
            this.tabStep3.TabIndex = 2;
            this.tabStep3.Text = "Anteprima";
            // 
            // dgvImport
            // 
            this.dgvImport.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvImport.BackgroundColor = System.Drawing.Color.White;
            this.dgvImport.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvImport.Location = new System.Drawing.Point(0, 68);
            this.dgvImport.Name = "dgvImport";
            this.dgvImport.Size = new System.Drawing.Size(892, 383);
            this.dgvImport.TabIndex = 5;
            // 
            // chkUpdateStorePrice
            // 
            this.chkUpdateStorePrice.AutoSize = true;
            this.chkUpdateStorePrice.Checked = true;
            this.chkUpdateStorePrice.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkUpdateStorePrice.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.chkUpdateStorePrice.Location = new System.Drawing.Point(10, 6);
            this.chkUpdateStorePrice.Name = "chkUpdateStorePrice";
            this.chkUpdateStorePrice.Size = new System.Drawing.Size(320, 19);
            this.chkUpdateStorePrice.TabIndex = 6;
            this.chkUpdateStorePrice.Text = "Applica Prezzo Excel come Prezzo Vendita Negozio";
            this.chkUpdateStorePrice.UseVisualStyleBackColor = true;
            // 
            // chkInsertNew
            // 
            this.chkInsertNew.AutoSize = true;
            this.chkInsertNew.Checked = true;
            this.chkInsertNew.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkInsertNew.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.chkInsertNew.Location = new System.Drawing.Point(400, 6);
            this.chkInsertNew.Name = "chkInsertNew";
            this.chkInsertNew.Size = new System.Drawing.Size(320, 19);
            this.chkInsertNew.TabIndex = 7;
            this.chkInsertNew.Text = "Inserisci articoli non trovati come nuove referenze";
            this.chkInsertNew.UseVisualStyleBackColor = true;
            // 
            // chkSendToCasse
            // 
            this.chkSendToCasse.AutoSize = true;
            this.chkSendToCasse.Checked = true;
            this.chkSendToCasse.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSendToCasse.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.chkSendToCasse.Location = new System.Drawing.Point(10, 26);
            this.chkSendToCasse.Name = "chkSendToCasse";
            this.chkSendToCasse.Size = new System.Drawing.Size(326, 19);
            this.chkSendToCasse.TabIndex = 8;
            this.chkSendToCasse.Text = "Genera variazioni per invio alle CASSE (Punto cassa)";
            this.chkSendToCasse.UseVisualStyleBackColor = true;
            // 
            // chkSendToStampa
            // 
            this.chkSendToStampa.AutoSize = true;
            this.chkSendToStampa.Checked = true;
            this.chkSendToStampa.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSendToStampa.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.chkSendToStampa.Location = new System.Drawing.Point(400, 26);
            this.chkSendToStampa.Name = "chkSendToStampa";
            this.chkSendToStampa.Size = new System.Drawing.Size(320, 19);
            this.chkSendToStampa.TabIndex = 9;
            this.chkSendToStampa.Text = "Genera segnalazioni per stampa ETICHETTE (Variazioni prezzi)";
            this.chkSendToStampa.UseVisualStyleBackColor = true;
            // 
            // chkSetActiveStatus
            // 
            this.chkSetActiveStatus.AutoSize = true;
            this.chkSetActiveStatus.Checked = true;
            this.chkSetActiveStatus.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSetActiveStatus.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.chkSetActiveStatus.Location = new System.Drawing.Point(10, 46);
            this.chkSetActiveStatus.Name = "chkSetActiveStatus";
            this.chkSetActiveStatus.Size = new System.Drawing.Size(320, 19);
            this.chkSetActiveStatus.TabIndex = 10;
            this.chkSetActiveStatus.Text = "Imposta stato articolo ATTIVO (\'A\') in anagrafica";
            this.chkSetActiveStatus.UseVisualStyleBackColor = true;
            // 
            // lblLegend
            // 
            this.lblLegend.AutoSize = true;
            this.lblLegend.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold);
            this.lblLegend.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(44)))), ((int)(((byte)(62)))), ((int)(((byte)(80)))));
            this.lblLegend.Location = new System.Drawing.Point(400, 48);
            this.lblLegend.Name = "lblLegend";
            this.lblLegend.Size = new System.Drawing.Size(460, 13);
            this.lblLegend.TabIndex = 11;
            this.lblLegend.Text = "Legenda: [ Blu = Nuova Referenza ]  [ Rosso = Aumento Costo ]  [ Verde = Calo Costo ]";
            // 
            // tabStep4
            // 
            this.tabStep4.BackColor = System.Drawing.Color.White;
            this.tabStep4.Controls.Add(this.btnAcquire);
            this.tabStep4.Controls.Add(this.lblSummary);
            this.tabStep4.Location = new System.Drawing.Point(4, 5);
            this.tabStep4.Name = "tabStep4";
            this.tabStep4.Size = new System.Drawing.Size(892, 456);
            this.tabStep4.TabIndex = 3;
            this.tabStep4.Text = "Fine";
            // 
            // btnAcquire
            // 
            this.btnAcquire.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(204)))), ((int)(((byte)(113)))));
            this.btnAcquire.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAcquire.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAcquire.ForeColor = System.Drawing.Color.White;
            this.btnAcquire.Location = new System.Drawing.Point(250, 200);
            this.btnAcquire.Name = "btnAcquire";
            this.btnAcquire.Size = new System.Drawing.Size(400, 70);
            this.btnAcquire.TabIndex = 6;
            this.btnAcquire.Text = "AVVIA ACQUISIZIONE DATI";
            this.btnAcquire.UseVisualStyleBackColor = false;
            this.btnAcquire.Click += new System.EventHandler(this.btnAcquire_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 0);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(900, 12);
            this.progressBar1.TabIndex = 4;
            this.progressBar1.Visible = false;
            // 
            // lblSummary
            // 
            this.lblSummary.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSummary.Location = new System.Drawing.Point(100, 80);
            this.lblSummary.Name = "lblSummary";
            this.lblSummary.Size = new System.Drawing.Size(700, 100);
            this.lblSummary.TabIndex = 8;
            this.lblSummary.Text = "Tutto pronto per l\'importazione.\r\nVerifica i dati nell\'anteprima e procedi.\r\n\r\nRi" +
    "ghe da elaborare: 0";
            this.lblSummary.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // frmUtyImportXls
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(900, 575);
            this.Controls.Add(this.tabWizard);
            this.Controls.Add(this.pnlBottom);
            this.Controls.Add(this.pnlTop);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "frmUtyImportXls";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Wizard Importazione Excel";
            this.Load += new System.EventHandler(this.frmUtyImportXls_Load);
            this.pnlTop.ResumeLayout(false);
            this.pnlTop.PerformLayout();
            this.pnlBottom.ResumeLayout(false);
            this.tabWizard.ResumeLayout(false);
            this.tabStep1.ResumeLayout(false);
            this.tabStep1.PerformLayout();
            this.tabStep2.ResumeLayout(false);
            this.tabStep2.PerformLayout();
            this.tabStep3.ResumeLayout(false);
            this.tabStep4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvImport)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlTop;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblStepInfo;
        private System.Windows.Forms.Panel pnlBottom;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Button btnBack;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.TabControl tabWizard;
        private System.Windows.Forms.TabPage tabStep1;
        private System.Windows.Forms.TabPage tabStep2;
        private System.Windows.Forms.TabPage tabStep3;
        private System.Windows.Forms.TabPage tabStep4;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Label lblFileStatus;
        private System.Windows.Forms.Button btnNewMapping;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbMapping;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private APOffice.APDataGridView dgvImport;
        private System.Windows.Forms.Button btnAcquire;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblSummary;
        private System.Windows.Forms.CheckBox chkUpdateStorePrice;
        private System.Windows.Forms.CheckBox chkInsertNew;
        private System.Windows.Forms.Button btnEditMapping;
        private System.Windows.Forms.CheckBox chkSendToCasse;
        private System.Windows.Forms.CheckBox chkSendToStampa;
        private System.Windows.Forms.Label lblFileSelected;
        public System.Windows.Forms.CheckBox chkSetActiveStatus;
        public System.Windows.Forms.Label lblLegend;
    }
}

