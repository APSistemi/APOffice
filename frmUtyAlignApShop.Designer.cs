namespace APOffice
{
    partial class frmUtyAlignApShop
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
            this.pnlTop = new System.Windows.Forms.Panel();
            this.btnAnalizza = new System.Windows.Forms.Button();
            this.btnBrowseMdb = new System.Windows.Forms.Button();
            this.txtMdbPath = new System.Windows.Forms.TextBox();
            this.lblMdbPath = new System.Windows.Forms.Label();
            this.pnlStats = new System.Windows.Forms.Panel();
            this.lblStrategy = new System.Windows.Forms.Label();
            this.cmbPriceStrategy = new System.Windows.Forms.ComboBox();
            this.chkKeepLongerDes = new System.Windows.Forms.CheckBox();
            this.chkOnlyDiff = new System.Windows.Forms.CheckBox();
            this.chkIgnoreNoEan = new System.Windows.Forms.CheckBox();
            this.chkIgnoreNoPrv = new System.Windows.Forms.CheckBox();
            this.chkIgnoreSamePrice = new System.Windows.Forms.CheckBox();
            this.lblSort = new System.Windows.Forms.Label();
            this.cmbSort = new System.Windows.Forms.ComboBox();
            this.lblAllineati = new System.Windows.Forms.Label();
            this.lblDiversi = new System.Windows.Forms.Label();
            this.lblMancanti = new System.Windows.Forms.Label();
            this.lblTotApShop = new System.Windows.Forms.Label();
            this.dgvGrid = new System.Windows.Forms.DataGridView();
            this.colCod = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDes = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colEan = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colRep = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colIva = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colUmi = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colPrv = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colPrvSql = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDiffPrv = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDtVenApShop = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDtVenAPOffice = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCos = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCosSql = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDiffCos = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDtAcqApShop = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDtAcqAPOffice = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colStato = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pnlBottom = new System.Windows.Forms.Panel();
            this.btnStop = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btnEsci = new System.Windows.Forms.Button();
            this.btnAllinea = new System.Windows.Forms.Button();
            this.pnlTop.SuspendLayout();
            this.pnlStats.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvGrid)).BeginInit();
            this.pnlBottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlTop
            // 
            this.pnlTop.BackColor = System.Drawing.SystemColors.ControlLight;
            this.pnlTop.Controls.Add(this.btnAnalizza);
            this.pnlTop.Controls.Add(this.btnBrowseMdb);
            this.pnlTop.Controls.Add(this.txtMdbPath);
            this.pnlTop.Controls.Add(this.lblMdbPath);
            this.pnlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlTop.Location = new System.Drawing.Point(0, 0);
            this.pnlTop.Name = "pnlTop";
            this.pnlTop.Size = new System.Drawing.Size(864, 52);
            this.pnlTop.TabIndex = 0;
            // 
            // btnAnalizza
            // 
            this.btnAnalizza.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAnalizza.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnalizza.Location = new System.Drawing.Point(685, 11);
            this.btnAnalizza.Name = "btnAnalizza";
            this.btnAnalizza.Size = new System.Drawing.Size(167, 30);
            this.btnAnalizza.TabIndex = 3;
            this.btnAnalizza.Text = "🔍 Analizza / Confronta";
            this.btnAnalizza.UseVisualStyleBackColor = true;
            this.btnAnalizza.Click += new System.EventHandler(this.btnAnalizza_Click);
            // 
            // btnBrowseMdb
            // 
            this.btnBrowseMdb.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowseMdb.Location = new System.Drawing.Point(595, 14);
            this.btnBrowseMdb.Name = "btnBrowseMdb";
            this.btnBrowseMdb.Size = new System.Drawing.Size(75, 25);
            this.btnBrowseMdb.TabIndex = 2;
            this.btnBrowseMdb.Text = "Sfoglia...";
            this.btnBrowseMdb.UseVisualStyleBackColor = true;
            this.btnBrowseMdb.Click += new System.EventHandler(this.btnBrowseMdb_Click);
            // 
            // txtMdbPath
            // 
            this.txtMdbPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtMdbPath.Location = new System.Drawing.Point(180, 16);
            this.txtMdbPath.Name = "txtMdbPath";
            this.txtMdbPath.Size = new System.Drawing.Size(409, 20);
            this.txtMdbPath.TabIndex = 1;
            this.txtMdbPath.Text = "C:\\ApProject\\ApShop\\DataBase\\DataBase.mdb";
            // 
            // lblMdbPath
            // 
            this.lblMdbPath.AutoSize = true;
            this.lblMdbPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMdbPath.Location = new System.Drawing.Point(12, 18);
            this.lblMdbPath.Name = "lblMdbPath";
            this.lblMdbPath.Size = new System.Drawing.Size(162, 15);
            this.lblMdbPath.TabIndex = 0;
            this.lblMdbPath.Text = "Database APShop POS (.mdb):";
            // 
            // pnlStats
            // 
            this.pnlStats.BackColor = System.Drawing.Color.WhiteSmoke;
            this.pnlStats.Controls.Add(this.chkKeepLongerDes);
            this.pnlStats.Controls.Add(this.cmbPriceStrategy);
            this.pnlStats.Controls.Add(this.lblStrategy);
            this.pnlStats.Controls.Add(this.cmbSort);
            this.pnlStats.Controls.Add(this.lblSort);
            this.pnlStats.Controls.Add(this.chkIgnoreSamePrice);
            this.pnlStats.Controls.Add(this.chkIgnoreNoEan);
            this.pnlStats.Controls.Add(this.chkIgnoreNoPrv);
            this.pnlStats.Controls.Add(this.chkOnlyDiff);
            this.pnlStats.Controls.Add(this.lblAllineati);
            this.pnlStats.Controls.Add(this.lblDiversi);
            this.pnlStats.Controls.Add(this.lblMancanti);
            this.pnlStats.Controls.Add(this.lblTotApShop);
            this.pnlStats.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlStats.Location = new System.Drawing.Point(0, 52);
            this.pnlStats.Name = "pnlStats";
            this.pnlStats.Size = new System.Drawing.Size(864, 88);
            this.pnlStats.TabIndex = 1;
            // 
            // lblStrategy
            // 
            this.lblStrategy.AutoSize = true;
            this.lblStrategy.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStrategy.Location = new System.Drawing.Point(12, 62);
            this.lblStrategy.Name = "lblStrategy";
            this.lblStrategy.Size = new System.Drawing.Size(104, 13);
            this.lblStrategy.TabIndex = 10;
            this.lblStrategy.Text = "Strategia Prezzo:";
            // 
            // cmbPriceStrategy
            // 
            this.cmbPriceStrategy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPriceStrategy.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbPriceStrategy.FormattingEnabled = true;
            this.cmbPriceStrategy.Items.AddRange(new object[] {
            "Sempre da APShop (Standard)",
            "Mantieni Prezzo Maggiore (Max APShop / APOffice)",
            "Mantieni Prezzo con Data più Recente"});
            this.cmbPriceStrategy.Location = new System.Drawing.Point(120, 58);
            this.cmbPriceStrategy.Name = "cmbPriceStrategy";
            this.cmbPriceStrategy.Size = new System.Drawing.Size(315, 21);
            this.cmbPriceStrategy.TabIndex = 11;
            this.cmbPriceStrategy.SelectedIndexChanged += new System.EventHandler(this.chkFilter_CheckedChanged);
            // 
            // chkKeepLongerDes
            // 
            this.chkKeepLongerDes.AutoSize = true;
            this.chkKeepLongerDes.Checked = true;
            this.chkKeepLongerDes.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkKeepLongerDes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkKeepLongerDes.Location = new System.Drawing.Point(448, 60);
            this.chkKeepLongerDes.Name = "chkKeepLongerDes";
            this.chkKeepLongerDes.Size = new System.Drawing.Size(262, 17);
            this.chkKeepLongerDes.TabIndex = 12;
            this.chkKeepLongerDes.Text = "Mantieni Descrizione più lunga / completa";
            this.chkKeepLongerDes.UseVisualStyleBackColor = true;
            this.chkKeepLongerDes.CheckedChanged += new System.EventHandler(this.chkFilter_CheckedChanged);
            // 
            // lblSort
            // 
            this.lblSort.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblSort.AutoSize = true;
            this.lblSort.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSort.Location = new System.Drawing.Point(540, 11);
            this.lblSort.Name = "lblSort";
            this.lblSort.Size = new System.Drawing.Size(79, 13);
            this.lblSort.TabIndex = 8;
            this.lblSort.Text = "Ordinamento:";
            // 
            // cmbSort
            // 
            this.cmbSort.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSort.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbSort.FormattingEnabled = true;
            this.cmbSort.Items.AddRange(new object[] {
            "Da Inserire ➔ Da Aggiornare ➔ Allineati",
            "Da Aggiornare ➔ Da Inserire ➔ Allineati",
            "Allineati ➔ Da Inserire ➔ Da Aggiornare",
            "Codice Articolo (A-Z)",
            "Descrizione Articolo (A-Z)"});
            this.cmbSort.Location = new System.Drawing.Point(625, 8);
            this.cmbSort.Name = "cmbSort";
            this.cmbSort.Size = new System.Drawing.Size(227, 21);
            this.cmbSort.TabIndex = 9;
            this.cmbSort.SelectedIndexChanged += new System.EventHandler(this.cmbSort_SelectedIndexChanged);
            // 
            // chkIgnoreSamePrice
            // 
            this.chkIgnoreSamePrice.AutoSize = true;
            this.chkIgnoreSamePrice.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkIgnoreSamePrice.Location = new System.Drawing.Point(448, 34);
            this.chkIgnoreSamePrice.Name = "chkIgnoreSamePrice";
            this.chkIgnoreSamePrice.Size = new System.Drawing.Size(248, 17);
            this.chkIgnoreSamePrice.TabIndex = 7;
            this.chkIgnoreSamePrice.Text = "Escludi Prezzi Uguali (APShop = APOffice)";
            this.chkIgnoreSamePrice.UseVisualStyleBackColor = true;
            this.chkIgnoreSamePrice.CheckedChanged += new System.EventHandler(this.chkFilter_CheckedChanged);
            // 
            // chkIgnoreNoEan
            // 
            this.chkIgnoreNoEan.AutoSize = true;
            this.chkIgnoreNoEan.Checked = true;
            this.chkIgnoreNoEan.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIgnoreNoEan.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkIgnoreNoEan.Location = new System.Drawing.Point(12, 34);
            this.chkIgnoreNoEan.Name = "chkIgnoreNoEan";
            this.chkIgnoreNoEan.Size = new System.Drawing.Size(195, 17);
            this.chkIgnoreNoEan.TabIndex = 5;
            this.chkIgnoreNoEan.Text = "Escludi senza Barcode (EAN)";
            this.chkIgnoreNoEan.UseVisualStyleBackColor = true;
            this.chkIgnoreNoEan.CheckedChanged += new System.EventHandler(this.chkFilter_CheckedChanged);
            // 
            // chkIgnoreNoPrv
            // 
            this.chkIgnoreNoPrv.AutoSize = true;
            this.chkIgnoreNoPrv.Checked = true;
            this.chkIgnoreNoPrv.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIgnoreNoPrv.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkIgnoreNoPrv.Location = new System.Drawing.Point(215, 34);
            this.chkIgnoreNoPrv.Name = "chkIgnoreNoPrv";
            this.chkIgnoreNoPrv.Size = new System.Drawing.Size(222, 17);
            this.chkIgnoreNoPrv.TabIndex = 6;
            this.chkIgnoreNoPrv.Text = "Escludi senza Prezzo POS (<= 0)";
            this.chkIgnoreNoPrv.UseVisualStyleBackColor = true;
            this.chkIgnoreNoPrv.CheckedChanged += new System.EventHandler(this.chkFilter_CheckedChanged);
            // 
            // chkOnlyDiff
            // 
            this.chkOnlyDiff.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.chkOnlyDiff.AutoSize = true;
            this.chkOnlyDiff.Checked = true;
            this.chkOnlyDiff.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkOnlyDiff.Location = new System.Drawing.Point(702, 34);
            this.chkOnlyDiff.Name = "chkOnlyDiff";
            this.chkOnlyDiff.Size = new System.Drawing.Size(150, 17);
            this.chkOnlyDiff.TabIndex = 4;
            this.chkOnlyDiff.Text = "Mostra solo da inserire / aggiornare";
            this.chkOnlyDiff.UseVisualStyleBackColor = true;
            this.chkOnlyDiff.CheckedChanged += new System.EventHandler(this.chkOnlyDiff_CheckedChanged);
            // 
            // lblAllineati
            // 
            this.lblAllineati.AutoSize = true;
            this.lblAllineati.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAllineati.ForeColor = System.Drawing.Color.ForestGreen;
            this.lblAllineati.Location = new System.Drawing.Point(475, 12);
            this.lblAllineati.Name = "lblAllineati";
            this.lblAllineati.Size = new System.Drawing.Size(80, 15);
            this.lblAllineati.TabIndex = 3;
            this.lblAllineati.Text = "Allineati: 0";
            // 
            // lblDiversi
            // 
            this.lblDiversi.AutoSize = true;
            this.lblDiversi.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDiversi.ForeColor = System.Drawing.Color.DarkOrange;
            this.lblDiversi.Location = new System.Drawing.Point(310, 12);
            this.lblDiversi.Name = "lblDiversi";
            this.lblDiversi.Size = new System.Drawing.Size(127, 15);
            this.lblDiversi.TabIndex = 2;
            this.lblDiversi.Text = "Da Aggiornare: 0";
            // 
            // lblMancanti
            // 
            this.lblMancanti.AutoSize = true;
            this.lblMancanti.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMancanti.ForeColor = System.Drawing.Color.Crimson;
            this.lblMancanti.Location = new System.Drawing.Point(165, 12);
            this.lblMancanti.Name = "lblMancanti";
            this.lblMancanti.Size = new System.Drawing.Size(107, 15);
            this.lblMancanti.TabIndex = 1;
            this.lblMancanti.Text = "Da Inserire: 0";
            // 
            // lblTotApShop
            // 
            this.lblTotApShop.AutoSize = true;
            this.lblTotApShop.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTotApShop.Location = new System.Drawing.Point(12, 12);
            this.lblTotApShop.Name = "lblTotApShop";
            this.lblTotApShop.Size = new System.Drawing.Size(117, 15);
            this.lblTotApShop.TabIndex = 0;
            this.lblTotApShop.Text = "Articoli APShop: 0";
            // 
            // dgvGrid
            // 
            this.dgvGrid.AllowUserToAddRows = false;
            this.dgvGrid.AllowUserToDeleteRows = false;
            this.dgvGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.None;
            this.dgvGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colCod,
            this.colDes,
            this.colEan,
            this.colRep,
            this.colIva,
            this.colUmi,
            this.colPrv,
            this.colPrvSql,
            this.colDiffPrv,
            this.colDtVenApShop,
            this.colDtVenAPOffice,
            this.colCos,
            this.colCosSql,
            this.colDiffCos,
            this.colDtAcqApShop,
            this.colDtAcqAPOffice,
            this.colStato});
            this.dgvGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvGrid.Location = new System.Drawing.Point(0, 140);
            this.dgvGrid.Name = "dgvGrid";
            this.dgvGrid.ReadOnly = true;
            this.dgvGrid.RowHeadersWidth = 25;
            this.dgvGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvGrid.Size = new System.Drawing.Size(864, 372);
            this.dgvGrid.TabIndex = 2;
            // 
            // colCod
            // 
            this.colCod.DataPropertyName = "art_cod";
            this.colCod.HeaderText = "Codice";
            this.colCod.Name = "colCod";
            this.colCod.ReadOnly = true;
            this.colCod.Width = 75;
            // 
            // colDes
            // 
            this.colDes.DataPropertyName = "art_des";
            this.colDes.HeaderText = "Descrizione APShop";
            this.colDes.Name = "colDes";
            this.colDes.ReadOnly = true;
            this.colDes.Width = 150;
            // 
            // colEan
            // 
            this.colEan.DataPropertyName = "ean_ean";
            this.colEan.HeaderText = "Barcode EAN";
            this.colEan.Name = "colEan";
            this.colEan.ReadOnly = true;
            this.colEan.Width = 95;
            // 
            // colRep
            // 
            this.colRep.DataPropertyName = "art_rep";
            this.colRep.HeaderText = "Rep.";
            this.colRep.Name = "colRep";
            this.colRep.ReadOnly = true;
            this.colRep.Width = 45;
            // 
            // colIva
            // 
            this.colIva.DataPropertyName = "art_iva";
            this.colIva.HeaderText = "IVA";
            this.colIva.Name = "colIva";
            this.colIva.ReadOnly = true;
            this.colIva.Width = 45;
            // 
            // colUmi
            // 
            this.colUmi.DataPropertyName = "art_umi";
            this.colUmi.HeaderText = "UM";
            this.colUmi.Name = "colUmi";
            this.colUmi.ReadOnly = true;
            this.colUmi.Width = 45;
            // 
            // colPrv
            // 
            this.colPrv.DataPropertyName = "art_prv";
            this.colPrv.HeaderText = "Prz APShop";
            this.colPrv.Name = "colPrv";
            this.colPrv.ReadOnly = true;
            this.colPrv.Width = 75;
            // 
            // colPrvSql
            // 
            this.colPrvSql.DataPropertyName = "apo_prv";
            this.colPrvSql.HeaderText = "Prz APOffice";
            this.colPrvSql.Name = "colPrvSql";
            this.colPrvSql.ReadOnly = true;
            this.colPrvSql.Width = 75;
            // 
            // colDiffPrv
            // 
            this.colDiffPrv.DataPropertyName = "diff_prv";
            this.colDiffPrv.HeaderText = "Diff. Prz";
            this.colDiffPrv.Name = "colDiffPrv";
            this.colDiffPrv.ReadOnly = true;
            this.colDiffPrv.Width = 70;
            // 
            // colDtVenApShop
            // 
            this.colDtVenApShop.DataPropertyName = "dt_ven_apshop";
            this.colDtVenApShop.HeaderText = "Dt.Ven APShop";
            this.colDtVenApShop.Name = "colDtVenApShop";
            this.colDtVenApShop.ReadOnly = true;
            this.colDtVenApShop.Width = 95;
            // 
            // colDtVenAPOffice
            // 
            this.colDtVenAPOffice.DataPropertyName = "dt_ven_apoffice";
            this.colDtVenAPOffice.HeaderText = "Dt.Ven APOffice";
            this.colDtVenAPOffice.Name = "colDtVenAPOffice";
            this.colDtVenAPOffice.ReadOnly = true;
            this.colDtVenAPOffice.Width = 95;
            // 
            // colCos
            // 
            this.colCos.DataPropertyName = "art_cos";
            this.colCos.HeaderText = "Cos APShop";
            this.colCos.Name = "colCos";
            this.colCos.ReadOnly = true;
            this.colCos.Width = 75;
            // 
            // colCosSql
            // 
            this.colCosSql.DataPropertyName = "apo_cos";
            this.colCosSql.HeaderText = "Cos APOffice";
            this.colCosSql.Name = "colCosSql";
            this.colCosSql.ReadOnly = true;
            this.colCosSql.Width = 75;
            // 
            // colDiffCos
            // 
            this.colDiffCos.DataPropertyName = "diff_cos";
            this.colDiffCos.HeaderText = "Diff. Cos";
            this.colDiffCos.Name = "colDiffCos";
            this.colDiffCos.ReadOnly = true;
            this.colDiffCos.Width = 70;
            // 
            // colDtAcqApShop
            // 
            this.colDtAcqApShop.DataPropertyName = "dt_acq_apshop";
            this.colDtAcqApShop.HeaderText = "Dt.Acq APShop";
            this.colDtAcqApShop.Name = "colDtAcqApShop";
            this.colDtAcqApShop.ReadOnly = true;
            this.colDtAcqApShop.Width = 95;
            // 
            // colDtAcqAPOffice
            // 
            this.colDtAcqAPOffice.DataPropertyName = "dt_acq_apoffice";
            this.colDtAcqAPOffice.HeaderText = "Dt.Acq APOffice";
            this.colDtAcqAPOffice.Name = "colDtAcqAPOffice";
            this.colDtAcqAPOffice.ReadOnly = true;
            this.colDtAcqAPOffice.Width = 95;
            // 
            // colStato
            // 
            this.colStato.DataPropertyName = "stato_sync";
            this.colStato.HeaderText = "Stato Allineamento";
            this.colStato.Name = "colStato";
            this.colStato.ReadOnly = true;
            this.colStato.Width = 105;
            // 
            // pnlBottom
            // 
            this.pnlBottom.BackColor = System.Drawing.SystemColors.Control;
            this.pnlBottom.Controls.Add(this.btnStop);
            this.pnlBottom.Controls.Add(this.progressBar1);
            this.pnlBottom.Controls.Add(this.btnEsci);
            this.pnlBottom.Controls.Add(this.btnAllinea);
            this.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnlBottom.Location = new System.Drawing.Point(0, 510);
            this.pnlBottom.Name = "pnlBottom";
            this.pnlBottom.Size = new System.Drawing.Size(864, 52);
            this.pnlBottom.TabIndex = 3;
            // 
            // btnStop
            // 
            this.btnStop.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnStop.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnStop.ForeColor = System.Drawing.Color.DarkRed;
            this.btnStop.Location = new System.Drawing.Point(395, 10);
            this.btnStop.Name = "btnStop";
            this.btnStop.Size = new System.Drawing.Size(110, 32);
            this.btnStop.TabIndex = 3;
            this.btnStop.Text = "🛑 Interrompi";
            this.btnStop.UseVisualStyleBackColor = true;
            this.btnStop.Visible = false;
            this.btnStop.Click += new System.EventHandler(this.btnStop_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(12, 15);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(370, 23);
            this.progressBar1.TabIndex = 2;
            this.progressBar1.Visible = false;
            // 
            // btnEsci
            // 
            this.btnEsci.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnEsci.Location = new System.Drawing.Point(760, 10);
            this.btnEsci.Name = "btnEsci";
            this.btnEsci.Size = new System.Drawing.Size(92, 32);
            this.btnEsci.TabIndex = 1;
            this.btnEsci.Text = "Esci";
            this.btnEsci.UseVisualStyleBackColor = true;
            this.btnEsci.Click += new System.EventHandler(this.btnEsci_Click);
            // 
            // btnAllinea
            // 
            this.btnAllinea.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAllinea.Enabled = false;
            this.btnAllinea.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAllinea.Location = new System.Drawing.Point(515, 10);
            this.btnAllinea.Name = "btnAllinea";
            this.btnAllinea.Size = new System.Drawing.Size(235, 32);
            this.btnAllinea.TabIndex = 0;
            this.btnAllinea.Text = "🚀 Allinea e Importa in APOffice";
            this.btnAllinea.UseVisualStyleBackColor = true;
            this.btnAllinea.Click += new System.EventHandler(this.btnAllinea_Click);
            // 
            // frmUtyAlignApShop
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(864, 562);
            this.Controls.Add(this.dgvGrid);
            this.Controls.Add(this.pnlBottom);
            this.Controls.Add(this.pnlStats);
            this.Controls.Add(this.pnlTop);
            this.Name = "frmUtyAlignApShop";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Confronto e Allineamento Anagrafica APShop POS -> APOffice";
            this.Load += new System.EventHandler(this.frmUtyAlignApShop_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyAlignApShop_KeyDown);
            this.pnlTop.ResumeLayout(false);
            this.pnlTop.PerformLayout();
            this.pnlStats.ResumeLayout(false);
            this.pnlStats.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvGrid)).EndInit();
            this.pnlBottom.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlTop;
        private System.Windows.Forms.Label lblMdbPath;
        private System.Windows.Forms.TextBox txtMdbPath;
        private System.Windows.Forms.Button btnBrowseMdb;
        private System.Windows.Forms.Button btnAnalizza;
        private System.Windows.Forms.Panel pnlStats;
        private System.Windows.Forms.Label lblTotApShop;
        private System.Windows.Forms.Label lblMancanti;
        private System.Windows.Forms.Label lblDiversi;
        private System.Windows.Forms.Label lblAllineati;
        private System.Windows.Forms.CheckBox chkOnlyDiff;
        private System.Windows.Forms.CheckBox chkIgnoreNoEan;
        private System.Windows.Forms.CheckBox chkIgnoreNoPrv;
        private System.Windows.Forms.CheckBox chkIgnoreSamePrice;
        private System.Windows.Forms.Label lblSort;
        private System.Windows.Forms.ComboBox cmbSort;
        private System.Windows.Forms.Label lblStrategy;
        private System.Windows.Forms.ComboBox cmbPriceStrategy;
        private System.Windows.Forms.CheckBox chkKeepLongerDes;
        private System.Windows.Forms.DataGridView dgvGrid;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCod;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDes;
        private System.Windows.Forms.DataGridViewTextBoxColumn colEan;
        private System.Windows.Forms.DataGridViewTextBoxColumn colRep;
        private System.Windows.Forms.DataGridViewTextBoxColumn colIva;
        private System.Windows.Forms.DataGridViewTextBoxColumn colUmi;
        private System.Windows.Forms.DataGridViewTextBoxColumn colPrv;
        private System.Windows.Forms.DataGridViewTextBoxColumn colPrvSql;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDiffPrv;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDtVenApShop;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDtVenAPOffice;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCos;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCosSql;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDiffCos;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDtAcqApShop;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDtAcqAPOffice;
        private System.Windows.Forms.DataGridViewTextBoxColumn colStato;
        private System.Windows.Forms.Panel pnlBottom;
        private System.Windows.Forms.Button btnAllinea;
        private System.Windows.Forms.Button btnEsci;
        private System.Windows.Forms.Button btnStop;
        private System.Windows.Forms.ProgressBar progressBar1;
    }
}
