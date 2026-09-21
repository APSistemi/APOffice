namespace APOffice
{
    partial class frmUtyEtiFormati
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            // ---- Contenitore principale a 3 colonne: Lista | Proprieta | Canvas ----
            this.tblMain = new System.Windows.Forms.TableLayoutPanel();

            // -- Colonna 1: lista formati + bottoni --
            this.pnlLista = new System.Windows.Forms.Panel();
            this.dgvFormati = new APOffice.APDataGridView();
            this.pnlBtnLista = new System.Windows.Forms.Panel();
            this.btnNuovo = new System.Windows.Forms.Button();
            this.btnDuplica = new System.Windows.Forms.Button();
            this.btnElimina = new System.Windows.Forms.Button();

            // -- Colonna 2: proprieta formato + elemento --
            this.pnlProp = new System.Windows.Forms.Panel();
            this.grpFormato = new System.Windows.Forms.GroupBox();
            this.lblCodice = new System.Windows.Forms.Label();
            this.txtCodice = new System.Windows.Forms.TextBox();
            this.lblDescrizione = new System.Windows.Forms.Label();
            this.txtDescrizione = new System.Windows.Forms.TextBox();
            this.lblLargh = new System.Windows.Forms.Label();
            this.nudLargh = new System.Windows.Forms.NumericUpDown();
            this.lblAlt = new System.Windows.Forms.Label();
            this.nudAlt = new System.Windows.Forms.NumericUpDown();
            this.lblCols = new System.Windows.Forms.Label();
            this.nudCols = new System.Windows.Forms.NumericUpDown();
            this.lblRighe = new System.Windows.Forms.Label();
            this.nudRighe = new System.Windows.Forms.NumericUpDown();
            this.grpOrient = new System.Windows.Forms.GroupBox();
            this.radOrientV = new System.Windows.Forms.RadioButton();
            this.radOrientH = new System.Windows.Forms.RadioButton();
            this.lblMarginTop = new System.Windows.Forms.Label();
            this.nudMarginTop = new System.Windows.Forms.NumericUpDown();
            this.lblMarginLeft = new System.Windows.Forms.Label();
            this.nudMarginLeft = new System.Windows.Forms.NumericUpDown();
            this.lblEtiSta = new System.Windows.Forms.Label();
            this.cmbEtiSta = new System.Windows.Forms.ComboBox();

            this.grpElemento = new System.Windows.Forms.GroupBox();
            this.lstCampi = new System.Windows.Forms.ListBox();
            this.chkVisibile = new System.Windows.Forms.CheckBox();
            this.lblFont = new System.Windows.Forms.Label();
            this.cmbFont = new System.Windows.Forms.ComboBox();
            this.lblFontSz = new System.Windows.Forms.Label();
            this.nudFontSz = new System.Windows.Forms.NumericUpDown();
            this.chkBold = new System.Windows.Forms.CheckBox();
            this.cmbAllinea = new System.Windows.Forms.ComboBox();
            this.lblPosX = new System.Windows.Forms.Label();
            this.nudPosX = new System.Windows.Forms.NumericUpDown();
            this.lblPosY = new System.Windows.Forms.Label();
            this.nudPosY = new System.Windows.Forms.NumericUpDown();
            this.lblPosW = new System.Windows.Forms.Label();
            this.nudPosW = new System.Windows.Forms.NumericUpDown();
            this.lblPosH = new System.Windows.Forms.Label();
            this.nudPosH = new System.Windows.Forms.NumericUpDown();
            this.lblInfo = new System.Windows.Forms.Label();
            this.pnlPropBottom = new System.Windows.Forms.Panel();
            this.btnSalva = new System.Windows.Forms.Button();
            this.btnEsci = new System.Windows.Forms.Button();

            // -- Colonna 3: anteprima canvas --
            this.pnlCanvas = new System.Windows.Forms.Panel();
            this.lblCanvasTip = new System.Windows.Forms.Label();
            this.pnlCanvasTop = new System.Windows.Forms.Panel();
            this.btnZoomIn = new System.Windows.Forms.Button();
            this.btnZoomOut = new System.Windows.Forms.Button();
            this.lblZoom = new System.Windows.Forms.Label();
            this.lblFormaTipo = new System.Windows.Forms.Label();
            this.cmbFormaTipo = new System.Windows.Forms.ComboBox();
            this.lblFormaColore = new System.Windows.Forms.Label();
            this.cmbFormaColore = new System.Windows.Forms.ComboBox();
            this.lblFormaSpess = new System.Windows.Forms.Label();
            this.nudFormaSpess = new System.Windows.Forms.NumericUpDown();
            this.chkFormaFill = new System.Windows.Forms.CheckBox();
            this.chkFormaZ = new System.Windows.Forms.CheckBox();
            this.btnCaricaLogo = new System.Windows.Forms.Button();
            this.lblLogo = new System.Windows.Forms.Label();
            this.txtLogoPath = new System.Windows.Forms.TextBox();
            this.lblTestoLibero = new System.Windows.Forms.Label();
            this.txtTestoLibero = new System.Windows.Forms.TextBox();
            this.picCanvas = new System.Windows.Forms.PictureBox();

            ((System.ComponentModel.ISupportInitialize)(this.dgvFormati)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudLargh)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudAlt)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudCols)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudRighe)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudMarginTop)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudMarginLeft)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudFontSz)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudPosX)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudPosY)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudPosW)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudPosH)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudFormaSpess)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picCanvas)).BeginInit();
            this.SuspendLayout();

            // ====================================================
            // TABLE LAYOUT PRINCIPALE  (3 colonne)
            // ====================================================
            this.tblMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblMain.ColumnCount = 3;
            this.tblMain.RowCount = 1;
            this.tblMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 230f));  // Colonna Lista
            this.tblMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 420f));  // Colonna Proprieta
            this.tblMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100f));  // Colonna Canvas
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100f));
            this.tblMain.Controls.Add(this.pnlLista, 0, 0);
            this.tblMain.Controls.Add(this.pnlProp, 1, 0);
            this.tblMain.Controls.Add(this.pnlCanvas, 2, 0);
            this.tblMain.Padding = new System.Windows.Forms.Padding(3);
            this.tblMain.BackColor = System.Drawing.Color.FromArgb(243, 244, 246);

            // ====================================================
            // COL 1 — LISTA FORMATI
            // ====================================================
            this.pnlLista.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlLista.BackColor = System.Drawing.Color.FromArgb(243, 244, 246);
            this.pnlLista.Padding = new System.Windows.Forms.Padding(2);

            // Header lista
            var lblLista = new System.Windows.Forms.Label();
            lblLista.Text = "FORMATI ETICHETTE";
            lblLista.Dock = System.Windows.Forms.DockStyle.Top;
            lblLista.Height = 26;
            lblLista.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            lblLista.BackColor = System.Drawing.Color.FromArgb(51, 65, 85);
            lblLista.ForeColor = System.Drawing.Color.White;
            lblLista.Font = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold);

            // Bottoni in cima
            this.pnlBtnLista.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlBtnLista.Height = 36;
            this.pnlBtnLista.Padding = new System.Windows.Forms.Padding(1);
            this.pnlBtnLista.BackColor = System.Drawing.Color.FromArgb(243, 244, 246);

            this.btnNuovo.Location = new System.Drawing.Point(2, 3);
            this.btnNuovo.Size = new System.Drawing.Size(72, 28);
            this.btnNuovo.Click += new System.EventHandler(this.btnNuovo_Click);

            this.btnDuplica.Location = new System.Drawing.Point(77, 3);
            this.btnDuplica.Size = new System.Drawing.Size(72, 28);
            this.btnDuplica.Click += new System.EventHandler(this.btnDuplica_Click);

            this.btnElimina.Location = new System.Drawing.Point(152, 3);
            this.btnElimina.Size = new System.Drawing.Size(72, 28);
            this.btnElimina.Click += new System.EventHandler(this.btnElimina_Click);

            this.pnlBtnLista.Controls.AddRange(new System.Windows.Forms.Control[] { btnNuovo, btnDuplica, btnElimina });

            // DataGrid
            this.dgvFormati.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvFormati.AllowUserToAddRows = false;
            this.dgvFormati.AllowUserToDeleteRows = false;
            this.dgvFormati.ReadOnly = true;
            this.dgvFormati.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvFormati.RowHeadersVisible = false;
            this.dgvFormati.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvFormati.BackgroundColor = System.Drawing.Color.White;

            this.pnlLista.Controls.Add(this.dgvFormati);
            this.pnlLista.Controls.Add(this.pnlBtnLista);
            this.pnlLista.Controls.Add(lblLista);

            // ====================================================
            // COL 2 — PROPRIETA'
            // ====================================================
            this.pnlProp.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlProp.BackColor = System.Drawing.Color.FromArgb(243, 244, 246);
            this.pnlProp.Padding = new System.Windows.Forms.Padding(4);
            this.pnlProp.AutoScroll = true;

            // --- GroupBox Formato ---
            this.grpFormato.Text = "Formato Etichetta";
            this.grpFormato.Location = new System.Drawing.Point(4, 4);
            this.grpFormato.Size = new System.Drawing.Size(404, 140);
            this.grpFormato.Font = new System.Drawing.Font("Segoe UI", 8.5f, System.Drawing.FontStyle.Bold);
            this.grpFormato.ForeColor = System.Drawing.Color.FromArgb(30, 41, 59);

            Lbl(this.lblCodice, "Codice:", 8, 24, 45);
            this.txtCodice.Location = new System.Drawing.Point(55, 22);
            this.txtCodice.Size = new System.Drawing.Size(45, 21);
            this.txtCodice.ReadOnly = true;
            this.txtCodice.BackColor = System.Drawing.SystemColors.Control;
            this.txtCodice.Font = new System.Drawing.Font("Segoe UI", 8.5f);

            Lbl(this.lblDescrizione, "Nome:", 108, 24, 42);
            this.txtDescrizione.Location = new System.Drawing.Point(152, 22);
            this.txtDescrizione.Size = new System.Drawing.Size(244, 21);
            this.txtDescrizione.Font = new System.Drawing.Font("Segoe UI", 8.5f);

            Lbl(this.lblLargh, "Largh:", 8, 52, 42);
            this.nudLargh.Location = new System.Drawing.Point(52, 50); this.nudLargh.Size = new System.Drawing.Size(48, 21);
            this.nudLargh.Minimum = 20; this.nudLargh.Maximum = 420; this.nudLargh.Value = 100;
            this.nudLargh.Font = new System.Drawing.Font("Segoe UI", 8.5f);

            Lbl(this.lblAlt, "Altez:", 106, 52, 40);
            this.nudAlt.Location = new System.Drawing.Point(148, 50); this.nudAlt.Size = new System.Drawing.Size(48, 21);
            this.nudAlt.Minimum = 10; this.nudAlt.Maximum = 420; this.nudAlt.Value = 50;
            this.nudAlt.Font = new System.Drawing.Font("Segoe UI", 8.5f);

            Lbl(this.lblCols, "Col:", 204, 52, 30);
            this.nudCols.Location = new System.Drawing.Point(236, 50); this.nudCols.Size = new System.Drawing.Size(46, 21);
            this.nudCols.Minimum = 1; this.nudCols.Maximum = 10; this.nudCols.Value = 1;
            this.nudCols.Font = new System.Drawing.Font("Segoe UI", 8.5f);
            this.nudCols.ValueChanged += new System.EventHandler(this.NudDim_ValueChanged);

            Lbl(this.lblRighe, "Rig:", 294, 52, 30);
            this.nudRighe.Location = new System.Drawing.Point(326, 50); this.nudRighe.Size = new System.Drawing.Size(46, 21);
            this.nudRighe.Minimum = 1; this.nudRighe.Maximum = 30; this.nudRighe.Value = 1;
            this.nudRighe.Font = new System.Drawing.Font("Segoe UI", 8.5f);
            this.nudRighe.ValueChanged += new System.EventHandler(this.NudDim_ValueChanged);

            // Orientamento e margini compatti
            this.radOrientV.Text = "↕ Verticale"; this.radOrientV.Location = new System.Drawing.Point(8, 78);
            this.radOrientV.Size = new System.Drawing.Size(92, 20); this.radOrientV.Checked = true;
            this.radOrientV.Font = new System.Drawing.Font("Segoe UI", 8.25f);

            this.radOrientH.Text = "↔ Orizzontale"; this.radOrientH.Location = new System.Drawing.Point(104, 78);
            this.radOrientH.Size = new System.Drawing.Size(105, 20);
            this.radOrientH.Font = new System.Drawing.Font("Segoe UI", 8.25f);

            Lbl(this.lblMarginTop, "↑:", 216, 80, 20);
            this.nudMarginTop.Location = new System.Drawing.Point(236, 78); this.nudMarginTop.Size = new System.Drawing.Size(46, 21);
            this.nudMarginTop.Minimum = 0; this.nudMarginTop.Maximum = 50; this.nudMarginTop.Value = 0;
            this.nudMarginTop.Font = new System.Drawing.Font("Segoe UI", 8.5f);

            Lbl(this.lblMarginLeft, "←:", 296, 80, 20);
            this.nudMarginLeft.Location = new System.Drawing.Point(326, 78); this.nudMarginLeft.Size = new System.Drawing.Size(46, 21);
            this.nudMarginLeft.Minimum = 0; this.nudMarginLeft.Maximum = 50; this.nudMarginLeft.Value = 0;
            this.nudMarginLeft.Font = new System.Drawing.Font("Segoe UI", 8.5f);

            Lbl(this.lblEtiSta, "Post-Stampa:", 8, 108, 75);
            this.cmbEtiSta.Location = new System.Drawing.Point(86, 106);
            this.cmbEtiSta.Size = new System.Drawing.Size(310, 21);
            this.cmbEtiSta.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbEtiSta.Font = new System.Drawing.Font("Segoe UI", 8.25f);

            this.grpFormato.Controls.AddRange(new System.Windows.Forms.Control[] {
                lblCodice, txtCodice, lblDescrizione, txtDescrizione,
                lblLargh, nudLargh, lblAlt, nudAlt,
                lblCols, nudCols, lblRighe, nudRighe,
                radOrientV, radOrientH, lblMarginTop, nudMarginTop, lblMarginLeft, nudMarginLeft,
                lblEtiSta, cmbEtiSta });

            // --- GroupBox Elemento ---
            this.grpElemento.Text = "Campo Selezionato";
            this.grpElemento.Location = new System.Drawing.Point(4, 148);
            this.grpElemento.Size = new System.Drawing.Size(404, 304);
            this.grpElemento.Font = new System.Drawing.Font("Segoe UI", 8.5f, System.Drawing.FontStyle.Bold);
            this.grpElemento.ForeColor = System.Drawing.Color.FromArgb(30, 41, 59);

            // ListBox campi
            this.lstCampi.Location = new System.Drawing.Point(8, 20);
            this.lstCampi.Size = new System.Drawing.Size(388, 90);
            this.lstCampi.Font = new System.Drawing.Font("Segoe UI", 8.25f);
            this.lstCampi.ItemHeight = 15;
            this.lstCampi.SelectedIndexChanged += new System.EventHandler(this.lstCampi_SelectedIndexChanged);

            // Checkbox visibile
            this.chkVisibile.Text = "✔  MOSTRA sull'etichetta";
            this.chkVisibile.Location = new System.Drawing.Point(8, 114);
            this.chkVisibile.Size = new System.Drawing.Size(220, 22);
            this.chkVisibile.Font = new System.Drawing.Font("Segoe UI", 8.5f, System.Drawing.FontStyle.Bold);
            this.chkVisibile.ForeColor = System.Drawing.Color.FromArgb(15, 23, 42);
            this.chkVisibile.CheckedChanged += new System.EventHandler(this.ControlloProprietà_Changed);

            // Separatore 1
            var sep = new System.Windows.Forms.Label();
            sep.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            sep.Location = new System.Drawing.Point(8, 140); sep.Size = new System.Drawing.Size(388, 2);

            // Controlli per Testo Libero Personalizzato
            Lbl(this.lblTestoLibero, "Testo:", 8, 150, 36);
            this.lblTestoLibero.Visible = false;

            this.txtTestoLibero.Location = new System.Drawing.Point(46, 148);
            this.txtTestoLibero.Size = new System.Drawing.Size(350, 21);
            this.txtTestoLibero.Font = new System.Drawing.Font("Segoe UI", 8.5f);
            this.txtTestoLibero.Visible = false;
            this.txtTestoLibero.TextChanged += new System.EventHandler(this.ControlloProprietà_Changed);

            // Font / Pt / Grassetto
            Lbl(this.lblFont, "Font:", 8, 150, 34);
            this.cmbFont.Location = new System.Drawing.Point(44, 148);
            this.cmbFont.Size = new System.Drawing.Size(120, 21);
            this.cmbFont.Font = new System.Drawing.Font("Segoe UI", 8.25f);
            this.cmbFont.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFont.Items.AddRange(new object[] { "Arial", "Verdana", "Courier New", "Impact", "Tahoma" });
            this.cmbFont.SelectedIndexChanged += new System.EventHandler(this.ControlloProprietà_Changed);

            Lbl(this.lblFontSz, "Pt:", 170, 150, 22);
            this.nudFontSz.Location = new System.Drawing.Point(194, 148);
            this.nudFontSz.Size = new System.Drawing.Size(48, 21);
            this.nudFontSz.Minimum = 5; this.nudFontSz.Maximum = 250; this.nudFontSz.Value = 9;
            this.nudFontSz.Font = new System.Drawing.Font("Segoe UI", 8.5f);
            this.nudFontSz.ValueChanged += new System.EventHandler(this.ControlloProprietà_Changed);

            this.chkBold.Text = "Grassetto";
            this.chkBold.Location = new System.Drawing.Point(250, 148); this.chkBold.Size = new System.Drawing.Size(80, 21);
            this.chkBold.Font = new System.Drawing.Font("Segoe UI", 8.5f);
            this.chkBold.CheckedChanged += new System.EventHandler(this.ControlloProprietà_Changed);

            // Controlli per Forme
            Lbl(this.lblFormaTipo, "Tipo:", 8, 150, 34);
            this.cmbFormaTipo.Location = new System.Drawing.Point(44, 148);
            this.cmbFormaTipo.Size = new System.Drawing.Size(115, 21);
            this.cmbFormaTipo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFormaTipo.Items.AddRange(new object[] { "Rettangolo", "Linea Orizz.", "Linea Vert." });
            this.cmbFormaTipo.Font = new System.Drawing.Font("Segoe UI", 8.25f);
            this.cmbFormaTipo.SelectedIndexChanged += new System.EventHandler(this.ControlloProprietà_Changed);

            Lbl(this.lblFormaSpess, "Spess:", 166, 150, 42);
            this.nudFormaSpess.Location = new System.Drawing.Point(210, 148);
            this.nudFormaSpess.Size = new System.Drawing.Size(46, 21);
            this.nudFormaSpess.Minimum = 1; this.nudFormaSpess.Maximum = 20; this.nudFormaSpess.Value = 1;
            this.nudFormaSpess.Font = new System.Drawing.Font("Segoe UI", 8.5f);
            this.nudFormaSpess.ValueChanged += new System.EventHandler(this.ControlloProprietà_Changed);

            this.chkFormaFill.Text = "Riempito";
            this.chkFormaFill.Location = new System.Drawing.Point(264, 148); this.chkFormaFill.Size = new System.Drawing.Size(80, 21);
            this.chkFormaFill.Font = new System.Drawing.Font("Segoe UI", 8.5f);
            this.chkFormaFill.CheckedChanged += new System.EventHandler(this.ControlloProprietà_Changed);

            // Controlli per Immagine / Logo
            this.btnCaricaLogo.Location = new System.Drawing.Point(8, 146);
            this.btnCaricaLogo.Size = new System.Drawing.Size(88, 26);
            this.btnCaricaLogo.Visible = false;
            this.btnCaricaLogo.Click += new System.EventHandler(this.btnCaricaLogo_Click);

            Lbl(this.lblLogo, "File:", 102, 150, 30);
            this.lblLogo.Visible = false;

            this.txtLogoPath.Location = new System.Drawing.Point(134, 148);
            this.txtLogoPath.Size = new System.Drawing.Size(262, 21);
            this.txtLogoPath.ReadOnly = true;
            this.txtLogoPath.Visible = false;
            this.txtLogoPath.Font = new System.Drawing.Font("Segoe UI", 7.5f);

            // Slot 2: Allineamento, Colore, Z-Index
            this.cmbAllinea.Location = new System.Drawing.Point(8, 176);
            this.cmbAllinea.Size = new System.Drawing.Size(90, 21);
            this.cmbAllinea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAllinea.Font = new System.Drawing.Font("Segoe UI", 8.25f);
            this.cmbAllinea.Items.AddRange(new object[] { "Sinistra", "Destra", "Centro" });
            this.cmbAllinea.SelectedIndexChanged += new System.EventHandler(this.ControlloProprietà_Changed);

            Lbl(this.lblFormaColore, "Colore:", 106, 178, 44);
            this.cmbFormaColore.Location = new System.Drawing.Point(152, 176);
            this.cmbFormaColore.Size = new System.Drawing.Size(82, 21);
            this.cmbFormaColore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFormaColore.Font = new System.Drawing.Font("Segoe UI", 8.25f);
            this.cmbFormaColore.Items.AddRange(new object[] { "Nero", "Rosso", "Blu", "Verde", "Grigio", "Bianco", "Giallo" });
            this.cmbFormaColore.SelectedIndexChanged += new System.EventHandler(this.ControlloProprietà_Changed);

            this.chkFormaZ.Text = "🚀 Primo Piano";
            this.chkFormaZ.Location = new System.Drawing.Point(240, 176);
            this.chkFormaZ.Size = new System.Drawing.Size(155, 21);
            this.chkFormaZ.Font = new System.Drawing.Font("Segoe UI", 8.25f);
            this.chkFormaZ.CheckedChanged += new System.EventHandler(this.ControlloProprietà_Changed);

            // Inizializza visibilità forme/logo nascosti di default
            this.lblFormaTipo.Visible = false; this.cmbFormaTipo.Visible = false;
            this.lblFormaSpess.Visible = false; this.nudFormaSpess.Visible = false;
            this.chkFormaFill.Visible = false;

            // Separatore 2
            var sep2 = new System.Windows.Forms.Label();
            sep2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            sep2.Location = new System.Drawing.Point(8, 230); sep2.Size = new System.Drawing.Size(388, 2);

            var lblPos = new System.Windows.Forms.Label();
            lblPos.Text = "Posizione e Dimensioni (millimetri)"; lblPos.Location = new System.Drawing.Point(8, 235);
            lblPos.Size = new System.Drawing.Size(388, 15); lblPos.Font = new System.Drawing.Font("Segoe UI", 7.5f, System.Drawing.FontStyle.Italic);
            lblPos.ForeColor = System.Drawing.Color.FromArgb(71, 85, 105);

            Lbl(this.lblPosX, "X ←→:", 8, 254, 38);
            this.nudPosX.Location = new System.Drawing.Point(48, 252); this.nudPosX.Size = new System.Drawing.Size(48, 21);
            this.nudPosX.Maximum = 500; this.nudPosX.Font = new System.Drawing.Font("Segoe UI", 8.5f);
            this.nudPosX.ValueChanged += new System.EventHandler(this.NudPos_ValueChanged);

            Lbl(this.lblPosY, "Y ↑↓:", 102, 254, 36);
            this.nudPosY.Location = new System.Drawing.Point(140, 252); this.nudPosY.Size = new System.Drawing.Size(48, 21);
            this.nudPosY.Maximum = 500; this.nudPosY.Font = new System.Drawing.Font("Segoe UI", 8.5f);
            this.nudPosY.ValueChanged += new System.EventHandler(this.NudPos_ValueChanged);

            Lbl(this.lblPosW, "Largh:", 196, 254, 44);
            this.nudPosW.Location = new System.Drawing.Point(242, 252); this.nudPosW.Size = new System.Drawing.Size(50, 21);
            this.nudPosW.Minimum = 3; this.nudPosW.Maximum = 500; this.nudPosW.Value = 40;
            this.nudPosW.Font = new System.Drawing.Font("Segoe UI", 8.5f);
            this.nudPosW.ValueChanged += new System.EventHandler(this.NudPos_ValueChanged);

            Lbl(this.lblPosH, "Altez:", 300, 254, 42);
            this.nudPosH.Location = new System.Drawing.Point(344, 252); this.nudPosH.Size = new System.Drawing.Size(50, 21);
            this.nudPosH.Minimum = 2; this.nudPosH.Maximum = 500; this.nudPosH.Value = 8;
            this.nudPosH.Font = new System.Drawing.Font("Segoe UI", 8.5f);
            this.nudPosH.ValueChanged += new System.EventHandler(this.NudPos_ValueChanged);

            this.lblInfo.Text = "💡 Clicca o trascina i blocchi nell'anteprima a destra.";
            this.lblInfo.Location = new System.Drawing.Point(8, 280);
            this.lblInfo.Size = new System.Drawing.Size(388, 18);
            this.lblInfo.Font = new System.Drawing.Font("Segoe UI", 7.5f, System.Drawing.FontStyle.Italic);
            this.lblInfo.ForeColor = System.Drawing.Color.FromArgb(67, 56, 202);

            this.grpElemento.Controls.AddRange(new System.Windows.Forms.Control[] {
                lstCampi, chkVisibile, sep, lblTestoLibero, txtTestoLibero, lblFont, cmbFont, lblFontSz, nudFontSz,
                chkBold, cmbAllinea, sep2, lblPos,
                lblPosX, nudPosX, lblPosY, nudPosY,
                lblPosW, nudPosW, lblPosH, nudPosH,
                lblFormaTipo, cmbFormaTipo, lblFormaColore, cmbFormaColore,
                lblFormaSpess, nudFormaSpess, chkFormaFill, chkFormaZ,
                btnCaricaLogo, lblLogo, txtLogoPath, lblInfo });

            // Bottoni SALVA / ESCI in un pannello compatto
            this.pnlPropBottom.Location = new System.Drawing.Point(4, 458);
            this.pnlPropBottom.Size = new System.Drawing.Size(404, 42);
            this.pnlPropBottom.BackColor = System.Drawing.Color.Transparent;

            this.btnSalva.Location = new System.Drawing.Point(0, 2);
            this.btnSalva.Size = new System.Drawing.Size(198, 38);
            this.btnSalva.Click += new System.EventHandler(this.btnSalva_Click);

            this.btnEsci.Location = new System.Drawing.Point(206, 2);
            this.btnEsci.Size = new System.Drawing.Size(198, 38);
            this.btnEsci.Click += new System.EventHandler(this.btnEsci_Click);

            this.pnlPropBottom.Controls.AddRange(new System.Windows.Forms.Control[] { btnSalva, btnEsci });

            this.pnlProp.Controls.AddRange(new System.Windows.Forms.Control[] { grpFormato, grpElemento, pnlPropBottom });

            // ====================================================
            // COL 3 — CANVAS ANTEPRIMA
            // ====================================================
            this.pnlCanvas.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlCanvas.BackColor = System.Drawing.Color.FromArgb(203, 213, 225);
            this.pnlCanvas.Padding = new System.Windows.Forms.Padding(4);
            this.pnlCanvas.AutoScroll = true;

            this.pnlCanvasTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlCanvasTop.Height = 28;
            this.pnlCanvasTop.BackColor = System.Drawing.Color.FromArgb(51, 65, 85);

            this.lblCanvasTip.Text = "ANTEPRIMA LIVE — Clicca e trascina i blocchi colorati";
            this.lblCanvasTip.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblCanvasTip.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblCanvasTip.ForeColor = System.Drawing.Color.White;
            this.lblCanvasTip.Font = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold);

            this.btnZoomOut.Text = "−";
            this.btnZoomOut.Dock = System.Windows.Forms.DockStyle.Left;
            this.btnZoomOut.Width = 32;
            this.btnZoomOut.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnZoomOut.FlatAppearance.BorderSize = 0;
            this.btnZoomOut.BackColor = System.Drawing.Color.FromArgb(37, 99, 235);
            this.btnZoomOut.ForeColor = System.Drawing.Color.White;
            this.btnZoomOut.Font = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold);
            this.btnZoomOut.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnZoomOut.Click += new System.EventHandler(this.btnZoomOut_Click);

            this.lblZoom.Text = "200%";
            this.lblZoom.Dock = System.Windows.Forms.DockStyle.Left;
            this.lblZoom.Width = 50;
            this.lblZoom.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblZoom.ForeColor = System.Drawing.Color.White;
            this.lblZoom.Font = new System.Drawing.Font("Segoe UI", 8.5f, System.Drawing.FontStyle.Bold);

            this.btnZoomIn.Text = "+";
            this.btnZoomIn.Dock = System.Windows.Forms.DockStyle.Left;
            this.btnZoomIn.Width = 32;
            this.btnZoomIn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnZoomIn.FlatAppearance.BorderSize = 0;
            this.btnZoomIn.BackColor = System.Drawing.Color.FromArgb(37, 99, 235);
            this.btnZoomIn.ForeColor = System.Drawing.Color.White;
            this.btnZoomIn.Font = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold);
            this.btnZoomIn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnZoomIn.Click += new System.EventHandler(this.btnZoomIn_Click);

            this.pnlCanvasTop.Controls.Add(this.btnZoomOut);
            this.pnlCanvasTop.Controls.Add(this.lblZoom);
            this.pnlCanvasTop.Controls.Add(this.btnZoomIn);
            this.pnlCanvasTop.Controls.Add(this.lblCanvasTip);

            this.picCanvas.Location = new System.Drawing.Point(10, 36);
            this.picCanvas.BackColor = System.Drawing.Color.White;
            this.picCanvas.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.picCanvas.Cursor = System.Windows.Forms.Cursors.Default;
            this.picCanvas.Name = "picCanvas";
            this.picCanvas.Paint += new System.Windows.Forms.PaintEventHandler(this.picCanvas_Paint);
            this.picCanvas.MouseDown += new System.Windows.Forms.MouseEventHandler(this.picCanvas_MouseDown);
            this.picCanvas.MouseMove += new System.Windows.Forms.MouseEventHandler(this.picCanvas_MouseMove);
            this.picCanvas.MouseUp += new System.Windows.Forms.MouseEventHandler(this.picCanvas_MouseUp);

            this.pnlCanvas.Controls.Add(this.picCanvas);
            this.pnlCanvas.Controls.Add(this.pnlCanvasTop);

            // ====================================================
            // FORM
            // ====================================================
            this.AutoScaleDimensions = new System.Drawing.SizeF(6f, 13f);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1200, 720);
            this.MinimumSize = new System.Drawing.Size(1050, 650);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Controls.Add(this.tblMain);
            this.Name = "frmUtyEtiFormati";
            this.Text = "Designer Etichette da Scaffale";
            this.Load += new System.EventHandler(this.frmUtyEtiFormati_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmUtyEtiFormati_FormClosing);
            this.KeyPreview = true;
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyEtiFormati_KeyDown);

            ((System.ComponentModel.ISupportInitialize)(this.dgvFormati)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudLargh)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudAlt)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudCols)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudRighe)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudMarginTop)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudMarginLeft)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudFontSz)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudPosX)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudPosY)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudPosW)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudPosH)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudFormaSpess)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picCanvas)).EndInit();
            this.ResumeLayout(false);
        }

        // Helper: crea label con font standard
        private void Lbl(System.Windows.Forms.Label l, string t, int x, int y, int w)
        {
            l.Text = t;
            l.Location = new System.Drawing.Point(x, y);
            l.Size = new System.Drawing.Size(w, 18);
            l.Font = new System.Drawing.Font("Segoe UI", 8.5f);
            l.ForeColor = System.Drawing.Color.FromArgb(30, 41, 59);
            l.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
        }

        // Campi privati
        private System.Windows.Forms.TableLayoutPanel tblMain;
        private System.Windows.Forms.Panel pnlLista, pnlBtnLista, pnlProp, pnlPropBottom, pnlCanvas;
        private APOffice.APDataGridView dgvFormati;
        private System.Windows.Forms.Button btnNuovo, btnDuplica, btnElimina, btnSalva, btnEsci;
        private System.Windows.Forms.GroupBox grpFormato, grpElemento;
        private System.Windows.Forms.Label lblCodice, lblDescrizione, lblLargh, lblAlt, lblCols, lblRighe;
        private System.Windows.Forms.TextBox txtCodice, txtDescrizione;
        private System.Windows.Forms.NumericUpDown nudLargh, nudAlt, nudCols, nudRighe;
        private System.Windows.Forms.NumericUpDown nudMarginTop, nudMarginLeft;
        private System.Windows.Forms.RadioButton radOrientV, radOrientH;
        private System.Windows.Forms.GroupBox grpOrient;
        private System.Windows.Forms.Label lblMarginTop, lblMarginLeft;
        private System.Windows.Forms.Label lblEtiSta;
        private System.Windows.Forms.ComboBox cmbEtiSta;
        private System.Windows.Forms.ListBox lstCampi;
        private System.Windows.Forms.CheckBox chkVisibile, chkBold;
        private System.Windows.Forms.ComboBox cmbAllinea;
        private System.Windows.Forms.Label lblFont, lblFontSz, lblPosX, lblPosY, lblPosW, lblPosH, lblInfo, lblCanvasTip;
        private System.Windows.Forms.Panel pnlCanvasTop;
        private System.Windows.Forms.Button btnZoomIn, btnZoomOut;
        private System.Windows.Forms.Label lblZoom;
        private System.Windows.Forms.ComboBox cmbFont;
        private System.Windows.Forms.NumericUpDown nudFontSz, nudPosX, nudPosY, nudPosW, nudPosH;
        private System.Windows.Forms.Button btnCaricaLogo;
        private System.Windows.Forms.Label lblLogo;
        private System.Windows.Forms.TextBox txtLogoPath;
        private System.Windows.Forms.Label lblTestoLibero;
        private System.Windows.Forms.TextBox txtTestoLibero;
        private System.Windows.Forms.Label lblFormaTipo, lblFormaColore, lblFormaSpess;
        private System.Windows.Forms.ComboBox cmbFormaTipo, cmbFormaColore;
        private System.Windows.Forms.NumericUpDown nudFormaSpess;
        private System.Windows.Forms.CheckBox chkFormaFill, chkFormaZ;
        private System.Windows.Forms.PictureBox picCanvas;
    }
}
