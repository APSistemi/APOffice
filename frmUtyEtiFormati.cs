using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using BarcodeLib;

namespace APOffice
{
    // =========================================================
    // Descrizione di ogni campo salvato nel DB:
    //   eti_s003 = Intestazione / Rag. Sociale
    //   eti_s004 = Descrizione Articolo
    //   eti_s005 = Info / Lotto / Ingredienti
    //   eti_s006 = Cod. Interno / Data confezionamento
    //   eti_s007 = Slogan / Data Promo
    //   eti_s008 = Prezzo al Kg / Litro
    //   eti_s009 = Avvertenze / Allergeni
    //   eti_s010 = Prezzo VECCHIO barrato (promo)
    //   eti_s011 = PREZZO FINALE (campo grande)
    //   eti_s012 = (riservato logo)
    //   eti_s013 = Codice a barre EAN
    //
    // Formato stringa campo: "Y,X,Larghezza,Altezza,Allineamento;FontNome:Dimensione:Stile"
    //   Allineamento: 0=Sinistra  1=Destra  2=Centro
    //   Stile: Regular / B / I
    // =========================================================

    public partial class frmUtyEtiFormati : Form
    {
        private clsFuncs _clsFun = new clsFuncs();
        public string _strConSql = "";

        private DataTable _tabFmt;
        private SqlDataAdapter _da;
        private int _selectedRowIndex = -1;
        private bool _suppressEvents = false;

        // Canvas scale factor: pixel per mm
        private float _pageScale = 2.0f;

        // Drag state
        private int _dragCampoIdx = -1;
        private Point _dragStartMouse;
        private FieldDef _dragStartState;
        private bool _isDragging = false;

        // Resize state (angoli)
        private enum ResizeHandle { None, TopLeft, TopRight, BottomLeft, BottomRight }
        private ResizeHandle _resizeHandle = ResizeHandle.None;
        private bool _isResizing = false;

        // Selezionato
        private int _selIdx = -1;

        // Definizione dei campi con metadati leggibili
        private static readonly string[] CAMPI_NOMI = {
            "Ragione Sociale", "Descrizione Articolo", "Info / Origine addizionale", "Codice Interno",
            "Slogan / Promo", "Prezzo al Kg", "Sconto Euro", "Prezzo Vecchio", "Prezzo Finale",
            "Logo Aziendale", "Barcode EAN", "Origine Articolo", "Calibro Articolo", "Tasto Bilancia (PLU)",
            "Immagine Articolo", "Unità di Misura", "Grammatura/Peso", "Pezzi", "Merceologia (Cat.)", "Reparto",
            "Data di Stampa", "Periodo Promo", "Prezzo Lt/Kg",
            "Testo Libero 1", "Testo Libero 2", "Testo Libero 3", "Testo Libero 4",
            "Testo Libero 5", "Testo Libero 6", "Testo Libero 7", "Testo Libero 8",
            "Rettangolo / Linea 1", "Rettangolo / Linea 2", "Rettangolo / Linea 3", "Rettangolo / Linea 4",
            "Rettangolo / Linea 5", "Rettangolo / Linea 6", "Rettangolo / Linea 7", "Rettangolo / Linea 8"
        };
        private static readonly string[] CAMPI_DB = {
            "eti_s003", "eti_s004", "eti_s005", "eti_s006", "eti_s007", "eti_s008", "eti_s009", "eti_s010", "eti_s011",
            "eti_s012", "eti_s013", "eti_s014", "eti_s015", "eti_s016", "eti_s017", "eti_s018", "eti_s019", "eti_s020", "eti_s021", "eti_s022",
            "eti_s023", "eti_s024", "eti_s025",
            "eti_t001", "eti_t002", "eti_t003", "eti_t004", "eti_t005", "eti_t006", "eti_t007", "eti_t008",
            "eti_h001", "eti_h002", "eti_h003", "eti_h004", "eti_h005", "eti_h006", "eti_h007", "eti_h008"
        };
        private static readonly Color[] CAMPI_COLORI = {
            Color.FromArgb(120,190,255), Color.FromArgb(130,200,130),
            Color.FromArgb(255,200,100), Color.FromArgb(200,160,255),
            Color.FromArgb(255,170,170), Color.FromArgb(100,220,200),
            Color.FromArgb(200,200,200), Color.FromArgb(255,100,100),
            Color.FromArgb(255,80,20),   Color.FromArgb(180,180,180), Color.FromArgb(60,60,60),
            Color.FromArgb(0,150,136),   Color.FromArgb(121,85,72),   Color.FromArgb(63,81,181),
            Color.FromArgb(255,165,0),   Color.FromArgb(100,149,237), Color.FromArgb(152,251,152),
            Color.FromArgb(255,192,203), Color.FromArgb(173,216,230), Color.FromArgb(255,228,196),
            Color.FromArgb(200,100,250), Color.FromArgb(100,255,100), Color.FromArgb(250,250,100),
            // Testi Liberi 1..8 (colori vivaci)
            Color.FromArgb(245,158,11),  Color.FromArgb(236,72,153),  Color.FromArgb(14,165,233),  Color.FromArgb(132,204,22),
            Color.FromArgb(168,85,247),  Color.FromArgb(20,184,166),  Color.FromArgb(249,115,22),  Color.FromArgb(239,68,68),
            // Forme geometriche h001..h008 (colore neutro grigio)
            Color.FromArgb(180,180,180), Color.FromArgb(160,160,160),
            Color.FromArgb(140,140,140), Color.FromArgb(120,120,120),
            Color.FromArgb(100,100,100), Color.FromArgb(80,80,80),
            Color.FromArgb(60,60,60),    Color.FromArgb(40,40,40)
        };

        // Dati runtime per il formato corrente
        private List<FieldDef> _campi = new List<FieldDef>();

        public frmUtyEtiFormati()
        {
            InitializeComponent();
        }

        // ----------------------------------------------------------------
        // CARICAMENTO & MODERN UI
        // ----------------------------------------------------------------
        private void frmUtyEtiFormati_Load(object sender, EventArgs e)
        {
            if (_strConSql == "") _strConSql = _clsFun.ConSql("");
            try { new clsDbMigration().CheckDb(_strConSql); } catch { }
            ApplyModernUi();
            clsUiIcons.RestoreFormBounds(this);
            if (this.Width < 1050) this.Width = 1200;
            if (this.Height < 680) this.Height = 720;
            LoadFormati();
            clsUiIcons.RestoreGridColumnWidths(dgvFormati, "frmUtyEtiFormati_dgv");
        }

        private void frmUtyEtiFormati_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgvFormati, "frmUtyEtiFormati_dgv");
        }

        private void ApplyModernUi()
        {
            try
            {
                this.BackColor = Color.FromArgb(243, 244, 246);

                // Bottoni Lista Formati:
                clsUiIcons.StyleStatButton(btnNuovo, "Nuovo", "plus", 12,
                    Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208), Color.FromArgb(52, 211, 153),
                    Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));

                clsUiIcons.StyleStatButton(btnDuplica, "Copia", "copy", 12,
                    Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254), Color.FromArgb(96, 165, 250),
                    Color.FromArgb(30, 58, 138), Color.FromArgb(37, 99, 235));

                clsUiIcons.StyleStatButton(btnElimina, "Canc", "trash", 12,
                    Color.FromArgb(254, 242, 242), Color.FromArgb(254, 202, 202), Color.FromArgb(248, 113, 113),
                    Color.FromArgb(127, 29, 29), Color.FromArgb(220, 38, 38));

                // Bottoni Azione Principali (F5 Salva ed Esci):
                clsUiIcons.StyleStatButton(btnSalva, "F5  SALVA FORMATO", "save", 16,
                    Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208), Color.FromArgb(52, 211, 153),
                    Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));

                clsUiIcons.StyleStatButton(btnEsci, "Esci", "exit", 16,
                    Color.FromArgb(254, 242, 242), Color.FromArgb(254, 202, 202), Color.FromArgb(248, 113, 113),
                    Color.FromArgb(127, 29, 29), Color.FromArgb(220, 38, 38));

                clsUiIcons.StyleStatButton(btnCaricaLogo, "Sfoglia...", "search", 12,
                    Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254), Color.FromArgb(96, 165, 250),
                    Color.FromArgb(30, 58, 138), Color.FromArgb(37, 99, 235));

                clsUiIcons.StyleDataGridView(dgvFormati);
            }
            catch { }
        }

        private void frmUtyEtiFormati_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5) { btnSalva_Click(null, null); e.Handled = true; }
            else if (e.KeyCode == Keys.Escape) { this.Close(); e.Handled = true; }
        }

        private void LoadFormati()
        {
            string sql = "SELECT * FROM TabEtiFormatiLayout WHERE eti_ann=0 ORDER BY eti_cod";
            SqlConnection cn = new SqlConnection(_strConSql);
            _da = new SqlDataAdapter(sql, cn);
            new SqlCommandBuilder(_da);
            _tabFmt = new DataTable();
            _da.Fill(_tabFmt);

            // Rilassa il limite in RAM nel caso la migrazione V004/V008 sia appena passata o non in sync
            if (_tabFmt.Columns.Contains("eti_s012") && _tabFmt.Columns["eti_s012"].MaxLength < 1000)
                _tabFmt.Columns["eti_s012"].MaxLength = 1000;
            if (_tabFmt.Columns.Contains("eti_s017") && _tabFmt.Columns["eti_s017"].MaxLength < 1000)
                _tabFmt.Columns["eti_s017"].MaxLength = 1000;
            for (int i = 1; i <= 8; i++)
            {
                string tCol = "eti_t00" + i;
                if (_tabFmt.Columns.Contains(tCol) && _tabFmt.Columns[tCol].MaxLength < 500)
                    _tabFmt.Columns[tCol].MaxLength = 500;
            }


            dgvFormati.AutoGenerateColumns = false;
            dgvFormati.Columns.Clear();
            var colCod = new DataGridViewTextBoxColumn(); colCod.DataPropertyName = "eti_cod"; colCod.HeaderText = "Cod"; colCod.Width = 45;
            var colDes = new DataGridViewTextBoxColumn(); colDes.DataPropertyName = "eti_des"; colDes.HeaderText = "Nome Formato"; colDes.Width = 180; colDes.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvFormati.Columns.AddRange(new DataGridViewColumn[] { colCod, colDes });
            dgvFormati.DataSource = _tabFmt;
            dgvFormati.SelectionChanged += DgvFormati_SelectionChanged;

            // Lista campi nella ListBox
            lstCampi.Items.Clear();
            foreach (string n in CAMPI_NOMI) lstCampi.Items.Add(n);

            // Popola Dropdown Stato Post-Stampa
            cmbEtiSta.Items.Clear();
            cmbEtiSta.Items.Add("S - Segnala come stampata (con richiesta)");
            cmbEtiSta.Items.Add("N - Mantieni da stampare");
            cmbEtiSta.Items.Add("E - Cancella variazioni stampate");
            cmbEtiSta.SelectedIndex = 0;

            if (_tabFmt.Rows.Count > 0)
            {
                dgvFormati.Rows[0].Selected = true;
                _selectedRowIndex = 0;
                CaricaDettaglioFormato(_tabFmt.Rows[0]);
            }
        }

        // ----------------------------------------------------------------
        // SELEZIONE FORMATO
        // ----------------------------------------------------------------
        private void DgvFormati_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvFormati.SelectedRows.Count == 0) return;
            _selectedRowIndex = dgvFormati.SelectedRows[0].Index;
            CaricaDettaglioFormato(_tabFmt.Rows[_selectedRowIndex]);
        }

        private void CaricaDettaglioFormato(DataRow r)
        {
            _suppressEvents = true;
            txtCodice.Text = r["eti_cod"].ToString();
            txtDescrizione.Text = r["eti_des"].ToString();

            // Stato post-stampa
            string sta = "S";
            if (r.Table.Columns.Contains("eti_sta") && !DBNull.Value.Equals(r["eti_sta"]))
                sta = r["eti_sta"].ToString().Trim().ToUpper();
            if (sta == "N") cmbEtiSta.SelectedIndex = 1;
            else if (sta == "E") cmbEtiSta.SelectedIndex = 2;
            else cmbEtiSta.SelectedIndex = 0;

            // Dimensioni etichetta
            string dim = r["eti_dim"].ToString();
            string[] dims = dim.Split(',');
            if (dims.Length >= 2)
            {
                decimal w, h;
                if (decimal.TryParse(dims[0], out w)) nudLargh.Value = Math.Min(nudLargh.Maximum, Math.Max(nudLargh.Minimum, w));
                if (decimal.TryParse(dims[1], out h)) nudAlt.Value = Math.Min(nudAlt.Maximum, Math.Max(nudAlt.Minimum, h));
            }

            // Colonne/Righe per foglio + margini + orientamento
            string pag = r["eti_pag"].ToString();
            string[] pags = pag.Split(',');
            if (pags.Length >= 2)
            {
                decimal cols, rows;
                if (decimal.TryParse(pags[0], out cols)) nudCols.Value = Math.Min(nudCols.Maximum, Math.Max(1, cols));
                if (decimal.TryParse(pags[1], out rows)) nudRighe.Value = Math.Min(nudRighe.Maximum, Math.Max(1, rows));
            }
            // Margine superiore (pos 2)
            if (pags.Length >= 3) { decimal mg; if (decimal.TryParse(pags[2], out mg)) nudMarginTop.Value = Math.Min(nudMarginTop.Maximum, Math.Max(0, mg)); else nudMarginTop.Value = 0; }
            // Margine sinistro (pos 3)
            if (pags.Length >= 4) { decimal mg; if (decimal.TryParse(pags[3], out mg)) nudMarginLeft.Value = Math.Min(nudMarginLeft.Maximum, Math.Max(0, mg)); else nudMarginLeft.Value = 0; }
            // Orientamento (pos 6 o 7)
            string orient = "V";
            if (pags.Length >= 7) orient = pags[6].Trim().ToUpper();
            radOrientV.Checked = (orient != "H");
            radOrientH.Checked = (orient == "H");
            // Parsing campi
            _campi.Clear();
            for (int i = 0; i < CAMPI_DB.Length; i++)
            {
                string raw = r.Table.Columns.Contains(CAMPI_DB[i]) ? r[CAMPI_DB[i]].ToString() : "";
                _campi.Add(FieldDef.Parse(raw, i));
            }

            _suppressEvents = false;

            if (lstCampi.SelectedIndex < 0) lstCampi.SelectedIndex = 8; // default: Prezzo Finale
            else AggiornaProprieta();

            AggiornaCanvas();
        }

        // ----------------------------------------------------------------
        // SELEZIONE CAMPO nella ListBox
        // ----------------------------------------------------------------
        private void lstCampi_SelectedIndexChanged(object sender, EventArgs e)
        {
            AggiornaProprieta();
        }

        private void AggiornaProprieta()
        {
            int idx = lstCampi.SelectedIndex;
            if (idx < 0 || idx >= _campi.Count) return;
            _selIdx = idx;
            FieldDef f = _campi[idx];
            string field = CAMPI_DB[idx];

            _suppressEvents = true;
            chkVisibile.Checked = f.Visibile;
            cmbFont.SelectedItem = f.FontNome;
            if (cmbFont.SelectedIndex < 0) cmbFont.SelectedIndex = 0;
            nudFontSz.Value = Math.Min(nudFontSz.Maximum, Math.Max(nudFontSz.Minimum, (decimal)f.FontSize));
            chkBold.Checked = f.Bold;
            if (f.Align >= 0 && f.Align <= 2) cmbAllinea.SelectedIndex = f.Align; else cmbAllinea.SelectedIndex = 0;
            nudPosX.Value = Math.Min(nudPosX.Maximum, (decimal)f.X);
            nudPosY.Value = Math.Min(nudPosY.Maximum, (decimal)f.Y);
            nudPosW.Value = Math.Min(nudPosW.Maximum, Math.Max(nudPosW.Minimum, (decimal)f.W));
            nudPosH.Value = Math.Min(nudPosH.Maximum, Math.Max(nudPosH.Minimum, (decimal)f.H));

            bool isImageField = (field == "eti_s012" || field == "eti_s017");
            bool isShape = field.StartsWith("eti_h0");
            bool isFreeText = field.StartsWith("eti_t0");

            // Testo Libero
            lblTestoLibero.Visible = isFreeText;
            txtTestoLibero.Visible = isFreeText;
            txtTestoLibero.Text = f.TestoLibero ?? "";

            // Logo / Immagine
            btnCaricaLogo.Visible = isImageField;
            lblLogo.Visible = isImageField;
            txtLogoPath.Visible = isImageField;
            txtLogoPath.Text = f.FontNome;

            // Font & Stile (visibile per campi testo standard e testo libero)
            bool showFont = (!isShape && !isImageField);
            lblFont.Visible = showFont; cmbFont.Visible = showFont;
            lblFontSz.Visible = showFont; nudFontSz.Visible = showFont;
            chkBold.Visible = showFont; cmbAllinea.Visible = showFont;

            // Forme geometriche
            lblFormaTipo.Visible = isShape; cmbFormaTipo.Visible = isShape;
            if (isShape) cmbFormaTipo.SelectedIndex = Math.Max(0, f.ShapeType - 1);

            lblFormaSpess.Visible = isShape; nudFormaSpess.Visible = isShape;
            if (isShape) nudFormaSpess.Value = Math.Max(nudFormaSpess.Minimum, Math.Min(nudFormaSpess.Maximum, (decimal)f.ShapeThick));
            chkFormaFill.Visible = isShape;
            if (isShape) chkFormaFill.Checked = f.ShapeFill;

            // Colore (visibile per campi testo, testo libero e forme)
            bool showColore = (!isImageField);
            lblFormaColore.Visible = showColore; cmbFormaColore.Visible = showColore;
            if (showColore)
            {
                int[] argo = { Color.Black.ToArgb(), Color.Red.ToArgb(), Color.Blue.ToArgb(), Color.Green.ToArgb(), Color.Gray.ToArgb(), Color.White.ToArgb(), Color.Yellow.ToArgb() };
                int tCol = Array.IndexOf(argo, f.ShapeColor);
                cmbFormaColore.SelectedIndex = (tCol >= 0) ? tCol : 0;
            }

            chkFormaZ.Visible = true;

            // Disposizione posizioni dinamiche controlli
            if (isFreeText)
            {
                // Row 1 (Y=148): Testo Libero Box
                lblTestoLibero.Location = new Point(8, 150);
                txtTestoLibero.Location = new Point(46, 148);
                txtTestoLibero.Size = new Size(350, 21);

                // Row 2 (Y=176): Font, Pt, Bold
                lblFont.Location = new Point(8, 178);
                cmbFont.Location = new Point(44, 176);
                cmbFont.Size = new Size(120, 21);
                lblFontSz.Location = new Point(170, 178);
                nudFontSz.Location = new Point(194, 176);
                chkBold.Location = new Point(250, 176);

                // Row 3 (Y=204): Allineamento, Colore, Z-Index
                cmbAllinea.Location = new Point(8, 204);
                cmbAllinea.Size = new Size(90, 21);
                lblFormaColore.Location = new Point(106, 206);
                cmbFormaColore.Location = new Point(152, 204);
                cmbFormaColore.Size = new Size(82, 21);
                chkFormaZ.Text = "🚀 Primo Piano";
                chkFormaZ.Checked = (f.ZIndex == 2);
                chkFormaZ.Location = new Point(240, 204);
                chkFormaZ.Size = new Size(155, 21);
            }
            else if (!isShape && !isImageField)
            {
                // Standard Data Text Fields
                // Row 1 (Y=148): Font, Pt, Bold
                lblFont.Location = new Point(8, 150);
                cmbFont.Location = new Point(44, 148);
                cmbFont.Size = new Size(120, 21);
                lblFontSz.Location = new Point(170, 150);
                nudFontSz.Location = new Point(194, 148);
                chkBold.Location = new Point(250, 148);

                // Row 2 (Y=176): Allineamento, Colore, Z-Index
                cmbAllinea.Location = new Point(8, 176);
                cmbAllinea.Size = new Size(90, 21);
                lblFormaColore.Location = new Point(106, 178);
                cmbFormaColore.Location = new Point(152, 176);
                cmbFormaColore.Size = new Size(82, 21);
                chkFormaZ.Text = "🚀 Primo Piano";
                chkFormaZ.Checked = (f.ZIndex == 2);
                chkFormaZ.Location = new Point(240, 176);
                chkFormaZ.Size = new Size(155, 21);
            }
            else if (isShape)
            {
                // Shapes
                // Row 1 (Y=148): Tipo, Spessore, Riempito
                lblFormaTipo.Location = new Point(8, 150);
                cmbFormaTipo.Location = new Point(44, 148);
                cmbFormaTipo.Size = new Size(115, 21);
                lblFormaSpess.Location = new Point(166, 150);
                nudFormaSpess.Location = new Point(210, 148);
                chkFormaFill.Location = new Point(264, 148);

                // Row 2 (Y=176): Colore, Z-Index
                lblFormaColore.Location = new Point(8, 178);
                cmbFormaColore.Location = new Point(56, 176);
                cmbFormaColore.Size = new Size(85, 21);
                chkFormaZ.Text = "🖼 Metti Dietro (sotto)";
                chkFormaZ.Checked = (f.ZIndex == 0);
                chkFormaZ.Location = new Point(150, 176);
                chkFormaZ.Size = new Size(240, 21);
            }
            else
            {
                // Image / Logo
                // Row 1 (Y=148): Sfoglia, File, Path
                btnCaricaLogo.Location = new Point(8, 146);
                btnCaricaLogo.Size = new Size(88, 26);
                lblLogo.Location = new Point(102, 150);
                txtLogoPath.Location = new Point(134, 148);
                txtLogoPath.Size = new Size(262, 21);

                // Row 2 (Y=176): Z-Index
                chkFormaZ.Text = "🖼 Metti Dietro (sotto gli altri)";
                chkFormaZ.Checked = (f.ZIndex == 0);
                chkFormaZ.Location = new Point(8, 176);
                chkFormaZ.Size = new Size(240, 21);
            }

            _suppressEvents = false;

            // Aggiorna abilitazione controlli
            bool abi = chkVisibile.Checked;
            txtTestoLibero.Enabled = abi;
            cmbFont.Enabled = chkBold.Enabled = cmbAllinea.Enabled = abi;
            nudFontSz.Enabled = abi;
            cmbFormaColore.Enabled = cmbFormaTipo.Enabled = abi;
            nudFormaSpess.Enabled = chkFormaFill.Enabled = chkFormaZ.Enabled = abi;
            btnCaricaLogo.Enabled = abi;
            nudPosX.Enabled = nudPosY.Enabled = abi;
            nudPosW.Enabled = nudPosH.Enabled = abi;

            AggiornaCanvas();
        }

        private void btnCaricaLogo_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Immagini (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtLogoPath.Text = ofd.FileName;
                    SalvaProprietaCampo();
                }
            }
        }

        private void SalvaProprietaCampo()
        {
            int idx = lstCampi.SelectedIndex;
            if (idx < 0 || idx >= _campi.Count) return;
            FieldDef f = _campi[idx];
            f.Visibile = chkVisibile.Checked;

            string field = CAMPI_DB[idx];
            bool isImageField = (field == "eti_s012" || field == "eti_s017");
            bool isShape = field.StartsWith("eti_h0");
            bool isFreeText = field.StartsWith("eti_t0");

            if (isFreeText)
            {
                f.TestoLibero = txtTestoLibero.Text;
                f.FontNome = cmbFont.SelectedItem != null ? cmbFont.SelectedItem.ToString() : "Arial";
                f.ZIndex = chkFormaZ.Checked ? 2 : 1; // 2=Sopra Tutto, 1=Normale
                int[] argo = { Color.Black.ToArgb(), Color.Red.ToArgb(), Color.Blue.ToArgb(), Color.Green.ToArgb(), Color.Gray.ToArgb(), Color.White.ToArgb(), Color.Yellow.ToArgb() };
                if (cmbFormaColore.SelectedIndex >= 0) f.ShapeColor = argo[cmbFormaColore.SelectedIndex];
            }
            else if (!isShape && !isImageField)
            {
                f.FontNome = cmbFont.SelectedItem != null ? cmbFont.SelectedItem.ToString() : "Arial";
                f.ZIndex = chkFormaZ.Checked ? 2 : 1; // 2=Sopra Tutto, 1=Normale
                int[] argo = { Color.Black.ToArgb(), Color.Red.ToArgb(), Color.Blue.ToArgb(), Color.Green.ToArgb(), Color.Gray.ToArgb(), Color.White.ToArgb(), Color.Yellow.ToArgb() };
                if (cmbFormaColore.SelectedIndex >= 0) f.ShapeColor = argo[cmbFormaColore.SelectedIndex];
            }
            else if (isImageField)
            {
                f.FontNome = txtLogoPath.Text;
                f.ZIndex = chkFormaZ.Checked ? 0 : 1; // 0=Dietro, 1=Fronte
            }
            else // Shape
            {
                int[] argo = { Color.Black.ToArgb(), Color.Red.ToArgb(), Color.Blue.ToArgb(), Color.Green.ToArgb(), Color.Gray.ToArgb(), Color.White.ToArgb(), Color.Yellow.ToArgb() };
                if (cmbFormaColore.SelectedIndex >= 0) f.ShapeColor = argo[cmbFormaColore.SelectedIndex];

                f.ShapeType = cmbFormaTipo.SelectedIndex + 1;
                f.ShapeThick = (int)nudFormaSpess.Value;
                f.ShapeFill = chkFormaFill.Checked;
                f.ZIndex = chkFormaZ.Checked ? 0 : 1;
            }

            f.FontSize = (double)nudFontSz.Value;
            f.Bold = chkBold.Checked;
            f.Align = cmbAllinea.SelectedIndex >= 0 ? cmbAllinea.SelectedIndex : 0;
            f.X = (int)nudPosX.Value;
            f.Y = (int)nudPosY.Value;
            f.W = (int)nudPosW.Value;
            f.H = (int)nudPosH.Value;

            AggiornaCanvas();
        }

        // ----------------------------------------------------------------
        // EVENTI MODIFICA CONTROLS PROPRIETA'
        // ----------------------------------------------------------------
        private void ControlloProprietà_Changed(object sender, EventArgs e)
        {
            if (_suppressEvents) return;
            SalvaProprietaCampo();
        }

        private void NudPos_ValueChanged(object sender, EventArgs e)
        {
            if (_suppressEvents) return;
            SalvaProprietaCampo();
        }

        // ----------------------------------------------------------------
        // EVENTS Dimensioni formato
        // ----------------------------------------------------------------
        private void NudDim_ValueChanged(object sender, EventArgs e)
        {
            if (!_suppressEvents) AggiornaCanvas();
        }

        // ----------------------------------------------------------------
        // CANVAS GDI+ — disegno live
        // ----------------------------------------------------------------
        // Dimensioni foglio A4 in mm
        private const int A4_W_MM = 210;
        private const int A4_H_MM = 297;

        private void AggiornaCanvas()
        {
            bool isH = radOrientH != null && radOrientH.Checked;
            int pageW = isH ? A4_H_MM : A4_W_MM;
            int pageH = isH ? A4_W_MM : A4_H_MM;

            // Il canvas mostra il foglio A4 intero (scala ridotta per il foglio)
            picCanvas.Width = Math.Max(10, (int)(pageW * _pageScale));
            picCanvas.Height = Math.Max(10, (int)(pageH * _pageScale));
            picCanvas.Tag = _pageScale; // salva scala corrente
            if (lblZoom != null) lblZoom.Text = (_pageScale * 100).ToString("0") + "%";
            picCanvas.Invalidate();
        }

        private void picCanvas_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            float ps = picCanvas.Tag is float ? (float)picCanvas.Tag : 2.0f; // page scale

            bool isH = radOrientH != null && radOrientH.Checked;
            int pageW = isH ? A4_H_MM : A4_W_MM;
            int pageH = isH ? A4_W_MM : A4_H_MM;
            int mmW = (int)nudLargh.Value;
            int mmH = (int)nudAlt.Value;
            int mTop = nudMarginTop != null ? (int)nudMarginTop.Value : 0;
            int mLeft = nudMarginLeft != null ? (int)nudMarginLeft.Value : 0;

            // --- Foglio A4 (sfondo) ---
            int pgW = (int)(pageW * ps);
            int pgH = (int)(pageH * ps);
            g.FillRectangle(new SolidBrush(Color.FromArgb(230, 230, 235)), 0, 0, pgW, pgH);
            g.DrawRectangle(new Pen(Color.DarkGray, 1), 0, 0, pgW - 1, pgH - 1);

            // Linee margine (guida tratteggiata)
            int mTopPx = (int)(mTop * ps);
            int mLeftPx = (int)(mLeft * ps);
            using (Pen pMarg = new Pen(Color.FromArgb(100, Color.SteelBlue), 1f) { DashStyle = DashStyle.Dash })
            {
                if (mTopPx > 0) g.DrawLine(pMarg, 0, mTopPx, pgW, mTopPx);
                if (mLeftPx > 0) g.DrawLine(pMarg, mLeftPx, 0, mLeftPx, pgH);
            }

            // --- Griglia etichette sul foglio A4 ---
            int cols = nudCols != null ? (int)nudCols.Value : 1;
            int rows = nudRighe != null ? (int)nudRighe.Value : 1;

            int etiOriginX = mLeftPx;
            int etiOriginY = mTopPx;
            int etiW = (int)(mmW * ps);
            int etiH = (int)(mmH * ps);

            // Disegna tutte le etichette vuote sul foglio per visualizzare gli ingombri
            for (int dr = 0; dr < rows; dr++)
            {
                for (int dc = 0; dc < cols; dc++)
                {
                    if (dr == 0 && dc == 0) continue; // La prima la disegniamo completa dopo
                    int ox = mLeftPx + dc * etiW;
                    int oy = mTopPx + dr * etiH;

                    // Ignora se esce fuori dal foglio
                    if (ox + etiW > pgW || oy + etiH > pgH) continue;

                    g.FillRectangle(Brushes.WhiteSmoke, ox, oy, etiW, etiH);
                    g.DrawRectangle(Pens.LightGray, ox, oy, etiW - 1, etiH - 1);
                }
            }

            // --- Prima etichetta (posizionata con i margini in alto a sx) ---

            // Sfondo etichetta bianco
            g.FillRectangle(Brushes.White, etiOriginX, etiOriginY, etiW, etiH);
            g.DrawRectangle(Pens.Gray, etiOriginX, etiOriginY, etiW - 1, etiH - 1);

            // Imposta clip sull'etichetta per il disegno dei campi
            g.SetClip(new Rectangle(etiOriginX, etiOriginY, etiW, etiH));

            // Ordina per Z-Index. Per emulare esattamente il PDF:
            // Sfondo (Z=0): Forme vengono per prime, poi Immagini
            // Altri (Z=1): Immagini prima, Forme per ultime, Testo in mezzo.
            var campiOrdinati = _campi.Where(c => c.Visibile).OrderBy(c => c.ZIndex)
                                      .ThenBy(c =>
                                      {
                                          bool isShape = CAMPI_DB[c.Indice].StartsWith("eti_h0");
                                          bool isImage = (CAMPI_DB[c.Indice] == "eti_s012" || CAMPI_DB[c.Indice] == "eti_s017");
                                          if (c.ZIndex == 2) return 0; // Text on top
                                          if (c.ZIndex == 0) return isShape ? 0 : (isImage ? 1 : 2);
                                          else return isImage ? 0 : (isShape ? 2 : 1);
                                      }).ToList();
            foreach (FieldDef f in campiOrdinati)
            {
                int idxArr = _campi.IndexOf(f); // L'indice nell'array originale, serve per la tinta
                int i = f.Indice;
                string fld = CAMPI_DB[idxArr];
                bool isShape = fld.StartsWith("eti_h0");
                bool isImage = (fld == "eti_s012" || fld == "eti_s017");
                bool isBarcode = (fld == "eti_s013");
                bool isFreeText = fld.StartsWith("eti_t0");

                int px = etiOriginX + (int)(f.X * ps);
                int py = etiOriginY + (int)(f.Y * ps);
                int pw = (int)(f.W * ps);
                int ph = (int)(f.H * ps);
                Rectangle rect = new Rectangle(px, py, pw, ph);

                // Highlight trasparente per chiudere dentro gli elementi in modo visivo ma sbiadito (solo in edit)
                bool sel = (idxArr == _selIdx);
                Color clrHint = CAMPI_COLORI[idxArr];
                if (sel || _dragCampoIdx == -1) // Se sto trascinando qualcos'altro non mostro mille hover bg
                {
                    Color bgColor = Color.FromArgb(sel ? 200 : 50, clrHint);
                    using (SolidBrush br = new SolidBrush(bgColor))
                        g.FillRectangle(br, px, py, pw, ph);
                    using (Pen pen = new Pen(sel ? Color.DarkBlue : Color.FromArgb(100, clrHint.R / 2, clrHint.G / 2, clrHint.B / 2), sel ? 2 : 1))
                        g.DrawRectangle(pen, px, py, pw, ph);
                }

                if (isShape)
                {
                    Color shapeCol = Color.FromArgb(f.ShapeColor == 0 ? Color.Black.ToArgb() : f.ShapeColor);
                    if (f.ShapeType == 2) // Linea H
                    {
                        using (Pen pen = new Pen(shapeCol, f.ShapeThick)) g.DrawLine(pen, px, py + ph / 2, px + pw, py + ph / 2);
                    }
                    else if (f.ShapeType == 3) // Linea V
                    {
                        using (Pen pen = new Pen(shapeCol, f.ShapeThick)) g.DrawLine(pen, px + pw / 2, py, px + pw / 2, py + ph);
                    }
                    else // Rettangolo (1 o default)
                    {
                        if (f.ShapeFill) { using (SolidBrush br = new SolidBrush(shapeCol)) g.FillRectangle(br, rect); }
                        else { using (Pen pen = new Pen(shapeCol, f.ShapeThick)) g.DrawRectangle(pen, rect); }
                    }
                }
                else if (isImage)
                {
                    string path = f.FontNome;
                    StringFormat centerSf = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                    if (!string.IsNullOrEmpty(path) && File.Exists(path))
                    {
                        try { using (Image img = Image.FromFile(path)) e.Graphics.DrawImage(img, rect); }
                        catch
                        {
                            using (Font font = new Font("Arial", 8, FontStyle.Bold))
                                e.Graphics.DrawString(fld == "eti_s012" ? "LOGO" : "IMG", font, Brushes.Red, rect, centerSf);
                        }
                    }
                    else
                    {
                        using (Font font = new Font("Arial", 8, FontStyle.Bold))
                            e.Graphics.DrawString(fld == "eti_s012" ? "LOGO" : "IMG", font, Brushes.Gray, rect, centerSf);
                    }
                }
                else if (isBarcode)
                {
                    BarcodeLib.Barcode b = new BarcodeLib.Barcode();
                    b.Alignment = BarcodeLib.AlignmentPositions.CENTER;
                    b.IncludeLabel = true;
                    b.LabelPosition = BarcodeLib.LabelPositions.BOTTOMCENTER;
                    try
                    {
                        using (Image bImg = b.Encode(BarcodeLib.TYPE.EAN13, "8001234567890", Color.Black, Color.Transparent, pw, ph))
                        {
                            e.Graphics.DrawImage(bImg, rect);
                        }
                    }
                    catch
                    {
                        using (Font fnt = new Font("Arial", 8))
                            e.Graphics.DrawString("▌▌ EAN-13 ▌▌", fnt, Brushes.Black, rect);
                    }
                }
                else
                {
                    string demoTesto = GetDemoTesto(idxArr);
                    float fs = Math.Max(5, (float)f.FontSize * 0.8f);
                    FontStyle fst = f.Bold ? FontStyle.Bold : FontStyle.Regular;
                    if (fld == "eti_s010") fst |= FontStyle.Strikeout;
                    StringFormat sf = new StringFormat();
                    sf.Alignment = f.Align == 2 ? StringAlignment.Center : (f.Align == 1 ? StringAlignment.Far : StringAlignment.Near);
                    sf.Trimming = StringTrimming.EllipsisCharacter;
                    sf.FormatFlags = StringFormatFlags.NoWrap;
                    try
                    {
                        using (Font fnt = new Font(f.FontNome, fs, fst))
                        {
                            Color tc = Color.FromArgb(f.ShapeColor == 0 ? Color.Black.ToArgb() : f.ShapeColor);
                            // Se è il prezzo finale e il colore è nero o default, usiamo il rosso scuro come aiuto visivo
                            if (f.ShapeColor == 0 && fld == "eti_s011") tc = Color.DarkRed;
                            // Se è il prezzo vecchio e il colore è nero o default, usiamo il grigio
                            if (f.ShapeColor == 0 && fld == "eti_s010") tc = Color.Gray;

                            using (SolidBrush br = new SolidBrush(tc))
                            {
                                RectangleF textRect = new RectangleF(px + 2, py + 1, pw - 4, ph - 2);
                                g.DrawString(demoTesto, fnt, br, textRect, sf);
                            }
                        }
                    }
                    catch { }
                }

                if (sel)
                {
                    int hs = 8;
                    int half = hs / 2;
                    // 4 handles agli angoli (quadratini blu scuro)
                    g.FillRectangle(Brushes.DarkBlue, px - half, py - half, hs, hs); // top-left
                    g.FillRectangle(Brushes.DarkBlue, px + pw - half, py - half, hs, hs); // top-right
                    g.FillRectangle(Brushes.DarkBlue, px - half, py + ph - half, hs, hs); // bottom-left
                    g.FillRectangle(Brushes.DarkBlue, px + pw - half, py + ph - half, hs, hs); // bottom-right
                    // Contorno handle
                    using (Pen hp = new Pen(Color.White, 1))
                    {
                        g.DrawRectangle(hp, px - half, py - half, hs - 1, hs - 1);
                        g.DrawRectangle(hp, px + pw - half, py - half, hs - 1, hs - 1);
                        g.DrawRectangle(hp, px - half, py + ph - half, hs - 1, hs - 1);
                        g.DrawRectangle(hp, px + pw - half, py + ph - half, hs - 1, hs - 1);
                    }
                }
            }

            g.ResetClip();

            // Griglia mm leggera sull'etichetta
            using (Pen pGrid = new Pen(Color.FromArgb(20, 0, 0, 200), 1f))
            {
                pGrid.DashStyle = DashStyle.Dot;
                for (int mx = 10; mx < mmW; mx += 10)
                    g.DrawLine(pGrid, etiOriginX + (int)(mx * ps), etiOriginY, etiOriginX + (int)(mx * ps), etiOriginY + etiH);
                for (int my = 10; my < mmH; my += 10)
                    g.DrawLine(pGrid, etiOriginX, etiOriginY + (int)(my * ps), etiOriginX + etiW, etiOriginY + (int)(my * ps));
            }

            // Label dimensioni (tooltip visivo)
            string dimLabel = string.Format("A4 {0}  |  Eti: {1}x{2}mm  |  Margini: ↑{3} ←{4} mm",
                isH ? "Orizzontale" : "Verticale", mmW, mmH, mTop, mLeft);
            using (Font fLbl = new Font("Arial", 7f))
                g.DrawString(dimLabel, fLbl, Brushes.DimGray, 4, pgH - 14);
        }

        private string GetDemoTesto(int i)
        {
            if (i < 0 || i >= CAMPI_DB.Length) return "...";
            string fld = CAMPI_DB[i];
            if (fld.StartsWith("eti_t0"))
            {
                FieldDef f = (i >= 0 && i < _campi.Count) ? _campi[i] : null;
                if (f != null && !string.IsNullOrEmpty(f.TestoLibero))
                    return f.TestoLibero;
                return "TESTO LIBERO " + fld.Substring(5);
            }

            switch (fld)
            {
                case "eti_s003": return "INTESTAZIONE DEL NEGOZIO S.r.l.";
                case "eti_s004": return "PROSCIUTTO COTTO AL FORNO";
                case "eti_s005": return "Origine: Italia";
                case "eti_s006": return "Cod: 0012345";
                case "eti_s007": return "PROMO dal 20 al 27 Marzo";
                case "eti_s008": return "€ 12,90/kg";
                case "eti_s009": return "Sconto € 1,00";
                case "eti_s010": return "€ 2,99";
                case "eti_s011": return "€ 1,99";
                case "eti_s012": return "LOGO";
                case "eti_s013": return "8001234567890";
                case "eti_s014": return "ITALIA";
                case "eti_s015": return "CALIBRO: 1";
                case "eti_s016": return "PLU: 123";
                case "eti_s017": return "[IMG ARTICOLO]";
                case "eti_s018": return "KG";
                case "eti_s019": return "0,500";
                case "eti_s020": return "12";
                case "eti_s021": return "SALUMI";
                case "eti_s022": return "BANCO FRIGO";
                case "eti_s023": return DateTime.Now.ToString("dd/MM/yy");
                case "eti_s024": return "Dal 01/03 al 15/03";
                case "eti_s025": return "€ 9,90/kg";
                default: return "...";
            }
        }

        // ----------------------------------------------------------------
        // DRAG & DROP sull'anteprima
        // ----------------------------------------------------------------
        private void picCanvas_MouseDown(object sender, MouseEventArgs e)
        {
            float ps = picCanvas.Tag is float ? (float)picCanvas.Tag : 2.0f;
            bool isH = radOrientH != null && radOrientH.Checked;
            int mTop = nudMarginTop != null ? (int)nudMarginTop.Value : 0;
            int mLeft = nudMarginLeft != null ? (int)nudMarginLeft.Value : 0;
            int etiOriginX = (int)(mLeft * ps);
            int etiOriginY = (int)(mTop * ps);
            int hs = 8; // handle size px

            // Controlla prima se clicco su un handle dell'elemento selezionato
            if (_selIdx >= 0 && _selIdx < _campi.Count)
            {
                FieldDef fSel = _campi[_selIdx];
                if (fSel.Visibile)
                {
                    int px = etiOriginX + (int)(fSel.X * ps);
                    int py = etiOriginY + (int)(fSel.Y * ps);
                    int pw = (int)(fSel.W * ps);
                    int ph = (int)(fSel.H * ps);
                    int half = hs / 2;

                    Rectangle htTL = new Rectangle(px - half, py - half, hs + 2, hs + 2);
                    Rectangle htTR = new Rectangle(px + pw - half, py - half, hs + 2, hs + 2);
                    Rectangle htBL = new Rectangle(px - half, py + ph - half, hs + 2, hs + 2);
                    Rectangle htBR = new Rectangle(px + pw - half, py + ph - half, hs + 2, hs + 2);

                    ResizeHandle hit = ResizeHandle.None;
                    if (htTL.Contains(e.Location)) hit = ResizeHandle.TopLeft;
                    else if (htTR.Contains(e.Location)) hit = ResizeHandle.TopRight;
                    else if (htBL.Contains(e.Location)) hit = ResizeHandle.BottomLeft;
                    else if (htBR.Contains(e.Location)) hit = ResizeHandle.BottomRight;

                    if (hit != ResizeHandle.None)
                    {
                        _resizeHandle = hit;
                        _isResizing = true;
                        _dragCampoIdx = _selIdx;
                        _dragStartMouse = e.Location;
                        _dragStartState = fSel.Clone();
                        Cursor cur = (hit == ResizeHandle.TopLeft || hit == ResizeHandle.BottomRight)
                                     ? Cursors.SizeNWSE : Cursors.SizeNESW;
                        picCanvas.Cursor = cur;
                        return;
                    }
                }
            }

            // Nessun handle colpito: comportamento originale drag-move
            _isResizing = false;
            _resizeHandle = ResizeHandle.None;
            for (int i = _campi.Count - 1; i >= 0; i--)
            {
                FieldDef f = _campi[i];
                if (!f.Visibile) continue;
                Rectangle r = new Rectangle(etiOriginX + (int)(f.X * ps), etiOriginY + (int)(f.Y * ps), (int)(f.W * ps), (int)(f.H * ps));
                if (r.Contains(e.Location))
                {
                    _dragCampoIdx = i;
                    _dragStartMouse = e.Location;
                    _dragStartState = f.Clone();
                    _isDragging = true;
                    _selIdx = i;
                    lstCampi.SelectedIndex = i;
                    picCanvas.Cursor = Cursors.SizeAll;
                    AggiornaCanvas();
                    return;
                }
            }
            _selIdx = -1;
            AggiornaCanvas();
        }

        private void picCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            float ps = picCanvas.Tag is float ? (float)picCanvas.Tag : 2.0f;
            int mTop = nudMarginTop != null ? (int)nudMarginTop.Value : 0;
            int mLeft = nudMarginLeft != null ? (int)nudMarginLeft.Value : 0;
            int etiOriginX = (int)(mLeft * ps);
            int etiOriginY = (int)(mTop * ps);
            int hs = 8;

            if (_isResizing && _dragCampoIdx >= 0 && _dragCampoIdx < _campi.Count)
            {
                int dxPx = e.Location.X - _dragStartMouse.X;
                int dyPx = e.Location.Y - _dragStartMouse.Y;
                int dxMm = (int)Math.Round(dxPx / ps);
                int dyMm = (int)Math.Round(dyPx / ps);

                FieldDef f = _campi[_dragCampoIdx];
                int startX = _dragStartState.X;
                int startY = _dragStartState.Y;
                int startW = _dragStartState.W;
                int startH = _dragStartState.H;

                switch (_resizeHandle)
                {
                    case ResizeHandle.TopLeft:
                        f.X = Math.Max(0, startX + dxMm);
                        f.Y = Math.Max(0, startY + dyMm);
                        f.W = Math.Max(3, startW - dxMm);
                        f.H = Math.Max(2, startH - dyMm);
                        break;
                    case ResizeHandle.TopRight:
                        f.Y = Math.Max(0, startY + dyMm);
                        f.W = Math.Max(3, startW + dxMm);
                        f.H = Math.Max(2, startH - dyMm);
                        break;
                    case ResizeHandle.BottomLeft:
                        f.X = Math.Max(0, startX + dxMm);
                        f.W = Math.Max(3, startW - dxMm);
                        f.H = Math.Max(2, startH + dyMm);
                        break;
                    case ResizeHandle.BottomRight:
                        f.W = Math.Max(3, startW + dxMm);
                        f.H = Math.Max(2, startH + dyMm);
                        break;
                }

                _suppressEvents = true;
                nudPosX.Value = Math.Min(nudPosX.Maximum, (decimal)f.X);
                nudPosY.Value = Math.Min(nudPosY.Maximum, (decimal)f.Y);
                nudPosW.Value = Math.Min(nudPosW.Maximum, Math.Max(nudPosW.Minimum, (decimal)f.W));
                nudPosH.Value = Math.Min(nudPosH.Maximum, Math.Max(nudPosH.Minimum, (decimal)f.H));
                _suppressEvents = false;

                AggiornaCanvas();
                return;
            }

            if (_isDragging && _dragCampoIdx >= 0 && _dragCampoIdx < _campi.Count)
            {
                int dxPx = e.Location.X - _dragStartMouse.X;
                int dyPx = e.Location.Y - _dragStartMouse.Y;
                int dxMm = (int)Math.Round(dxPx / ps);
                int dyMm = (int)Math.Round(dyPx / ps);

                FieldDef f = _campi[_dragCampoIdx];
                f.X = Math.Max(0, _dragStartState.X + dxMm);
                f.Y = Math.Max(0, _dragStartState.Y + dyMm);

                _suppressEvents = true;
                nudPosX.Value = Math.Min(nudPosX.Maximum, (decimal)f.X);
                nudPosY.Value = Math.Min(nudPosY.Maximum, (decimal)f.Y);
                _suppressEvents = false;

                AggiornaCanvas();
                return;
            }

            // Hover: aggiorna il cursore del mouse in base a dove si trova
            if (_selIdx >= 0 && _selIdx < _campi.Count)
            {
                FieldDef fSel = _campi[_selIdx];
                if (fSel.Visibile)
                {
                    int px = etiOriginX + (int)(fSel.X * ps);
                    int py = etiOriginY + (int)(fSel.Y * ps);
                    int pw = (int)(fSel.W * ps);
                    int ph = (int)(fSel.H * ps);
                    int half = hs / 2;

                    Rectangle htTL = new Rectangle(px - half, py - half, hs + 2, hs + 2);
                    Rectangle htTR = new Rectangle(px + pw - half, py - half, hs + 2, hs + 2);
                    Rectangle htBL = new Rectangle(px - half, py + ph - half, hs + 2, hs + 2);
                    Rectangle htBR = new Rectangle(px + pw - half, py + ph - half, hs + 2, hs + 2);

                    if (htTL.Contains(e.Location) || htBR.Contains(e.Location)) { picCanvas.Cursor = Cursors.SizeNWSE; return; }
                    if (htTR.Contains(e.Location) || htBL.Contains(e.Location)) { picCanvas.Cursor = Cursors.SizeNESW; return; }

                    Rectangle rSel = new Rectangle(px, py, pw, ph);
                    if (rSel.Contains(e.Location)) { picCanvas.Cursor = Cursors.SizeAll; return; }
                }
            }

            // Se è su qualsiasi altro elemento
            for (int i = _campi.Count - 1; i >= 0; i--)
            {
                FieldDef f = _campi[i];
                if (!f.Visibile) continue;
                Rectangle r = new Rectangle(etiOriginX + (int)(f.X * ps), etiOriginY + (int)(f.Y * ps), (int)(f.W * ps), (int)(f.H * ps));
                if (r.Contains(e.Location)) { picCanvas.Cursor = Cursors.Hand; return; }
            }

            picCanvas.Cursor = Cursors.Default;
        }

        private void picCanvas_MouseUp(object sender, MouseEventArgs e)
        {
            _isDragging = false;
            _isResizing = false;
            _dragCampoIdx = -1;
            _resizeHandle = ResizeHandle.None;
            picCanvas.Cursor = Cursors.Default;
        }

        // ----------------------------------------------------------------
        // ZOOM ANTEPRIMA
        // ----------------------------------------------------------------
        private void btnZoomIn_Click(object sender, EventArgs e)
        {
            if (_pageScale < 6.0f)
            {
                _pageScale += 0.5f;
                AggiornaCanvas();
            }
        }

        private void btnZoomOut_Click(object sender, EventArgs e)
        {
            if (_pageScale > 0.5f)
            {
                _pageScale -= 0.5f;
                AggiornaCanvas();
            }
        }

        // ----------------------------------------------------------------
        // PULSANTI
        // ----------------------------------------------------------------
        private void btnNuovo_Click(object sender, EventArgs e)
        {
            int max = 899;
            foreach (DataRow r in _tabFmt.Rows)
            {
                if (r.RowState == DataRowState.Deleted) continue;
                int c; if (int.TryParse(r["eti_cod"].ToString(), out c) && c > max) max = c;
            }
            DataRow nr = _tabFmt.NewRow();
            nr["eti_cod"] = (max + 1).ToString("000");
            nr["eti_des"] = "Nuova Etichetta";
            nr["eti_tip"] = "PDF"; nr["eti_pri"] = ""; nr["eti_dim"] = "100,50";
            nr["eti_imm"] = "N"; nr["eti_pag"] = "1,1,0,0,0,0";
            if (_tabFmt.Columns.Contains("eti_sta")) nr["eti_sta"] = "S";
            foreach (string col in CAMPI_DB)
            {
                if (_tabFmt.Columns.Contains(col)) nr[col] = "";
            }
            nr["eti_ann"] = false;
            _tabFmt.Rows.Add(nr);
            dgvFormati.Rows[_tabFmt.Rows.Count - 1].Selected = true;
        }

        private void btnDuplica_Click(object sender, EventArgs e)
        {
            if (_selectedRowIndex < 0) return;
            DataRow src = _tabFmt.Rows[_selectedRowIndex];
            int max = 899;
            foreach (DataRow r in _tabFmt.Rows) { int c; if (int.TryParse(r["eti_cod"].ToString(), out c) && c > max) max = c; }
            DataRow nr = _tabFmt.NewRow();
            foreach (DataColumn col in _tabFmt.Columns)
                nr[col.ColumnName] = src[col.ColumnName];
            nr["eti_cod"] = (max + 1).ToString("000");
            nr["eti_des"] = "Copia - " + src["eti_des"];
            _tabFmt.Rows.Add(nr);
            dgvFormati.Rows[_tabFmt.Rows.Count - 1].Selected = true;
        }

        private void btnElimina_Click(object sender, EventArgs e)
        {
            if (_selectedRowIndex < 0) return;
            if (MessageBox.Show("Eliminare il formato selezionato?", "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                _tabFmt.Rows[_selectedRowIndex]["eti_ann"] = true;
                _tabFmt.Rows[_selectedRowIndex].Delete();
                _da.Update(_tabFmt);
                _tabFmt.AcceptChanges();
                LoadFormati();
            }
        }

        private void btnSalva_Click(object sender, EventArgs e)
        {
            if (_selectedRowIndex < 0) return;
            try
            {
                // Aggiorna riga selezionata nel DataTable
                DataRow r = _tabFmt.Rows[_selectedRowIndex];
                r["eti_des"] = txtDescrizione.Text;
                r["eti_dim"] = string.Format("{0},{1}", (int)nudLargh.Value, (int)nudAlt.Value);
                // pag: Cols,Rows,topMargin,leftMargin,xStep,yStep,orientation(V/H)
                string ori = radOrientH.Checked ? "H" : "V";
                r["eti_pag"] = string.Format("{0},{1},{2},{3},{4},{5},{6}",
                    (int)nudCols.Value, (int)nudRighe.Value,
                    (int)nudMarginTop.Value, (int)nudMarginLeft.Value,
                    (int)nudLargh.Value, (int)nudAlt.Value, ori);

                // Stato post-stampa (S/N/E)
                string sta = "S";
                if (cmbEtiSta.SelectedIndex == 1) sta = "N";
                else if (cmbEtiSta.SelectedIndex == 2) sta = "E";
                if (r.Table.Columns.Contains("eti_sta"))
                    r["eti_sta"] = sta;

                // Salva i campi
                for (int i = 0; i < _campi.Count; i++)
                {
                    if (r.Table.Columns.Contains(CAMPI_DB[i]))
                        r[CAMPI_DB[i]] = _campi[i].ToDb();
                }

                _da.Update(_tabFmt);
                _tabFmt.AcceptChanges();
                MessageBox.Show("Formato salvato!", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore: " + ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnEsci_Click(object sender, EventArgs e) { this.Close(); }

        // txtDescrizione/Dimensioni eventi
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            txtDescrizione.TextChanged += (s, ev) => { if (!_suppressEvents && _selectedRowIndex >= 0) _tabFmt.Rows[_selectedRowIndex]["eti_des"] = txtDescrizione.Text; };
            nudLargh.ValueChanged += NudDim_ValueChanged;
            nudAlt.ValueChanged += NudDim_ValueChanged;
            nudMarginTop.ValueChanged += NudDim_ValueChanged;
            nudMarginLeft.ValueChanged += NudDim_ValueChanged;
            radOrientV.CheckedChanged += (s, ev) => { if (!_suppressEvents) AggiornaCanvas(); };
            radOrientH.CheckedChanged += (s, ev) => { if (!_suppressEvents) AggiornaCanvas(); };
        }
    }

    // --------------------------------------------------------
    // Modello dati per un singolo campo dell'etichetta
    // --------------------------------------------------------
    internal class FieldDef
    {
        public bool Visibile { get; set; }
        public int X { get; set; }
        public int Y { get; set; }
        public int W { get; set; }
        public int H { get; set; }
        public int Align { get; set; }     // 0=Sx 1=Dx 2=Centro
        public string FontNome { get; set; }
        public double FontSize { get; set; }
        public bool Bold { get; set; }
        public int Indice { get; set; }

        public int ZIndex { get; set; }    // 0 = Sfondo, 1 = Primo Piano, 2 = Sopra Tutto
        public int ShapeColor { get; set; }
        public int ShapeThick { get; set; }
        public bool ShapeFill { get; set; }
        public int ShapeType { get; set; } // 1 = Rettangolo, 2 = Linea H, 3 = Linea V
        public string TestoLibero { get; set; }

        public static FieldDef Parse(string raw, int indice)
        {
            var f = new FieldDef { Visibile = false, X = 5, Y = 5, W = 60, H = 10, Align = 0, FontNome = "Arial", FontSize = 9, Bold = false, Indice = indice, ZIndex = 1, ShapeColor = Color.Black.ToArgb(), ShapeThick = 1, ShapeFill = false, ShapeType = 1, TestoLibero = "" };
            if (string.IsNullOrEmpty(raw)) return f;

            string[] parts = raw.Split(';');
            if (parts.Length < 3) return f; // Formato non valido o vecchio

            f.Visibile = (parts[0] == "S");

            // Coordinate (secondo segmento)
            string[] dim = parts[1].Split(',');
            double d;
            if (dim.Length > 0 && double.TryParse(dim[0], out d)) f.Y = (int)Math.Round(d);
            if (dim.Length > 1 && double.TryParse(dim[1], out d)) f.X = (int)Math.Round(d);
            if (dim.Length > 2 && double.TryParse(dim[2], out d)) f.W = (int)Math.Round(d);
            if (dim.Length > 3 && double.TryParse(dim[3], out d)) f.H = (int)Math.Round(d);
            int a; if (dim.Length > 4 && int.TryParse(dim[4], out a)) f.Align = a;

            // Proprietà Font/Path (terzo segmento)
            string[] fnt = parts[2].Split('|');
            if (fnt.Length > 0 && fnt[0] != "") f.FontNome = fnt[0];
            if (fnt.Length > 1 && double.TryParse(fnt[1], out d)) f.FontSize = d;
            if (fnt.Length > 2) f.Bold = (fnt[2].ToUpper() == "B");

            // Proprietà Forma / Colore / ZIndex (quarto segmento)
            if (parts.Length > 3)
            {
                string[] shp = parts[3].Split('|');
                int val;
                if (shp.Length > 0 && int.TryParse(shp[0], out val)) f.ZIndex = val;
                if (shp.Length > 1 && int.TryParse(shp[1], out val)) f.ShapeColor = val;
                if (shp.Length > 2 && int.TryParse(shp[2], out val)) f.ShapeThick = val;
                f.ShapeFill = (shp.Length > 3 && shp[3] == "F");
                if (shp.Length > 4 && int.TryParse(shp[4], out val)) f.ShapeType = val;
            }

            // Testo Libero (quinto segmento)
            if (parts.Length > 4)
            {
                f.TestoLibero = parts[4];
            }

            return f;
        }

        public string ToDb()
        {
            string vis = Visibile ? "S" : "N";
            string tLib = (TestoLibero ?? "").Replace(';', ',');
            return string.Format("{0};{1},{2},{3},{4},{5};{6}|{7}|{8};{9}|{10}|{11}|{12}|{13};{14}",
                vis, Y, X, W, H, Align, FontNome, FontSize.ToString("0.##"), Bold ? "B" : "N",
                ZIndex, ShapeColor, ShapeThick, ShapeFill ? "F" : "E", ShapeType, tLib);
        }

        public FieldDef Clone()
        {
            return new FieldDef
            {
                Visibile = this.Visibile,
                X = this.X,
                Y = this.Y,
                W = this.W,
                H = this.H,
                Align = this.Align,
                FontNome = this.FontNome,
                FontSize = this.FontSize,
                Bold = this.Bold,
                Indice = this.Indice,
                ZIndex = this.ZIndex,
                ShapeColor = this.ShapeColor,
                ShapeThick = this.ShapeThick,
                ShapeFill = this.ShapeFill,
                ShapeType = this.ShapeType,
                TestoLibero = this.TestoLibero
            };
        }
    }
}
