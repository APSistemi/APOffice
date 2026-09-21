using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;

namespace APOffice
{
    public partial class frmAnaArticolo : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsVariazioni _clsVar = new clsVariazioni();
        clsQuery _clsQry = new clsQuery();

        private const string TABECRLV1 = "TabEcrLv1";
        private const string TABECRLV2 = "TabEcrLv2";
        private const string TABECRLV3 = "TabEcrLv3";

        private const string TABANAART = "AnaArticoli";
        private const string TABANAEAN = "AnaBarcode";
        private const string TABLISACQ = "GesLisAcquisto";
        private const string TABLISVEN = "GesLisVendita";

        public string _strArtCod = "";
        private string _strArtDes = "";

        private string _strConSql = "";
        public string _strSqlArt = "";

        private string _strPar016PathDivNegozi = "";
        private string _strPar021GesIngredienti = "";

        private DataSet _dasGen = new DataSet();

        public DataTable _tabArt = new DataTable();
        private DataTable _tabEan = new DataTable();  // Cache AnaBarcode — riusato da SalvaEan
        private DataTable _tabCos = new DataTable();  // Cache GesLisAcquisto — riusato da SalvaCos
        private DataTable _tabVen = new DataTable();  // Cache GesLisVendita — riusato da SalvaVen


        private bool _bolLoading = false;
        private bool _bolIsDirty = false;

        private void SetDirty(object sender, EventArgs e)
        {
            if (!_bolLoading)
            {
                _bolIsDirty = true;
            }
        }

        private void WireEvents(Control.ControlCollection controls)
        {
            foreach (Control c in controls)
            {
                if (c == txtSeek) continue;
                if (c is TextBox tb)
                    tb.TextChanged += SetDirty;
                else if (c is ComboBox cb)
                    cb.SelectedIndexChanged += SetDirty;
                else if (c is CheckBox chk)
                    chk.CheckedChanged += SetDirty;
                else if (c is DateTimePicker dtp)
                    dtp.ValueChanged += SetDirty;

                if (c.HasChildren)
                    WireEvents(c.Controls);
            }
        }

        private void EnableDoubleBuffering(Control ctrl)
        {
            if (ctrl == null) return;
            try
            {
                typeof(Control).InvokeMember("DoubleBuffered",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty,
                    null, ctrl, new object[] { true });
            }
            catch { }
        }

        public frmAnaArticolo()
        {
            InitializeComponent();
            this.DoubleBuffered = true;
            EnableDoubleBuffering(dgv1);
            EnableDoubleBuffering(dgv2);
            EnableDoubleBuffering(dgv3);
            EnableDoubleBuffering(dgv4);
            new clsGesGraph().SetGraph(this, 0);
        }

        private static readonly Dictionary<string, Image> _iconCache = new Dictionary<string, Image>();

        private Image GetIcon(string unicodeChar, Color fallbackColor)
        {
            string key = unicodeChar + "_" + fallbackColor.ToArgb();
            if (_iconCache.TryGetValue(key, out Image cachedImg))
                return cachedImg;

            Bitmap bmp = new Bitmap(24, 24);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                using (Font f = new Font("Segoe UI Emoji", 12, FontStyle.Regular))
                using (SolidBrush b = new SolidBrush(fallbackColor))
                {
                    StringFormat fmt = new StringFormat();
                    fmt.Alignment = StringAlignment.Center;
                    fmt.LineAlignment = StringAlignment.Center;
                    g.DrawString(unicodeChar, f, b, new RectangleF(0, 0, 24, 24), fmt);
                }
            }
            _iconCache[key] = bmp;
            return bmp;
        }

        private void frmAnaArticolo_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");

            _strPar016PathDivNegozi = _clsFun.ParGet(clsDefine.enuParametri.Par016PathDivNegozi, _strConSql);
            _strPar021GesIngredienti = _clsFun.ParGet(clsDefine.enuParametri.Par021GesIngredienti, _strConSql);

            if (_clsFun.ParGet(clsDefine.enuParametri.Par018NuovoArticoloNo, _strConSql) == "S")
            {
                nuovoToolStripMenuItem.Enabled = false;
                inserimentoVeloceToolStripMenuItem.Enabled = false;
            }

            FillTabs("");
            SetDgv1();
            SetDgv2();
            SetDgv3();
            SetDgv4();

            EnableDoubleBuffering(dgv1);
            EnableDoubleBuffering(dgv2);
            EnableDoubleBuffering(dgv3);
            EnableDoubleBuffering(dgv4);
            WireEvents(this.Controls);
            FillArt(false);
            txtSeek.Select();
            if (_strArtCod == "")
                CtrlPanels("INI");
            else
                CtrlPanels("");

            InizializzaIconeMenu();
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmAnaArticolo_FormClosing);
            this.Shown += new System.EventHandler(this.frmAnaArticolo_Shown);
            this.Activated += new System.EventHandler(this.frmAnaArticolo_Activated);
        }

        private void EnableDoubleBuffering(DataGridView dgv)
        {
            if (dgv == null) return;
            try
            {
                typeof(DataGridView).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                    ?.SetValue(dgv, true, null);
            }
            catch { }
        }

        private void SetFocusToSeek()
        {
            try
            {
                if (txtSeek != null && txtSeek.CanFocus)
                {
                    this.ActiveControl = txtSeek;
                    txtSeek.Focus();
                    txtSeek.SelectAll();
                }
            }
            catch { }
        }

        private void frmAnaArticolo_Shown(object sender, EventArgs e)
        {
            this.BeginInvoke(new Action(SetFocusToSeek));
        }

        private void frmAnaArticolo_Activated(object sender, EventArgs e)
        {
            this.BeginInvoke(new Action(SetFocusToSeek));
        }

        private void InizializzaIconeMenu()
        {
            try
            {
                // Abilita la visualizzazione dei ToolTip per il MenuStrip
                menuStrip1.ShowItemToolTips = true;

                // Icone Menu superiore e ToolTips con tasti di scelta rapida
                esciToolStripMenuItem.Image = GetIcon("🚪", Color.DarkRed);
                esciToolStripMenuItem.ToolTipText = "Esci (ESC) - Chiudi maschera anagrafica";

                salvaAggiornaToolStripMenuItem.Image = GetIcon("🔄", Color.DarkGreen);
                salvaAggiornaToolStripMenuItem.ToolTipText = "Aggiorna (F5) - Salva e aggiorna scheda articolo";

                nuovoToolStripMenuItem.Image = GetIcon("➕", Color.Navy);
                nuovoToolStripMenuItem.ToolTipText = "Nuovo - Crea un nuovo articolo";

                inserimentoVeloceToolStripMenuItem.Text = "Ins. Veloce";
                inserimentoVeloceToolStripMenuItem.Image = GetIcon("⚡", Color.DarkOrange);
                inserimentoVeloceToolStripMenuItem.ToolTipText = "Inserimento veloce (F3) - Ricerca e creazione rapida da barcode";

                ricercaToolStripMenuItem.Image = GetIcon("🔍", Color.DarkSlateBlue);
                ricercaToolStripMenuItem.ToolTipText = "Ricerca (F4) - Cerca articolo nell'archivio";

                invioInCassaToolStripMenuItem.Text = "Invia Cassa";
                invioInCassaToolStripMenuItem.Image = GetIcon("🛒", Color.DarkGreen);
                invioInCassaToolStripMenuItem.ToolTipText = "Variazione cassa - Invia variazioni al registratore di cassa / POS";

                generazioneEtichettaToolStripMenuItem.Text = "Etichette";
                generazioneEtichettaToolStripMenuItem.Image = GetIcon("🏷️", Color.Chocolate);
                generazioneEtichettaToolStripMenuItem.ToolTipText = "Generazione etichetta - Stampa etichette per questo articolo";

                gestioneVariazioniToolStripMenuItem.Text = "Variazioni";
                gestioneVariazioniToolStripMenuItem.Image = GetIcon("⚙️", Color.DimGray);
                gestioneVariazioniToolStripMenuItem.ToolTipText = "Gestione variazioni - Consultazione ed allineamenti";

                statisticheToolStripMenuItem.Text = "Statistiche";
                statisticheToolStripMenuItem.Image = GetIcon("📈", Color.RoyalBlue);
                statisticheToolStripMenuItem.ToolTipText = "Statistiche - Visualizza storico e statistiche di vendita";

                // Icone Pulsanti principali (Mantenendo allineamento e posizione)
                btnIng.Image = GetIcon("🥗", Color.SeaGreen);
                btnIng.TextImageRelation = TextImageRelation.ImageBeforeText;
                btnIng.ImageAlign = ContentAlignment.MiddleLeft;

                btnPluEan.Image = GetIcon("🔢", Color.DarkSlateGray);
                btnPluEan.TextImageRelation = TextImageRelation.ImageBeforeText;
                btnPluEan.ImageAlign = ContentAlignment.MiddleLeft;

                btnForCos.Image = GetIcon("🪙", Color.Goldenrod);
                btnForCos.TextImageRelation = TextImageRelation.ImageBeforeText;
                btnForCos.ImageAlign = ContentAlignment.MiddleLeft;
            }
            catch { }
        }

        private void frmAnaArticolo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                Esci();
            }
            else if (e.KeyCode == Keys.F3)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                inserimentoVeloceToolStripMenuItem_Click(sender, e);
            }
            else if (e.KeyCode == Keys.F5)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                salvaAggiornaToolStripMenuItem_Click(sender, e);
            }
            else if (e.KeyCode == Keys.F8)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                refreshToolStripMenuItem_Click(sender, e);
            }
        }

        private void ResetGridModifiedFlags()
        {
            try
            {
                DataTable dt1 = GetGridDataTable(dgv1);
                if (dt1 != null && dt1.Columns.Contains("EanMdy"))
                {
                    foreach (DataRow r in dt1.Rows)
                        r["EanMdy"] = "";
                }

                DataTable dt2 = GetGridDataTable(dgv2);
                if (dt2 != null && dt2.Columns.Contains("LiaMdy"))
                {
                    foreach (DataRow r in dt2.Rows)
                        r["LiaMdy"] = "";
                }

                DataTable dt3 = GetGridDataTable(dgv3);
                if (dt3 != null && dt3.Columns.Contains("LivMdy"))
                {
                    foreach (DataRow r in dt3.Rows)
                        r["LivMdy"] = "";
                }

                DataTable dt4 = GetGridDataTable(dgv4);
                if (dt4 != null && dt4.Columns.Contains("LivMdy"))
                {
                    foreach (DataRow r in dt4.Rows)
                        r["LivMdy"] = "";
                }
            }
            catch { }
        }

        private void salvaAggiornaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string sCod = txtArtCod.Text != null ? txtArtCod.Text.Trim() : "";

            if (string.IsNullOrEmpty(sCod))
            {
                MessageBox.Show(this, "Codice articolo non valido o non definito. NON POSSIBILE salvare o aggiornare l'articolo!", "NON POSSIBILE", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            salvaAggiornaToolStripMenuItem.Text = "Salvataggio...";
            Application.UseWaitCursor = true;

            try
            {
                Boolean b = Salva();

                if (b)
                {
                    _bolIsDirty = false;
                    ResetGridModifiedFlags();
                    salvaAggiornaToolStripMenuItem.Text = "Aggiorna";
                    Application.UseWaitCursor = false;
                    MessageBox.Show(this, "Scheda articolo '" + txtArtDes.Text + "' (Cod. " + txtArtCod.Text + ") salvata e aggiornata con successo!", "SCHEDA AGGIORNATA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    Application.UseWaitCursor = false;
                    MessageBox.Show(this, "Impossibile salvare l'articolo. Verificare i dati inseriti.", "NON POSSIBILE", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            finally
            {
                salvaAggiornaToolStripMenuItem.Text = "Aggiorna";
                Application.UseWaitCursor = false;
            }
        }

        private void frmAnaArticolo_FormClosing(object sender, FormClosingEventArgs e)
        {
            AutoSaveIfDirty();
        }

        private void AutoSaveIfDirty()
        {
            try
            {
                if (_bolIsDirty && !string.IsNullOrEmpty(txtArtDes.Text.Trim()) && txtArtCod.Text.Trim() != "" && txtArtCod.Text.Trim() != _clsDef.CODNEW)
                {
                    Salva();
                    _bolIsDirty = false;
                }
            }
            catch { }
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void Esci()
        {
            _strArtCod = txtArtCod.Text;
            _strArtDes = txtArtDes.Text;

            Boolean b = Salva();
            if (!b)
            {
                if (MessageBox.Show("Abbandono delle modifiche, confermi?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                      == DialogResult.Yes)
                    this.Close();
            }
            else
                this.Close();
        }

        private void nuovoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (txtArtCod.Text == _clsDef.CODNEW)
                MessageBox.Show(this, "Codice articolo già in inseriemnto!", "NUOVO ARTICOLO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                Salva();

                string s = _clsFun.ParGet(clsDefine.enuParametri.Par011InsArtCodMod, _strConSql);
                if (s == "S")
                    txtArtCod.Enabled = true;

                NewArt();
                pnlArt.Enabled = true;
                CtrlPanels("");
            }
        }

        private void inserimentoVeloceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Salva();

            //            s = _clsFun.ParGet(clsDefine.enuParametri.Par036ArtInsVeloce, _strConSql);
            //if(s.Length > 0 && s.Substring(0,1) == "S")
            //{
            frmAnaArtNuovo f = new frmAnaArtNuovo();
            if (_clsFun.Numerico(txtSeek.Text, "0123456789") && txtSeek.Text.Length > 7 && txtSeek.Text.Length < 14)
                f._strEan = txtSeek.Text;
            f.ShowDialog();
            if (f._strRes != "")
                FillNewArt(f._strRes);
            //}
            //else
            //{

        }

        private void NewArt()
        {
            Boolean b = true;

            if (txtArtCod.Text == _clsDef.CODNEW)
            {
                if (MessageBox.Show("Articolo già in fase di inserimento, confermi nuovo inserimento?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    b = false;
            }

            if (b)
            {
                if (txtArtCod.Text != _clsDef.CODNEW && txtArtCod.Text != "NEW" && _bolIsDirty)
                {
                    Salva();
                }

                pnlArt.Enabled = true;
                txtArtCod.Text = _clsDef.CODNEW;
                _strArtCod = _clsDef.CODNEW;

                // Pulizia campi anagrafici
                txtArtDes.Text = "";
                txtArtDeb.Text = "";
                txtArtPlu.Text = "";
                txtArtPxc.Text = "1";
                txtArtPne.Text = "0,00";
                txtArtCou.Text = "0,000";
                txtArtCol.Text = "0,000";
                txtArtCof.Text = "0,000";
                txtArtGsc.Text = "0";
                txtArtSfr.Text = "0";
                txtArtTar.Text = "0";
                txtArtGia.Text = "0";
                txtArtTas.Text = "0";
                txtArtImg.Text = "";
                chkArtBpz.Checked = false;
                chkArtBil.Checked = false;
                chkArtWeb.Checked = false;
                chkArtCel.Checked = false;
                chkTra.Checked = false;
                cmbArtSta.SelectedValue = _clsDef.STAATT;

                // Inizializza tutte le griglie con schemi vuoti validi
                FillArt(true);

                string s = _clsFun.ParGet(clsDefine.enuParametri.ParEcrDefault, _strConSql);
                string[] aEcr = s.Split(',');

                if (Convert.ToString(cmbArtEc1.SelectedValue) == "" && aEcr.Length > 2)
                {
                    cmbArtEc1.SelectedValue = aEcr[0];
                    cmbArtEc2.SelectedValue = aEcr[1];
                    cmbArtEc3.SelectedValue = aEcr[2];
                }
                txtArtDes.Select();
                _bolIsDirty = false;
            }
        }

        private void ricercaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Ricerca();
        }

        private void CtrlPanels(string strTip)
        {
            if (strTip == "INI")
            {
                pnlArt.Enabled = false;
                groupBox1.Enabled = false;
            }
            else
            {
                pnlArt.Enabled = true;
                if (groupBox1.Text.Length > 10)
                    groupBox1.Enabled = true;
            }
        }

        private void ArtStato()
        {
            if (cmbArtSta.SelectedValue != null)
            {
                cmbArtSta.BackColor = Color.Aquamarine;
                cmbArtSta.ForeColor = Color.Black;

                string s = Convert.ToString(cmbArtSta.SelectedValue);
                if (s == _clsDef.STAATT)
                    cmbArtSta.BackColor = Color.PaleGreen;
                else if (s == _clsDef.STANOA)
                    cmbArtSta.BackColor = Color.Crimson;
                else if (s == _clsDef.STANOA)
                    cmbArtSta.BackColor = Color.Crimson;
                else if (s == _clsDef.STACAN)
                {
                    cmbArtSta.BackColor = Color.Black;
                    cmbArtSta.ForeColor = Color.White;
                }
            }
        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;
            dgv1.SelectionMode = DataGridViewSelectionMode.CellSelect;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ean_ean";
            cTbc.Name = "Barcode";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ean_qta";
            cTbc.Name = "Q.tà";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ean_prv";
            cTbc.Name = "Prezzo vendita";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "ean_bil";
            cCbc.Name = "Bilancia";
            cCbc.Width = 25;
            dgv1.Columns.Add(cCbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "ean_ecp";
            cCbc.Name = "Ean con peso";
            cCbc.Width = 25;
            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ean_pun";
            cTbc.Name = "Punti fidelity";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0";

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "ean_ann";
            cCbc.Name = "Annullato";
            cCbc.Width = 25;
            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ean_dtm";
            cTbc.Name = "Data Modifica";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "EanMdy";
            cTbc.Name = "Modificato";
            cTbc.Width = 1;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void SetDgv2()
        {
            dgv2.AutoGenerateColumns = false;
            //dgv2.VirtualMode = true;
            //dgv2.Dock = DockStyle.Fill;
            dgv2.AllowUserToAddRows = false;
            dgv2.ReadOnly = false;
            dgv2.AllowUserToDeleteRows = false;
            //dgv2.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_tip";
            cTbc.Name = "Tipo";
            cTbc.Width = 15;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.ToolTipText = "Manuale, Divulgazione, Fattura";
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "CosFod";
            cTbc.Name = "Fornitore";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_stf";
            cTbc.Name = "Stato";
            cTbc.Width = 15;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_arf";
            cTbc.Name = "Articolo fornitore";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 9;
            dgv2.Columns.Add(cTbc);
            dgv2.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv2.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.000";

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "lia_ann";
            cCbc.Name = "Annullato";
            cCbc.Width = 25;
            dgv2.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_dti";
            cTbc.Name = "Data Modifica";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);
            dgv2.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_dtf";
            cTbc.Name = "Data fine";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);
            dgv2.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_pxc";
            cTbc.Name = "P x c";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 9;
            dgv2.Columns.Add(cTbc);
            dgv2.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv2.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.000";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_cxp";
            cTbc.Name = "C x p";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 9;
            dgv2.Columns.Add(cTbc);
            dgv2.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv2.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.000";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_af2";
            cTbc.Name = "Articolo consegnatario";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_fo2";
            cTbc.Name = "Consegnatario";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_prv";
            cTbc.Name = "P.consigliato";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 9;
            dgv2.Columns.Add(cTbc);
            dgv2.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv2.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "LiaMdy";
            cTbc.Name = "Modificato";
            cTbc.Width = 5;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);
        }

        private void SetDgv3()
        {
            dgv3.AutoGenerateColumns = false;
            //dgv3.VirtualMode = true;
            //dgv3.Dock = DockStyle.Fill;
            dgv3.AllowUserToAddRows = false;
            dgv3.ReadOnly = false;
            dgv3.AllowUserToDeleteRows = false;
            //dgv3.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "liv_lis";
            //cTbc.Name = "Listino";
            //cTbc.Width = 0;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = true;
            //dgv3.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "LivDes";
            cTbc.Name = "Listino";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv3.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "liv_prv";
            cTbc.Name = "Prezzo";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 9;
            dgv3.Columns.Add(cTbc);
            dgv3.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv3.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "liv_dti";
            cTbc.Name = "Data inizio";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = false;
            dgv3.Columns.Add(cTbc);
            dgv3.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "liv_dtf";
            cTbc.Name = "Data fine";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = false;
            dgv3.Columns.Add(cTbc);
            dgv3.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "liv_ann";
            cCbc.Name = "Annullato";
            cCbc.Width = 20;
            dgv3.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "liv_sta";
            cTbc.Name = "Stato";
            cTbc.Width = 15;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv3.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "LivMdy";
            cTbc.Name = "Modificato";
            cTbc.Width = 5;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv3.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "LivTip";
            cTbc.Name = "Tp";
            cTbc.Width = 25;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.Visible = false;
            cTbc.DefaultCellStyle.Font = new Font("Arial", 10.0F, GraphicsUnit.Pixel);
            cTbc.ToolTipText = "PC=p.consigliato PN=p.negozio PR=promozione Lx=listino negozio";
            dgv3.Columns.Add(cTbc);
        }

        private void SetDgv4()
        {
            dgv4.AutoGenerateColumns = false;
            //dgv4.VirtualMode = true;
            //dgv4.Dock = DockStyle.Fill;
            dgv4.AllowUserToAddRows = false;
            dgv4.ReadOnly = false;
            dgv4.AllowUserToDeleteRows = false;
            //dgv4.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "liv_lis";
            //cTbc.Name = "Listino";
            //cTbc.Width = 0;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = true;
            //dgv4.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "LivTip";
            cTbc.Name = "Tp";
            cTbc.Width = 25;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Font = new Font("Arial", 10.0F, GraphicsUnit.Pixel);
            cTbc.ToolTipText = "PC=p.consigliato PN=p.negozio PR=promozione Lx=listino negozio";
            dgv4.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "LivPrv";
            cTbc.Name = "Prezzo";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 9;
            dgv4.Columns.Add(cTbc);
            dgv4.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv4.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";
            dgv4.Columns[cTbc.Name].DefaultCellStyle.BackColor = Color.PaleGreen;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "LivMav";
            cTbc.Name = "Mrg val";
            cTbc.Width = 35;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 9;
            dgv4.Columns.Add(cTbc);
            dgv4.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv4.Columns[cTbc.Name].DefaultCellStyle.Font = new Font("Arial", 10.0F, GraphicsUnit.Pixel);
            dgv4.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "LivMap";
            cTbc.Name = "Mrg %";
            cTbc.Width = 35;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 9;
            dgv4.Columns.Add(cTbc);
            dgv4.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv4.Columns[cTbc.Name].DefaultCellStyle.Font = new Font("Arial", 10.0F, GraphicsUnit.Pixel);
            dgv4.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "LivMao";
            cTbc.Name = "Mrg obbiettivo";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 9;
            dgv4.Columns.Add(cTbc);
            dgv4.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv4.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "LivDti";
            cTbc.Name = "Data inizio";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = false;
            dgv4.Columns.Add(cTbc);
            dgv4.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";
            dgv4.Columns[cTbc.Name].DefaultCellStyle.Font = new Font("Arial", 10.0F, GraphicsUnit.Pixel);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "LivDtf";
            cTbc.Name = "Data fine";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = false;
            dgv4.Columns.Add(cTbc);
            dgv4.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";
            dgv4.Columns[cTbc.Name].DefaultCellStyle.Font = new Font("Arial", 10.0F, GraphicsUnit.Pixel);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "LivAnn";
            cCbc.Name = "Annullato";
            cCbc.Width = 20;
            cCbc.ReadOnly = false;
            dgv4.Columns.Add(cCbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "LivMdy";
            //cTbc.Name = "Modificato";
            //cTbc.Width = 5;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = true;
            //dgv4.Columns.Add(cTbc);
        }

        private static DataSet _dsTabsCache = null;
        public static void ClearTabsCache() { _dsTabsCache = null; }

        private void FillTabs(string strTab)
        {
            if (strTab != "") return; // Non ottimizziamo ricaricamenti parziali per ora

            string[] tables = {
                "TabStato", "TabIva", "TabUmi", "TabTipoGrammatura", "TabEcrLv1",
                "TabEcrLv2", "TabEcrLv3", "TabReparti", "TabBilance", "TabOrigine",
                "TabCalibro", "TabCategoria", "TabRepBilance", "TabListini",
                "TabEtichette", "TabMarchio", "TabEquPrezzi"
            };

            if (_dsTabsCache == null)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("SELECT * FROM TabStato WHERE tab_art=1;");
                sb.Append("SELECT * FROM TabIva ORDER BY tab_cod;");
                sb.Append("SELECT * FROM TabUmi ORDER BY tab_des;");
                sb.Append("SELECT * FROM TabTipoGrammatura ORDER BY tab_des;");
                sb.Append("SELECT * FROM TabEcrLv1 ORDER BY tab_des;");
                sb.Append("SELECT * FROM TabEcrLv2 ORDER BY tab_des;");
                sb.Append("SELECT * FROM TabEcrLv3 ORDER BY tab_des;");
                sb.Append("SELECT * FROM TabReparti WHERE tab_ann=0 ORDER BY tab_des;");
                sb.Append("SELECT * FROM TabBilance WHERE tab_ann=0;");
                sb.Append("SELECT * FROM TabOrigine ORDER BY tab_des;");
                sb.Append("SELECT * FROM TabCalibro;");
                sb.Append("SELECT * FROM TabCategoria;");
                sb.Append("SELECT * FROM TabRepBilance;");
                sb.Append("SELECT * FROM TabListini WHERE tab_ann=0 AND tab_tip<>'O';");
                sb.Append("SELECT * FROM TabEtichette WHERE tab_ann=0;");
                sb.Append("SELECT * FROM TabMarchio ORDER BY tab_des;");
                sb.Append("SELECT * FROM TabEquPrezzi;");

                DataSet dsRaw = _clsFun.FillDataSetSql(sb.ToString(), _strConSql);
                for (int i = 0; i < tables.Length; i++)
                {
                    if (dsRaw.Tables.Count <= i) break;
                    DataTable t = dsRaw.Tables[i];
                    t.TableName = tables[i];

                    // Aggiunta riga vuota dove necessario
                    if (t.TableName != "TabBilance" && t.TableName != "TabEcrLv2" && t.TableName != "TabEcrLv3")
                    {
                        DataRow x = t.NewRow();
                        x["tab_cod"] = "";
                        if (t.Columns.Contains("tab_des")) x["tab_des"] = (t.TableName == "TabUmi" || t.TableName == "TabTipoGrammatura") ? "  " : "  Non definito";
                        t.Rows.InsertAt(x, 0);
                    }
                }
                _dsTabsCache = dsRaw;
            }

            DataSet ds = _dsTabsCache.Copy();

            for (int i = 0; i < tables.Length; i++)
            {
                if (ds.Tables.Count <= i) break;
                DataTable t = ds.Tables[i];

                // Binding ai controlli
                switch (t.TableName)
                {
                    case "TabStato": BindCombo(cmbArtSta, t); break;
                    case "TabIva": BindCombo(cmbArtIva, t); break;
                    case "TabUmi": BindCombo(cmbArtUmi, t); break;
                    case "TabTipoGrammatura": BindCombo(cmbArtTgr, t); break;
                    case "TabEcrLv1": BindCombo(cmbArtEc1, t); break;
                    case "TabReparti": BindCombo(cmbArtRep, t); break;
                    case "TabBilance": if (t.Rows.Count > 0) groupBox1.Text = "Bilancia " + t.Rows[0]["tab_cod"] + " " + t.Rows[0]["tab_des"]; break;
                    case "TabOrigine": BindCombo(cmbArtOri, t); break;
                    case "TabCalibro": BindCombo(cmbArtCal, t); break;
                    case "TabCategoria": BindCombo(cmbArtCat, t); break;
                    case "TabRepBilance": BindCombo(cmbArtReb, t); break;
                    case "TabListini": BindCombo(cmbLisVen, t); break;
                    case "TabEtichette": BindCombo(cmbArtEti, t); break;
                    case "TabMarchio": BindCombo(cmbArtMar, t); break;
                    case "TabEquPrezzi": BindCombo(cmbArtEqp, t); break;
                }

                if (t.TableName.StartsWith("TabEcrLv"))
                {
                    if (_dasGen.Tables.Contains(t.TableName)) _dasGen.Tables.Remove(t.TableName);
                    _dasGen.Tables.Add(t.Copy());
                }
            }
        }

        private void BindCombo(ComboBox cmb, DataTable t)
        {
            cmb.DataSource = t;
            cmb.DisplayMember = "tab_des";
            cmb.ValueMember = "tab_cod";
            cmb.SelectedValue = "";
        }

        private void cmbArtEcr_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                ComboBox cmb = (ComboBox)sender;
                String s = "";

                if (cmb.Name == "cmbArtEc1")
                {
                    s = "tab_lv1='" + Convert.ToString(cmbArtEc1.SelectedValue) + "'";
                    if (_dasGen.Tables.Contains(TABECRLV2))
                    {
                        DataView v = new DataView(_dasGen.Tables[TABECRLV2], s, "", DataViewRowState.CurrentRows);
                        cmbArtEc2.DataSource = v;
                        cmbArtEc2.DisplayMember = "tab_des";
                        cmbArtEc2.ValueMember = "tab_cod";
                        cmbArtEc2.SelectedValue = "";
                    }
                }
                if (cmb.Name == "cmbArtEc2")
                {
                    s = "tab_lv1='" + Convert.ToString(cmbArtEc1.SelectedValue) + "' AND tab_lv2='" + Convert.ToString(cmbArtEc2.SelectedValue) + "'";
                    if (_dasGen.Tables.Contains(TABECRLV3))
                    {
                        DataView v = new DataView(_dasGen.Tables[TABECRLV3], s, "", DataViewRowState.CurrentRows);
                        cmbArtEc3.DataSource = v;
                        cmbArtEc3.DisplayMember = "tab_des";
                        cmbArtEc3.ValueMember = "tab_cod";
                        cmbArtEc3.SelectedValue = "";
                    }
                }
            }
            catch { }
        }

        private void cmbArtEcr_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ComboBox cmb = (ComboBox)sender;
                String s = "";

                if (cmb.Name == "cmbArtEc1")
                {
                    s = "tab_lv1='" + Convert.ToString(cmbArtEc1.SelectedValue) + "'";
                    if (_dasGen.Tables.Contains(TABECRLV2))
                    {
                        DataView v = new DataView(_dasGen.Tables[TABECRLV2], s, "", DataViewRowState.CurrentRows);
                        cmbArtEc2.DataSource = v;
                        cmbArtEc2.DisplayMember = "tab_des";
                        cmbArtEc2.ValueMember = "tab_cod";
                        cmbArtEc2.SelectedValue = "";
                    }
                }
                if (cmb.Name == "cmbArtEc2")
                {
                    s = "tab_lv1='" + Convert.ToString(cmbArtEc1.SelectedValue) + "' AND tab_lv2='" + Convert.ToString(cmbArtEc2.SelectedValue) + "'";
                    if (_dasGen.Tables.Contains(TABECRLV3))
                    {
                        DataView v = new DataView(_dasGen.Tables[TABECRLV3], s, "", DataViewRowState.CurrentRows);
                        cmbArtEc3.DataSource = v;
                        cmbArtEc3.DisplayMember = "tab_des";
                        cmbArtEc3.ValueMember = "tab_cod";
                        cmbArtEc3.SelectedValue = "";
                    }
                }
            }
            catch { }
        }

        private void FillArt(Boolean bolAncheSeVuoto)
        {
            _bolLoading = true;
            try
            {
                txtEan.Text = "";

                if (_strArtCod == _clsDef.CODNEW)
                {
                    txtArtCod.Text = _clsDef.CODNEW;
                }

                if (_strArtCod != _clsDef.CODNEW || bolAncheSeVuoto)
                {
                    CtrlPanels("");
                    if (_tabArt != null && _tabArt.Rows.Count > 0)
                    {
                        if (_tabArt.Columns.Contains("tmp_art") && !DBNull.Value.Equals(_tabArt.Rows[0]["tmp_art"]))
                            _strArtCod = Convert.ToString(_tabArt.Rows[0]["tmp_art"]);
                        else if (_tabArt.Columns.Contains("art_cod") && !DBNull.Value.Equals(_tabArt.Rows[0]["art_cod"]))
                            _strArtCod = Convert.ToString(_tabArt.Rows[0]["art_cod"]);
                    }

                    if (string.IsNullOrEmpty(_strArtCod) || _strArtCod == _clsDef.CODNEW)
                    {
                        if (!string.IsNullOrEmpty(txtArtCod.Text) && txtArtCod.Text.Trim() != _clsDef.CODNEW)
                            _strArtCod = txtArtCod.Text.Trim();
                    }

                    txtArtCod.Text = _strArtCod;
                    string sArt = string.IsNullOrEmpty(_strArtCod) ? "9999999999999" : _strArtCod;

                    string sBatch =
                        "SELECT * FROM AnaArticoli WHERE art_cod = '" + sArt + "'; " +
                        "SELECT * FROM " + TABANAEAN + " WHERE ean_art = '" + sArt + "' ORDER BY ean_dtm DESC; " +
                        "SELECT GesLisAcquisto.*, AnaFornitori.for_des AS CosFod FROM GesLisAcquisto LEFT OUTER JOIN AnaFornitori ON GesLisAcquisto.lia_for = AnaFornitori.for_cod WHERE GesLisAcquisto.lia_art = '" + sArt + "' ORDER BY lia_dti DESC; " +
                        "SELECT liv_idx, liv_lis, liv_art, liv_prv, liv_dti, liv_dtf, liv_day, tab_des AS LivDes, tab_tip AS LivTip, liv_sta, liv_ann FROM GesLisVendita LEFT OUTER JOIN TabListini ON GesLisVendita.liv_lis = TabListini.tab_cod WHERE liv_art = '" + sArt + "' AND liv_lis <> '" + _clsDef.LISOFF + "' ORDER BY liv_dti DESC, liv_lis; " +
                        "SELECT * FROM AnaArtIngredienti WHERE ing_art = '" + sArt + "'; " +
                        // OTT.2: ArtInOff come 6° statement — elimina la query separata sincrona
                        "SELECT GesOffArticoli.ofa_yea, GesOffArticoli.ofa_cod, GesOffArticoli.ofa_art, GesOffArticoli.ofa_val, GesOffArticoli.ofa_tip, GesOffTestate.oft_des, GesOffTestate.oft_dti, GesOffTestate.oft_dtf, GesOffTestate.oft_cod " +
                        "FROM GesOffArticoli INNER JOIN GesOffTestate ON GesOffArticoli.ofa_yea = GesOffTestate.oft_yea AND GesOffArticoli.ofa_cod = GesOffTestate.oft_cod " +
                        "WHERE GesOffArticoli.ofa_art = '" + sArt + "' AND GesOffArticoli.ofa_ann = 0 " +
                        "AND GesOffTestate.oft_dti <= " + _clsFun.DaySql(DateTime.Today) + " AND GesOffTestate.oft_dtf >= " + _clsFun.DaySql(DateTime.Today) + ";";

                    DataSet ds = _clsFun.FillDataSetSql(sBatch, _strConSql);

                    if (ds != null && ds.Tables.Count >= 5)
                    {
                        // 1. AnaArticoli
                        DataTable t = ds.Tables[0];
                        t.TableName = TABANAART;
                        if (t.Rows.Count > 0)
                        {
                            txtArtCod.Text = Convert.ToString(t.Rows[0]["art_cod"]);
                            txtArtDes.Text = Convert.ToString(t.Rows[0]["art_des"]);
                            txtArtDeb.Text = Convert.ToString(t.Rows[0]["art_deb"]);
                            txtArtPxc.Text = _clsFun.Dec2Txt(t.Rows[0]["art_pxc"], 0);
                            txtArtPne.Text = _clsFun.Dec2Txt(t.Rows[0]["art_pne"], 2);
                            txtArtGsc.Text = Convert.ToString(t.Rows[0]["art_gsc"]);
                            txtArtPlu.Text = Convert.ToString(t.Rows[0]["art_plu"]);
                            txtArtSfr.Text = _clsFun.Dec2Txt(t.Rows[0]["art_sfr"], 0);
                            txtArtTar.Text = _clsFun.Dec2Txt(t.Rows[0]["art_tar"], 0);
                            txtArtGia.Text = _clsFun.Dec2Txt(t.Rows[0]["art_gia"], 0);
                            txtArtTas.Text = Convert.ToString(t.Rows[0]["art_tas"]);
                            txtArtImg.Text = Convert.ToString(t.Rows[0]["art_img"]);

                            cmbArtSta.SelectedValue = _clsDef.STANOA;
                            if (!DBNull.Value.Equals(t.Rows[0]["art_sta"])) cmbArtSta.SelectedValue = t.Rows[0]["art_sta"];
                            if (!DBNull.Value.Equals(t.Rows[0]["art_iva"])) cmbArtIva.SelectedValue = t.Rows[0]["art_iva"];
                            if (!DBNull.Value.Equals(t.Rows[0]["art_eqp"])) cmbArtEqp.SelectedValue = t.Rows[0]["art_eqp"];
                            if (!DBNull.Value.Equals(t.Rows[0]["art_umi"])) cmbArtUmi.SelectedValue = t.Rows[0]["art_umi"];
                            if (!DBNull.Value.Equals(t.Rows[0]["art_tgr"])) cmbArtTgr.SelectedValue = t.Rows[0]["art_tgr"];
                            if (!DBNull.Value.Equals(t.Rows[0]["art_rep"])) cmbArtRep.SelectedValue = t.Rows[0]["art_rep"];
                            if (!DBNull.Value.Equals(t.Rows[0]["art_ec1"])) cmbArtEc1.SelectedValue = t.Rows[0]["art_ec1"];
                            if (!DBNull.Value.Equals(t.Rows[0]["art_ec2"])) cmbArtEc2.SelectedValue = t.Rows[0]["art_ec2"];
                            if (!DBNull.Value.Equals(t.Rows[0]["art_ec3"])) cmbArtEc3.SelectedValue = t.Rows[0]["art_ec3"];
                            if (!DBNull.Value.Equals(t.Rows[0]["art_ori"])) cmbArtOri.SelectedValue = t.Rows[0]["art_ori"];
                            if (!DBNull.Value.Equals(t.Rows[0]["art_cal"])) cmbArtCal.SelectedValue = t.Rows[0]["art_cal"];
                            if (!DBNull.Value.Equals(t.Rows[0]["art_reb"])) cmbArtReb.SelectedValue = t.Rows[0]["art_reb"];
                            if (!DBNull.Value.Equals(t.Rows[0]["art_cat"])) cmbArtCat.SelectedValue = t.Rows[0]["art_cat"];
                            if (!DBNull.Value.Equals(t.Rows[0]["art_eti"])) cmbArtEti.SelectedValue = t.Rows[0]["art_eti"];
                            if (!DBNull.Value.Equals(t.Rows[0]["art_mar"])) cmbArtMar.SelectedValue = t.Rows[0]["art_mar"];

                            chkArtBpz.Checked = !DBNull.Value.Equals(t.Rows[0]["art_bpz"]) && Convert.ToBoolean(t.Rows[0]["art_bpz"]);
                            chkArtBil.Checked = !DBNull.Value.Equals(t.Rows[0]["art_bil"]) && Convert.ToBoolean(t.Rows[0]["art_bil"]);

                            chkTra.Checked = false;
                            if (!DBNull.Value.Equals(t.Rows[0]["art_tra"]) && Convert.ToString(t.Rows[0]["art_tra"]).Trim() == "S")
                                chkTra.Checked = true;

                            chkArtWeb.Checked = !DBNull.Value.Equals(t.Rows[0]["art_web"]) && Convert.ToBoolean(t.Rows[0]["art_web"]);
                            chkArtCel.Checked = !DBNull.Value.Equals(t.Rows[0]["art_cel"]) && Convert.ToBoolean(t.Rows[0]["art_cel"]);

                            if (!DBNull.Value.Equals(t.Rows[0]["art_dti"]))
                                dtpArtDti.Value = Convert.ToDateTime(t.Rows[0]["art_dti"]);
                            ArtStato();
                        }
                        _tabArt = t.Copy();

                        // 2. AnaBarcode — OTT.4: salvato in _tabEan per riuso in SalvaEan (zero SELECT)
                        DataTable tEan = ds.Tables[1];
                        tEan.TableName = TABANAEAN;
                        foreach (DataRow r in tEan.Rows)
                        {
                            if (tEan.Columns.Contains("ean_bil") && DBNull.Value.Equals(r["ean_bil"])) r["ean_bil"] = false;
                            if (tEan.Columns.Contains("ean_ecp") && DBNull.Value.Equals(r["ean_ecp"])) r["ean_ecp"] = false;
                            if (tEan.Columns.Contains("ean_ann") && DBNull.Value.Equals(r["ean_ann"])) r["ean_ann"] = false;
                        }
                        _tabEan = tEan.Copy(); // cache per SalvaEan
                        tEan.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "EanMdy", Caption = "Mdy", MaxLength = 1, DefaultValue = "" });
                        dgv1.DataSource = new DataView(tEan, "", "ean_ean", DataViewRowState.CurrentRows);

                        // 3. GesLisAcquisto — OTT.4: salvato in _tabCos per riuso in SalvaCos (zero SELECT)
                        DataTable tCos = ds.Tables[2];
                        tCos.TableName = TABLISACQ;
                        foreach (DataRow y in tCos.Rows)
                        {
                            if (DBNull.Value.Equals(y["lia_ann"])) y["lia_ann"] = false;
                            if (!DBNull.Value.Equals(y["lia_dtf"]) && ((DateTime)y["lia_dtf"]).Year == 2050) y["lia_dtf"] = DBNull.Value;
                        }
                        _tabCos = tCos.Copy(); // cache per SalvaCos
                        tCos.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "LiaMdy", Caption = "Mdy", MaxLength = 1, DefaultValue = "" });
                        dgv2.DataSource = new DataView(tCos, "", "lia_dti DESC", DataViewRowState.CurrentRows);

                        // 4. GesLisVendita — OTT.4: salvato in _tabVen per riuso in SalvaVen (zero SELECT)
                        DataTable tVen = ds.Tables[3];
                        tVen.TableName = TABLISVEN;

                        // OTT.2: offerte già nel batch (tabella [5]) — nessuna query separata
                        if (ds.Tables.Count >= 6 && !string.IsNullOrEmpty(txtArtCod.Text))
                        {
                            DataTable tOff = ds.Tables[5];
                            if (tOff.Rows.Count > 0)
                            {
                                DataRow x = tVen.NewRow();
                                x["liv_lis"] = _clsDef.LISOFF;
                                x["liv_art"] = txtArtCod.Text;
                                x["liv_prv"] = tOff.Rows[0]["ofa_val"];
                                x["liv_dti"] = tOff.Rows[0]["oft_dti"];
                                x["liv_dtf"] = tOff.Rows[0]["oft_dtf"];
                                x["liv_ann"] = false;
                                x["LivTip"] = "O";
                                x["LivDes"] = "Offerta " + Convert.ToString(tOff.Rows[0]["ofa_cod"]);
                                tVen.Rows.Add(x);
                            }
                        }

                        _tabVen = tVen.Copy(); // cache per SalvaVen
                        tVen.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "LivMdy", Caption = "Mdy", MaxLength = 1, DefaultValue = "" });
                        foreach (DataRow y in tVen.Rows)
                        {
                            if (!DBNull.Value.Equals(y["liv_dtf"]) && ((DateTime)y["liv_dtf"]).Year == 2050) y["liv_dtf"] = DBNull.Value;
                        }
                        dgv3.DataSource = new DataView(tVen, "", "liv_dti DESC", DataViewRowState.CurrentRows);

                        FillVal();

                        // 5. AnaArtIngredienti
                        DataTable tIng = ds.Tables[4];
                        btnIng.BackColor = tIng.Rows.Count > 0 ? Color.PaleGreen : Color.Gainsboro;
                    }
                }
            }
            finally
            {
                _bolLoading = false;
                _bolIsDirty = false;
            }
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_bolIsDirty)
            {
                if (MessageBox.Show("Ci sono modifiche non salvate nella scheda articolo.\nVuoi ricaricare i dati dal database annullando le modifiche?", "RICARICA DATI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }
            }
            try
            {
                frmWait.ShowWait("Ricarica dati articolo " + txtArtCod.Text.Trim() + " in corso... ATTENDERE.");
                FillArt(false);
            }
            finally
            {
                frmWait.CloseWait();
            }
        }

        private void FillEan()
        {
            string s = "SELECT * FROM " + TABANAEAN + " WHERE ean_art = '" + _strArtCod + "' ORDER BY ean_dtm DESC";
            DataTable t = _clsFun.FillTabSql(TABANAEAN, s, false, _strConSql);

            foreach (DataRow r in t.Rows)
            {
                if (t.Columns.Contains("ean_bil") && DBNull.Value.Equals(r["ean_bil"]))
                    r["ean_bil"] = false;
                if (t.Columns.Contains("ean_ecp") && DBNull.Value.Equals(r["ean_ecp"]))
                    r["ean_ecp"] = false;
                if (t.Columns.Contains("ean_ann") && DBNull.Value.Equals(r["ean_ann"]))
                    r["ean_ann"] = false;
            }

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "EanMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            DataView v = new DataView(t, "", "ean_dtm DESC", DataViewRowState.CurrentRows);

            dgv1.DataSource = v;
        }

        private void FillCos()
        {
            string s = "SELECT ";
            s += "GesLisAcquisto.*, ";
            //s += "GesLisAcquisto.lia_art, ";
            //s += "GesLisAcquisto.lia_for, ";
            //s += "GesLisAcquisto.lia_dti, ";
            //s += "GesLisAcquisto.lia_tip, ";
            //s += "GesLisAcquisto.lia_arf, ";
            //s += "GesLisAcquisto.lia_cos, ";
            //s += "GesLisAcquisto.lia_pxc, ";
            //s += "GesLisAcquisto.lia_cxp, ";
            s += "AnaFornitori.for_des AS CosFod ";
            s += "FROM GesLisAcquisto LEFT OUTER JOIN AnaFornitori ON GesLisAcquisto.lia_for = AnaFornitori.for_cod ";
            s += "WHERE GesLisAcquisto.lia_art ='" + txtArtCod.Text + "' ";
            s += "ORDER BY lia_dti DESC";

            DataTable t = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "LiaMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            foreach (DataRow y in t.Rows)
            {
                if (DBNull.Value.Equals(y["lia_ann"]))
                    y["lia_ann"] = false;
                if (!DBNull.Value.Equals(y["lia_dtf"]) && ((DateTime)y["lia_dtf"]).Year == 2050)
                    y["lia_dtf"] = DBNull.Value;
            }

            DataView v = new DataView(t, "", "lia_dti DESC", DataViewRowState.CurrentRows);

            dgv2.DataSource = v;
        }

        private void FillVen()
        {
            DataRow x;

            string s = "SELECT ";
            s += "liv_idx, ";
            s += "liv_lis, ";
            s += "liv_art, ";
            s += "liv_prv, ";
            s += "liv_dti, ";
            s += "liv_dtf, ";
            s += "liv_day, ";
            s += "tab_des AS LivDes, ";
            s += "tab_tip AS LivTip, ";
            s += "liv_sta, ";
            s += "liv_ann ";
            s += "FROM GesLisVendita LEFT OUTER JOIN TabListini ON GesLisVendita.liv_lis = TabListini.tab_cod ";
            s += "WHERE liv_art ='" + txtArtCod.Text + "' AND liv_lis<>'" + _clsDef.LISOFF + "'";
            s += "ORDER BY liv_dti DESC, liv_lis";

            DataTable t = _clsFun.FillTabSql(TABLISVEN, s, false, _strConSql);

            if (txtArtCod.Text != "")
            {
                DataTable tOff = _clsQry.ArtInOff(txtArtCod.Text, DateTime.Today, _clsDef.DAYOUT);
                if (tOff.Rows.Count > 0)
                {
                    Console.WriteLine("aaaaaaaaa");

                    x = t.NewRow();
                    x["liv_lis"] = _clsDef.LISOFF;
                    x["liv_art"] = txtArtCod.Text;
                    x["liv_prv"] = tOff.Rows[0]["ofa_val"];
                    x["liv_dti"] = tOff.Rows[0]["oft_dti"];
                    x["liv_dtf"] = tOff.Rows[0]["oft_dtf"];
                    x["liv_ann"] = false;
                    x["LivTip"] = "O";

                    s = "Offerta " + (string)tOff.Rows[0]["ofa_cod"]; // +" " + ((DateTime)tOff.Rows[0]["oft_dti"]).ToString("dd/MM") + "-" + ((DateTime)tOff.Rows[0]["oft_dtf"]).ToString("dd/MM");
                    x["LivDes"] = s;
                    t.Rows.Add(x);
                }
            }
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "LivMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            foreach (DataRow y in t.Rows)
            {
                if (!DBNull.Value.Equals(y["liv_dtf"]) && ((DateTime)y["liv_dtf"]).Year == 2050)
                    y["liv_dtf"] = DBNull.Value;
            }

            DataView v = new DataView(t, "", "liv_dti DESC", DataViewRowState.CurrentRows);

            dgv3.DataSource = v;
        }

        private void FillVal()
        {
            try
            {
                DataTable tLiv = new clsGenTabTmp().TabTmpAnaArtPVendita("tVen");

                DataTable t = GetGridDataTable(dgv2);
                DateTime dDti = new DateTime(2000, 1, 1);
                decimal dVal = 0m;
                decimal dIva = 0m;
                DataRow x;
                ArrayList a = new ArrayList();

                txtArtCou.Text = "0";
                txtArtCol.Text = "0";
                txtArtCof.Text = "0";

                if (cmbArtIva.SelectedValue != null && cmbArtIva.DataSource is DataTable dtIva)
                {
                    DataRow[] j = dtIva.Select("tab_cod='" + cmbArtIva.SelectedValue.ToString().Replace("'", "''") + "'");
                    if (j.Length > 0 && !DBNull.Value.Equals(j[0]["tab_ali"]))
                        decimal.TryParse(Convert.ToString(j[0]["tab_ali"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dIva);
                }

                if (t != null)
                {
                    foreach (DataRow y in t.Rows)
                    {
                        if (y.RowState == DataRowState.Deleted) continue;
                        if (t.Columns.Contains("lia_ann") && !DBNull.Value.Equals(y["lia_ann"]) && !Convert.ToBoolean(y["lia_ann"]))
                        {
                            decimal liaCosVal = 0m;
                            if (!DBNull.Value.Equals(y["lia_cos"]))
                                decimal.TryParse(Convert.ToString(y["lia_cos"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out liaCosVal);

                            DateTime liaDtiVal = DateTime.MinValue;
                            if (!DBNull.Value.Equals(y["lia_dti"]))
                                DateTime.TryParse(Convert.ToString(y["lia_dti"]), out liaDtiVal);

                            if (liaDtiVal > dDti && liaCosVal > 0)
                            {
                                dDti = liaDtiVal;
                                dVal = liaCosVal;
                            }
                        }
                    }
                    txtArtCou.Text = dVal.ToString("#0.000");

                    dDti = new DateTime(2000, 1, 1);
                    dVal = 0m;
                    foreach (DataRow y in t.Rows)
                    {
                        if (y.RowState == DataRowState.Deleted) continue;
                        if (t.Columns.Contains("lia_ann") && !DBNull.Value.Equals(y["lia_ann"]) && !Convert.ToBoolean(y["lia_ann"]))
                        {
                            if (t.Columns.Contains("lia_tip") && Convert.ToString(y["lia_tip"]) == "F")
                            {
                                decimal liaCosVal = 0m;
                                if (!DBNull.Value.Equals(y["lia_cos"]))
                                    decimal.TryParse(Convert.ToString(y["lia_cos"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out liaCosVal);

                                DateTime liaDtiVal = DateTime.MinValue;
                                if (!DBNull.Value.Equals(y["lia_dti"]))
                                    DateTime.TryParse(Convert.ToString(y["lia_dti"]), out liaDtiVal);

                                if (liaDtiVal > dDti && liaCosVal > 0)
                                {
                                    dDti = liaDtiVal;
                                    dVal = liaCosVal;
                                }
                            }
                        }
                    }
                    txtArtCof.Text = dVal.ToString("#0.000");

                    dDti = new DateTime(2000, 1, 1);
                    dVal = 0m;
                    foreach (DataRow y in t.Rows)
                    {
                        if (y.RowState == DataRowState.Deleted) continue;
                        if (t.Columns.Contains("lia_ann") && !DBNull.Value.Equals(y["lia_ann"]) && !Convert.ToBoolean(y["lia_ann"]))
                        {
                            if (t.Columns.Contains("lia_tip") && Convert.ToString(y["lia_tip"]) == "L")
                            {
                                decimal liaCosVal = 0m;
                                if (!DBNull.Value.Equals(y["lia_cos"]))
                                    decimal.TryParse(Convert.ToString(y["lia_cos"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out liaCosVal);

                                DateTime liaDtiVal = DateTime.MinValue;
                                if (!DBNull.Value.Equals(y["lia_dti"]))
                                    DateTime.TryParse(Convert.ToString(y["lia_dti"]), out liaDtiVal);

                                if (liaDtiVal > dDti && liaCosVal > 0)
                                {
                                    dDti = liaDtiVal;
                                    dVal = liaCosVal;
                                }
                            }
                        }
                    }
                    txtArtCol.Text = dVal.ToString("#0.000");
                }

                DataView vVen = GetGridDataView(dgv3);
                if (vVen != null)
                {
                    // Priorità 1: Cerca record attivi non annullati per ciascun listino
                    foreach (DataRowView r in vVen)
                    {
                        if (r.Row.RowState == DataRowState.Deleted) continue;
                        if (!DBNull.Value.Equals(r["liv_lis"]) && Convert.ToString(r["liv_lis"]) != _clsDef.LISOFF)
                        {
                            string sLisCod = Convert.ToString(r["liv_lis"]).Trim();
                            bool isAnn = r.Row.Table.Columns.Contains("liv_ann") && !DBNull.Value.Equals(r["liv_ann"]) && Convert.ToBoolean(r["liv_ann"]);
                            if (!isAnn && !a.Contains(sLisCod))
                            {
                                a.Add(sLisCod);
                                x = tLiv.NewRow();
                                x["LivLis"] = sLisCod;

                                x["LivTip"] = "PC";
                                x["LivIdx"] = "000";
                                if (r.Row.Table.Columns.Contains("LivTip") && !DBNull.Value.Equals(r["LivTip"]) && Convert.ToString(r["LivTip"]) == "R")
                                {
                                    x["LivIdx"] = sLisCod;
                                    x["LivTip"] = "LR";
                                }
                                else if (sLisCod == _clsDef.LISPOS)
                                {
                                    x["LivIdx"] = sLisCod;
                                    x["LivTip"] = "PN";
                                }
                                else if (sLisCod == _clsDef.LISPRO)
                                {
                                    x["LivIdx"] = sLisCod;
                                    x["LivTip"] = "PR";
                                }
                                else if (sLisCod != _clsDef.LISFOR)
                                {
                                    x["LivIdx"] = sLisCod;
                                    x["LivTip"] = "L" + sLisCod.Substring(Math.Max(0, sLisCod.Length - 1));
                                }

                                decimal dPrvVal = 0m;
                                if (!DBNull.Value.Equals(r["liv_prv"]))
                                    decimal.TryParse(Convert.ToString(r["liv_prv"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dPrvVal);

                                x["LivPrv"] = dPrvVal;
                                x["LivMav"] = _clsFun.Margine(dPrvVal, _clsFun.Txt2Dec(txtArtCou.Text), dIva, _clsFun.Txt2Dec(txtArtSfr.Text), "V");
                                x["LivMap"] = _clsFun.Margine(dPrvVal, _clsFun.Txt2Dec(txtArtCou.Text), dIva, _clsFun.Txt2Dec(txtArtSfr.Text), "P");

                                DateTime dtDtiVal = DateTime.Today;
                                if (!DBNull.Value.Equals(r["liv_dti"]))
                                    DateTime.TryParse(Convert.ToString(r["liv_dti"]), out dtDtiVal);
                                x["LivDti"] = dtDtiVal;

                                if (!DBNull.Value.Equals(r["liv_dtf"]))
                                {
                                    DateTime dtDtfVal;
                                    if (DateTime.TryParse(Convert.ToString(r["liv_dtf"]), out dtDtfVal))
                                        x["LivDtf"] = dtDtfVal;
                                    else
                                        x["LivDtf"] = DBNull.Value;
                                }
                                else
                                    x["LivDtf"] = DBNull.Value;

                                x["LivAnn"] = isAnn;
                                tLiv.Rows.Add(x);
                            }
                        }
                    }
                }

                DataView v = new DataView(tLiv, "", "LivTip", DataViewRowState.CurrentRows);
                dgv4.DataSource = v;

                MargRica();
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArticolo.FillVal", ex.Message);
            }
        }

        private void MargRica()
        {
            try
            {
                string s = "";
                decimal dVal = 0m;
                if (cmbArtIva.SelectedValue != null && cmbArtIva.DataSource is DataTable dtIva)
                {
                    DataRow[] j = dtIva.Select("tab_cod='" + cmbArtIva.SelectedValue.ToString().Replace("'", "''") + "'");
                    if (j.Length > 0 && !DBNull.Value.Equals(j[0]["tab_ali"]))
                        decimal.TryParse(Convert.ToString(j[0]["tab_ali"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dVal);
                }

                decimal dPrv = 0m;
                s = PrezzoValido();
                string[] a = s.Split('|');
                dPrv = a.Length > 0 ? _clsFun.Txt2Dec(a[0]) : 0m;

                txtArtCouRiv.Text = _clsFun.Dec2Txt(_clsFun.Ricarico(dPrv, _clsFun.Txt2Dec(txtArtCou.Text), dVal, "V"), 2);
                txtArtCouRip.Text = _clsFun.Dec2Txt(_clsFun.Ricarico(dPrv, _clsFun.Txt2Dec(txtArtCou.Text), dVal, "P"), 2);
                txtArtCofRiv.Text = _clsFun.Dec2Txt(_clsFun.Ricarico(dPrv, _clsFun.Txt2Dec(txtArtCof.Text), dVal, "V"), 2);
                txtArtCofRip.Text = _clsFun.Dec2Txt(_clsFun.Ricarico(dPrv, _clsFun.Txt2Dec(txtArtCof.Text), dVal, "P"), 2);
                txtArtColRiv.Text = _clsFun.Dec2Txt(_clsFun.Ricarico(dPrv, _clsFun.Txt2Dec(txtArtCol.Text), dVal, "V"), 2);
                txtArtColRip.Text = _clsFun.Dec2Txt(_clsFun.Ricarico(dPrv, _clsFun.Txt2Dec(txtArtCol.Text), dVal, "P"), 2);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArticolo.MargRica", ex.Message);
            }
        }

        private string PrezzoValido()
        {
            decimal d = 0;
            string s = _clsDef.LISPOS;

            DataTable dt4 = GetGridDataTable(dgv4);
            if (dt4 != null && dt4.Rows.Count > 0)
            {
                // Priorità 1: Cerca Listino Cassa (001) con prezzo > 0
                DataRow[] rPos = dt4.Select("LivLis='" + _clsDef.LISPOS + "'");
                if (rPos.Length > 0 && !DBNull.Value.Equals(rPos[0]["LivPrv"]) && Convert.ToDecimal(rPos[0]["LivPrv"]) > 0)
                {
                    return Convert.ToDecimal(rPos[0]["LivPrv"]).ToString("#0.00") + "|" + _clsDef.LISPOS;
                }

                // Priorità 2: Qualsiasi listino PC o PN con prezzo > 0
                foreach (DataRow y in dt4.Rows)
                {
                    if (y.RowState != DataRowState.Deleted)
                    {
                        string tip = Convert.ToString(y["LivTip"]);
                        if (tip == "PC" || tip == "PN")
                        {
                            if (!DBNull.Value.Equals(y["LivPrv"]) && Convert.ToDecimal(y["LivPrv"]) > 0)
                            {
                                d = Convert.ToDecimal(y["LivPrv"]);
                                s = Convert.ToString(y["LivLis"]);
                                break;
                            }
                        }
                    }
                }
            }

            return d.ToString("#0.00") + "|" + s;
        }

        private DataView GetGridDataView(DataGridView dgv)
        {
            if (dgv == null || dgv.DataSource == null) return null;
            if (dgv.DataSource is DataView dv) return dv;
            if (dgv.DataSource is BindingSource bs)
            {
                if (bs.DataSource is DataView bsdv) return bsdv;
                if (bs.DataSource is DataTable bsdt) return bsdt.DefaultView;
                if (bs.List is DataView bsldv) return bsldv;
            }
            if (dgv.DataSource is DataTable dt) return dt.DefaultView;
            return null;
        }

        private DataTable GetGridDataTable(DataGridView dgv)
        {
            DataView dv = GetGridDataView(dgv);
            if (dv != null) return dv.Table;
            if (dgv.DataSource is DataTable dt) return dt;
            return null;
        }

        private Boolean Salva()
        {
            string sCodCheck = txtArtCod.Text != null ? txtArtCod.Text.Trim() : "";

            if (string.IsNullOrEmpty(sCodCheck))
            {
                return false;
            }
            _strArtCod = sCodCheck;

            bool isNewArticle = (_tabArt == null || _tabArt.Rows.Count == 0 || !_tabArt.Columns.Contains("art_cod") || sCodCheck == _clsDef.CODNEW || sCodCheck == "NEW" || sCodCheck == "0000000");

            try
            {
                if (dgv1 != null)
                {
                    dgv1.EndEdit();
                    if (dgv1.DataSource != null)
                        this.BindingContext[dgv1.DataSource]?.EndCurrentEdit();
                }
                if (dgv2 != null)
                {
                    dgv2.EndEdit();
                    if (dgv2.DataSource != null)
                        this.BindingContext[dgv2.DataSource]?.EndCurrentEdit();
                }
                if (dgv3 != null)
                {
                    dgv3.EndEdit();
                    if (dgv3.DataSource != null)
                        this.BindingContext[dgv3.DataSource]?.EndCurrentEdit();
                }
                if (dgv4 != null)
                {
                    dgv4.EndEdit();
                    if (dgv4.DataSource != null)
                        this.BindingContext[dgv4.DataSource]?.EndCurrentEdit();

                    // CRITICO: Sincronizza SEMPRE tutte le righe dei prezzi correnti (dgv4) in GesLisVendita (dgv3)
                    DataTable dt4 = GetGridDataTable(dgv4);
                    if (dt4 != null)
                    {
                        foreach (DataRow r4 in dt4.Rows)
                        {
                            if (r4.RowState != DataRowState.Deleted)
                                LivAggPre(r4);
                        }
                    }
                }
            }
            catch { }

            bool hasGridChanges = false;
            DataTable dt1Check = GetGridDataTable(dgv1);
            if (dt1Check != null)
            {
                foreach (DataRow r in dt1Check.Rows)
                {
                    if (Convert.ToString(r["EanMdy"]) == "S")
                    {
                        hasGridChanges = true;
                        break;
                    }
                }
            }
            DataTable dt2Check = GetGridDataTable(dgv2);
            if (dt2Check != null && !hasGridChanges)
            {
                foreach (DataRow r in dt2Check.Rows)
                {
                    if (Convert.ToString(r["LiaMdy"]) == "S")
                    {
                        hasGridChanges = true;
                        break;
                    }
                }
            }
            DataTable dt3Check = GetGridDataTable(dgv3);
            if (dt3Check != null && !hasGridChanges)
            {
                foreach (DataRow r in dt3Check.Rows)
                {
                    if (Convert.ToString(r["LivMdy"]) == "S")
                    {
                        hasGridChanges = true;
                        break;
                    }
                }
            }

            if (!_bolIsDirty && !hasGridChanges)
                return true;

            _bolIsDirty = true;

            Boolean b = false;
            string sMsg = "";

            if (string.IsNullOrEmpty(txtArtDes.Text.Trim()))
            {
                sMsg += "Descrizione non definita, " + _clsDef.CRLF;
                b = false;
            }
            if (cmbArtEc1.SelectedValue == null || cmbArtEc2.SelectedValue == null || cmbArtEc3.SelectedValue == null)
                sMsg += "Merceologia non completa, " + _clsDef.CRLF;
            else if (cmbArtEc1.SelectedValue.ToString() == "" || cmbArtEc2.SelectedValue.ToString() == "" || cmbArtEc3.SelectedValue.ToString() == "")
                sMsg += "Merceologia non completa, " + _clsDef.CRLF;
            if (cmbArtRep.SelectedValue == null || cmbArtRep.SelectedValue.ToString() == "")
                sMsg += "Reparto non definito, " + _clsDef.CRLF;
            if (cmbArtIva.SelectedValue == null || cmbArtIva.SelectedValue.ToString() == "")
                sMsg += "IVA non definito, " + _clsDef.CRLF;
            if (cmbArtSta.SelectedValue == null || cmbArtSta.SelectedValue.ToString() == "")
                sMsg += "Stato articolo non definito, " + _clsDef.CRLF;
            if (cmbArtTgr.SelectedValue == null || cmbArtTgr.SelectedValue.ToString() == "")
                sMsg += "Tipo grammatura non definito, " + _clsDef.CRLF;
            if (cmbArtUmi.SelectedValue == null || cmbArtUmi.SelectedValue.ToString() == "")
                sMsg += "Unità di misura non definito, " + _clsDef.CRLF;
            else
                b = true;

            if (sMsg != "")
            {
                sMsg = sMsg.Substring(0, sMsg.Length - 2);
                MessageBox.Show(this, sMsg, "DATI MANCANTI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!b)
                return false;

            string sVarAgg = "";

            string sCurCod = txtArtCod.Text != null ? txtArtCod.Text.Trim() : "";
            if (!isNewArticle)
                isNewArticle = (sCurCod == _clsDef.CODNEW || sCurCod == "NEW" || sCurCod == "0000000");

            if (isNewArticle || sCurCod == _clsDef.CODNEW || sCurCod == "NEW" || sCurCod == "0000000")
            {
                txtArtCod.Text = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaArticoli, 7, _strConSql);
                _strArtCod = txtArtCod.Text;
                isNewArticle = true;

                // Propaga il nuovo codice generato alle righe in memoria
                DataTable dt1Sync = GetGridDataTable(dgv1);
                if (dt1Sync != null)
                {
                    foreach (DataRow r in dt1Sync.Rows)
                        if (r.RowState != DataRowState.Deleted) r["ean_art"] = _strArtCod;
                }
                DataTable dt2Sync = GetGridDataTable(dgv2);
                if (dt2Sync != null)
                {
                    foreach (DataRow r in dt2Sync.Rows)
                        if (r.RowState != DataRowState.Deleted) r["lia_art"] = _strArtCod;
                }
                DataTable dt3Sync = GetGridDataTable(dgv3);
                if (dt3Sync != null)
                {
                    foreach (DataRow r in dt3Sync.Rows)
                        if (r.RowState != DataRowState.Deleted) r["liv_art"] = _strArtCod;
                }
            }

            // SOLUZIONE 2: Riuso _tabArt in memoria — elimina SELECT ridondante su AnaArticoli
            DataTable t;
            if (isNewArticle)
            {
                // Nuovo articolo: legge schema dal DB con codice impossibile
                t = _clsFun.FillTabSql(TABANAART, "SELECT * FROM AnaArticoli WHERE art_cod='9999999999999'", false, _strConSql);
            }
            else
            {
                // Articolo esistente: clone schema + dati da _tabArt (ZERO roundtrip DB)
                t = _tabArt.Clone();
                DataRow importedRow = t.NewRow();
                foreach (DataColumn col in _tabArt.Columns)
                    importedRow[col.ColumnName] = _tabArt.Rows[0][col.ColumnName];
                t.Rows.Add(importedRow);
            }

            // Garantisce MaxLength aggiornato per migrazione V003
            if (t.Columns.Contains("art_img") && t.Columns["art_img"].MaxLength < 255)
                t.Columns["art_img"].MaxLength = 255;

            string s;
            DataRow y = t.NewRow();
            y["art_cod"] = txtArtCod.Text;
            y["art_des"] = txtArtDes.Text;
            y["art_deb"] = txtArtDeb.Text;
            y["art_sta"] = Convert.ToString(cmbArtSta.SelectedValue);

            y["art_iva"] = "";
            if (cmbArtIva.SelectedValue != null)
                y["art_iva"] = cmbArtIva.SelectedValue.ToString();

            y["art_eqp"] = "";
            if (cmbArtEqp.SelectedValue != null)
                y["art_eqp"] = cmbArtEqp.SelectedValue.ToString();

            y["art_umi"] = "";
            if (cmbArtUmi.SelectedValue != null)
                y["art_umi"] = cmbArtUmi.SelectedValue.ToString();
            y["art_tgr"] = "";
            if (cmbArtTgr.SelectedValue != null)
                y["art_tgr"] = cmbArtTgr.SelectedValue.ToString();

            y["art_pxc"] = _clsFun.Txt2Dec(txtArtPxc.Text);
            y["art_pne"] = _clsFun.Txt2Dec(txtArtPne.Text);

            y["art_rep"] = "";
            if (cmbArtRep.SelectedValue != null)
                y["art_rep"] = cmbArtRep.SelectedValue.ToString();

            y["art_ec1"] = Convert.ToString(cmbArtEc1.SelectedValue);
            if (cmbArtEc2.SelectedValue != null)
                y["art_ec2"] = cmbArtEc2.SelectedValue.ToString();
            if (cmbArtEc3.SelectedValue != null)
                y["art_ec3"] = cmbArtEc3.SelectedValue.ToString();
            if (cmbArtOri.SelectedValue != null)
                y["art_ori"] = cmbArtOri.SelectedValue.ToString();
            if (cmbArtCal.SelectedValue != null)
                y["art_cal"] = cmbArtCal.SelectedValue.ToString();

            y["art_reb"] = "";
            if (cmbArtReb.SelectedValue != null)
                y["art_reb"] = cmbArtReb.SelectedValue.ToString();

            if (cmbArtCat.SelectedValue != null)
                y["art_cat"] = cmbArtCat.SelectedValue.ToString();

            y["art_tra"] = "";
            if (chkTra.Checked)
                y["art_tra"] = "S";

            y["art_gsc"] = _clsFun.Txt2Dec(txtArtGsc.Text);
            y["art_bil"] = chkArtBil.Checked;
            y["art_web"] = chkArtWeb.Checked;
            y["art_cel"] = chkArtCel.Checked;

            y["art_plu"] = "";
            if (txtArtPlu.Text != "" && _clsFun.Numerico(txtArtPlu.Text))
                y["art_plu"] = txtArtPlu.Text.PadLeft(4, Convert.ToChar("0"));

            y["art_sfr"] = 0;
            if (_clsFun.Numerico(txtArtSfr.Text))
                y["art_sfr"] = _clsFun.Txt2Dec(txtArtSfr.Text);

            y["art_tar"] = 0;
            if (_clsFun.Numerico(txtArtTar.Text))
                y["art_tar"] = _clsFun.Txt2Dec(txtArtTar.Text);

            y["art_gia"] = 0;
            if (_clsFun.Numerico(txtArtGia.Text))
                y["art_gia"] = _clsFun.Txt2Dec(txtArtGia.Text);

            y["art_eti"] = "";
            if (cmbArtEti.SelectedValue != null)
                y["art_eti"] = cmbArtEti.SelectedValue.ToString();

            y["art_mar"] = "";
            if (cmbArtMar.SelectedValue != null)
                y["art_mar"] = cmbArtMar.SelectedValue.ToString();

            y["art_tas"] = _clsFun.Txt2Dec(txtArtTas.Text);
            y["art_img"] = txtArtImg.Text;
            y["art_bpz"] = chkArtBpz.Checked;
            y["art_dtm"] = DateTime.Now.ToString("dd/MM/yyyy");

            y["art_cos"] = 0;
            DataView vCos = GetGridDataView(dgv2);
            if (vCos != null && vCos.Count > 0 && !DBNull.Value.Equals(vCos[0]["lia_cos"]))
            {
                try { y["art_cos"] = Convert.ToDecimal(vCos[0]["lia_cos"]); } catch { }
            }

            y["art_prv"] = 0;
            DataTable dt4Save = GetGridDataTable(dgv4);
            if (dt4Save != null && dt4Save.Rows.Count > 0)
            {
                DataRow[] rowsPos = dt4Save.Select("LivLis='" + _clsDef.LISPOS + "'");
                if (rowsPos.Length > 0 && !DBNull.Value.Equals(rowsPos[0]["LivPrv"]) && Convert.ToDecimal(rowsPos[0]["LivPrv"]) > 0)
                {
                    y["art_prv"] = Convert.ToDecimal(rowsPos[0]["LivPrv"]);
                }
                if (Convert.ToDecimal(y["art_prv"]) == 0)
                {
                    foreach (DataRow r4 in dt4Save.Rows)
                    {
                        if (r4.RowState != DataRowState.Deleted && !DBNull.Value.Equals(r4["LivPrv"]) && Convert.ToDecimal(r4["LivPrv"]) > 0)
                        {
                            y["art_prv"] = Convert.ToDecimal(r4["LivPrv"]);
                            break;
                        }
                    }
                }
            }

            if (Convert.ToDecimal(y["art_prv"]) == 0)
            {
                DataView vVen = GetGridDataView(dgv3);
                if (vVen != null && vVen.Count > 0)
                {
                    decimal d1 = 0;
                    decimal d2 = 0;

                    foreach (DataRowView r in vVen)
                    {
                        if (!DBNull.Value.Equals(r["liv_prv"]) && (!r.Row.Table.Columns.Contains("liv_ann") || DBNull.Value.Equals(r["liv_ann"]) || !Convert.ToBoolean(r["liv_ann"])))
                        {
                            decimal prvVal = Convert.ToDecimal(r["liv_prv"]);
                            string lisCod = Convert.ToString(r["liv_lis"]);
                            if (lisCod == _clsDef.LISPOS && prvVal > 0 && d1 == 0)
                            {
                                d1 = prvVal;
                                break;
                            }
                            else if (lisCod == _clsDef.LISFOR && prvVal > 0 && d2 == 0)
                                d2 = prvVal;
                        }
                    }

                    y["art_prv"] = d1 > 0 ? d1 : d2;
                }
            }

            if (isNewArticle)
            {
                y["art_dti"] = DateTime.Today;
                s = _clsFun.SqlInsertRow(TABANAART, t, y);
            }
            else
            {
                ArrayList aWhe = new ArrayList();
                aWhe.Add("art_cod");
                ArrayList aExl = new ArrayList();
                aExl.Add("art_dti");
                aExl.Add("art_dtm");
                s = _clsFun.SqlUpdRow(TABANAART, t, t.Rows[0], y, aWhe, aExl);
            }

            if (s != "")
            {
                _clsFun.SqlWrite(s, _strConSql);
                Variazioni(txtArtCod.Text, s, "Anagrafica articolo", _clsDef.VARALL);

                sVarAgg += "A";
                _strSqlArt = s;

                string[] aLog = {
                        "ANAART",                                     //  log_tip
                        Convert.ToString(y["art_cod"]),               //  log_art
                        "",                                           //  log_ean
                        Convert.ToString(y["art_des"]),               //  log_des
                        "",                                           //  log_fil
                        Convert.ToDecimal(y["art_cos"]).ToString(),   //  log_acq
                        Convert.ToDecimal(y["art_prv"]).ToString(),   //  log_prv
                        ""                                            //  log_msg
                            };
                _clsQry.LogSql(aLog);
            }
            if (t.Rows.Count > 0)
                t.Rows[0].Delete();
            t.Rows.Add(y);

            _tabArt = t.Copy();

            bool hasEanChanges = isNewArticle;
            dt1Check = GetGridDataTable(dgv1);
            if (dt1Check != null && dt1Check.Rows.Count > 0)
            {
                foreach (DataRow r in dt1Check.Rows)
                {
                    if (Convert.ToString(r["EanMdy"]) == "S" || isNewArticle || !string.IsNullOrEmpty(Convert.ToString(r["ean_ean"]))) { hasEanChanges = true; break; }
                }
            }
            if (hasEanChanges && SalvaEan(isNewArticle))
                sVarAgg += "E";

            bool hasCosChanges = isNewArticle;
            dt2Check = GetGridDataTable(dgv2);
            if (dt2Check != null && dt2Check.Rows.Count > 0)
            {
                foreach (DataRow r in dt2Check.Rows)
                {
                    if (Convert.ToString(r["LiaMdy"]) == "S" || isNewArticle) { hasCosChanges = true; break; }
                }
            }
            if (hasCosChanges && SalvaCos(isNewArticle))
                sVarAgg += "C";

            bool hasVenChanges = isNewArticle;
            dt3Check = GetGridDataTable(dgv3);
            if (dt3Check != null && dt3Check.Rows.Count > 0)
            {
                foreach (DataRow r in dt3Check.Rows)
                {
                    if (Convert.ToString(r["LivMdy"]) == "S" || isNewArticle) { hasVenChanges = true; break; }
                }
            }
            if (hasVenChanges && SalvaVen(isNewArticle))
                sVarAgg += "V";

            // SOLUZIONE 3: DivNegozi asincrona — I/O file su disco/rete non blocca più la GUI
            if (_strPar016PathDivNegozi != "" && sVarAgg != "")
                System.Threading.ThreadPool.QueueUserWorkItem(_ => { try { DivNegozi(); } catch { } });

            _bolIsDirty = false;
            return b;
        }

        private Boolean SalvaEan(bool isNewArticle = false)
        {
            // OTT.4: riuso _tabEan in memoria — nessun SELECT su AnaBarcode
            DataTable t = (_tabEan != null && _tabEan.Columns.Count > 0 && !isNewArticle)
                ? _tabEan.Copy()
                : _clsFun.FillTabSql(TABANAEAN, "SELECT * FROM " + TABANAEAN + " WHERE ean_art = '" + _strArtCod + "' ORDER BY ean_dtm DESC", false, _strConSql);
            DataRow[] j;
            DataRow x;
            Boolean bVar = false;

            ArrayList aWhe = new ArrayList();
            aWhe.Add("ean_ean");
            ArrayList aExl = new ArrayList();
            aExl.Add("ean_dti");
            aExl.Add("ean_dtm");
            string s = "";


            // OTT.1: raccogli tutti i barcode modificati e fai UN SOLO SELECT IN(...) per il check duplicati
            DataTable dt1 = GetGridDataTable(dgv1);
            var modifiedEans = new List<string>();
            if (dt1 != null)
                foreach (DataRow y in dt1.Rows)
                    if (isNewArticle || Convert.ToString(y["EanMdy"]) == "S" || string.IsNullOrEmpty(Convert.ToString(y["ean_art"])) || Convert.ToString(y["ean_art"]) == _clsDef.CODNEW || Convert.ToString(y["ean_art"]) == "NEW")
                        if (!string.IsNullOrEmpty(Convert.ToString(y["ean_ean"])))
                            modifiedEans.Add(Convert.ToString(y["ean_ean"]));

            // Pre-load barcode duplicati su altri articoli con una sola query IN(...)
            var dupLookup = new Dictionary<string, DataRow>(); // ean_ean -> row su altro art.
            if (modifiedEans.Count > 0)
            {
                string inList = string.Join(",", modifiedEans.Select(e => "'" + e.Replace("'", "''") + "'"));
                string sDup = "SELECT * FROM " + TABANAEAN + " WHERE ean_art<>" + (isNewArticle ? "'9999999999'" : "'" + _strArtCod + "'") + " AND ean_ean IN (" + inList + ")";
                DataTable tDup = _clsFun.FillTabSql(TABANAEAN, sDup, false, _strConSql);
                foreach (DataRow r in tDup.Rows)
                {
                    string eanKey = Convert.ToString(r["ean_ean"]);
                    if (!dupLookup.ContainsKey(eanKey)) dupLookup[eanKey] = r;
                }
            }

            // Batch: collect all SQL statements for a single connection
            var batchSql = new System.Collections.Generic.List<string>();

            if (dt1 != null)
            {
                foreach (DataRow y in dt1.Rows)
                {
                    string eanVal = Convert.ToString(y["ean_ean"]).Trim();
                    if (string.IsNullOrEmpty(eanVal)) continue;

                    if (isNewArticle || Convert.ToString(y["EanMdy"]) == "S" || string.IsNullOrEmpty(Convert.ToString(y["ean_art"])) || Convert.ToString(y["ean_art"]) == _clsDef.CODNEW || Convert.ToString(y["ean_art"]) == "NEW")
                    {
                        y["ean_art"] = txtArtCod.Text;
                        string eanKey = Convert.ToString(y["ean_ean"]);

                        // OTT.1: lookup in-memory invece di SELECT per ogni barcode
                        if (dupLookup.ContainsKey(eanKey))
                        {
                            x = t.NewRow();
                            x["ean_art"] = y["ean_art"];
                            x["ean_ean"] = y["ean_ean"];
                            x["ean_qta"] = y["ean_qta"];
                            x["ean_prv"] = y["ean_prv"];
                            x["ean_dti"] = y["ean_dti"];
                            x["ean_dtm"] = DateTime.Today;
                            x["ean_ann"] = y["ean_ann"];
                            x["ean_bil"] = y["ean_bil"];
                            x["ean_ecp"] = y["ean_ecp"];

                            s = _clsFun.SqlUpdRow(TABANAEAN, t, dupLookup[eanKey], x, aWhe, aExl);
                            Variazioni(txtArtCod.Text, "Spostato barcode " + eanKey, "Anagrafica barcode", _clsDef.VARALL);
                        }
                        else
                        {
                            s = "ean_ean='" + y["ean_ean"] + "'";
                            j = t.Select(s);
                            if (j.Length > 0)
                            {
                                y["ean_dtm"] = DateTime.Today;
                                s = _clsFun.SqlUpdRow(TABANAEAN, t, j[0], y, aWhe, aExl);
                            }
                            else
                                s = _clsFun.SqlInsertRow(TABANAEAN, t, y);
                        }
                        if (s != "")
                        {
                            batchSql.Add(s);
                            bVar = true;
                            string[] aLog = { "ANAEAN", Convert.ToString(y["ean_art"]), Convert.ToString(y["ean_ean"]), "", "", "", "", "" };
                            _clsQry.LogSql(aLog);
                        }
                    }
                }

                // Execute all EAN SQL with a single DB connection
                if (batchSql.Count > 0)
                    _clsFun.SqlWriteBatch(batchSql, _strConSql);

                if (bVar)
                {
                    Variazioni(txtArtCod.Text, "Forza", "Anagrafica articolo", _clsDef.VARPOS);
                }
            }

            return bVar;
        }



        private Boolean SalvaCos(bool isNewArticle = false)
        {
            Boolean bVar = false;
            if (_strArtCod != "")
            {
                // OTT.4: riuso _tabCos in memoria — nessun SELECT su GesLisAcquisto
                DataTable t = (_tabCos != null && _tabCos.Columns.Count > 0 && !isNewArticle)
                    ? _tabCos.Copy()
                    : _clsFun.FillTabSql(TABLISACQ, "SELECT * FROM GesLisAcquisto WHERE lia_art = '" + _strArtCod + "' ORDER BY lia_for, lia_dti DESC", false, _strConSql);
                DataRow[] j;
                string s = "";


                ArrayList aWhe = new ArrayList();
                aWhe.Add("lia_tip");
                aWhe.Add("lia_for");
                aWhe.Add("lia_arf");
                aWhe.Add("lia_dti");
                aWhe.Add("lia_dtf");
                ArrayList aExl = new ArrayList();

                // Batch: collect all SQL statements to write with a single connection
                var batchSql = new System.Collections.Generic.List<string>();
                string lastVarSql = "";

                DataTable dt2 = GetGridDataTable(dgv2);
                if (dt2 != null)
                {
                    foreach (DataRow y in dt2.Rows)
                    {
                        if (isNewArticle || Convert.ToString(y["LiaMdy"]) == "S" || string.IsNullOrEmpty(Convert.ToString(y["lia_art"])) || Convert.ToString(y["lia_art"]) == _clsDef.CODNEW)
                        {
                            if (DBNull.Value.Equals(y["lia_ann"]))
                                y["lia_ann"] = false;
                            if (DBNull.Value.Equals(y["lia_arf"]))
                                y["lia_arf"] = "";

                            if (Convert.ToString(y["lia_arf"]).Trim() != "" && !Convert.ToBoolean(y["lia_ann"]))
                            {
                                s = "SELECT * FROM GesLisAcquisto WHERE ";
                                s += "lia_for='" + Convert.ToString(y["lia_for"]) + "' AND ";
                                s += "lia_art<>'" + _strArtCod + "' AND ";
                                s += "lia_arf='" + Convert.ToString(y["lia_arf"]) + "' ";
                                s += "ORDER BY lia_dti DESC";
                                DataTable tTmp = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);
                                if (tTmp.Rows.Count > 0)
                                {
                                    foreach (DataRow yy in tTmp.Rows)
                                    {
                                        batchSql.Add("DELETE FROM GesLisAcquisto WHERE lia_idx=" + Convert.ToString(yy["lia_idx"]));

                                        string[] aLog = {
                                        "ANAART",
                                        Convert.ToString(yy["lia_art"]),
                                        "",
                                        "",
                                        "",
                                        Convert.ToDecimal(yy["lia_cos"]).ToString(),
                                        "",
                                        "Spostamento arf " + Convert.ToString(yy["lia_arf"]) + " su " + _strArtCod
                                        };
                                        _clsQry.LogSql(aLog);
                                    }
                                }
                            }

                            y["lia_art"] = _strArtCod;

                            if (DBNull.Value.Equals(y["lia_prv"]))
                                y["lia_prv"] = 0;

                            s = "lia_tip='" + y["lia_tip"] + "' AND lia_art='" + _strArtCod + "' AND lia_for='" + y["lia_for"] + "' AND lia_arf='" + y["lia_arf"] + "' AND lia_dti=" + _clsFun.DayMdb((DateTime)y["lia_dti"]);
                            j = t.Select(s);
                            if (j.Length > 0)
                            {
                                s = _clsFun.SqlUpdRowIdx(TABLISACQ, t, j[0], y, aExl);
                            }
                            else
                                s = _clsFun.SqlInsertRow(TABLISACQ, t, y);

                            if (s != "")
                            {
                                batchSql.Add(s);
                                lastVarSql = s;
                                bVar = true;
                            }
                        }
                    }
                }

                // Execute all cost SQL with a single DB connection
                if (batchSql.Count > 0)
                {
                    _clsFun.SqlWriteBatch(batchSql, _strConSql);
                    if (!string.IsNullOrEmpty(lastVarSql))
                        Variazioni(txtArtCod.Text, lastVarSql, "Anagrafica costi", _clsDef.VARALL);
                }
            }

            return bVar;
        }

        private Boolean SalvaVen(bool isNewArticle = false)
        {
            Boolean bVar = false;

            Boolean b = false;
            // OTT.4: riuso _tabVen in memoria — nessun SELECT su GesLisVendita
            DataTable t = (_tabVen != null && _tabVen.Columns.Count > 0 && !isNewArticle)
                ? _tabVen.Copy()
                : _clsFun.FillTabSql(TABLISVEN, "SELECT * FROM GesLisVendita WHERE liv_art = '" + _strArtCod + "' ORDER BY liv_dti DESC", false, _strConSql);
            DataRow[] j;
            string s = "";


            if (_strArtCod != "")
            {
                ArrayList aWhe = new ArrayList();
                aWhe.Add("liv_lis");
                aWhe.Add("liv_art");
                aWhe.Add("liv_dti");
                ArrayList aExl = new ArrayList();

                // Batch: collect all SQL statements to write with a single connection
                var batchSql = new System.Collections.Generic.List<string>();
                string lastVarSql = "";

                DataTable dt3 = GetGridDataTable(dgv3);
                if (dt3 != null)
                {
                    foreach (DataRow y in dt3.Rows)
                    {
                        if (Convert.ToString(y["LivTip"]) != "O" && (Convert.ToString(y["LivMdy"]) == "S" || isNewArticle || string.IsNullOrEmpty(Convert.ToString(y["liv_art"])) || Convert.ToString(y["liv_art"]) == _clsDef.CODNEW))
                        {
                            y["liv_art"] = _strArtCod;

                            if (DBNull.Value.Equals(y["liv_dtf"]))
                                y["liv_dtf"] = _clsDef.DAYOUT;

                            decimal currPrv = !DBNull.Value.Equals(y["liv_prv"]) ? Convert.ToDecimal(y["liv_prv"]) : 0m;
                            int livIdx = (!DBNull.Value.Equals(y["liv_idx"]) && Convert.ToInt32(y["liv_idx"]) > 0) ? Convert.ToInt32(y["liv_idx"]) : 0;

                            if (!isNewArticle && livIdx > 0)
                            {
                                DataRow[] origRows = t.Select("liv_idx = " + livIdx);
                                if (origRows.Length > 0)
                                {
                                    DataRow origRow = origRows[0];
                                    DateTime origDti = !DBNull.Value.Equals(origRow["liv_dti"]) ? Convert.ToDateTime(origRow["liv_dti"]).Date : DateTime.Today;
                                    decimal origPrv = !DBNull.Value.Equals(origRow["liv_prv"]) ? Convert.ToDecimal(origRow["liv_prv"]) : 0m;

                                    if (currPrv != origPrv)
                                    {
                                        if (origDti < DateTime.Today)
                                        {
                                            // 1. Chiude il vecchio record a ieri
                                            DateTime dIeri = DateTime.Today.AddDays(-1);
                                            string sClose = "UPDATE GesLisVendita SET liv_dtf = " + _clsFun.DaySql(dIeri) + " WHERE liv_idx = " + livIdx;
                                            batchSql.Add(sClose);

                                            // 2. Verifica se esiste già una riga per oggi per questo articolo/listino
                                            DataRow[] candRows = t.Select("liv_lis='" + Convert.ToString(y["liv_lis"]).Replace("'", "''") + "' AND liv_art='" + _strArtCod.Replace("'", "''") + "'");
                                            DataRow todayRow = null;
                                            foreach (DataRow rCand in candRows)
                                            {
                                                if (!DBNull.Value.Equals(rCand["liv_dti"]) && Convert.ToDateTime(rCand["liv_dti"]).Date == DateTime.Today)
                                                {
                                                    todayRow = rCand;
                                                    break;
                                                }
                                            }

                                            if (todayRow != null && !DBNull.Value.Equals(todayRow["liv_idx"]))
                                            {
                                                string sUpdToday = "UPDATE GesLisVendita SET liv_prv = " + currPrv.ToString(CultureInfo.InvariantCulture) + ", liv_dtf = " + _clsFun.DaySql(_clsDef.DAYOUT) + " WHERE liv_idx = " + todayRow["liv_idx"];
                                                batchSql.Add(sUpdToday);
                                                lastVarSql = sUpdToday;
                                            }
                                            else
                                            {
                                                // Inserisce nuova riga storica con data inizio oggi
                                                string sIns = "INSERT INTO GesLisVendita (liv_lis, liv_art, liv_prv, liv_dti, liv_dtf, liv_day, liv_sta, liv_ann) VALUES (" +
                                                              "'" + Convert.ToString(y["liv_lis"]).Replace("'", "''") + "', " +
                                                              "'" + _strArtCod.Replace("'", "''") + "', " +
                                                              currPrv.ToString(CultureInfo.InvariantCulture) + ", " +
                                                              _clsFun.DaySql(DateTime.Today) + ", " +
                                                              _clsFun.DaySql(_clsDef.DAYOUT) + ", " +
                                                              _clsFun.DaySql(DateTime.Today) + ", " +
                                                              "'A', 0)";
                                                batchSql.Add(sIns);
                                                lastVarSql = sIns;
                                            }
                                            bVar = true;
                                            continue;
                                        }
                                        else if (origDti == DateTime.Today)
                                        {
                                            // Modifica multipla nello stesso giorno: aggiorna la riga odierna
                                            string sUpdToday = "UPDATE GesLisVendita SET liv_prv = " + currPrv.ToString(CultureInfo.InvariantCulture) + ", liv_dtf = " + _clsFun.DaySql(_clsDef.DAYOUT) + " WHERE liv_idx = " + livIdx;
                                            batchSql.Add(sUpdToday);
                                            lastVarSql = sUpdToday;
                                            bVar = true;
                                            continue;
                                        }
                                    }
                                }
                            }

                            s = "liv_lis='" + y["liv_lis"] + "' AND liv_art='" + y["liv_art"] + "' AND liv_dti='" + y["liv_dti"] + "'";
                            j = t.Select(s);
                            if (j.Length > 0)
                                s = _clsFun.SqlUpdRow(TABLISVEN, t, j[0], y, aWhe, aExl);
                            else
                            {
                                b = false;
                                if (!DBNull.Value.Equals(y["liv_idx"]) && Convert.ToInt32(y["liv_idx"]) > 0)
                                {
                                    s = "liv_idx='" + y["liv_idx"] + "'";
                                    j = t.Select(s);
                                    if (j.Length > 0)
                                    {
                                        s = _clsFun.SqlUpdRowIdx(TABLISVEN, t, j[0], y, aExl);
                                        b = true;
                                    }
                                }

                                if (!b)
                                    s = _clsFun.SqlInsertRow(TABLISVEN, t, y);
                            }
                            if (s != "")
                            {
                                batchSql.Add(s);
                                lastVarSql = s;
                                bVar = true;
                            }
                        }
                    }
                }

                // Execute all price SQL with a single DB connection
                if (batchSql.Count > 0)
                {
                    _clsFun.SqlWriteBatch(batchSql, _strConSql);
                    if (!string.IsNullOrEmpty(lastVarSql))
                        Variazioni(txtArtCod.Text, lastVarSql, "Anagrafica prezzi", _clsDef.VARALL);

                    string sVenQry = "SELECT liv_idx, liv_lis, liv_art, liv_prv, liv_dti, liv_dtf, liv_day, tab_des AS LivDes, tab_tip AS LivTip, liv_sta, liv_ann FROM GesLisVendita LEFT OUTER JOIN TabListini ON GesLisVendita.liv_lis = TabListini.tab_cod WHERE liv_art = '" + _strArtCod + "' AND liv_lis <> '" + _clsDef.LISOFF + "' ORDER BY liv_dti DESC, liv_lis";
                    _tabVen = _clsFun.FillTabSql(TABLISVEN, sVenQry, false, _strConSql);
                    if (_tabVen != null)
                    {
                        DataTable tVen = _tabVen.Copy();
                        if (!tVen.Columns.Contains("LivMdy"))
                            tVen.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "LivMdy", Caption = "Mdy", MaxLength = 1, DefaultValue = "" });
                        foreach (DataRow yRow in tVen.Rows)
                        {
                            if (tVen.Columns.Contains("liv_dtf") && !DBNull.Value.Equals(yRow["liv_dtf"]) && ((DateTime)yRow["liv_dtf"]).Year == 2050)
                                yRow["liv_dtf"] = DBNull.Value;
                        }
                        dgv3.DataSource = new DataView(tVen, "", "liv_dti DESC", DataViewRowState.CurrentRows);
                        FillVal();
                    }
                }

            }
            return bVar;
        }

        private void txtEan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                EanNew(false);
        }

        private Boolean EanNew(Boolean bolBil)
        {
            Boolean b = true;
            string s = "";
            DataRow[] j;

            if (string.IsNullOrEmpty(txtEan.Text))
            {
                MessageBox.Show("Barcode non definito!");
                return false;
            }
            if (!_clsFun.Numerico(txtEan.Text))
            {
                MessageBox.Show("Barcode non corretto!");
                return false;
            }

            b = true;
            s = txtEan.Text.Trim();
            if (!(s.Substring(0, 1) == "2" && s.Length == 13 && s.Substring(8, 5) == "00000"))
            {
                if (Convert.ToDouble(s) > 799999)
                {
                    string checkDigit = new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                    string expected = s.Substring(0, s.Length - 1) + checkDigit;
                    if (expected != s)
                    {
                        b = false;
                        if (MessageBox.Show("Barcode non corretto! (" + expected + "), forzi l'inserimento?", "INSERIMENTO EAN", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            b = true;
                    }
                }
            }

            if (!b) return false;

            DataTable dtDgv = GetGridDataTable(dgv1);
            if (dtDgv == null)
            {
                DataTable tEan = _clsFun.FillTabSql(TABANAEAN, "SELECT * FROM " + TABANAEAN + " WHERE 1=0", false, _strConSql);
                if (!tEan.Columns.Contains("EanMdy"))
                    tEan.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "EanMdy", Caption = "Mdy", MaxLength = 1, DefaultValue = "" });
                _tabEan = tEan.Copy();
                dgv1.DataSource = new DataView(tEan, "", "ean_ean", DataViewRowState.CurrentRows);
                dtDgv = tEan;
            }

            if (dtDgv != null)
            {
                j = dtDgv.Select("ean_ean='" + s.Replace("'", "''") + "'");
                if (j.Length > 0)
                {
                    MessageBox.Show("Barcode già presente!");
                    return false;
                }
            }

            s = "SELECT ";
            s += "AnaBarcode.ean_ean, ";
            s += "AnaBarcode.ean_art, ";
            s += "AnaArticoli.art_des ";
            s += "FROM AnaBarcode INNER JOIN AnaArticoli ON AnaBarcode.ean_art = AnaArticoli.art_cod ";
            s += " WHERE ean_ean='" + txtEan.Text.Trim().Replace("'", "''") + "'";
            DataTable t = _clsFun.FillTabSql(TABANAEAN, s, false, _strConSql);
            if (t != null && t.Rows.Count > 0)
            {
                string sArtCodDup = Convert.ToString(t.Rows[0]["ean_art"]);
                string sArtDesDup = Convert.ToString(t.Rows[0]["art_des"]);
                string sArtInfo = sArtCodDup + "-" + sArtDesDup;
                if (MessageBox.Show("Barcode già presente sull'articolo " + sArtInfo + ", vuoi spostarlo?", "BARCODE GIA' PRESENTE SU ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return false;
            }

            if (dtDgv != null)
            {
                if (bolBil)
                {
                    foreach (DataRow y in dtDgv.Rows)
                    {
                        if (y.RowState != DataRowState.Deleted && !DBNull.Value.Equals(y["ean_bil"]) && Convert.ToBoolean(y["ean_bil"]))
                        {
                            y["ean_bil"] = false;
                            y["EanMdy"] = "S";
                        }
                    }
                }

                DataRow x = dtDgv.NewRow();
                x["ean_art"] = txtArtCod.Text;
                x["ean_ean"] = txtEan.Text.Trim();
                x["ean_qta"] = 0;
                x["ean_prv"] = 0;
                x["ean_dti"] = DateTime.Today;
                x["ean_dtm"] = DateTime.Today;
                x["ean_bil"] = bolBil;
                x["ean_ecp"] = false;
                x["ean_ann"] = false;
                x["EanMdy"] = "S";
                dtDgv.Rows.Add(x);

                _bolIsDirty = true;
                txtEan.Text = "";

                if (dgv1 != null && dgv1.Rows.Count > 0)
                {
                    dgv1.Refresh();
                    int newRowIndex = dgv1.Rows.Count - 1;
                    try
                    {
                        dgv1.ClearSelection();
                        dgv1.Rows[newRowIndex].Selected = true;
                        dgv1.FirstDisplayedScrollingRowIndex = newRowIndex;
                    }
                    catch { }
                }

                txtEan.Focus();
            }

            // Non invocare Salva() se l'articolo è ancora in fase di creazione [NEW]
            if (txtArtCod.Text != _clsDef.CODNEW && txtArtCod.Text != "NEW" && !string.IsNullOrEmpty(txtArtCod.Text))
            {
                Salva();
            }
            return true;
        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                if (dgv1.CurrentCell is DataGridViewCheckBoxCell)
                {
                    _bolIsDirty = true;
                    if (dgv1.Rows[e.RowIndex].DataBoundItem is DataRowView drv)
                    {
                        drv["EanMdy"] = "S";
                    }
                    else if (dgv1.Columns.Contains("Modificato") && dgv1.Rows[e.RowIndex].Cells["Modificato"] != null)
                    {
                        dgv1.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
                    }
                }
            }
        }

        private void dgv1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgv1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private bool _isUpdatingBilanciaFlag = false;

        private void dgv1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_bolLoading || _isUpdatingBilanciaFlag || e.RowIndex < 0 || e.ColumnIndex < 0) return;

            _bolIsDirty = true;

            if (dgv1.Rows[e.RowIndex].DataBoundItem is DataRowView drv)
            {
                drv["EanMdy"] = "S";

                string colName = dgv1.Columns[e.ColumnIndex].Name;
                string dataProp = dgv1.Columns[e.ColumnIndex].DataPropertyName;

                if (dataProp == "ean_bil" || colName == "Bilancia")
                {
                    bool isBil = false;
                    if (!DBNull.Value.Equals(drv["ean_bil"]))
                        isBil = Convert.ToBoolean(drv["ean_bil"]);

                    if (isBil && dgv1.DataSource is DataView dv)
                    {
                        _isUpdatingBilanciaFlag = true;
                        try
                        {
                            for (int i = 0; i < dv.Count; i++)
                            {
                                if (i != e.RowIndex)
                                {
                                    DataRowView otherDrv = dv[i];
                                    if (!DBNull.Value.Equals(otherDrv["ean_bil"]) && Convert.ToBoolean(otherDrv["ean_bil"]))
                                    {
                                        otherDrv["ean_bil"] = false;
                                        otherDrv["EanMdy"] = "S";
                                    }
                                }
                            }
                            dgv1.Refresh();
                        }
                        finally
                        {
                            _isUpdatingBilanciaFlag = false;
                        }
                    }
                }
            }
            else if (dgv1.Columns.Contains("Modificato") && dgv1.Rows[e.RowIndex].Cells["Modificato"] != null)
            {
                dgv1.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
            }
        }

        private void dgv1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                _bolIsDirty = true;
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm != null && cm.Position >= 0 && e.RowIndex < cm.List.Count)
                {
                    DataRowView r = cm.List[e.RowIndex] as DataRowView;
                    if (r != null)
                    {
                        r.Row["EanMdy"] = "S";
                    }
                }
            }
        }

        private void dgv1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.CellStyle.BackColor = Color.Aquamarine;

            if (e.Control is TextBox)
            {
                var txt = e.Control as TextBox;
                if (txt != null)
                {
                    e.Control.KeyPress += new KeyPressEventHandler(txtArtPne_KeyPress);
                }
            }
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= dgv1.Rows.Count) return;

            string sArt = txtArtCod.Text != null ? txtArtCod.Text.Trim() : "";
            if (string.IsNullOrEmpty(sArt) || sArt == _clsDef.CODNEW || sArt == "NEW")
            {
                MessageBox.Show("Nessun articolo caricato. Impossibile cancellare il codice a barre.", "Articolo non caricato", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!(dgv1.Rows[e.RowIndex].DataBoundItem is DataRowView drv)) return;

            DataRow row = drv.Row;
            string strEan = row["ean_ean"] != null ? row["ean_ean"].ToString().Trim() : "";
            string strRowArt = row.Table.Columns.Contains("ean_art") && row["ean_art"] != null ? row["ean_art"].ToString().Trim() : sArt;
            if (string.IsNullOrEmpty(strRowArt)) strRowArt = sArt;

            if (string.IsNullOrEmpty(strEan)) return;

            string msg = "Sei sicuro di voler CANCELLARE DEFINITIVAMENTE il codice a barre '" + strEan + "' dal database per l'articolo " + sArt + (txtArtDes.Text.Trim() != "" ? " (" + txtArtDes.Text.Trim() + ")" : "") + "?\n\nL'operazione è irreversibile.";
            if (MessageBox.Show(msg, "Conferma Cancellazione Barcode", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
            {
                return;
            }

            try
            {
                string sSqlDel = "DELETE FROM " + TABANAEAN + " WHERE ean_ean = '" + strEan.Replace("'", "''") + "' AND (ean_art = '" + strRowArt.Replace("'", "''") + "' OR ean_art = '" + sArt.Replace("'", "''") + "')";
                bool ok = _clsFun.SqlWrite(sSqlDel, _strConSql);

                if (ok)
                {
                    Variazioni(sArt, "Cancellato barcode " + strEan, "Anagrafica barcode", _clsDef.VARALL);

                    // Rimuovi la riga dalla tabella sorgente associata alla griglia
                    DataTable parentTable = row.Table;
                    if (parentTable != null)
                    {
                        parentTable.Rows.Remove(row);
                        parentTable.AcceptChanges();
                    }

                    // Rimuovi dalla cache _tabEan se presente
                    if (_tabEan != null && _tabEan.Columns.Contains("ean_ean"))
                    {
                        DataRow[] arrEan = _tabEan.Select("ean_ean = '" + strEan.Replace("'", "''") + "'");
                        foreach (DataRow rEan in arrEan)
                        {
                            _tabEan.Rows.Remove(rEan);
                        }
                        _tabEan.AcceptChanges();
                    }

                    dgv1.Refresh();

                    MessageBox.Show("Codice a barre '" + strEan + "' cancellato definitivamente dal database.", "Barcode Cancellato", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Errore durante la cancellazione del codice a barre dal database.", "Errore Cancellazione", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore durante la cancellazione del barcode: " + ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dgv2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dgv2.DataSource == null) return;
                CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
                if (cm != null && cm.Position >= 0 && dgv2.CurrentRow != null && dgv2.CurrentRow.Index < cm.List.Count)
                {
                    DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
                    if (r != null)
                    {
                        DataRow x = r.Row;
                        string sArf = Convert.ToString(x["lia_arf"]).Trim();
                        string sFor = Convert.ToString(x["lia_for"]).Trim();

                        if (e.ColumnIndex == 3)
                        {
                            if (sArf != "")
                            {
                                string s = "";
                                DataTable t = _clsQry.ArtSeek(sArf, "ARF" + sFor);
                                if (t != null && t.Rows.Count > 0 && Convert.ToString(t.Rows[0]["tmp_art"]) != txtArtCod.Text)
                                {
                                    s = Convert.ToString(t.Rows[0]["tmp_art"]) + " " + Convert.ToString(t.Rows[0]["tmp_ard"]);
                                    if (MessageBox.Show("Codice fornitore già presente sull'articolo " + s + ", vuoi spostarlo?", "ARTICOLO FORNITORE GIA' PRESENTE SU ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                    {
                                        x["lia_arf"] = "";
                                    }
                                }
                            }
                        }

                        if (Convert.ToString(x["lia_arf"]).Trim() == "")
                            MessageBox.Show("Codice fornitore sulla riga costo mancante!", "COSTO NON COMPLETO", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        x["LiaMdy"] = "S";
                        _bolIsDirty = true;
                        FillVal();
                    }
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArticolo.dgv2_CellEndEdit", ex.Message);
            }
        }

        private void dgv2_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.CellStyle.BackColor = Color.Aquamarine;

            if (e.Control is TextBox)
            {
                var txt = e.Control as TextBox;
                if (txt != null)
                {
                    e.Control.KeyPress += new KeyPressEventHandler(txtArtPne_KeyPress);
                }
            }
        }

        private void dgv2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv2.Columns[e.ColumnIndex].Name == "Ann." && dgv2.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv2.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
            }
        }

        private void dgv2_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgv2.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv2.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dgv2_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (dgv2.DataSource == null) return;
                CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
                if (cm != null && cm.Position >= 0 && dgv2.CurrentRow != null && dgv2.CurrentRow.Index < cm.List.Count)
                {
                    DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
                    if (r != null)
                    {
                        DataRow x = r.Row;

                        string s = "";
                        decimal dIva = 10m;
                        if (cmbArtIva.SelectedValue != null && cmbArtIva.DataSource is DataTable dtIva)
                        {
                            DataRow[] j = dtIva.Select("tab_cod='" + cmbArtIva.SelectedValue.ToString().Replace("'", "''") + "'");
                            if (j.Length > 0 && !DBNull.Value.Equals(j[0]["tab_ali"]))
                                decimal.TryParse(Convert.ToString(j[0]["tab_ali"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dIva);
                        }

                        decimal dSfr = 0m;
                        if (_clsFun.Numerico(txtArtSfr.Text))
                            dSfr = _clsFun.Txt2Dec(txtArtSfr.Text);

                        s = PrezzoValido();

                        string[] a = s.Split('|');
                        decimal dPrv = a.Length > 0 ? _clsFun.Txt2Dec(a[0]) : 0m;
                        string sLis = a.Length > 1 ? a[1] : _clsDef.LISPOS;

                        decimal decLiaCos = 0m;
                        if (!DBNull.Value.Equals(x["lia_cos"]))
                            decimal.TryParse(Convert.ToString(x["lia_cos"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decLiaCos);

                        frmAnaArtMargini f = new frmAnaArtMargini();
                        f._tabLiv = cmbLisVen.DataSource as DataTable;
                        f._strLis = sLis;
                        f._decCos = decLiaCos;
                        f._decPrv = dPrv;
                        f._decIva = dIva;
                        f._decSfr = dSfr;
                        f.ShowDialog();

                        if (f._decPrv > 0)
                        {
                            string targetLis = !string.IsNullOrEmpty(f._strLis) ? f._strLis : _clsDef.LISPOS;
                            DataTable dt4 = GetGridDataTable(dgv4);
                            if (dt4 != null)
                            {
                                DataRow[] v = dt4.Select("LivLis='" + targetLis.Replace("'", "''") + "'");
                                if (v.Length > 0)
                                {
                                    v[0]["LivPrv"] = f._decPrv;
                                    v[0]["LivMav"] = _clsFun.Margine(f._decPrv, _clsFun.Txt2Dec(txtArtCou.Text), dIva, _clsFun.Txt2Dec(txtArtSfr.Text), "V");
                                    v[0]["LivMap"] = _clsFun.Margine(f._decPrv, _clsFun.Txt2Dec(txtArtCou.Text), dIva, _clsFun.Txt2Dec(txtArtSfr.Text), "P");
                                    LivAggPre(v[0]);
                                }
                                else
                                {
                                    DataRow newRow = dt4.NewRow();
                                    newRow["LivLis"] = targetLis;
                                    newRow["LivTip"] = (targetLis == _clsDef.LISPOS) ? "PC" : "L";
                                    newRow["LivPrv"] = f._decPrv;
                                    newRow["LivMav"] = _clsFun.Margine(f._decPrv, _clsFun.Txt2Dec(txtArtCou.Text), dIva, _clsFun.Txt2Dec(txtArtSfr.Text), "V");
                                    newRow["LivMap"] = _clsFun.Margine(f._decPrv, _clsFun.Txt2Dec(txtArtCou.Text), dIva, _clsFun.Txt2Dec(txtArtSfr.Text), "P");
                                    newRow["LivDti"] = DateTime.Today;
                                    newRow["LivDtf"] = DBNull.Value;
                                    newRow["LivAnn"] = false;
                                    dt4.Rows.Add(newRow);
                                    LivAggPre(newRow);
                                }

                                _bolIsDirty = true;
                                if (targetLis == _clsDef.LISPOS)
                                    txtArtPne.Text = f._decPrv.ToString("#0.00");

                                if (cmbArtEqp.SelectedValue != null && cmbArtEqp.SelectedValue.ToString() != "")
                                    AnaEqp();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArticolo.dgv2_DoubleClick", ex.Message);
            }
        }

        private void dgv3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv3.Columns[e.ColumnIndex].Name == "Ann." && dgv3.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv3.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
            }
        }

        private void dgv3_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgv3.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv3.CommitEdit(DataGridViewDataErrorContexts.Commit);
                FillVal();
            }
        }

        private void dgv3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dgv3.DataSource == null) return;
                CurrencyManager cm = dgv3.BindingContext[dgv3.DataSource, dgv3.DataMember] as CurrencyManager;
                if (cm != null && cm.Position >= 0 && dgv3.CurrentRow != null && dgv3.CurrentRow.Index < cm.List.Count)
                {
                    DataRowView r = cm.List[dgv3.CurrentRow.Index] as DataRowView;
                    if (r != null)
                    {
                        DataRow x = r.Row;

                        string s = Convert.ToString(x["LivTip"]).Trim();
                        if (!"NP".Contains(s) && !DBNull.Value.Equals(x["liv_dtf"]))
                        {
                            x["liv_dtf"] = DBNull.Value;
                            MessageBox.Show("Solo promo o prezzo negozio!", "CONTROLLO PREZZO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            x["LivMdy"] = "S";
                            _bolIsDirty = true;
                        }
                        FillVal();
                    }
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArticolo.dgv3_CellEndEdit", ex.Message);
            }
        }

        private void dgv3_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.CellStyle.BackColor = Color.Aquamarine;

            if (e.Control is TextBox)
            {
                var txt = e.Control as TextBox;
                if (txt != null)
                {
                    e.Control.KeyPress += new KeyPressEventHandler(txtArtPne_KeyPress);
                }
            }
        }

        private void dgv4_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.CellStyle.BackColor = Color.Aquamarine;

            if (e.Control is TextBox)
            {
                var txt = e.Control as TextBox;
                if (txt != null)
                {
                    e.Control.KeyPress += new KeyPressEventHandler(txt_KeyPress);
                }
            }
        }

        private void txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((sender as TextBox).Text != null && (sender as TextBox).Text.Trim() != "")
            {
                if (e.KeyChar == '.')
                    e.KeyChar = ',';
            }
        }

        private void dgv4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dgv4.DataSource == null) return;
                CurrencyManager cm = dgv4.BindingContext[dgv4.DataSource, dgv4.DataMember] as CurrencyManager;
                if (cm != null && cm.Position >= 0 && dgv4.CurrentRow != null && dgv4.CurrentRow.Index < cm.List.Count)
                {
                    DataRowView r = cm.List[dgv4.CurrentRow.Index] as DataRowView;
                    if (r != null)
                    {
                        DataRow x = r.Row;

                        decimal dIva = 10m;
                        if (cmbArtIva.SelectedValue != null && cmbArtIva.DataSource is DataTable dtIva)
                        {
                            DataRow[] j = dtIva.Select("tab_cod='" + cmbArtIva.SelectedValue.ToString().Replace("'", "''") + "'");
                            if (j.Length > 0 && !DBNull.Value.Equals(j[0]["tab_ali"]))
                                decimal.TryParse(Convert.ToString(j[0]["tab_ali"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dIva);
                        }

                        decimal dPrv = 0m;
                        if (!DBNull.Value.Equals(x["LivPrv"]))
                            decimal.TryParse(Convert.ToString(x["LivPrv"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dPrv);

                        x["LivMav"] = _clsFun.Margine(dPrv, _clsFun.Txt2Dec(txtArtCou.Text), dIva, _clsFun.Txt2Dec(txtArtSfr.Text), "V");
                        x["LivMap"] = _clsFun.Margine(dPrv, _clsFun.Txt2Dec(txtArtCou.Text), dIva, _clsFun.Txt2Dec(txtArtSfr.Text), "P");

                        _bolIsDirty = true;
                        LivAggPre(x);

                        if (cmbArtEqp.SelectedValue != null && cmbArtEqp.SelectedValue.ToString() != "")
                            AnaEqp();
                    }
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArticolo.dgv4_CellEndEdit", ex.Message);
            }
        }

        private void dgv4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv4.Columns[e.ColumnIndex].Name == "Ann." && dgv4.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv4.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
            }
        }

        private void dgv4_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgv4.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv4.CommitEdit(DataGridViewDataErrorContexts.Commit);
                FillVal();
            }
        }

        private void txtArtPne_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((sender as TextBox).Text != null && (sender as TextBox).Text.Trim() != "")
            {
                if (e.KeyChar == '.')
                    e.KeyChar = ',';
            }
        }

        private void btnForCos_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtArtCod.Text != _clsDef.CODNEW && txtArtCod.Text != "NEW" && _bolIsDirty)
                    Salva();

                frmSeekFor f = new frmSeekFor();
                f._bolAnagra = false;
                f.ShowDialog();
                if (f._strCod != "")
                {
                    if (f._tabTmp.Rows.Count > 0)
                    {
                        DataTable t = GetGridDataTable(dgv2);
                        if (t == null)
                        {
                            DataTable tCos = _clsFun.FillTabSql(TABLISACQ, "SELECT GesLisAcquisto.*, AnaFornitori.for_des AS CosFod FROM GesLisAcquisto LEFT OUTER JOIN AnaFornitori ON GesLisAcquisto.lia_for = AnaFornitori.for_cod WHERE 1=0", false, _strConSql);
                            if (!tCos.Columns.Contains("LiaMdy"))
                                tCos.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "LiaMdy", Caption = "Mdy", MaxLength = 1, DefaultValue = "" });
                            _tabCos = tCos.Copy();
                            dgv2.DataSource = new DataView(tCos, "", "lia_dti DESC", DataViewRowState.CurrentRows);
                            t = tCos;
                        }

                        DataRow x = t.NewRow();
                        x["lia_for"] = f._tabTmp.Rows[0]["tmp_for"];
                        x["CosFod"] = f._tabTmp.Rows[0]["tmp_fod"];
                        x["lia_dti"] = DateTime.Today;
                        x["lia_tip"] = "M";
                        x["lia_arf"] = "";
                        x["lia_cos"] = 0m;
                        x["lia_ann"] = false;
                        x["LiaMdy"] = "S";
                        _bolIsDirty = true;
                        t.Rows.Add(x);

                        if (dgv2.Rows.Count > 0)
                        {
                            dgv2.Select();
                            try
                            {
                                dgv2.CurrentCell = dgv2[3, dgv2.Rows.Count - 1];
                            }
                            catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArticolo.btnForCos_Click", ex.Message);
            }
        }

        private void cmbLisVen_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                if (cmbLisVen.SelectedValue == null || string.IsNullOrEmpty(cmbLisVen.SelectedValue.ToString()))
                    return;

                string selCod = cmbLisVen.SelectedValue.ToString();
                string sTip = "C";
                string sLis = "001";
                string sDes = cmbLisVen.Text;

                if (cmbLisVen.DataSource is DataTable dtCmb)
                {
                    DataRow[] jCmb = dtCmb.Select("tab_cod='" + selCod.Replace("'", "''") + "'");
                    if (jCmb.Length > 0)
                    {
                        if (!DBNull.Value.Equals(jCmb[0]["tab_tip"])) sTip = Convert.ToString(jCmb[0]["tab_tip"]);
                        if (!DBNull.Value.Equals(jCmb[0]["tab_cod"])) sLis = Convert.ToString(jCmb[0]["tab_cod"]);
                        if (!DBNull.Value.Equals(jCmb[0]["tab_des"])) sDes = Convert.ToString(jCmb[0]["tab_des"]);
                    }
                }

                DataTable t = GetGridDataTable(dgv3);
                if (t == null)
                {
                    DataTable tVen = _clsFun.FillTabSql(TABLISVEN, "SELECT liv_idx, liv_lis, liv_art, liv_prv, liv_dti, liv_dtf, liv_day, tab_des AS LivDes, tab_tip AS LivTip, liv_sta, liv_ann FROM GesLisVendita WHERE 1=0", false, _strConSql);
                    if (!tVen.Columns.Contains("LivMdy"))
                        tVen.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "LivMdy", Caption = "Mdy", MaxLength = 1, DefaultValue = "" });
                    _tabVen = tVen.Copy();
                    dgv3.DataSource = new DataView(tVen, "", "liv_dti DESC", DataViewRowState.CurrentRows);
                    t = tVen;
                }

                if (sTip != "P")
                {
                    DataRow[] j = t.Select("liv_lis='" + sLis.Replace("'", "''") + "'");
                    if (j.Length > 0)
                    {
                        MessageBox.Show("Tipo listino '" + cmbLisVen.Text + "' già presente!", "CONTROLLO LISTINO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }

                DataRow x = t.NewRow();
                if (t.Columns.Contains("LivTip")) x["LivTip"] = sTip;
                if (t.Columns.Contains("liv_lis")) x["liv_lis"] = sLis;
                if (t.Columns.Contains("LivDes")) x["LivDes"] = sDes;
                if (t.Columns.Contains("liv_art")) x["liv_art"] = txtArtCod.Text;
                if (t.Columns.Contains("liv_prv")) x["liv_prv"] = 0m;
                if (t.Columns.Contains("liv_dti")) x["liv_dti"] = DateTime.Today; // DateTime, NOT string!
                if (t.Columns.Contains("liv_dtf")) x["liv_dtf"] = DBNull.Value;
                if (t.Columns.Contains("liv_ann")) x["liv_ann"] = false;
                if (t.Columns.Contains("liv_sta")) x["liv_sta"] = "A";
                if (t.Columns.Contains("LivMdy")) x["LivMdy"] = "S";
                t.Rows.Add(x);

                _bolIsDirty = true;
                FillVal();
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArticolo.cmbLisVen_SelectionChangeCommitted", ex.Message);
            }
        }

        private void txtArtDes_Validated(object sender, EventArgs e)
        {
            if (txtArtDes.Text != "")
                txtArtDes.Text = txtArtDes.Text.ToUpper();
        }

        private void LivAggPre(DataRow rowLis)
        {
            if (rowLis == null) return;
            try
            {
                DataTable tLis = GetGridDataTable(dgv3);
                if (tLis == null) return;

                string sLis = Convert.ToString(rowLis["LivLis"]).Trim();
                if (string.IsNullOrEmpty(sLis)) return;

                decimal dPrv = 0m;
                if (!DBNull.Value.Equals(rowLis["LivPrv"]))
                    decimal.TryParse(Convert.ToString(rowLis["LivPrv"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dPrv);

                DateTime dDti = DateTime.Today;
                if (!DBNull.Value.Equals(rowLis["LivDti"]))
                    DateTime.TryParse(Convert.ToString(rowLis["LivDti"]), out dDti);

                DateTime dDtf = _clsDef.DAYOUT;
                if (!DBNull.Value.Equals(rowLis["LivDtf"]))
                    DateTime.TryParse(Convert.ToString(rowLis["LivDtf"]), out dDtf);

                bool bAnn = false;
                if (!DBNull.Value.Equals(rowLis["LivAnn"]))
                    bAnn = Convert.ToBoolean(rowLis["LivAnn"]);

                decimal dp = 0m;
                DateTime di = DateTime.MinValue;
                DateTime df = _clsDef.DAYOUT;
                bool ba = false;

                DataView dv3 = GetGridDataView(dgv3);
                if (dv3 != null)
                {
                    foreach (DataRowView r in dv3)
                    {
                        if (Convert.ToString(r["liv_lis"]).Trim() == sLis)
                        {
                            if (!DBNull.Value.Equals(r["liv_prv"]))
                                decimal.TryParse(Convert.ToString(r["liv_prv"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dp);

                            if (!DBNull.Value.Equals(r["liv_dti"]))
                                DateTime.TryParse(Convert.ToString(r["liv_dti"]), out di);

                            df = _clsDef.DAYOUT;
                            if (!DBNull.Value.Equals(r["liv_dtf"]))
                                DateTime.TryParse(Convert.ToString(r["liv_dtf"]), out df);

                            if (!DBNull.Value.Equals(r["liv_ann"]))
                                ba = Convert.ToBoolean(r["liv_ann"]);

                            if (di <= dDti)
                                break;
                        }
                    }
                }

                if (dp != dPrv || di.ToString("yyyyMMdd") != dDti.ToString("yyyyMMdd") || df.ToString("yyyyMMdd") != dDtf.ToString("yyyyMMdd") || ba != bAnn)
                {
                    DataRow[] j = cmbLisVen.DataSource is DataTable dtCmb ? dtCmb.Select("tab_cod='" + sLis.Replace("'", "''") + "'") : null;

                    string s = "liv_lis='" + sLis.Replace("'", "''") + "'";
                    DataRow[] jLiv = tLis.Select(s, "liv_dti DESC");

                    if (jLiv.Length > 0)
                    {
                        jLiv[0]["liv_prv"] = dPrv;
                        jLiv[0]["liv_dtf"] = (dDtf == _clsDef.DAYOUT) ? (object)DBNull.Value : dDtf;
                        jLiv[0]["LivMdy"] = "S";
                        jLiv[0]["liv_ann"] = bAnn;

                        rowLis["LivDti"] = DateTime.Today;

                        if (Convert.ToString(jLiv[0]["liv_lis"]) == _clsDef.LISPRO)
                        {
                            if (!DBNull.Value.Equals(jLiv[0]["liv_dtf"]))
                            {
                                DateTime dtFine;
                                if (DateTime.TryParse(Convert.ToString(jLiv[0]["liv_dtf"]), out dtFine) && dtFine >= DateTime.Today && Convert.ToString(jLiv[0]["liv_sta"]) == "")
                                {
                                    jLiv[0]["liv_sta"] = "D";
                                }
                            }
                        }
                    }
                    else
                    {
                        DataRow x = tLis.NewRow();
                        x["liv_lis"] = sLis;
                        x["liv_art"] = txtArtCod.Text;
                        x["liv_prv"] = dPrv;
                        x["liv_dti"] = DateTime.Today;
                        x["liv_dtf"] = (dDtf == _clsDef.DAYOUT) ? (object)DBNull.Value : dDtf;

                        string sDes = "";
                        string sTip = "";
                        if (j != null && j.Length > 0)
                        {
                            sDes = Convert.ToString(j[0]["tab_des"]);
                            sTip = Convert.ToString(j[0]["tab_tip"]);
                        }
                        x["LivDes"] = sDes;
                        x["LivTip"] = sTip;
                        x["liv_ann"] = bAnn;
                        x["LivMdy"] = "S";
                        tLis.Rows.Add(x);

                        rowLis["LivDti"] = DateTime.Today;
                    }

                    _bolIsDirty = true;
                    if (sLis == _clsDef.LISPOS && dPrv > 0)
                        txtArtPne.Text = dPrv.ToString("#0.00");
                    MargRica();
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArticolo.LivAggPre", ex.Message);
            }
        }

        private void txtArtCou_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                decimal d = _clsFun.Txt2Dec(txtArtCou.Text);
                if (d > 0)
                    LiaAgg("M", d);
            }
        }

        private void LiaAgg(string strTip, decimal decPra)
        {
            try
            {
                string s = _clsFun.ParGet(clsDefine.enuParametri.ParForDefault, _strConSql);
                string sFor = "";
                string sFod = "";
                string[] a = s.Split(',');
                if (a.Length > 0)
                    sFor = a[0];
                if (a.Length > 1)
                    sFod = a[1];

                if (sFor.Length == 0 || sFor.Length > 5)
                    MessageBox.Show("Parametro fornitore principale " + s + " da sistemare!", "CONTROLLO FORNITORE", MessageBoxButtons.OK, MessageBoxIcon.Question);
                else
                {
                    DataTable t = GetGridDataTable(dgv2);
                    if (t == null)
                    {
                        DataTable tCos = _clsFun.FillTabSql(TABLISACQ, "SELECT GesLisAcquisto.*, AnaFornitori.for_des AS CosFod FROM GesLisAcquisto LEFT OUTER JOIN AnaFornitori ON GesLisAcquisto.lia_for = AnaFornitori.for_cod WHERE 1=0", false, _strConSql);
                        if (!tCos.Columns.Contains("LiaMdy"))
                            tCos.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "LiaMdy", Caption = "Mdy", MaxLength = 1, DefaultValue = "" });
                        _tabCos = tCos.Copy();
                        dgv2.DataSource = new DataView(tCos, "", "lia_dti DESC", DataViewRowState.CurrentRows);
                        t = tCos;
                    }

                    DataRow[] j = t.Select("lia_for='" + sFor.Replace("'", "''") + "'");
                    DataRow[] jLia = t.Select("lia_tip='" + strTip.Replace("'", "''") + "' AND lia_dti=#" + DateTime.Today.ToString("MM/dd/yyyy") + "#");

                    if (jLia.Length > 0)
                    {
                        jLia[0]["lia_cos"] = decPra;
                        jLia[0]["LiaMdy"] = "S";
                        _bolIsDirty = true;
                    }
                    else
                    {
                        if (sFor.Length > 5)
                            sFor = sFor.Substring(0, 5);

                        DataRow x = t.NewRow();
                        x["lia_tip"] = strTip;
                        x["lia_art"] = txtArtCod.Text;
                        x["lia_for"] = sFor;
                        x["CosFod"] = sFod;
                        x["lia_dti"] = DateTime.Today;
                        x["lia_cos"] = decPra;

                        x["lia_arf"] = "";
                        x["lia_pxc"] = 1;
                        x["lia_cxp"] = 0;
                        if (j.Length > 0)
                        {
                            x["lia_arf"] = j[0]["lia_arf"];
                            x["lia_pxc"] = j[0]["lia_pxc"];
                            x["lia_cxp"] = j[0]["lia_cxp"];
                        }
                        x["lia_ann"] = false;
                        x["LiaMdy"] = "S";
                        _bolIsDirty = true;
                        t.Rows.Add(x);
                    }
                    MargRica();
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArticolo.LiaAgg", ex.Message);
            }
        }

        private void btnSeek_Click(object sender, EventArgs e)
        {
            Ricerca();
            if (txtArtCod.Text.Trim() != "")
                CtrlPanels("");
        }

        //private void OldRicerca()
        //{
        //    Salva();
        //    frmSeekArt f = new frmSeekArt();
        //    //f._strUltSeek = _strUltSeek;
        //    f.ShowDialog();
        //    if (f._tabArt != null && f._tabArt.Rows.Count > 0)
        //    {
        //        //_strUltSeek = f._strUltSeek;
        //        if (f._tabArt.Rows.Count > 0)
        //        {
        //            _strArtCod = (string)f._tabArt.Rows[0]["tmp_art"];
        //            _tabArt = f._tabArt.Copy();
        //        }
        //        FillArt();
        //    }
        //}

        private void Ricerca()
        {
            Salva();

            using (frmSeekArt f = new frmSeekArt())
            {
                f._bolValori = true;
                f.ShowDialog();
                if (f._tabArt != null && f._tabArt.Rows.Count > 0)
                {
                    _strArtCod = (string)f._tabArt.Rows[0]["tmp_art"];
                    _tabArt = f._tabArt.Copy();
                    FillArt(false);
                }
            }
        }

        private void txtSeek_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                Salva();

                if (txtSeek.Text != "")
                {
                    DataTable t = _clsQry.ArtSeek(txtSeek.Text, "SEEK");
                    if (t != null && t.Rows.Count > 0)
                    {
                        txtArtCod.Text = (string)t.Rows[0]["tmp_art"];
                        _tabArt = t.Copy();

                        FillArt(false);
                    }
                }
                if (txtArtCod.Text.Trim() != "")
                    pnlArt.Enabled = true;
                
                txtSeek.Text = "";
                SetFocusToSeek();
            }
        }

        private void txtSeek_Enter(object sender, EventArgs e)
        {
            txtSeek.SelectAll();
        }

        private void invioInCassaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtArtCod.Text.Trim()))
            {
                MessageBox.Show(this, "Nessun articolo selezionato.", "INVIO CASSA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                Salva();
                string sArtCod = txtArtCod.Text.Trim().PadLeft(7, '0');
                Variazioni(sArtCod, "Forza", "Anagrafica articolo", _clsDef.VARPOS);
                if (_strPar016PathDivNegozi != "")
                    DivNegozi();

                MessageBox.Show(this, "Variazione Cassa generata con successo per l'articolo " + sArtCod + ".", "INVIO CASSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void generazioneEtichettaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtArtCod.Text.Trim()) || txtArtCod.Text.Trim() == _clsDef.CODNEW || txtArtCod.Text.Trim() == "NEW")
            {
                MessageBox.Show(this, "Nessun articolo valido selezionato.", "GENERAZIONE ETICHETTA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                Salva();
                string sArtCod = txtArtCod.Text.Trim().PadLeft(7, '0');
                DataTable t = GetGridDataTable(dgv3);
                if (t != null)
                {
                    DataView v = new DataView(t, "LivTip='P' AND liv_dti <= #" + DateTime.Now.ToString("MM/dd/yyyy") + "#", "liv_dti DESC", DataViewRowState.CurrentRows);

                    if (v.Count > 0)
                        Variazioni(sArtCod, "PRO", "Anagrafica articolo", _clsDef.VARETI);
                    else
                        Variazioni(sArtCod, "Forza", "Anagrafica articolo", _clsDef.VARETI);

                    MessageBox.Show(this, "Variazione Etichetta generata con successo per l'articolo " + sArtCod + ".", "GENERAZIONE ETICHETTA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void Variazioni(string strArt, string strSql, string strDes, string strTip)
        {
            if (string.IsNullOrEmpty(strArt)) strArt = txtArtCod.Text;
            string sArt = strArt.Trim().PadLeft(7, '0');

            _clsVar.Variazioni(sArt, strSql, "Anagrafica articolo", strTip);
            _clsFun.FileLog("AnaArticoli", sArt, strSql);
        }

        private void btnPluNew_Click(object sender, EventArgs e)
        {
            if (cmbArtReb.SelectedValue.ToString() == "")
                MessageBox.Show("Reparto bilancia non definito");
            else if (txtArtPlu.Text != "")
                MessageBox.Show("Codice PLU già definito!");
            else
            {
                string s = _clsQry.NewPlu(cmbArtReb.SelectedValue.ToString());
                if (_clsFun.Numerico(s))
                    txtArtPlu.Text = s;
                else
                    MessageBox.Show("Codici PLU non univoci!");
            }
        }

        private void btnPluEan_Click(object sender, EventArgs e)
        {
            string s = "";
            Boolean b = false;
            if (cmbArtReb.SelectedValue.ToString() == "")
                MessageBox.Show("Reparto bilancia non definito");
            else if (!_clsFun.Numerico(txtArtPlu.Text))
                MessageBox.Show("PLU bilancia non definito");
            else
            {
                string sBil = "";
                s = groupBox1.Text;
                if (s.Length > 10)
                    sBil = s.Substring(9, 2);

                groupBox1.Text.Substring(10, 2);
                string sPre = _clsFun.ParGet(clsDefine.enuParametri.ParPrefixEabBilancia, _strConSql);
                string sLun = _clsFun.ParGet(clsDefine.enuParametri.ParOmegaBarcodeLun, _strConSql);
                if (!_clsFun.Numerico(sLun))
                    sLun = "";
                if (sLun == "8")
                    sLun = "4";

                if (sPre != "" && (!_clsFun.Numerico(sPre) || sPre.Length != 2))
                {
                    MessageBox.Show("Parametro prefisso EAN per bilance non corretto!");
                    sPre = "";
                }

                if (sPre.Trim() == "")
                    sPre = "21";

                //s = txtArtPlu.Text;
                //if(_clsFun.Numerico(sLun,"1234567890"))
                //    s = txtArtPlu.Text.PadLeft(3, Convert.ToChar('0'));

                txtEan.Text = sPre + cmbArtReb.SelectedValue.ToString() + txtArtPlu.Text + "000000";

                if (sBil == Convert.ToInt16((object)clsDefine.enuBilance.bilBizerba).ToString("00"))
                {
                    s = txtArtPlu.Text;
                    if (s.Length > 3)
                        s = s.Substring(s.Length - 3);

                    txtEan.Text = sPre + cmbArtReb.SelectedValue.ToString() + s + "0000000";
                }
                if (sBil == Convert.ToInt16((object)clsDefine.enuBilance.bilBizWinVarp).ToString("00"))
                {
                    s = txtArtPlu.Text;

                    if (sLun == "" || s.Length > Convert.ToInt16(sLun))
                        MessageBox.Show("PLU non corretto " + s + "!", "CONTROLLO PLU", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    else
                    {
                        s = s.PadLeft(Convert.ToInt16(sLun), Convert.ToChar("0"));
                        txtEan.Text = sPre + cmbArtReb.SelectedValue.ToString() + s + "00000000";
                        txtEan.Text = txtEan.Text.Substring(0, 13);
                    }
                }
                else if (sLun != "" && _clsFun.Numerico(sLun, "0123456789"))
                {
                    int i = Convert.ToInt16(sLun);

                    s = sPre + cmbArtReb.SelectedValue.ToString() + txtArtPlu.Text.PadLeft(i, Convert.ToChar('0')) + "0000000000";

                    txtEan.Text = s.Substring(0, 13);
                }
                txtEan.Select();
                EanNew(true);
                b = true;
            }
            if (!b)
                txtArtPlu.Text = "";
        }
        private void txtArtPlu_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                CtrlArtPlu();
        }

        private void txtArtPlu_Validated(object sender, EventArgs e)
        {
            CtrlArtPlu();
        }

        private void CtrlArtPlu()
        {
            if (txtArtPlu.Text != "")
            {
                string s = "";

                string sBil = "";
                s = groupBox1.Text;
                if (s.Length > 10)
                    sBil = s.Substring(9, 2);

                //string ss = Convert.ToInt16((object)clsDefine.enuBilance.bilBizerba).ToString("00");

                if (sBil != Convert.ToInt16((object)clsDefine.enuBilance.bilBizWinVarp).ToString("00"))
                    txtArtPlu.Text = txtArtPlu.Text.PadLeft(4, Convert.ToChar("0"));

                if (cmbArtReb.SelectedValue == null || cmbArtReb.SelectedValue.ToString() == "")
                {
                    MessageBox.Show("Reparto bilancia non definito!");
                    txtArtPlu.Text = "";
                }
                else
                {
                    DataTable t = _clsQry.PluExist(cmbArtReb.SelectedValue.ToString(), txtArtPlu.Text, txtArtCod.Text);

                    if (t.Rows.Count > 0 && (string)t.Rows[0]["art_cod"] != txtArtCod.Text)
                    {
                        //MessageBox.Show("PLU già presente sull'articolo " + (string)t.Rows[0]["art_cod"] + " " + (string)t.Rows[0]["art_des"] + "!");

                        if (MessageBox.Show("PLU già presente sull'articolo " + (string)t.Rows[0]["art_cod"] + " " + (string)t.Rows[0]["art_des"] + "! Vuoi spostarlo?", "CONTROLLO PLU", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            s = "UPDATE AnaArticoli SET art_plu='' WHERE art_cod='" + (string)t.Rows[0]["art_cod"] + "'";
                            _clsFun.SqlWrite(s, _strConSql);
                        }
                        else
                            txtArtPlu.Text = "";
                    }
                }
            }
        }

        private void statisticheToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (txtArtCod.Text == "")
                MessageBox.Show("Articolo non definito!", "CONTROLLO STATISTICHE", MessageBoxButtons.OK, MessageBoxIcon.Question);
            else
            {
                frmGesStatArt f = new frmGesStatArt();
                f._strArtCod = txtArtCod.Text;
                f._strArtDes = txtArtDes.Text;
                f._strArtUmi = cmbArtUmi.SelectedValue.ToString();
                f.ShowDialog();
            }
        }

        private void btnIng_Click(object sender, EventArgs e)
        {
            Salva();

            if (_strPar021GesIngredienti.Length > 0 && _strPar021GesIngredienti.Substring(0, 1) == "2")
            {
                frmAnaArtIngredient2 f = new frmAnaArtIngredient2();
                f._strArtCod = txtArtCod.Text;
                f.ShowDialog();
            }
            else
            {
                frmAnaArtIngredienti f = new frmAnaArtIngredienti();
                f._strArtCod = txtArtCod.Text;
                f.ShowDialog();
            }
        }

        private void gestioneVariazioniToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Salva();
            new frmGesVariazioni().ShowDialog();
            FillArt(false);
        }

        private void cmbArtSta_SelectionChangeCommitted(object sender, EventArgs e)
        {
            ArtStato();
        }

        private void DivNegozi()
        {
            string artCod = txtArtCod.Text;
            string pathDiv = _strPar016PathDivNegozi;
            if (string.IsNullOrEmpty(artCod) || string.IsNullOrEmpty(pathDiv)) return;
            string pNorm = pathDiv.Trim().ToUpper();
            if (pNorm == "N" || pNorm == "NO" || pNorm == "0" || pNorm == "FALSE") return;

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    ArrayList a = new ArrayList { artCod };
                    new clsVariazioni().DivNegArticoli(pathDiv, a, null, "", "");
                }
                catch (Exception ex)
                {
                    _clsFun.ErrorLog("frmAnaArticolo.DivNegozi", ex.Message);
                }
            });
        }

        private void dgv3_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= dgv3.Rows.Count) return;
            string s = dgv3.Rows[e.RowIndex].Cells["Tp"].Value != null ? dgv3.Rows[e.RowIndex].Cells["Tp"].Value.ToString() : "";

            if (s != "P")
            {
                if (dgv3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != DBNull.Value)
                {
                    dgv3.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
                    dgv3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = DBNull.Value;
                }
                e.Cancel = true;
            }
        }

        private void btnEqp_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabEquPrezzi";
            f._strTabDes = "Equivalenza prezzi";
            f._bolCodAlf = false;
            f._intCodLen = 4;
            f._intDesLen = 50;
            f._bolColAnn = true;
            f.ShowDialog();

            FillTabs("TabEquPrezzi");
        }

        private void btnArtEqp_Click(object sender, EventArgs e)
        {
            if (cmbArtEqp.SelectedValue == null || cmbArtEqp.SelectedValue.ToString() == "")
                MessageBox.Show("Equivalenta non definita", "DATI MANCANTI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
                AnaEqp();
        }

        private void AnaEqp()
        {
            Salva();

            decimal d = 0m;
            DataTable dt4 = GetGridDataTable(dgv4);
            if (dt4 != null)
            {
                DataRow[] j = dt4.Select("LivIdx='" + _clsDef.LISPOS + "'");
                if (j.Length > 0 && !DBNull.Value.Equals(j[0]["LivPrv"]))
                    decimal.TryParse(Convert.ToString(j[0]["LivPrv"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out d);

                string sEqp = cmbArtEqp.SelectedValue != null ? cmbArtEqp.SelectedValue.ToString() : "";
                frmAnaArtEquivalenze f = new frmAnaArtEquivalenze();
                f._strTip = "Prezzi";
                f._strEqp = sEqp;
                f._strArt = txtArtCod.Text;
                f._decPrv = d;
                f.ShowDialog();

                d = f._decPrv;

                if (j.Length > 0)
                    j[0]["LivPrv"] = d;
            }
        }

        private void btnEanAutomatico_Click(object sender, EventArgs e)
        {
            string s = "SELECT TOP (5000) ean_ean FROM AnaBarcode WHERE ean_ean >= '8000000000000' AND ean_ean <= '8999999999999'";
            DataTable t = _clsFun.FillTabSql(TABANAEAN, s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = t.Columns["ean_ean"];
            t.PrimaryKey = keys;

            if (t.Rows.Count > 0)
            {
                for (int i = 1; i < 5000; i++)
                {
                    s = (800000000000 + i).ToString() + "0";

                    s = s.Substring(0, 12) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));

                    DataRow[] j = t.Select("ean_ean='" + s + "'");
                    if (j.Length == 0)
                    {
                        txtEan.Text = s;
                        break;
                    }
                }
            }
        }

        private void FillNewArt(string strVar)
        {
            string s = "";
            string[] a = strVar.Split(';');

            if (a.Length > 0 && a.Length < 3)
            {
                foreach (string ss in a)
                {
                    if (ss != "")
                    {
                        string sFld = ss.Length >= 7 ? ss.Substring(0, 7) : "";
                        string sVal = ss.Length >= 8 ? ss.Substring(8) : "";

                        if (sFld == "art_cod" && sVal != _clsDef.CODNEW)
                        {
                            txtSeek.Text = sVal;

                            DataTable t = _clsQry.ArtSeek(txtSeek.Text, "SEEK");
                            if (t != null && t.Rows.Count > 0)
                            {
                                txtArtCod.Text = Convert.ToString(t.Rows[0]["tmp_art"]);
                                _tabArt = t.Copy();

                                FillArt(false);
                                txtSeek.Text = "";
                            }
                        }
                    }
                    else
                        break;
                }
            }
            else
            {
                _strArtCod = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaArticoli, 7, _strConSql);

                if (_tabArt != null) _tabArt.Clear();

                FillArt(true);

                txtArtCod.Text = _strArtCod;
                cmbArtSta.SelectedValue = "A";

                foreach (string ss in a)
                {
                    if (ss.Length > 5)
                    {
                        string sFld = ss.Length >= 7 ? ss.Substring(0, 7) : "";
                        string sVal = ss.Length >= 8 ? ss.Substring(8) : "";

                        if (sFld == "art_des")
                            txtArtDes.Text = sVal;
                        else if (sFld == "art_iva")
                            cmbArtIva.SelectedValue = sVal;
                        else if (sFld == "art_umi")
                            cmbArtUmi.SelectedValue = sVal;
                        else if (sFld == "art_tgr")
                            cmbArtTgr.SelectedValue = sVal;
                        else if (sFld == "art_pne")
                            txtArtPne.Text = sVal;
                        else if (sFld == "art_ecr" && sVal.Length >= 9)
                        {
                            string sEc1 = sVal.Substring(0, 3);
                            string sEc2 = sVal.Substring(3, 3);
                            string sEc3 = sVal.Substring(6, 3);

                            cmbArtEc1.SelectedValue = sEc1;
                            cmbArtEc2.SelectedValue = sEc2;
                            cmbArtEc3.SelectedValue = sEc3;
                        }
                        else if (sFld == "art_rep")
                            cmbArtRep.SelectedValue = sVal;
                        else if (sFld == "ean_ean")
                        {
                            txtEan.Text = sVal;
                            EanNew(false);
                            txtEan.Text = "";
                        }
                        else if (sFld == "lia_prc" && sVal != "")
                        {
                            string[] aa = sVal.Split(':');
                            if (aa.Length >= 3)
                            {
                                string sFor = aa[0];
                                string sFod = aa[1];
                                decimal dPrc = 0m;
                                decimal.TryParse(aa[2].Replace('.', ','), out dPrc);
                                if (sFor != "" && dPrc > 0)
                                {
                                    DataTable t = GetGridDataTable(dgv2);
                                    if (t != null)
                                    {
                                        DataRow x = t.NewRow();
                                        x["lia_for"] = sFor;
                                        x["CosFod"] = sFod;
                                        x["lia_dti"] = DateTime.Today;
                                        x["lia_tip"] = "M";
                                        x["lia_arf"] = _strArtCod;
                                        x["lia_cos"] = dPrc;
                                        x["lia_ann"] = false;
                                        x["LiaMdy"] = "S";
                                        t.Rows.Add(x);
                                    }
                                }
                            }
                        }
                        else if (sFld == "liv_prv" && sVal != "")
                        {
                            string[] aa = sVal.Split(':');
                            if (aa.Length >= 3)
                            {
                                string sLis = aa[0];
                                string sLid = aa[1];
                                decimal dPrv = 0m;
                                decimal.TryParse(aa[2].Replace('.', ','), out dPrv);

                                DataTable tLiv = GetGridDataTable(dgv4);
                                if (tLiv != null)
                                {
                                    DataRow x = tLiv.NewRow();
                                    x["LivLis"] = sLis;
                                    x["LivTip"] = "PN";
                                    x["LivPrv"] = dPrv;
                                    x["LivDti"] = DateTime.Today;
                                    x["LivDtf"] = _clsDef.DAYOUT;
                                    x["LivAnn"] = false;
                                    tLiv.Rows.Add(x);

                                    LivAggPre(x);
                                }
                            }
                        }
                    }
                }
            }
        }

        // ─── Immagine: Sfoglia, Ricerca Web, Drag&Drop, Anteprima ─────────────────

        private void btnWebSearchImg_Click(object sender, EventArgs e)
        {
            string artCod = txtArtCod.Text.Trim();
            if (string.IsNullOrEmpty(artCod) || artCod == _clsDef.CODNEW)
            {
                MessageBox.Show("Salvare prima l'articolo o inserire un codice articolo valido per cercare le foto online.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Raccogli tutti i barcode associati all'articolo (da dgv1 e da database)
            List<string> eanList = new List<string>();
            DataTable dtDgv = GetGridDataTable(dgv1);
            if (dtDgv != null)
            {
                foreach (DataRow r in dtDgv.Rows)
                {
                    if (dtDgv.Columns.Contains("ean_ean") && !DBNull.Value.Equals(r["ean_ean"]))
                    {
                        string ean = r["ean_ean"].ToString().Trim();
                        if (!string.IsNullOrEmpty(ean) && !eanList.Contains(ean))
                            eanList.Add(ean);
                    }
                }
            }

            try
            {
                DataTable tEanDb = _clsFun.FillTabSql("AnaBarcode", "SELECT ean_ean FROM AnaBarcode WHERE ean_art='" + artCod.Replace("'", "''") + "' AND ean_ann=0", false, _strConSql);
                if (tEanDb != null)
                {
                    foreach (DataRow r in tEanDb.Rows)
                    {
                        string ean = r["ean_ean"].ToString().Trim();
                        if (!string.IsNullOrEmpty(ean) && !eanList.Contains(ean))
                            eanList.Add(ean);
                    }
                }
            }
            catch { }

            string artDes = txtArtDes.Text.Trim();

            using (frmAnaArtImgWebSearch dlg = new frmAnaArtImgWebSearch(artCod, artDes, eanList))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK && !string.IsNullOrEmpty(dlg.SelectedImagePath))
                {
                    txtArtImg.Text = dlg.SelectedImagePath;
                    _bolIsDirty = true;

                    // Aggiorna subito il database AnaArticoli
                    try
                    {
                        string sSql = "UPDATE AnaArticoli SET art_img='" + dlg.SelectedImagePath.Replace("'", "''") + "' WHERE art_cod='" + artCod.Replace("'", "''") + "'";
                        _clsFun.SqlWrite(sSql, _strConSql);
                    }
                    catch { }
                }
            }
        }

        private void btnBrowseImg_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Title = "Seleziona immagine articolo";
                dlg.Filter = "Immagini|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.webp|Tutti i file|*.*";
                dlg.CheckFileExists = true;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtArtImg.Text = dlg.FileName;
                }
            }
        }

        private void txtArtImg_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void txtArtImg_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
                    if (files != null && files.Length > 0)
                    {
                        string sourceFile = files[0];
                        string artCod = txtArtCod.Text.Trim();

                        if (string.IsNullOrEmpty(artCod) || artCod == "NEWA")
                        {
                            MessageBox.Show("Salvare prima l'articolo per generare un codice valido o inserire un codice prima di trascinare l'immagine.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        string dirIni = System.IO.Path.GetDirectoryName(_clsDef.FILEINI);
                        string dirImg = System.IO.Path.GetFullPath(System.IO.Path.Combine(dirIni, @"..\..\Immagini\Articoli"));

                        if (!System.IO.Directory.Exists(dirImg))
                        {
                            try { System.IO.Directory.CreateDirectory(dirImg); } catch { dirImg = dirIni; }
                        }

                        string ext = System.IO.Path.GetExtension(sourceFile);
                        string newFileName = artCod + ext;
                        string destFile = System.IO.Path.Combine(dirImg, newFileName);

                        System.IO.File.Copy(sourceFile, destFile, true);
                        txtArtImg.Text = destFile;
                    }
                }
                else if (e.Data.GetDataPresent(DataFormats.Bitmap))
                {
                    // Se trascinano un'immagine pura (es. da browser web senza file fisico)
                    System.Drawing.Image img = e.Data.GetData(DataFormats.Bitmap) as System.Drawing.Image;
                    if (img != null)
                    {
                        string artCod = txtArtCod.Text.Trim();
                        if (string.IsNullOrEmpty(artCod) || artCod == "NEWA")
                        {
                            MessageBox.Show("Salvare prima l'articolo per generare un codice valido.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        string dirIni = System.IO.Path.GetDirectoryName(_clsDef.FILEINI);
                        string dirImg = System.IO.Path.GetFullPath(System.IO.Path.Combine(dirIni, @"..\..\Immagini\Articoli"));
                        if (!System.IO.Directory.Exists(dirImg))
                        {
                            try { System.IO.Directory.CreateDirectory(dirImg); } catch { dirImg = dirIni; }
                        }

                        string newFileName = artCod + ".png";
                        string destFile = System.IO.Path.Combine(dirImg, newFileName);
                        img.Save(destFile, System.Drawing.Imaging.ImageFormat.Png);
                        txtArtImg.Text = destFile;
                    }
                }
                else if (e.Data.GetDataPresent(DataFormats.StringFormat))
                {
                    // Se trascinano un link (URL) o un testo da browser
                    string urlOrPath = (string)e.Data.GetData(DataFormats.StringFormat);
                    if (!string.IsNullOrEmpty(urlOrPath))
                    {
                        if (urlOrPath.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                        {
                            // Tentativo di download basico per le URL
                            string artCod = txtArtCod.Text.Trim();
                            if (string.IsNullOrEmpty(artCod) || artCod == "NEWA")
                            {
                                MessageBox.Show("Salvare l'articolo prima per generare un codice valido.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                            string dirIni = System.IO.Path.GetDirectoryName(_clsDef.FILEINI);
                            string dirImg = System.IO.Path.GetFullPath(System.IO.Path.Combine(dirIni, @"..\..\Immagini\Articoli"));
                            if (!System.IO.Directory.Exists(dirImg))
                            {
                                try { System.IO.Directory.CreateDirectory(dirImg); } catch { dirImg = dirIni; }
                            }

                            // Cerca di inferire l'estensione dalla URL, di default .jpg
                            string ext = ".jpg";
                            if (urlOrPath.IndexOf(".png", StringComparison.OrdinalIgnoreCase) >= 0) ext = ".png";
                            else if (urlOrPath.IndexOf(".gif", StringComparison.OrdinalIgnoreCase) >= 0) ext = ".gif";

                            string destFile = System.IO.Path.Combine(dirImg, artCod + ext);
                            using (System.Net.WebClient wc = new System.Net.WebClient())
                            {
                                wc.DownloadFile(urlOrPath, destFile);
                            }
                            txtArtImg.Text = destFile;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore durante l'acquisizione dell'immagine: " + ex.Message, "Errore Drag&Drop", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtArtImg_DoubleClick(object sender, EventArgs e)
        {
            string sPath = txtArtImg.Text.Trim();
            if (string.IsNullOrEmpty(sPath))
            {
                MessageBox.Show("Nessun percorso immagine impostato.", "Anteprima", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!File.Exists(sPath))
            {
                MessageBox.Show("File immagine non trovato:\n" + sPath, "Anteprima", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Mostra anteprima in un dialog dedicato
            using (Form fPrev = new Form())
            {
                fPrev.Text = "Anteprima immagine — " + Path.GetFileName(sPath);
                fPrev.StartPosition = FormStartPosition.CenterParent;
                fPrev.Size = new Size(520, 460);
                fPrev.MinimizeBox = false;
                fPrev.MaximizeBox = false;
                fPrev.FormBorderStyle = FormBorderStyle.FixedDialog;
                fPrev.BackColor = Color.Black;

                PictureBox pb = new PictureBox();
                pb.Dock = DockStyle.Fill;
                pb.SizeMode = PictureBoxSizeMode.Zoom;
                pb.BackColor = Color.Black;
                try { pb.Image = Image.FromFile(sPath); }
                catch (Exception ex)
                {
                    MessageBox.Show("Impossibile aprire l'immagine:\n" + ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                fPrev.Controls.Add(pb);
                fPrev.ShowDialog(this);
                pb.Image.Dispose();
            }
        }

    }
}
