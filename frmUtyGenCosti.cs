using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmUtyGenCosti : Form
    {
        private clsDefine _clsDef = new clsDefine();
        private clsFuncs _clsFun = new clsFuncs();
        private clsQuery _clsQry = new clsQuery();
        private clsVariazioni _clsVar = new clsVariazioni();

        private string _strConSql = "";
        private DataTable _tabPreview = null;

        public frmUtyGenCosti()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
        }

        private void frmUtyGenCosti_Load(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_strConSql))
                    _strConSql = _clsFun.ConSql("");

                ApplyModernUi();
                clsUiIcons.RestoreFormBounds(this);

                dtpData.Value = DateTime.Today;
                nudPercScorporo.Value = 0.00m;
                chkSoloAttivi.Checked = true;
                chkAggiornaAna.Checked = true;

                SetDgvColumns();
                FillFornitori();
                clsUiIcons.RestoreGridColumnWidths(dgv1, "frmUtyGenCosti_dgv1");

                // Esegue subito l'estrazione anteprima all'avvio
                EstraiAnteprima();
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmUtyGenCosti.Load", ex.Message);
            }
        }

        private void ApplyModernUi()
        {
            try
            {
                this.BackColor = Color.FromArgb(243, 244, 246);

                if (menuStrip1 != null)
                {
                    menuStrip1.Renderer = clsUiIcons.GetModernMenuRenderer();
                    if (esciToolStripMenuItem != null)
                        esciToolStripMenuItem.Image = clsUiIcons.GetIcon("exit", 16);
                    if (estraiToolStripMenuItem != null)
                        estraiToolStripMenuItem.Image = clsUiIcons.GetIcon("search", 16);
                    if (generaToolStripMenuItem != null)
                        generaToolStripMenuItem.Image = clsUiIcons.GetIcon("check", 16);
                }

                if (panelHeader != null)
                {
                    panelHeader.BackColor = Color.FromArgb(248, 250, 252);
                }

                // Pulsante Estrai Anteprima: Sky Blue Glossy
                clsUiIcons.StyleStatButton(btnEstrai, "Estrai Anteprima", "search", 16,
                    Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254), Color.FromArgb(96, 165, 250),
                    Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));

                // Pulsante Genera Costi: Emerald Green Glossy
                clsUiIcons.StyleStatButton(btnGenera, "Genera Costi", "check", 16,
                    Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208), Color.FromArgb(52, 211, 153),
                    Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));

                // Pulsante Esci: Soft Red Glossy
                clsUiIcons.StyleStatButton(btnEsci, "Esci", "exit", 16,
                    Color.FromArgb(254, 242, 242), Color.FromArgb(254, 202, 202), Color.FromArgb(248, 113, 113),
                    Color.FromArgb(127, 29, 29), Color.FromArgb(220, 38, 38));

                clsUiIcons.StyleDataGridView(dgv1);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmUtyGenCosti.ApplyModernUi", ex.Message);
            }
        }

        private decimal CalcolaCoefficienteRicarico(decimal inputVal)
        {
            if (inputVal <= 0m) return 1.0m;
            // Se l'utente inserisce un valore >= 5.0 (es. 20, 30, 45), è espresso in percentuale (%)
            if (inputVal >= 5.0m)
            {
                return 1.0m + (inputVal / 100.0m);
            }
            // Se l'utente inserisce un valore tra 1.0 e 5.0 (es. 1.20, 1.30, 1.45), è il coefficiente diretto
            if (inputVal >= 1.0m)
            {
                return inputVal;
            }
            // Se l'utente inserisce un valore tra 0.01 e 0.99 (es. 0.30 per indicare +30%)
            return 1.0m + inputVal;
        }

        private void SetDgvColumns()
        {
            dgv1.AutoGenerateColumns = false;
            dgv1.Columns.Clear();

            DataGridViewTextBoxColumn cTbc;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_cod";
            cTbc.Name = "art_cod";
            cTbc.HeaderText = "Codice";
            cTbc.Width = 85;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "art_des";
            cTbc.HeaderText = "Descrizione Articolo";
            cTbc.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_ali";
            cTbc.Name = "tab_ali";
            cTbc.HeaderText = "IVA %";
            cTbc.Width = 70;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "prz_ven";
            cTbc.Name = "prz_ven";
            cTbc.HeaderText = "P. Vendita";
            cTbc.Width = 95;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "prz_net";
            cTbc.Name = "prz_net";
            cTbc.HeaderText = "P. Netto";
            cTbc.Width = 90;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "coeff_ric";
            cTbc.Name = "coeff_ric";
            cTbc.HeaderText = "Coeff. Ric.";
            cTbc.Width = 85;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "cos_net";
            cTbc.Name = "cos_net";
            cTbc.HeaderText = "Costo Calc.";
            cTbc.Width = 100;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0.000";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "cos_att";
            cTbc.Name = "cos_att";
            cTbc.HeaderText = "Costo Att.";
            cTbc.Width = 90;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0.000";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillFornitori()
        {
            try
            {
                string s = "SELECT for_cod, for_des FROM AnaFornitori WHERE (for_ann=0 OR for_ann IS NULL) ORDER BY for_des";
                DataTable t = _clsFun.FillTabSql("AnaFornitori", s, false, _strConSql);

                DataRow rVuoto = t.NewRow();
                rVuoto["for_cod"] = "";
                rVuoto["for_des"] = "-- Seleziona Fornitore --";
                t.Rows.InsertAt(rVuoto, 0);

                cmbFornitore.DataSource = t;
                cmbFornitore.DisplayMember = "for_des";
                cmbFornitore.ValueMember = "for_cod";
                cmbFornitore.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmUtyGenCosti.FillFornitori", ex.Message);
            }
        }

        private void btnEstrai_Click(object sender, EventArgs e)
        {
            EstraiAnteprima();
        }

        private void dtpData_ValueChanged(object sender, EventArgs e)
        {
            EstraiAnteprima();
        }

        private void nudPercScorporo_ValueChanged(object sender, EventArgs e)
        {
            EstraiAnteprima();
        }

        private void chkSoloAttivi_CheckedChanged(object sender, EventArgs e)
        {
            EstraiAnteprima();
        }

        private void EstraiAnteprima()
        {
            try
            {
                if (string.IsNullOrEmpty(_strConSql))
                    _strConSql = _clsFun.ConSql("");

                lblStatus.Text = "Estrazione articoli in corso...";
                progressBar1.Value = 0;
                this.Cursor = Cursors.WaitCursor;
                Application.DoEvents();

                // 1. Carica TabIva per il mapping aliquote
                Dictionary<string, decimal> dictIva = new Dictionary<string, decimal>();
                string sqlIva = "SELECT tab_cod, tab_ali FROM TabIva";
                DataTable tIva = _clsFun.FillTabSql("TabIva", sqlIva, false, _strConSql);
                if (tIva != null)
                {
                    foreach (DataRow r in tIva.Rows)
                    {
                        string cod = Convert.ToString(r["tab_cod"]).Trim();
                        decimal ali = 0m;
                        if (r.Table.Columns.Contains("tab_ali") && r["tab_ali"] != DBNull.Value)
                        {
                            ali = Convert.ToDecimal(r["tab_ali"]);
                        }
                        if (!string.IsNullOrEmpty(cod) && !dictIva.ContainsKey(cod))
                            dictIva[cod] = ali;
                    }
                }

                // 2. Carica listino vendita cassa (liv_lis = '001') più recente attivo con prezzo > 0
                Dictionary<string, decimal> dictLisVen = new Dictionary<string, decimal>();
                string sSqlVen = "SELECT liv_art, liv_prv FROM (" +
                                 "  SELECT ROW_NUMBER() OVER (PARTITION BY liv_art ORDER BY liv_dti DESC) AS ROW, " +
                                 "  liv_art, liv_prv " +
                                 "  FROM GesLisVendita " +
                                 "  WHERE liv_lis='" + _clsDef.LISPOS + "' " +
                                 "  AND (liv_ann=0 OR liv_ann IS NULL) " +
                                 "  AND (liv_prv > 0) " +
                                 ") AS A WHERE ROW = 1";
                DataTable tVen = _clsFun.FillTabSql("GesLisVendita", sSqlVen, false, _strConSql);
                if (tVen != null)
                {
                    foreach (DataRow r in tVen.Rows)
                    {
                        string art = Convert.ToString(r["liv_art"]).Trim();
                        decimal prv = (r["liv_prv"] != DBNull.Value) ? Convert.ToDecimal(r["liv_prv"]) : 0m;
                        if (!string.IsNullOrEmpty(art) && !dictLisVen.ContainsKey(art))
                            dictLisVen[art] = prv;
                    }
                }

                // 3. Carica Anagrafica Articoli
                string sqlArt = "SELECT art_cod, art_des, art_iva, art_prv, art_cos, art_sta FROM AnaArticoli";
                if (chkSoloAttivi.Checked)
                {
                    sqlArt += " WHERE art_sta = '" + _clsDef.STAATT + "'";
                }
                sqlArt += " ORDER BY art_des";

                DataTable tArt = _clsFun.FillTabSql("AnaArticoli", sqlArt, false, _strConSql);

                // 4. Creazione tabella anteprima
                _tabPreview = new DataTable("AnteprimaCosti");
                _tabPreview.Columns.Add("art_cod", typeof(string));
                _tabPreview.Columns.Add("art_des", typeof(string));
                _tabPreview.Columns.Add("art_iva", typeof(string));
                _tabPreview.Columns.Add("tab_ali", typeof(decimal));
                _tabPreview.Columns.Add("prz_ven", typeof(decimal));
                _tabPreview.Columns.Add("prz_net", typeof(decimal));
                _tabPreview.Columns.Add("coeff_ric", typeof(decimal));
                _tabPreview.Columns.Add("cos_net", typeof(decimal));
                _tabPreview.Columns.Add("cos_att", typeof(decimal));

                decimal inputRicarico = nudPercScorporo.Value;
                decimal coeffRic = CalcolaCoefficienteRicarico(inputRicarico);

                if (tArt != null && tArt.Rows.Count > 0)
                {
                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = tArt.Rows.Count;
                    progressBar1.Value = 0;

                    int count = 0;
                    foreach (DataRow r in tArt.Rows)
                    {
                        count++;
                        if (count % 300 == 0)
                        {
                            progressBar1.Value = count;
                            Application.DoEvents();
                        }

                        string cod = Convert.ToString(r["art_cod"]).Trim();
                        string des = Convert.ToString(r["art_des"]);
                        string iva = r.Table.Columns.Contains("art_iva") ? Convert.ToString(r["art_iva"]).Trim() : "";
                        decimal cosAtt = (r.Table.Columns.Contains("art_cos") && r["art_cos"] != DBNull.Value) ? Convert.ToDecimal(r["art_cos"]) : 0m;
                        decimal prvAna = (r.Table.Columns.Contains("art_prv") && r["art_prv"] != DBNull.Value) ? Convert.ToDecimal(r["art_prv"]) : 0m;

                        // Prezzo di vendita: prioritario da GesLisVendita '001', fallback ad AnaArticoli.art_prv
                        decimal przVen = prvAna;
                        if (dictLisVen.ContainsKey(cod) && dictLisVen[cod] > 0)
                        {
                            przVen = dictLisVen[cod];
                        }

                        // Aliquota IVA da TabIva
                        decimal aliIva = dictIva.ContainsKey(iva) ? dictIva[iva] : 0m;

                        // 1. Scorpora IVA dal Prezzo di Vendita Ivato per ottenere il Prezzo Netto: PrezzoNetto = P_vendita / (1 + IVA/100)
                        decimal przNetto = przVen;
                        if (aliIva > 0)
                        {
                            przNetto = przVen / (1.0m + (aliIva / 100.0m));
                        }

                        // 2. Calcola il Costo Netto di Acquisto in base al Coefficiente di Ricarico: CostoNetto = PrezzoNetto / K
                        decimal costoNetto = przNetto;
                        if (coeffRic > 0)
                        {
                            costoNetto = przNetto / coeffRic;
                        }
                        costoNetto = Math.Round(costoNetto, 3, MidpointRounding.AwayFromZero);

                        DataRow rowPrev = _tabPreview.NewRow();
                        rowPrev["art_cod"] = cod;
                        rowPrev["art_des"] = des;
                        rowPrev["art_iva"] = iva;
                        rowPrev["tab_ali"] = aliIva;
                        rowPrev["prz_ven"] = przVen;
                        rowPrev["prz_net"] = Math.Round(przNetto, 2, MidpointRounding.AwayFromZero);
                        rowPrev["coeff_ric"] = Math.Round(coeffRic, 2);
                        rowPrev["cos_net"] = costoNetto;
                        rowPrev["cos_att"] = cosAtt;

                        _tabPreview.Rows.Add(rowPrev);
                    }

                    progressBar1.Value = progressBar1.Maximum;
                }
                else
                {
                    progressBar1.Value = 0;
                }

                dgv1.DataSource = _tabPreview;

                lblRiepilogo.Text = "Articoli estratti: " + _tabPreview.Rows.Count.ToString("N0", CultureInfo.CurrentCulture);
                lblStatus.Text = string.Format("Anteprima completata ({0} articoli, Coeff. Ricarico: {1:0.00}). Selezionare il fornitore e premere 'Genera Costi'.", _tabPreview.Rows.Count, coeffRic);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmUtyGenCosti.EstraiAnteprima", ex.Message);
                MessageBox.Show("Errore durante l'estrazione dei dati:\n\n" + ex.Message, "ERRORE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "Errore durante l'estrazione dei dati.";
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void btnGenera_Click(object sender, EventArgs e)
        {
            GeneraCosti();
        }

        private void GeneraCosti()
        {
            try
            {
                // Controllo selezione fornitore
                string forCod = cmbFornitore.SelectedValue != null ? cmbFornitore.SelectedValue.ToString().Trim() : "";
                if (string.IsNullOrEmpty(forCod))
                {
                    MessageBox.Show("Selezionare un Fornitore valido dall'elenco prima di procedere.", "SELEZIONE FORNITORE", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    cmbFornitore.Focus();
                    return;
                }

                string forDes = cmbFornitore.Text;

                // Se l'anteprima non è stata ancora caricata o è vuota, la esegue ora
                if (_tabPreview == null || _tabPreview.Rows.Count == 0)
                {
                    EstraiAnteprima();
                }

                if (_tabPreview == null || _tabPreview.Rows.Count == 0)
                {
                    MessageBox.Show("Nessun articolo disponibile per la generazione dei costi.", "ATTENZIONE", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                decimal coeffRic = CalcolaCoefficienteRicarico(nudPercScorporo.Value);

                string msg = string.Format(
                    "Confermi la generazione e il salvataggio dei costi nel listino di acquisto per il fornitore:\n\n" +
                    "• Fornitore: {0} ({1})\n" +
                    "• Data Registrazione: {2}\n" +
                    "• Coeff. Ricarico applicato: {3:0.00}\n" +
                    "• Numero articoli da elaborare: {4}\n" +
                    "• Aggiornamento costo anagrafica: {5}\n\n" +
                    "Continuare?",
                    forDes,
                    forCod,
                    dtpData.Value.ToString("dd/MM/yyyy"),
                    coeffRic,
                    _tabPreview.Rows.Count,
                    chkAggiornaAna.Checked ? "SI" : "NO");

                DialogResult dr = MessageBox.Show(msg, "CONFERMA GENERAZIONE COSTI", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr != DialogResult.Yes) return;

                this.Cursor = Cursors.WaitCursor;
                lblStatus.Text = "Salvataggio costi in corso nel listino acquisti...";
                progressBar1.Minimum = 0;
                progressBar1.Maximum = _tabPreview.Rows.Count;
                progressBar1.Value = 0;
                Application.DoEvents();

                string dtiSql = _clsFun.DaySql(dtpData.Value);
                string dtfSql = _clsFun.DaySql(_clsDef.DAYOUT);
                bool aggAna = chkAggiornaAna.Checked;

                List<string> batchSql = new List<string>();
                int totalProcessed = 0;
                bool hasErrors = false;

                foreach (DataRow r in _tabPreview.Rows)
                {
                    string artCod = Convert.ToString(r["art_cod"]).Trim();
                    decimal cosNet = Convert.ToDecimal(r["cos_net"]);
                    decimal przVen = Convert.ToDecimal(r["prz_ven"]);

                    string cosNetSql = cosNet.ToString(CultureInfo.InvariantCulture);
                    string przVenSql = przVen.ToString(CultureInfo.InvariantCulture);

                    // 1. Inserimento / Aggiornamento in GesLisAcquisto per il fornitore e la data specificata (con codice articolo fornitore = codice articolo e lia_day)
                    string sSqlLia = "IF EXISTS (SELECT 1 FROM GesLisAcquisto WHERE lia_tip='L' AND lia_art='" + artCod + "' AND lia_for='" + forCod + "' AND lia_dti=" + dtiSql + ") " +
                                     "UPDATE GesLisAcquisto SET lia_cos=" + cosNetSql + ", lia_prv=" + przVenSql + ", lia_arf='" + artCod + "', lia_ann=0, lia_dtf=" + dtfSql + ", lia_day=GETDATE() WHERE lia_tip='L' AND lia_art='" + artCod + "' AND lia_for='" + forCod + "' AND lia_dti=" + dtiSql + " " +
                                     "ELSE " +
                                     "INSERT INTO GesLisAcquisto (lia_tip, lia_art, lia_for, lia_arf, lia_cos, lia_prv, lia_dti, lia_dtf, lia_pxc, lia_cxp, lia_stf, lia_ann, lia_day) " +
                                     "VALUES ('L', '" + artCod + "', '" + forCod + "', '" + artCod + "', " + cosNetSql + ", " + przVenSql + ", " + dtiSql + ", " + dtfSql + ", 1, 0, 'A', 0, GETDATE())";

                    batchSql.Add(sSqlLia);

                    // 2. Aggiornamento contestuale in AnaArticoli (art_cos) se richiesto
                    if (aggAna)
                    {
                        string sSqlAna = "UPDATE AnaArticoli SET art_cos=" + cosNetSql + " WHERE art_cod='" + artCod + "'";
                        batchSql.Add(sSqlAna);
                    }

                    totalProcessed++;

                    // Esecuzione a blocchi di 300 istruzioni per performance ottimali
                    if (batchSql.Count >= 300)
                    {
                        bool ok = _clsFun.SqlWriteBatch(batchSql, _strConSql);
                        if (!ok) hasErrors = true;
                        batchSql.Clear();

                        progressBar1.Value = totalProcessed;
                        lblStatus.Text = string.Format("Elaborati {0} di {1} articoli...", totalProcessed, _tabPreview.Rows.Count);
                        Application.DoEvents();
                    }
                }

                // Esegui eventuali query residue
                if (batchSql.Count > 0)
                {
                    bool ok = _clsFun.SqlWriteBatch(batchSql, _strConSql);
                    if (!ok) hasErrors = true;
                    batchSql.Clear();
                }

                progressBar1.Value = progressBar1.Maximum;

                // Log dell'operazione
                string[] aLog = {
                    "GENCOSTI",
                    forCod,
                    "",
                    "",
                    "",
                    nudPercScorporo.Value.ToString(CultureInfo.InvariantCulture),
                    "",
                    "Generati costi da listino vendita per fornitore " + forDes + " su " + totalProcessed + " articoli con scorporo " + nudPercScorporo.Value.ToString("0.00") + "%"
                };
                _clsQry.LogSql(aLog);

                if (hasErrors)
                {
                    lblStatus.Text = "Operazione completata con alcune segnalazioni.";
                    MessageBox.Show(
                        string.Format("Generazione costi completata con alcune anomalie!\n\nVerificare il file Error.log. Elaborati {0} articoli per il fornitore {1}.", totalProcessed, forDes),
                        "ATTENZIONE",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                else
                {
                    lblStatus.Text = "Operazione completata con successo!";
                    MessageBox.Show(
                        string.Format("Generazione costi completata con successo!\n\nElaborati {0} articoli per il fornitore {1}.", totalProcessed, forDes),
                        "OPERAZIONE COMPLETATA",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmUtyGenCosti.GeneraCosti", ex.Message);
                MessageBox.Show("Errore durante la generazione dei costi:\n\n" + ex.Message, "ERRORE SALVATAGGIO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void btnEsci_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private void frmUtyGenCosti_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
            else if (e.KeyCode == Keys.F5)
                EstraiAnteprima();
            else if (e.KeyCode == Keys.F8)
                GeneraCosti();
        }

        private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                string colName = dgv1.Columns[e.ColumnIndex].Name;
                if (colName == "art_des" || colName == "art_cod")
                {
                    ApriAnagraficaArticolo(e.RowIndex);
                }
            }
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                ApriAnagraficaArticolo(e.RowIndex);
            }
        }

        private void ApriAnagraficaArticolo(int rowIndex)
        {
            try
            {
                if (rowIndex >= 0 && rowIndex < dgv1.Rows.Count)
                {
                    string artCod = Convert.ToString(dgv1.Rows[rowIndex].Cells["art_cod"].Value).Trim();
                    if (!string.IsNullOrEmpty(artCod))
                    {
                        frmAnaArticolo f = new frmAnaArticolo();
                        f._strArtCod = artCod;
                        f.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmUtyGenCosti.ApriAnagraficaArticolo", ex.Message);
            }
        }

        private void frmUtyGenCosti_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                clsUiIcons.SaveFormBounds(this);
                clsUiIcons.SaveGridColumnWidths(dgv1, "frmUtyGenCosti_dgv1");
            }
            catch { }
        }
    }
}
