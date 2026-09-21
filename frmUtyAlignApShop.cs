using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Globalization;

namespace APOffice
{
    public partial class frmUtyAlignApShop : Form
    {
        private clsDefine _clsDef = new clsDefine();
        private clsFuncs _clsFun = new clsFuncs();
        private string _strConSql = "";
        private DataTable _dtResult = null;
        private bool _bolCancel = false;

        public frmUtyAlignApShop() : this(@"C:\ApProject\ApShop\DataBase\DataBase.mdb")
        {
        }

        public frmUtyAlignApShop(string sInitialPath)
        {
            InitializeComponent();
            this.DoubleBuffered = true;
            this.KeyPreview = true;

            if (!string.IsNullOrEmpty(sInitialPath))
                txtMdbPath.Text = sInitialPath;

            dgvGrid.CellFormatting += new DataGridViewCellFormattingEventHandler(dgvGrid_CellFormatting);
            txtMdbPath.DoubleClick += new EventHandler(txtMdbPath_DoubleClick);

            if (cmbSort.Items.Count > 0)
                cmbSort.SelectedIndex = 0;

            if (cmbPriceStrategy.Items.Count > 0)
                cmbPriceStrategy.SelectedIndex = 0;

            this.FormClosing += new FormClosingEventHandler(this.frmUtyAlignApShop_FormClosing);
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmUtyAlignApShop_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            if (File.Exists(txtMdbPath.Text))
            {
                // Auto-analizza all'avvio se il file MDB esiste nel percorso specificato
                EseguiConfronto();
            }
        }

        private void frmUtyAlignApShop_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                this.Close();
            }
        }

        private void btnBrowseMdb_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Database MS Access (*.mdb)|*.mdb|Tutti i file (*.*)|*.*";
            ofd.Title = "Seleziona il database APShop MDB";
            if (!string.IsNullOrEmpty(txtMdbPath.Text) && File.Exists(txtMdbPath.Text))
            {
                try { ofd.InitialDirectory = Path.GetDirectoryName(txtMdbPath.Text); } catch { }
            }
            else
            {
                ofd.InitialDirectory = @"C:\ApProject\ApShop\DataBase\";
            }

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtMdbPath.Text = ofd.FileName;
                EseguiConfronto();
            }
        }

        private void txtMdbPath_DoubleClick(object sender, EventArgs e)
        {
            PromptChangePath();
        }

        private void PromptChangePath()
        {
            frmUtyInput1 fInput = new frmUtyInput1();
            fInput._strDes = "Percorso Database APShop (MDB):";
            fInput._strTxt = txtMdbPath.Text.Trim();
            fInput.ShowDialog(this);

            string sNewPath = fInput._strTxt.Trim();
            if (!string.IsNullOrEmpty(sNewPath) && sNewPath != txtMdbPath.Text)
            {
                txtMdbPath.Text = sNewPath;
                if (File.Exists(sNewPath))
                {
                    EseguiConfronto();
                }
            }
        }

        private void btnAnalizza_Click(object sender, EventArgs e)
        {
            EseguiConfronto();
        }

        private OleDbConnection GetMdbConnection(string mdbPath)
        {
            string[] providers = new string[] {
                "Microsoft.Jet.OLEDB.4.0",
                "Microsoft.ACE.OLEDB.12.0",
                "Microsoft.ACE.OLEDB.16.0"
            };

            Exception lastEx = null;
            foreach (string provider in providers)
            {
                try
                {
                    string connStr = string.Format("Provider={0};Data Source=\"{1}\";Mode=Share Deny None;Persist Security Info=False;", provider, mdbPath);
                    OleDbConnection cn = new OleDbConnection(connStr);
                    cn.Open();
                    return cn;
                }
                catch (Exception ex)
                {
                    lastEx = ex;
                }
            }

            // Fallback via _clsFun.ConMdb
            try
            {
                string sCon = _clsFun.ConMdb(mdbPath);
                OleDbConnection cnFallback = new OleDbConnection(sCon);
                cnFallback.Open();
                return cnFallback;
            }
            catch { }

            throw lastEx ?? new Exception("Impossibile connettersi al file database MDB:\n" + mdbPath);
        }

        private DataTable FillTabMdbSafe(string tableName, string sql, string mdbPath)
        {
            try
            {
                using (OleDbConnection cn = GetMdbConnection(mdbPath))
                {
                    DataTable schemaTables = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    HashSet<string> mdbTables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    if (schemaTables != null)
                    {
                        foreach (DataRow r in schemaTables.Rows)
                        {
                            string tName = Convert.ToString(r["TABLE_NAME"]).Trim();
                            if (!string.IsNullOrEmpty(tName)) mdbTables.Add(tName);
                        }
                    }

                    string actualTableToQuery = null;
                    if (mdbTables.Contains(tableName))
                    {
                        actualTableToQuery = tableName;
                    }
                    else
                    {
                        if (tableName.Equals("GesLisVendita", StringComparison.OrdinalIgnoreCase))
                        {
                            if (mdbTables.Contains("LisVendita")) actualTableToQuery = "LisVendita";
                            else if (mdbTables.Contains("TabLisVendita")) actualTableToQuery = "TabLisVendita";
                            else if (mdbTables.Contains("GesLisVen")) actualTableToQuery = "GesLisVen";
                        }
                        else if (tableName.Equals("GesLisAcquisto", StringComparison.OrdinalIgnoreCase))
                        {
                            if (mdbTables.Contains("LisAcquisto")) actualTableToQuery = "LisAcquisto";
                            else if (mdbTables.Contains("TabLisAcquisto")) actualTableToQuery = "TabLisAcquisto";
                            else if (mdbTables.Contains("GesLisAcq")) actualTableToQuery = "GesLisAcq";
                        }
                    }

                    if (string.IsNullOrEmpty(actualTableToQuery))
                    {
                        // Se la tabella non esiste nel database MDB di questa versione APShop, ritorna DataTable vuota in silenzio senza errori
                        return new DataTable(tableName);
                    }

                    string actualSql = sql.Replace(tableName, actualTableToQuery);
                    using (OleDbDataAdapter da = new OleDbDataAdapter(actualSql, cn))
                    {
                        DataTable dt = new DataTable(tableName);
                        da.Fill(dt);
                        cn.Close();
                        return dt;
                    }
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("FillTabMdbSafe " + tableName, ex.Message);
                return new DataTable(tableName);
            }
        }

        private string Format3Digits(string sInput, string sDefault = "000")
        {
            if (string.IsNullOrEmpty(sInput)) return sDefault;
            string sClean = sInput.Trim();
            if (int.TryParse(sClean, out int val))
                return val.ToString("000");
            return sClean.PadLeft(3, '0');
        }

        private string FormatDateStr(object objDate)
        {
            if (objDate == null || DBNull.Value.Equals(objDate)) return "-";
            if (objDate is DateTime dt)
            {
                return dt.Year < 1990 ? "-" : dt.ToString("dd/MM/yyyy");
            }
            string s = Convert.ToString(objDate).Trim();
            if (string.IsNullOrEmpty(s)) return "-";
            if (DateTime.TryParse(s, out DateTime dtParsed))
            {
                return dtParsed.Year < 1990 ? "-" : dtParsed.ToString("dd/MM/yyyy");
            }
            return s;
        }

        private DateTime? ParseDate(string sDate)
        {
            if (string.IsNullOrEmpty(sDate) || sDate == "-") return null;
            if (DateTime.TryParseExact(sDate, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dtExact))
            {
                return dtExact;
            }
            if (DateTime.TryParse(sDate, out DateTime dtParsed))
            {
                return dtParsed;
            }
            return null;
        }

        private bool CheckMdbBilValue(DataRow rMdb, DataRow rMdbEan)
        {
            if (rMdb != null && rMdb.Table != null)
            {
                foreach (DataColumn col in rMdb.Table.Columns)
                {
                    if (col.ColumnName.IndexOf("bil", StringComparison.OrdinalIgnoreCase) >= 0 && !DBNull.Value.Equals(rMdb[col]))
                    {
                        try
                        {
                            if (Convert.ToBoolean(rMdb[col]) || Convert.ToInt32(rMdb[col]) == 1)
                                return true;
                        }
                        catch { }
                    }
                }
            }
            if (rMdbEan != null && rMdbEan.Table != null)
            {
                foreach (DataColumn col in rMdbEan.Table.Columns)
                {
                    if (col.ColumnName.IndexOf("bil", StringComparison.OrdinalIgnoreCase) >= 0 && !DBNull.Value.Equals(rMdbEan[col]))
                    {
                        try
                        {
                            if (Convert.ToBoolean(rMdbEan[col]) || Convert.ToInt32(rMdbEan[col]) == 1)
                                return true;
                        }
                        catch { }
                    }
                }
            }
            return false;
        }

        private void EseguiConfronto()
        {
            string sMdbPath = txtMdbPath.Text.Trim();
            if (!File.Exists(sMdbPath))
            {
                MessageBox.Show("Il file del database APShop non è stato trovato nel percorso specificato:\n" + sMdbPath,
                    "DATABASE NON TROVATO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                this.Cursor = Cursors.WaitCursor;
                lblTotApShop.Text = "⚡ APERTURA IN CORSO... Caricamento e confronto anagrafiche in corso, attendere...";
                lblTotApShop.ForeColor = Color.DarkBlue;
                progressBar1.Style = ProgressBarStyle.Marquee;
                progressBar1.Visible = true;
                Application.DoEvents();

                // 1. Carica dati da APShop MDB
                DataTable tMdbArt = FillTabMdbSafe("AnaArticoli", "SELECT * FROM AnaArticoli", sMdbPath);
                DataTable tMdbEan = FillTabMdbSafe("AnaBarcode", "SELECT * FROM AnaBarcode WHERE ean_ean IS NOT NULL AND ean_ean <> ''", sMdbPath);
                DataTable tMdbVen = FillTabMdbSafe("GesLisVendita", "SELECT * FROM GesLisVendita", sMdbPath);
                DataTable tMdbAcq = FillTabMdbSafe("GesLisAcquisto", "SELECT * FROM GesLisAcquisto", sMdbPath);

                if (tMdbArt == null || tMdbArt.Rows.Count == 0)
                {
                    MessageBox.Show("Nessun articolo trovato nella tabella AnaArticoli del file APShop MDB.",
                        "TABELLA VUOTA O NON TROVATA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                // 2. Carica dati da APOffice SQL Server
                DataTable tSqlArt = _clsFun.FillTabSql("AnaArticoli", "SELECT * FROM AnaArticoli", false, _strConSql);
                DataTable tSqlEan = _clsFun.FillTabSql("AnaBarcode", "SELECT * FROM AnaBarcode", false, _strConSql);
                DataTable tSqlVen = _clsFun.FillTabSql("GesLisVendita", "SELECT * FROM GesLisVendita", false, _strConSql);
                DataTable tSqlAcq = _clsFun.FillTabSql("GesLisAcquisto", "SELECT * FROM GesLisAcquisto", false, _strConSql);

                // 3. Prepara strutture di confronto rapide (Dictionary)
                Dictionary<string, DataRow> dictSqlArt = new Dictionary<string, DataRow>();
                if (tSqlArt != null)
                {
                    foreach (DataRow r in tSqlArt.Rows)
                    {
                        string cod = Convert.ToString(r["art_cod"]).Trim();
                        if (!string.IsNullOrEmpty(cod) && !dictSqlArt.ContainsKey(cod))
                            dictSqlArt.Add(cod, r);
                    }
                }

                HashSet<string> setSqlEan = new HashSet<string>();
                if (tSqlEan != null)
                {
                    foreach (DataRow r in tSqlEan.Rows)
                    {
                        string ean = Convert.ToString(r["ean_ean"]).Trim();
                        if (!string.IsNullOrEmpty(ean) && !setSqlEan.Contains(ean))
                            setSqlEan.Add(ean);
                    }
                }

                Dictionary<string, string> dictMdbFirstEan = new Dictionary<string, string>();
                Dictionary<string, decimal> dictMdbEanPrv = new Dictionary<string, decimal>();
                if (tMdbEan != null && tMdbEan.Columns.Contains("ean_art") && tMdbEan.Columns.Contains("ean_ean"))
                {
                    foreach (DataRow r in tMdbEan.Rows)
                    {
                        string art = Convert.ToString(r["ean_art"]).Trim();
                        string ean = Convert.ToString(r["ean_ean"]).Trim();
                        if (!string.IsNullOrEmpty(art) && !string.IsNullOrEmpty(ean))
                        {
                            if (!dictMdbFirstEan.ContainsKey(art))
                                dictMdbFirstEan.Add(art, ean);

                            if (tMdbEan.Columns.Contains("ean_prv") && !DBNull.Value.Equals(r["ean_prv"]))
                            {
                                decimal pEan = Convert.ToDecimal(r["ean_prv"]);
                                if (pEan > 0 && !dictMdbEanPrv.ContainsKey(art))
                                    dictMdbEanPrv.Add(art, pEan);
                            }
                        }
                    }
                }

                Dictionary<string, DataRow> dictMdbVenRow = new Dictionary<string, DataRow>();
                Dictionary<string, decimal> dictMdbPrv = new Dictionary<string, decimal>();
                if (tMdbVen != null && tMdbVen.Columns.Contains("liv_art"))
                {
                    foreach (DataRow r in tMdbVen.Rows)
                    {
                        string art = Convert.ToString(r["liv_art"]).Trim();
                        if (!string.IsNullOrEmpty(art))
                        {
                            if (!dictMdbVenRow.ContainsKey(art))
                                dictMdbVenRow.Add(art, r);

                            if (tMdbVen.Columns.Contains("liv_prv") && !DBNull.Value.Equals(r["liv_prv"]))
                            {
                                decimal prv = Convert.ToDecimal(r["liv_prv"]);
                                if (prv > 0 && !dictMdbPrv.ContainsKey(art))
                                    dictMdbPrv.Add(art, prv);
                            }
                        }
                    }
                }

                Dictionary<string, DataRow> dictMdbAcqRow = new Dictionary<string, DataRow>();
                Dictionary<string, decimal> dictMdbCos = new Dictionary<string, decimal>();
                if (tMdbAcq != null && tMdbAcq.Columns.Contains("lia_art"))
                {
                    foreach (DataRow r in tMdbAcq.Rows)
                    {
                        string art = Convert.ToString(r["lia_art"]).Trim();
                        if (!string.IsNullOrEmpty(art))
                        {
                            if (!dictMdbAcqRow.ContainsKey(art))
                                dictMdbAcqRow.Add(art, r);

                            if (tMdbAcq.Columns.Contains("lia_cos") && !DBNull.Value.Equals(r["lia_cos"]))
                            {
                                decimal cos = Convert.ToDecimal(r["lia_cos"]);
                                if (!dictMdbCos.ContainsKey(art))
                                    dictMdbCos.Add(art, cos);
                            }
                        }
                    }
                }

                // Strutture APOffice SQL Server per date e listini
                Dictionary<string, DataRow> dictSqlVenRow = new Dictionary<string, DataRow>();
                if (tSqlVen != null && tSqlVen.Columns.Contains("liv_art"))
                {
                    foreach (DataRow r in tSqlVen.Rows)
                    {
                        string art = Convert.ToString(r["liv_art"]).Trim();
                        if (!string.IsNullOrEmpty(art) && !dictSqlVenRow.ContainsKey(art))
                            dictSqlVenRow.Add(art, r);
                    }
                }

                Dictionary<string, DataRow> dictSqlAcqRow = new Dictionary<string, DataRow>();
                if (tSqlAcq != null && tSqlAcq.Columns.Contains("lia_art"))
                {
                    foreach (DataRow r in tSqlAcq.Rows)
                    {
                        string art = Convert.ToString(r["lia_art"]).Trim();
                        if (!string.IsNullOrEmpty(art) && !dictSqlAcqRow.ContainsKey(art))
                            dictSqlAcqRow.Add(art, r);
                    }
                }

                // 4. Prepara DataTable per la griglia di risultato
                DataTable dtRes = new DataTable();
                dtRes.Columns.Add("art_cod", typeof(string));
                dtRes.Columns.Add("art_des", typeof(string));
                dtRes.Columns.Add("ean_ean", typeof(string));
                dtRes.Columns.Add("art_rep", typeof(string));
                dtRes.Columns.Add("art_iva", typeof(string));
                dtRes.Columns.Add("art_umi", typeof(string));
                dtRes.Columns.Add("art_prv", typeof(decimal));
                dtRes.Columns.Add("apo_prv", typeof(decimal));
                dtRes.Columns.Add("diff_prv", typeof(string));
                dtRes.Columns.Add("dt_ven_apshop", typeof(string));
                dtRes.Columns.Add("dt_ven_apoffice", typeof(string));
                dtRes.Columns.Add("art_cos", typeof(decimal));
                dtRes.Columns.Add("apo_cos", typeof(decimal));
                dtRes.Columns.Add("diff_cos", typeof(string));
                dtRes.Columns.Add("dt_acq_apshop", typeof(string));
                dtRes.Columns.Add("dt_acq_apoffice", typeof(string));
                dtRes.Columns.Add("stato_sync", typeof(string));
                dtRes.Columns.Add("sort_rank", typeof(int));

                int nMancanti = 0;
                int nDiversi = 0;
                int nAllineati = 0;

                if (tMdbArt != null)
                {
                    foreach (DataRow rMdb in tMdbArt.Rows)
                    {
                        string rawCod = Convert.ToString(rMdb["art_cod"]).Trim();
                        if (string.IsNullOrEmpty(rawCod)) continue;
                        string cod = TruncateField(tSqlArt, "art_cod", rawCod);

                        string desApShop = Convert.ToString(rMdb["art_des"]).Trim();
                        string desAPOffice = dictSqlArt.ContainsKey(cod) ? Convert.ToString(dictSqlArt[cod]["art_des"]).Trim() : "";

                        // MANTIENI DESCRIZIONE PIÙ LUNGA E COMPLETA
                        string desFinal = desApShop;
                        if (chkKeepLongerDes.Checked && !string.IsNullOrEmpty(desAPOffice) && desAPOffice.Length > desApShop.Length)
                        {
                            desFinal = desAPOffice;
                        }

                        string ean = dictMdbFirstEan.ContainsKey(cod) ? dictMdbFirstEan[cod] : "";

                        // Reparto ed IVA formattati a 3 cifre con zeri (es: 001, 022, 004)
                        string rawRep = tMdbArt.Columns.Contains("art_rep") && !DBNull.Value.Equals(rMdb["art_rep"]) ? Convert.ToString(rMdb["art_rep"]).Trim() : "001";
                        string rep = Format3Digits(rawRep, "001");

                        string rawIva = tMdbArt.Columns.Contains("art_iva") && !DBNull.Value.Equals(rMdb["art_iva"]) ? Convert.ToString(rMdb["art_iva"]).Trim() : "022";
                        string iva = Format3Digits(rawIva, "022");

                        // UMI (Unità di Misura): default 'NR' se non specificato
                        string umi = tMdbArt.Columns.Contains("art_umi") && !DBNull.Value.Equals(rMdb["art_umi"]) ? Convert.ToString(rMdb["art_umi"]).Trim() : "";
                        if (string.IsNullOrEmpty(umi)) umi = "NR";

                        // Estrazione Prezzo Cassa APShop (POS Selling Price Hierarchy)
                        decimal prvApShop = 0m;
                        if (tMdbArt.Columns.Contains("art_pve") && !DBNull.Value.Equals(rMdb["art_pve"]))
                            prvApShop = Convert.ToDecimal(rMdb["art_pve"]);
                        if (prvApShop <= 0m && dictMdbPrv.ContainsKey(cod))
                            prvApShop = dictMdbPrv[cod];
                        if (prvApShop <= 0m && dictMdbEanPrv.ContainsKey(cod))
                            prvApShop = dictMdbEanPrv[cod];
                        if (prvApShop <= 0m && tMdbArt.Columns.Contains("art_prv") && !DBNull.Value.Equals(rMdb["art_prv"]))
                            prvApShop = Convert.ToDecimal(rMdb["art_prv"]);

                        // FILTRI ESCLUSIONE INIZIALI
                        if (chkIgnoreNoEan.Checked && string.IsNullOrEmpty(ean))
                            continue;
                        if (chkIgnoreNoPrv.Checked && prvApShop <= 0m)
                            continue;

                        decimal cos = 0m;
                        if (dictMdbCos.ContainsKey(cod))
                            cos = dictMdbCos[cod];
                        else if (tMdbArt.Columns.Contains("art_cos") && !DBNull.Value.Equals(rMdb["art_cos"]))
                            cos = Convert.ToDecimal(rMdb["art_cos"]);

                        // Controllo completo date di ultimo aggiornamento in APShop
                        string dtVenApShop = "-";
                        if (dictMdbVenRow.ContainsKey(cod))
                        {
                            DataRow rVenMdb = dictMdbVenRow[cod];
                            if (rVenMdb.Table.Columns.Contains("liv_dtf") && !DBNull.Value.Equals(rVenMdb["liv_dtf"]))
                                dtVenApShop = FormatDateStr(rVenMdb["liv_dtf"]);
                            else if (rVenMdb.Table.Columns.Contains("liv_dti") && !DBNull.Value.Equals(rVenMdb["liv_dti"]))
                                dtVenApShop = FormatDateStr(rVenMdb["liv_dti"]);
                        }
                        if (dtVenApShop == "-" && rMdb.Table.Columns.Contains("art_dtm") && !DBNull.Value.Equals(rMdb["art_dtm"]))
                            dtVenApShop = FormatDateStr(rMdb["art_dtm"]);
                        if (dtVenApShop == "-" && rMdb.Table.Columns.Contains("art_dti") && !DBNull.Value.Equals(rMdb["art_dti"]))
                            dtVenApShop = FormatDateStr(rMdb["art_dti"]);

                        string dtAcqApShop = "-";
                        if (dictMdbAcqRow.ContainsKey(cod))
                        {
                            DataRow rAcqMdb = dictMdbAcqRow[cod];
                            if (rAcqMdb.Table.Columns.Contains("lia_dti") && !DBNull.Value.Equals(rAcqMdb["lia_dti"]))
                                dtAcqApShop = FormatDateStr(rAcqMdb["lia_dti"]);
                        }
                        if (dtAcqApShop == "-" && rMdb.Table.Columns.Contains("art_dtm") && !DBNull.Value.Equals(rMdb["art_dtm"]))
                            dtAcqApShop = FormatDateStr(rMdb["art_dtm"]);
                        if (dtAcqApShop == "-" && rMdb.Table.Columns.Contains("art_dti") && !DBNull.Value.Equals(rMdb["art_dti"]))
                            dtAcqApShop = FormatDateStr(rMdb["art_dti"]);

                        // Valori e Date APOffice SQL Server
                        decimal prvSql = 0m;
                        decimal cosSql = 0m;
                        string dtVenAPOffice = "-";
                        string dtAcqAPOffice = "-";

                        if (dictSqlArt.ContainsKey(cod))
                        {
                            DataRow rSql = dictSqlArt[cod];
                            if (rSql.Table.Columns.Contains("art_pve") && !DBNull.Value.Equals(rSql["art_pve"]) && Convert.ToDecimal(rSql["art_pve"]) > 0)
                                prvSql = Convert.ToDecimal(rSql["art_pve"]);
                            else if (rSql.Table.Columns.Contains("art_prv") && !DBNull.Value.Equals(rSql["art_prv"]))
                                prvSql = Convert.ToDecimal(rSql["art_prv"]);

                            if (dictSqlVenRow.ContainsKey(cod))
                            {
                                DataRow rVenSql = dictSqlVenRow[cod];
                                if (rVenSql.Table.Columns.Contains("liv_prv") && !DBNull.Value.Equals(rVenSql["liv_prv"]) && Convert.ToDecimal(rVenSql["liv_prv"]) > 0)
                                    prvSql = Convert.ToDecimal(rVenSql["liv_prv"]);
                            }
                        }

                        // FLAG FILTRO 3: Escludi prodotti con Prezzo Vendita uguale tra APShop e APOffice
                        if (chkIgnoreSamePrice.Checked && dictSqlArt.ContainsKey(cod) && Math.Abs(prvApShop - prvSql) <= 0.001m)
                            continue;

                        // STRATEGIA PREZZO DA ALLINEARE
                        decimal targetPrv = prvApShop;
                        int selStrat = cmbPriceStrategy.SelectedIndex;
                        if (selStrat == 1) // Mantieni Prezzo Maggiore (Max tra APShop e APOffice)
                        {
                            if (dictSqlArt.ContainsKey(cod) && prvSql > prvApShop)
                                targetPrv = prvSql;
                        }
                        else if (selStrat == 2) // Mantieni Prezzo con Data più Recente
                        {
                            DateTime? dtApShop = ParseDate(dtVenApShop);
                            DateTime? dtAPOffice = ParseDate(dtVenAPOffice);
                            if (dtAPOffice.HasValue && dtApShop.HasValue && dtAPOffice.Value > dtApShop.Value && prvSql > 0)
                            {
                                targetPrv = prvSql;
                            }
                        }

                        string stato = "ALLINEATO";

                        if (!dictSqlArt.ContainsKey(cod))
                        {
                            stato = "INSERIRE";
                            nMancanti++;
                        }
                        else
                        {
                            DataRow rSql = dictSqlArt[cod];
                            string desSql = Convert.ToString(rSql["art_des"]).Trim();

                            string rawRepSql = rSql.Table.Columns.Contains("art_rep") && !DBNull.Value.Equals(rSql["art_rep"]) ? Convert.ToString(rSql["art_rep"]).Trim() : "";
                            string repSql = Format3Digits(rawRepSql, "001");

                            string rawIvaSql = rSql.Table.Columns.Contains("art_iva") && !DBNull.Value.Equals(rSql["art_iva"]) ? Convert.ToString(rSql["art_iva"]).Trim() : "";
                            string ivaSql = Format3Digits(rawIvaSql, "022");

                            string umiSql = rSql.Table.Columns.Contains("art_umi") && !DBNull.Value.Equals(rSql["art_umi"]) ? Convert.ToString(rSql["art_umi"]).Trim() : "";

                            if (dictSqlVenRow.ContainsKey(cod))
                            {
                                DataRow rVenSql = dictSqlVenRow[cod];
                                if (rVenSql.Table.Columns.Contains("liv_dtf") && !DBNull.Value.Equals(rVenSql["liv_dtf"]))
                                    dtVenAPOffice = FormatDateStr(rVenSql["liv_dtf"]);
                                else if (rVenSql.Table.Columns.Contains("liv_dti") && !DBNull.Value.Equals(rVenSql["liv_dti"]))
                                    dtVenAPOffice = FormatDateStr(rVenSql["liv_dti"]);
                            }
                            if (dtVenAPOffice == "-" && rSql.Table.Columns.Contains("art_dtm") && !DBNull.Value.Equals(rSql["art_dtm"]))
                                dtVenAPOffice = FormatDateStr(rSql["art_dtm"]);
                            if (dtVenAPOffice == "-" && rSql.Table.Columns.Contains("art_dti") && !DBNull.Value.Equals(rSql["art_dti"]))
                                dtVenAPOffice = FormatDateStr(rSql["art_dti"]);

                            if (dictSqlAcqRow.ContainsKey(cod))
                            {
                                DataRow rAcqSql = dictSqlAcqRow[cod];
                                if (rAcqSql.Table.Columns.Contains("lia_cos") && !DBNull.Value.Equals(rAcqSql["lia_cos"]))
                                    cosSql = Convert.ToDecimal(rAcqSql["lia_cos"]);
                                if (rAcqSql.Table.Columns.Contains("lia_dti") && !DBNull.Value.Equals(rAcqSql["lia_dti"]))
                                    dtAcqAPOffice = FormatDateStr(rAcqSql["lia_dti"]);
                            }
                            if (cosSql <= 0m && rSql.Table.Columns.Contains("art_cos") && !DBNull.Value.Equals(rSql["art_cos"]))
                                cosSql = Convert.ToDecimal(rSql["art_cos"]);
                            if (dtAcqAPOffice == "-" && rSql.Table.Columns.Contains("art_dtm") && !DBNull.Value.Equals(rSql["art_dtm"]))
                                dtAcqAPOffice = FormatDateStr(rSql["art_dtm"]);
                            if (dtAcqAPOffice == "-" && rSql.Table.Columns.Contains("art_dti") && !DBNull.Value.Equals(rSql["art_dti"]))
                                dtAcqAPOffice = FormatDateStr(rSql["art_dti"]);

                            bool hasEanInSql = string.IsNullOrEmpty(ean) || setSqlEan.Contains(ean);

                            if (!desFinal.Equals(desSql, StringComparison.OrdinalIgnoreCase) ||
                                !rep.Equals(repSql, StringComparison.OrdinalIgnoreCase) ||
                                !iva.Equals(ivaSql, StringComparison.OrdinalIgnoreCase) ||
                                !umi.Equals(umiSql, StringComparison.OrdinalIgnoreCase) ||
                                Math.Abs(targetPrv - prvSql) > 0.001m ||
                                Math.Abs(cos - cosSql) > 0.001m ||
                                !hasEanInSql)
                            {
                                stato = "AGGIORNARE";
                                nDiversi++;
                            }
                            else
                            {
                                nAllineati++;
                            }
                        }

                        // Calcolo Differenze Prezzo e Costo basato sul prezzo target calcolato
                        decimal dPrvDiff = targetPrv - prvSql;
                        string sDiffPrv = dPrvDiff == 0m ? "0.00" : (dPrvDiff > 0 ? "+" + dPrvDiff.ToString("0.00") : dPrvDiff.ToString("0.00"));

                        decimal dCosDiff = cos - cosSql;
                        string sDiffCos = dCosDiff == 0m ? "0.00" : (dCosDiff > 0 ? "+" + dCosDiff.ToString("0.00") : dCosDiff.ToString("0.00"));

                        DataRow rowRes = dtRes.NewRow();
                        rowRes["art_cod"] = cod;
                        rowRes["art_des"] = desFinal;
                        rowRes["ean_ean"] = ean;
                        rowRes["art_rep"] = rep;
                        rowRes["art_iva"] = iva;
                        rowRes["art_umi"] = umi;
                        rowRes["art_prv"] = targetPrv;
                        rowRes["apo_prv"] = prvSql;
                        rowRes["diff_prv"] = sDiffPrv;
                        rowRes["dt_ven_apshop"] = dtVenApShop;
                        rowRes["dt_ven_apoffice"] = dtVenAPOffice;
                        rowRes["art_cos"] = cos;
                        rowRes["apo_cos"] = cosSql;
                        rowRes["diff_cos"] = sDiffCos;
                        rowRes["dt_acq_apshop"] = dtAcqApShop;
                        rowRes["dt_acq_apoffice"] = dtAcqAPOffice;
                        rowRes["stato_sync"] = stato;
                        rowRes["sort_rank"] = (stato == "INSERIRE") ? 1 : ((stato == "AGGIORNARE") ? 2 : 3);
                        dtRes.Rows.Add(rowRes);
                    }
                }

                _dtResult = dtRes;

                lblTotApShop.Text = "Articoli APShop: " + (tMdbArt != null ? dtRes.Rows.Count.ToString("N0") : "0");
                lblMancanti.Text = "Da Inserire: " + nMancanti.ToString("N0");
                lblDiversi.Text = "Da Aggiornare: " + nDiversi.ToString("N0");
                lblAllineati.Text = "Allineati: " + nAllineati.ToString("N0");

                btnAllinea.Enabled = (nMancanti + nDiversi) > 0;

                ApplyFilter();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore durante l'analisi del database APShop:\n" + ex.Message,
                    "ERRORE CONFRONTO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _clsFun.ErrorLog("EseguiConfronto APShop", ex.Message);
            }
            finally
            {
                progressBar1.Style = ProgressBarStyle.Blocks;
                progressBar1.Visible = false;
                this.Cursor = Cursors.Default;
            }
        }

        private void cmbSort_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (_dtResult == null) return;

            DataView dv = new DataView(_dtResult);
            if (chkOnlyDiff.Checked)
            {
                dv.RowFilter = "stato_sync <> 'ALLINEATO'";
            }
            else
            {
                dv.RowFilter = "";
            }

            int selSort = cmbSort.SelectedIndex;
            if (selSort == 0) // Da Inserire -> Da Aggiornare -> Allineati
            {
                foreach (DataRowView drv in dv)
                {
                    string st = Convert.ToString(drv["stato_sync"]);
                    drv["sort_rank"] = (st == "INSERIRE") ? 1 : ((st == "AGGIORNARE") ? 2 : 3);
                }
                dv.Sort = "sort_rank ASC, art_cod ASC";
            }
            else if (selSort == 1) // Da Aggiornare -> Da Inserire -> Allineati
            {
                foreach (DataRowView drv in dv)
                {
                    string st = Convert.ToString(drv["stato_sync"]);
                    drv["sort_rank"] = (st == "AGGIORNARE") ? 1 : ((st == "INSERIRE") ? 2 : 3);
                }
                dv.Sort = "sort_rank ASC, art_cod ASC";
            }
            else if (selSort == 2) // Allineati -> Da Inserire -> Da Aggiornare
            {
                foreach (DataRowView drv in dv)
                {
                    string st = Convert.ToString(drv["stato_sync"]);
                    drv["sort_rank"] = (st == "ALLINEATO") ? 1 : ((st == "INSERIRE") ? 2 : 3);
                }
                dv.Sort = "sort_rank ASC, art_cod ASC";
            }
            else if (selSort == 3) // Codice Articolo A-Z
            {
                dv.Sort = "art_cod ASC";
            }
            else if (selSort == 4) // Descrizione A-Z
            {
                dv.Sort = "art_des ASC";
            }

            dgvGrid.DataSource = dv;
        }

        private void dgvGrid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < dgvGrid.Rows.Count)
            {
                DataGridViewRow row = dgvGrid.Rows[e.RowIndex];
                if (row.Cells["colStato"].Value != null)
                {
                    string st = row.Cells["colStato"].Value.ToString();
                    if (st == "INSERIRE")
                    {
                        row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235);
                        row.DefaultCellStyle.ForeColor = Color.DarkRed;
                    }
                    else if (st == "AGGIORNARE")
                    {
                        row.DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 220);
                        row.DefaultCellStyle.ForeColor = Color.DarkOrange;
                    }
                    else
                    {
                        row.DefaultCellStyle.BackColor = Color.White;
                        row.DefaultCellStyle.ForeColor = Color.Black;
                    }
                }
            }
        }

        private void chkOnlyDiff_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter();
        }

        private void chkFilter_CheckedChanged(object sender, EventArgs e)
        {
            if (File.Exists(txtMdbPath.Text.Trim()))
            {
                EseguiConfronto();
            }
        }

        private void btnAllinea_Click(object sender, EventArgs e)
        {
            if (_dtResult == null) return;

            DataRow[] rowsToSync = _dtResult.Select("stato_sync <> 'ALLINEATO'");
            if (rowsToSync.Length == 0)
            {
                MessageBox.Show("Nessun articolo da inserire o aggiornare.", "ALLINEAMENTO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DialogResult res = MessageBox.Show(
                "Verranno allineati " + rowsToSync.Length.ToString("N0") + " articoli da APShop a APOffice.\n\n" +
                "I prodotti verranno sincronizzati con i seguenti valori predefiniti ove mancanti:\n" +
                "• art_tgr = 'PZ', art_umi = 'NR', art_pxc = 1, art_eti = '043'\n" +
                "• art_bil ed ean_bil = 1 per prodotti bilancia o con EAN che inizia per 2, altrimenti 0\n" +
                "• ean_ecp = 0, ean_dti ed ean_dtm valorizzati a data odierna\n" +
                "• Stato Attivo ('A'), data inserimento/modifica odierna.\n\n" +
                "Confermi l'allineamento?",
                "CONFERMA ALLINEAMENTO ANAGRAFICA",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (res != DialogResult.Yes) return;

            string sMdbPath = txtMdbPath.Text.Trim();

            try
            {
                this.Cursor = Cursors.WaitCursor;
                progressBar1.Value = 0;
                progressBar1.Maximum = rowsToSync.Length;
                progressBar1.Visible = true;

                // 1. Carica dati sorgente MDB completi per estrarre eventuali campi avanzati e righe barcode
                DataTable tMdbArt = FillTabMdbSafe("AnaArticoli", "SELECT * FROM AnaArticoli", sMdbPath);
                DataTable tMdbEan = FillTabMdbSafe("AnaBarcode", "SELECT * FROM AnaBarcode WHERE ean_ean IS NOT NULL AND ean_ean <> ''", sMdbPath);

                Dictionary<string, DataRow> dictMdbArtRow = new Dictionary<string, DataRow>();
                if (tMdbArt != null)
                {
                    foreach (DataRow r in tMdbArt.Rows)
                    {
                        string c = Convert.ToString(r["art_cod"]).Trim();
                        if (!string.IsNullOrEmpty(c) && !dictMdbArtRow.ContainsKey(c))
                            dictMdbArtRow.Add(c, r);
                    }
                }

                Dictionary<string, DataRow> dictMdbEanRow = new Dictionary<string, DataRow>();
                if (tMdbEan != null && tMdbEan.Columns.Contains("ean_art"))
                {
                    foreach (DataRow r in tMdbEan.Rows)
                    {
                        string art = Convert.ToString(r["ean_art"]).Trim();
                        if (!string.IsNullOrEmpty(art) && !dictMdbEanRow.ContainsKey(art))
                            dictMdbEanRow.Add(art, r);
                    }
                }

                // 2. Carica tabelle di schema SQL Server per rilevamento dinamico delle colonne
                DataTable tArtSchema = _clsFun.FillTabSql("AnaArticoli", "SELECT TOP 1 * FROM AnaArticoli WHERE 1=0", false, _strConSql);
                DataTable tEanSchema = _clsFun.FillTabSql("AnaBarcode", "SELECT TOP 1 * FROM AnaBarcode WHERE 1=0", false, _strConSql);
                DataTable tLivSchema = _clsFun.FillTabSql("GesLisVendita", "SELECT TOP 1 * FROM GesLisVendita WHERE 1=0", false, _strConSql);
                DataTable tLiaSchema = _clsFun.FillTabSql("GesLisAcquisto", "SELECT TOP 1 * FROM GesLisAcquisto WHERE 1=0", false, _strConSql);

                bool hasArtPve = (tArtSchema != null && tArtSchema.Columns.Contains("art_pve"));
                bool hasArtTgr = (tArtSchema != null && tArtSchema.Columns.Contains("art_tgr"));
                bool hasArtPxc = (tArtSchema != null && tArtSchema.Columns.Contains("art_pxc"));
                bool hasArtBil = (tArtSchema != null && tArtSchema.Columns.Contains("art_bil"));
                bool hasArtEti = (tArtSchema != null && tArtSchema.Columns.Contains("art_eti"));

                bool hasEanBil = (tEanSchema != null && tEanSchema.Columns.Contains("ean_bil"));
                bool hasEanEcp = (tEanSchema != null && tEanSchema.Columns.Contains("ean_ecp"));

                int nIns = 0;
                int nUpd = 0;
                List<string> lstErrors = new List<string>();

                // Pre-carica i codici articolo e EAN esistenti per prevenire errori di chiave duplicata (IX_AnaArticoli)
                HashSet<string> existingSqlArtCods = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                DataTable tAllArt = _clsFun.FillTabSql("AllArt", "SELECT art_cod FROM AnaArticoli", false, _strConSql);
                if (tAllArt != null)
                {
                    foreach (DataRow r in tAllArt.Rows)
                    {
                        string c = Convert.ToString(r["art_cod"]).Trim();
                        if (!string.IsNullOrEmpty(c) && !existingSqlArtCods.Contains(c))
                            existingSqlArtCods.Add(c);
                    }
                }

                _bolCancel = false;
                btnStop.Visible = true;
                btnStop.Enabled = true;
                btnAllinea.Enabled = false;
                btnAnalizza.Enabled = false;

                using (SqlConnection cn = new SqlConnection(_strConSql))
                {
                    cn.Open();
                    using (SqlTransaction tr = cn.BeginTransaction())
                    {
                        try
                        {
                            int counterSync = 0;
                            foreach (DataRow rSync in rowsToSync)
                            {
                                counterSync++;
                                if (counterSync % 15 == 0 || counterSync == rowsToSync.Length)
                                {
                                    progressBar1.Value = Math.Min(counterSync, progressBar1.Maximum);
                                    Application.DoEvents();
                                }

                                if (_bolCancel)
                                {
                                    tr.Rollback();
                                    MessageBox.Show("Procedura di allineamento interrotta dall'utente.\nTutte le modifiche della sessione sono state annullate.",
                                        "PROCEDURA INTERROTTA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    return;
                                }

                                string rawArtCod = Convert.ToString(rSync["art_cod"]).Trim();
                                if (string.IsNullOrEmpty(rawArtCod)) continue;
                                string artCod = TruncateField(tArtSchema, "art_cod", rawArtCod);

                                DataRow rMdb = dictMdbArtRow.ContainsKey(rawArtCod) ? dictMdbArtRow[rawArtCod] : (dictMdbArtRow.ContainsKey(artCod) ? dictMdbArtRow[artCod] : null);
                                DataRow rMdbEan = dictMdbEanRow.ContainsKey(rawArtCod) ? dictMdbEanRow[rawArtCod] : (dictMdbEanRow.ContainsKey(artCod) ? dictMdbEanRow[artCod] : null);

                                string st = Convert.ToString(rSync["stato_sync"]).Trim();
                                // Se il codice articolo esiste già in SQL Server, passa in automatico ad AGGIORNARE!
                                if (st == "INSERIRE" && existingSqlArtCods.Contains(artCod))
                                {
                                    st = "AGGIORNARE";
                                }

                                string des = Convert.ToString(rSync["art_des"]).Trim();
                                string rep = Format3Digits(Convert.ToString(rSync["art_rep"]), "001");
                                string iva = Format3Digits(Convert.ToString(rSync["art_iva"]), "022");

                                // UMI / UMC: 'NR' se mancanti
                                string umi = Convert.ToString(rSync["art_umi"]).Trim();
                                if (string.IsNullOrEmpty(umi)) umi = "NR";

                                string umc = "NR";
                                if (rMdb != null && rMdb.Table.Columns.Contains("art_umc") && !DBNull.Value.Equals(rMdb["art_umc"]))
                                {
                                    string sUmcMdb = Convert.ToString(rMdb["art_umc"]).Trim();
                                    if (!string.IsNullOrEmpty(sUmcMdb)) umc = sUmcMdb;
                                }

                                decimal prv = 0m;
                                decimal cos = 0m;
                                try { prv = Convert.ToDecimal(rSync["art_prv"]); } catch { }
                                try { cos = Convert.ToDecimal(rSync["art_cos"]); } catch { }
                                string ean = Convert.ToString(rSync["ean_ean"]).Trim();

                                // Pezzatura (art_tgr): 'PZ' se mancante
                                string tgr = "PZ";
                                if (rMdb != null && rMdb.Table.Columns.Contains("art_tgr") && !DBNull.Value.Equals(rMdb["art_tgr"]))
                                {
                                    string sTgrMdb = Convert.ToString(rMdb["art_tgr"]).Trim();
                                    if (!string.IsNullOrEmpty(sTgrMdb)) tgr = sTgrMdb;
                                }

                                // Pezzi per confezione (art_pxc): 1 se mancante o <= 0
                                int pxc = 1;
                                if (rMdb != null && rMdb.Table.Columns.Contains("art_pxc") && !DBNull.Value.Equals(rMdb["art_pxc"]))
                                {
                                    try
                                    {
                                        int pMdb = Convert.ToInt32(rMdb["art_pxc"]);
                                        if (pMdb > 0) pxc = pMdb;
                                    }
                                    catch { }
                                }

                                // Gestito da bilancia (art_bil / ean_bil): 1 se EAN inizia per 2 o se bilancia in APShop (*bil*), altrimenti 0
                                int bil = 0;
                                if (ean.StartsWith("2") || CheckMdbBilValue(rMdb, rMdbEan))
                                {
                                    bil = 1;
                                }

                                // Etichetta (art_eti): '043' se mancante
                                string eti = "043";
                                if (rMdb != null && rMdb.Table.Columns.Contains("art_eti") && !DBNull.Value.Equals(rMdb["art_eti"]))
                                {
                                    string sEtiMdb = Convert.ToString(rMdb["art_eti"]).Trim();
                                    if (!string.IsNullOrEmpty(sEtiMdb)) eti = Format3Digits(sEtiMdb, "043");
                                }

                                // TRONCAMENTO SICURO LUNGHEZZE CAMPI
                                if (artCod.Length > 20) artCod = artCod.Substring(0, 20);
                                if (des.Length > 100) des = des.Substring(0, 100);
                                if (rep.Length > 3) rep = rep.Substring(0, 3);
                                if (iva.Length > 3) iva = iva.Substring(0, 3);
                                if (umi.Length > 3) umi = umi.Substring(0, 3);
                                if (umc.Length > 3) umc = umc.Substring(0, 3);
                                if (tgr.Length > 5) tgr = tgr.Substring(0, 5);
                                if (eti.Length > 3) eti = eti.Substring(0, 3);
                                if (ean.Length > 20) ean = ean.Substring(0, 20);

                                string artCodEsc = artCod.Replace("'", "''");
                                string desEsc = des.Replace("'", "''");
                                string ivaEsc = iva.Replace("'", "''");
                                string repEsc = rep.Replace("'", "''");
                                string umiEsc = umi.Replace("'", "''");
                                string umcEsc = umc.Replace("'", "''");
                                string tgrEsc = tgr.Replace("'", "''");
                                string etiEsc = eti.Replace("'", "''");
                                string eanEsc = ean.Replace("'", "''");

                                string sCosStr = cos.ToString("0.0000", CultureInfo.InvariantCulture);
                                string sPrvStr = prv.ToString("0.0000", CultureInfo.InvariantCulture);

                                try
                                {
                                    DataRow rNewArt = tArtSchema.NewRow();
                                    foreach (DataColumn col in tArtSchema.Columns)
                                    {
                                        if (col.DataType == typeof(string)) rNewArt[col] = "";
                                        else if (col.DataType == typeof(decimal) || col.DataType == typeof(double) || col.DataType == typeof(float)) rNewArt[col] = 0m;
                                        else if (col.DataType == typeof(int) || col.DataType == typeof(short) || col.DataType == typeof(long) || col.DataType == typeof(byte)) rNewArt[col] = 0;
                                        else if (col.DataType == typeof(DateTime)) rNewArt[col] = DateTime.Today;
                                        else if (col.DataType == typeof(bool)) rNewArt[col] = false;
                                    }

                                    // Copia eventuali campi presenti in APShop (rMdb)
                                    if (rMdb != null)
                                    {
                                        foreach (DataColumn col in tArtSchema.Columns)
                                        {
                                            if (rMdb.Table.Columns.Contains(col.ColumnName) && !DBNull.Value.Equals(rMdb[col.ColumnName]))
                                            {
                                                string sValMdb = rMdb[col.ColumnName].ToString().Trim();
                                                if (!string.IsNullOrEmpty(sValMdb))
                                                {
                                                    if (col.DataType == typeof(string))
                                                        rNewArt[col] = TruncateField(tArtSchema, col.ColumnName, sValMdb);
                                                    else if (col.DataType == typeof(decimal) || col.DataType == typeof(double) || col.DataType == typeof(float))
                                                        rNewArt[col] = _clsFun.Txt2Dec(sValMdb);
                                                    else if (col.DataType == typeof(int) || col.DataType == typeof(short) || col.DataType == typeof(long) || col.DataType == typeof(byte))
                                                        rNewArt[col] = Convert.ToInt32(_clsFun.Txt2Dec(sValMdb));
                                                }
                                            }
                                        }
                                    }

                                    // Regole ed Inquadramento Obbligatori APShop
                                    if (tArtSchema.Columns.Contains("art_cod")) rNewArt["art_cod"] = TruncateField(tArtSchema, "art_cod", artCod);
                                    if (tArtSchema.Columns.Contains("art_des")) rNewArt["art_des"] = TruncateField(tArtSchema, "art_des", des);
                                    if (tArtSchema.Columns.Contains("art_deb")) rNewArt["art_deb"] = TruncateField(tArtSchema, "art_deb", (des.Length > 20 ? des.Substring(0, 20) : des));
                                    if (tArtSchema.Columns.Contains("art_iva")) rNewArt["art_iva"] = TruncateField(tArtSchema, "art_iva", iva);
                                    if (tArtSchema.Columns.Contains("art_rep")) rNewArt["art_rep"] = TruncateField(tArtSchema, "art_rep", rep);
                                    if (tArtSchema.Columns.Contains("art_umi")) rNewArt["art_umi"] = TruncateField(tArtSchema, "art_umi", string.IsNullOrEmpty(umi) ? "NR" : umi);
                                    if (tArtSchema.Columns.Contains("art_umc")) rNewArt["art_umc"] = TruncateField(tArtSchema, "art_umc", string.IsNullOrEmpty(umc) ? "NR" : umc);
                                    if (tArtSchema.Columns.Contains("art_tgr")) rNewArt["art_tgr"] = TruncateField(tArtSchema, "art_tgr", string.IsNullOrEmpty(tgr) ? "PZ" : tgr);
                                    if (tArtSchema.Columns.Contains("art_eti")) rNewArt["art_eti"] = TruncateField(tArtSchema, "art_eti", string.IsNullOrEmpty(eti) ? "043" : eti);
                                    if (tArtSchema.Columns.Contains("art_sta")) rNewArt["art_sta"] = "A";
                                    if (tArtSchema.Columns.Contains("art_cos")) rNewArt["art_cos"] = cos;
                                    if (tArtSchema.Columns.Contains("art_prv")) rNewArt["art_prv"] = prv;
                                    if (tArtSchema.Columns.Contains("art_pve")) rNewArt["art_pve"] = prv;
                                    if (tArtSchema.Columns.Contains("art_pxc")) rNewArt["art_pxc"] = (pxc > 0 ? pxc : 1);
                                    if (tArtSchema.Columns.Contains("art_bil")) rNewArt["art_bil"] = (bil == 1);
                                    if (tArtSchema.Columns.Contains("art_dti")) rNewArt["art_dti"] = DateTime.Today;
                                    if (tArtSchema.Columns.Contains("art_dtm")) rNewArt["art_dtm"] = DateTime.Now.ToString("dd/MM/yyyy");

                                    if (st == "INSERIRE")
                                    {
                                        try
                                        {
                                            string sSqlInsArt = _clsFun.SqlInsertRow("AnaArticoli", tArtSchema, rNewArt);
                                            using (SqlCommand cm = new SqlCommand(sSqlInsArt, cn, tr)) { cm.ExecuteNonQuery(); }
                                            existingSqlArtCods.Add(artCod);
                                            nIns++;
                                        }
                                        catch (SqlException exSql) when (exSql.Number == 2601 || exSql.Number == 2627)
                                        {
                                            // Se la chiave esiste già (IX_AnaArticoli), esegui UPDATE di protezione!
                                            st = "AGGIORNARE";
                                        }
                                    }

                                    if (st == "AGGIORNARE")
                                    {
                                        List<string> updSets = new List<string>() {
                                            "art_des = '" + desEsc + "'",
                                            "art_deb = '" + desEsc + "'",
                                            "art_iva = '" + ivaEsc + "'",
                                            "art_rep = '" + repEsc + "'",
                                            "art_umi = '" + umiEsc + "'",
                                            "art_umc = ISNULL(NULLIF(art_umc, ''), 'NR')",
                                            "art_sta = 'A'",
                                            "art_cos = " + sCosStr,
                                            "art_prv = " + sPrvStr,
                                            "art_dtm = CONVERT(varchar, GETDATE(), 103)"
                                        };

                                        if (hasArtPve) updSets.Add("art_pve = " + sPrvStr);
                                        if (hasArtTgr) updSets.Add("art_tgr = '" + tgrEsc + "'");
                                        if (hasArtPxc) updSets.Add("art_pxc = " + pxc.ToString());
                                        if (hasArtBil) updSets.Add("art_bil = " + bil.ToString());
                                        if (hasArtEti) updSets.Add("art_eti = '" + etiEsc + "'");

                                        string sSqlUpdArt = string.Format(@"
UPDATE AnaArticoli SET 
    {1}
WHERE art_cod = '{0}'", artCodEsc, string.Join(",\n    ", updSets.ToArray()));

                                        using (SqlCommand cm = new SqlCommand(sSqlUpdArt, cn, tr)) 
                                        { 
                                            int rUpdCount = cm.ExecuteNonQuery(); 
                                            if (rUpdCount == 0)
                                            {
                                                string sSqlInsArt = _clsFun.SqlInsertRow("AnaArticoli", tArtSchema, rNewArt);
                                                using (SqlCommand cmIns = new SqlCommand(sSqlInsArt, cn, tr)) { cmIns.ExecuteNonQuery(); }
                                                existingSqlArtCods.Add(artCod);
                                                nIns++;
                                            }
                                            else
                                            {
                                                nUpd++;
                                            }
                                        }
                                    }

                                    // 2. Synch Barcode (AnaBarcode) con pre-inizializzazione campi vuoti
                                    if (!string.IsNullOrEmpty(ean))
                                    {
                                        DataRow rNewEan = tEanSchema.NewRow();
                                        foreach (DataColumn col in tEanSchema.Columns)
                                        {
                                            if (col.DataType == typeof(string)) rNewEan[col] = "";
                                            else if (col.DataType == typeof(decimal) || col.DataType == typeof(double) || col.DataType == typeof(float)) rNewEan[col] = 0m;
                                            else if (col.DataType == typeof(int) || col.DataType == typeof(short) || col.DataType == typeof(long) || col.DataType == typeof(byte)) rNewEan[col] = 0;
                                            else if (col.DataType == typeof(DateTime)) rNewEan[col] = DateTime.Today;
                                            else if (col.DataType == typeof(bool)) rNewEan[col] = false;
                                        }

                                        if (tEanSchema.Columns.Contains("ean_ean")) rNewEan["ean_ean"] = TruncateField(tEanSchema, "ean_ean", ean);
                                        if (tEanSchema.Columns.Contains("ean_art")) rNewEan["ean_art"] = TruncateField(tEanSchema, "ean_art", artCod);
                                        if (tEanSchema.Columns.Contains("ean_qta")) rNewEan["ean_qta"] = 1;
                                        if (tEanSchema.Columns.Contains("ean_prv")) rNewEan["ean_prv"] = prv;
                                        if (tEanSchema.Columns.Contains("ean_bil")) rNewEan["ean_bil"] = (bil == 1);
                                        if (tEanSchema.Columns.Contains("ean_ecp")) rNewEan["ean_ecp"] = 0;
                                        if (tEanSchema.Columns.Contains("ean_dti")) rNewEan["ean_dti"] = DateTime.Today;
                                        if (tEanSchema.Columns.Contains("ean_dtm")) rNewEan["ean_dtm"] = DateTime.Today;
                                        if (tEanSchema.Columns.Contains("ean_ann")) rNewEan["ean_ann"] = 0;

                                        string sCheckEan = "SELECT COUNT(*) FROM AnaBarcode WHERE ean_ean=@ean";
                                        using (SqlCommand cmCheckEan = new SqlCommand(sCheckEan, cn, tr))
                                        {
                                            cmCheckEan.Parameters.AddWithValue("@ean", ean);
                                            int cntEan = (int)cmCheckEan.ExecuteScalar();
                                            if (cntEan == 0)
                                            {
                                                string sInsEan = _clsFun.SqlInsertRow("AnaBarcode", tEanSchema, rNewEan);
                                                using (SqlCommand cmInsEan = new SqlCommand(sInsEan, cn, tr)) { cmInsEan.ExecuteNonQuery(); }
                                            }
                                            else
                                            {
                                                string sUpdEan = "UPDATE AnaBarcode SET ean_prv=@prv, ean_dtm=GETDATE(), ean_bil=@bil, ean_ecp=0 WHERE ean_ean=@ean";
                                                using (SqlCommand cmUpdEan = new SqlCommand(sUpdEan, cn, tr))
                                                {
                                                    cmUpdEan.Parameters.AddWithValue("@prv", prv);
                                                    cmUpdEan.Parameters.AddWithValue("@bil", bil);
                                                    cmUpdEan.Parameters.AddWithValue("@ean", ean);
                                                    cmUpdEan.ExecuteNonQuery();
                                                }
                                            }
                                        }
                                    }

                                    // 3. Synch Prezzo Vendita (GesLisVendita) con pre-inizializzazione
                                    DataRow rNewLiv = tLivSchema.NewRow();
                                    foreach (DataColumn col in tLivSchema.Columns)
                                    {
                                        if (col.DataType == typeof(string)) rNewLiv[col] = "";
                                        else if (col.DataType == typeof(decimal) || col.DataType == typeof(double) || col.DataType == typeof(float)) rNewLiv[col] = 0m;
                                        else if (col.DataType == typeof(int) || col.DataType == typeof(short) || col.DataType == typeof(long) || col.DataType == typeof(byte)) rNewLiv[col] = 0;
                                        else if (col.DataType == typeof(DateTime)) rNewLiv[col] = DateTime.Today;
                                        else if (col.DataType == typeof(bool)) rNewLiv[col] = false;
                                    }
                                    if (tLivSchema.Columns.Contains("liv_lis")) rNewLiv["liv_lis"] = _clsDef.LISPOS;
                                    if (tLivSchema.Columns.Contains("liv_art")) rNewLiv["liv_art"] = TruncateField(tLivSchema, "liv_art", artCod);
                                    if (tLivSchema.Columns.Contains("liv_qta")) rNewLiv["liv_qta"] = 1;
                                    if (tLivSchema.Columns.Contains("liv_prv")) rNewLiv["liv_prv"] = prv;
                                    if (tLivSchema.Columns.Contains("liv_dti")) rNewLiv["liv_dti"] = DateTime.Today;
                                    if (tLivSchema.Columns.Contains("liv_dtf")) rNewLiv["liv_dtf"] = _clsDef.DAYOUT;
                                    if (tLivSchema.Columns.Contains("liv_sta")) rNewLiv["liv_sta"] = "A";
                                    if (tLivSchema.Columns.Contains("liv_ann")) rNewLiv["liv_ann"] = 0;

                                    string sCheckLiv = "SELECT COUNT(*) FROM GesLisVendita WHERE liv_art=@art AND liv_lis=@lis AND liv_ann=0";
                                    using (SqlCommand cmCheckLiv = new SqlCommand(sCheckLiv, cn, tr))
                                    {
                                        cmCheckLiv.Parameters.AddWithValue("@art", artCod);
                                        cmCheckLiv.Parameters.AddWithValue("@lis", _clsDef.LISPOS);
                                        int cntLiv = (int)cmCheckLiv.ExecuteScalar();
                                        if (cntLiv == 0)
                                        {
                                            string sInsLiv = _clsFun.SqlInsertRow("GesLisVendita", tLivSchema, rNewLiv);
                                            using (SqlCommand cmInsLiv = new SqlCommand(sInsLiv, cn, tr)) { cmInsLiv.ExecuteNonQuery(); }
                                        }
                                        else
                                        {
                                            string sUpdLiv = "UPDATE GesLisVendita SET liv_prv=@prv, liv_dtf=" + _clsFun.DaySql(_clsDef.DAYOUT) + " WHERE liv_art=@art AND liv_lis=@lis AND liv_ann=0";
                                            using (SqlCommand cmUpdLiv = new SqlCommand(sUpdLiv, cn, tr))
                                            {
                                                cmUpdLiv.Parameters.AddWithValue("@prv", prv);
                                                cmUpdLiv.Parameters.AddWithValue("@art", artCod);
                                                cmUpdLiv.Parameters.AddWithValue("@lis", _clsDef.LISPOS);
                                                cmUpdLiv.ExecuteNonQuery();
                                            }
                                        }
                                    }

                                    // 4. Synch Costo Acquisto (GesLisAcquisto) con pre-inizializzazione
                                    if (cos > 0)
                                    {
                                        DataRow rNewLia = tLiaSchema.NewRow();
                                        foreach (DataColumn col in tLiaSchema.Columns)
                                        {
                                            if (col.DataType == typeof(string)) rNewLia[col] = "";
                                            else if (col.DataType == typeof(decimal) || col.DataType == typeof(double) || col.DataType == typeof(float)) rNewLia[col] = 0m;
                                            else if (col.DataType == typeof(int) || col.DataType == typeof(short) || col.DataType == typeof(long) || col.DataType == typeof(byte)) rNewLia[col] = 0;
                                            else if (col.DataType == typeof(DateTime)) rNewLia[col] = DateTime.Today;
                                            else if (col.DataType == typeof(bool)) rNewLia[col] = false;
                                        }
                                        if (tLiaSchema.Columns.Contains("lia_tip")) rNewLia["lia_tip"] = "F";
                                        if (tLiaSchema.Columns.Contains("lia_for")) rNewLia["lia_for"] = "000001";
                                        if (tLiaSchema.Columns.Contains("lia_art")) rNewLia["lia_art"] = TruncateField(tLiaSchema, "lia_art", artCod);
                                        if (tLiaSchema.Columns.Contains("lia_cos")) rNewLia["lia_cos"] = cos;
                                        if (tLiaSchema.Columns.Contains("lia_cxp")) rNewLia["lia_cxp"] = 1;
                                        if (tLiaSchema.Columns.Contains("lia_day")) rNewLia["lia_day"] = DateTime.Today;
                                        if (tLiaSchema.Columns.Contains("lia_dti")) rNewLia["lia_dti"] = DateTime.Today;
                                        if (tLiaSchema.Columns.Contains("lia_dtf")) rNewLia["lia_dtf"] = _clsDef.DAYOUT;
                                        if (tLiaSchema.Columns.Contains("lia_ann")) rNewLia["lia_ann"] = 0;

                                        string sCheckLia = "SELECT COUNT(*) FROM GesLisAcquisto WHERE lia_art=@art AND lia_ann=0";
                                        using (SqlCommand cmCheckLia = new SqlCommand(sCheckLia, cn, tr))
                                        {
                                            cmCheckLia.Parameters.AddWithValue("@art", artCod);
                                            int cntLia = (int)cmCheckLia.ExecuteScalar();
                                            if (cntLia == 0)
                                            {
                                                string sInsLia = _clsFun.SqlInsertRow("GesLisAcquisto", tLiaSchema, rNewLia);
                                                using (SqlCommand cmInsLia = new SqlCommand(sInsLia, cn, tr)) { cmInsLia.ExecuteNonQuery(); }
                                            }
                                            else
                                            {
                                                string sUpdLia = "UPDATE GesLisAcquisto SET lia_cos=@cos, lia_dti=GETDATE() WHERE lia_art=@art AND lia_ann=0";
                                                using (SqlCommand cmUpdLia = new SqlCommand(sUpdLia, cn, tr))
                                                {
                                                    cmUpdLia.Parameters.AddWithValue("@cos", cos);
                                                    cmUpdLia.Parameters.AddWithValue("@art", artCod);
                                                    cmUpdLia.ExecuteNonQuery();
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception exRow)
                                {
                                    string errItem = string.Format("Codice {0} ({1}): {2}", artCod, des, exRow.Message);
                                    lstErrors.Add(errItem);
                                    _clsFun.ErrorLog("btnAllinea_Click Item " + artCod, exRow.Message);
                                }
                            }

                            tr.Commit();
                            this.DialogResult = DialogResult.OK;
                        }
                        catch (Exception ex)
                        {
                            tr.Rollback();
                            throw ex;
                        }
                    }
                }

                string sMsg = string.Format(
                    "Allineamento da APShop a APOffice completato!\n\n" +
                    "• Articoli inseriti: {0:N0}\n" +
                    "• Articoli aggiornati: {1:N0}\n" +
                    "• Valori predefiniti applicati: art_tgr='PZ', art_umi='NR', art_pxc=1, art_eti='043'\n" +
                    "• Bilancia: art_bil/ean_bil=1 per barcode che iniziano per 2 o da APShop (_bil)\n" +
                    "• Barcode: ean_ecp=0, ean_dti/ean_dtm a data odierna, Stato Attivo ('A').", nIns, nUpd);

                if (lstErrors.Count > 0)
                {
                    sMsg += string.Format("\n\n⚠️ {0} articoli non sono stati allineati per errori di formato/dati (dettagli nei log). Primo errore:\n{1}",
                        lstErrors.Count, lstErrors[0]);
                }

                MessageBox.Show(sMsg, "ALLINEAMENTO COMPLETATO", MessageBoxButtons.OK, lstErrors.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);

                // Ri-esegue il confronto per aggiornare la griglia
                EseguiConfronto();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore durante l'allineamento dei dati:\n" + ex.Message,
                    "ERRORE ALLINEAMENTO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _clsFun.ErrorLog("btnAllinea_Click APShop", ex.Message);
            }
            finally
            {
                btnStop.Visible = false;
                btnAllinea.Enabled = true;
                btnAnalizza.Enabled = true;
                btnBrowseMdb.Enabled = true;
                progressBar1.Visible = false;
                this.Cursor = Cursors.Default;
                _bolCancel = false;
            }
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (progressBar1.Visible)
            {
                if (MessageBox.Show("Procedura di allineamento in corso.\nSei sicuro di voler interrompere l'operazione?",
                    "CONFERMA INTERRUZIONE", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    _bolCancel = true;
                }
            }
        }

        private string TruncateField(DataTable schemaTable, string colName, string val)
        {
            if (string.IsNullOrEmpty(val)) return "";
            int maxLen = -1;
            if (schemaTable != null && schemaTable.Columns.Contains(colName))
            {
                maxLen = schemaTable.Columns[colName].MaxLength;
            }

            // Limiti di sicurezza garantiti se MaxLength non è stato valorizzato da ADO.NET
            if (maxLen <= 0)
            {
                string cLow = colName.ToLower().Trim();
                if (cLow == "art_cod" || cLow == "ean_art" || cLow == "liv_art" || cLow == "lia_art") maxLen = 20;
                else if (cLow == "art_des") maxLen = 100;
                else if (cLow == "art_deb") maxLen = 20;
                else if (cLow == "art_rep" || cLow == "art_iva" || cLow == "art_umi" || cLow == "art_umc" || cLow == "art_eti" || cLow == "liv_lis") maxLen = 3;
                else if (cLow == "art_tgr") maxLen = 5;
                else if (cLow == "art_sta" || cLow == "liv_sta" || cLow == "lia_tip") maxLen = 1;
                else if (cLow == "ean_ean") maxLen = 20;
                else if (cLow == "lia_for") maxLen = 6;
            }

            if (maxLen > 0 && val.Length > maxLen)
            {
                return val.Substring(0, maxLen);
            }
            return val;
        }

        private void btnEsci_Click(object sender, EventArgs e)
        {
            if (progressBar1.Visible)
            {
                btnStop_Click(sender, e);
                return;
            }
            this.Close();
        }

        private void frmUtyAlignApShop_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (progressBar1.Visible)
            {
                if (MessageBox.Show("Procedura in corso. Sei sicuro di voler annullare e uscire?", "CONFERMA CHIUSURA", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    _bolCancel = true;
                }
                else
                {
                    e.Cancel = true;
                    return;
                }
            }

            try
            {
                dgvGrid.DataSource = null;
                _dtResult = null;
                OleDbConnection.ReleaseObjectPool();
            }
            catch { }
        }
    }
}
