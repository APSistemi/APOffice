using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;

namespace APOffice
{
    public partial class frmAnaArtNuovo : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsVariazioni _clsVar = new clsVariazioni();
        clsQuery _clsQry = new clsQuery();

        public string _strEan = "";
        public string _strRes = "";

        private string _strConSql = "";

        private string _strFilVoci = "C:\\ApProject\\ApOffice\\DataBase\\ArtInsVelox.ini";

        private System.Windows.Forms.Timer _tmrEanInput;
        private string _strLastEanWebSearched = "";
        private string _strLastWebImageUrl = "";

        public frmAnaArtNuovo()
        {
            InitializeComponent(); 
            ApplyHighContrastTheme();

            _tmrEanInput = new System.Windows.Forms.Timer();
            _tmrEanInput.Interval = 250;
            _tmrEanInput.Tick += (s, e) =>
            {
                _tmrEanInput.Stop();
                if (chkRicercaAuto != null && !chkRicercaAuto.Checked) return;
                string ean = txtEanEan.Text.Trim();
                if (ean.Length >= 6 && ean != _strLastEanWebSearched)
                {
                    _strLastEanWebSearched = ean;
                    CercaWeb(ean);
                }
            };
        }

        private void frmAnaArtNuovo_Load(object sender, EventArgs e)
        {
            ApplyHighContrastTheme();
            FillTabs();
            txtEanEan.Text = _strEan;
            txtEanEan.Select();

            UpdateWebStatusLabelPrompt();

            if (chkRicercaAuto != null && chkRicercaAuto.Checked && !string.IsNullOrEmpty(_strEan) && _strEan.Trim().Length >= 6)
            {
                _strLastEanWebSearched = _strEan.Trim();
                CercaWeb(_strEan);
            }
        }

        private void chkRicercaAuto_CheckedChanged(object sender, EventArgs e)
        {
            if (_tmrEanInput != null && chkRicercaAuto != null && !chkRicercaAuto.Checked)
            {
                _tmrEanInput.Stop();
            }
            UpdateWebStatusLabelPrompt();
            SalvaVoci();
        }

        private void UpdateWebStatusLabelPrompt()
        {
            if (lblWebStatus == null) return;
            if (chkRicercaAuto != null && chkRicercaAuto.Checked)
            {
                lblWebStatus.Text = "💡 Inserisci o scansiona il Barcode (ricerca automatica attiva)";
                lblWebStatus.ForeColor = Color.FromArgb(37, 99, 235);
            }
            else
            {
                lblWebStatus.Text = "💡 Inserisci il Barcode e premi Cerca Web (F2)";
                lblWebStatus.ForeColor = Color.FromArgb(71, 85, 105);
            }
        }

        private void frmAnaArtNuovo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
            else if (e.KeyCode == Keys.F2)
            {
                e.Handled = true;
                btnCercaWeb_Click(sender, e);
            }
            else if (e.KeyCode == Keys.F10)
            {
                e.Handled = true;
                btnSalvaNuovo_Click(sender, e);
            }
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private void FillTabs()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabIva";
            s = "SELECT * FROM TabIva ORDER BY tab_cod";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbArtIva.DataSource = t;
            cmbArtIva.DisplayMember = "tab_des";
            cmbArtIva.ValueMember = "tab_cod";
            cmbArtIva.SelectedValue = "";

            p = "TabUmi";
            s = "SELECT * FROM TabUmi ORDER BY tab_des";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  ";
            t.Rows.InsertAt(x, 0);
            cmbArtUmi.DataSource = t;
            cmbArtUmi.DisplayMember = "tab_des";
            cmbArtUmi.ValueMember = "tab_cod";
            cmbArtUmi.SelectedValue = "";

            p = "TabTipoGrammatura";
            s = "SELECT * FROM TabTipoGrammatura ORDER BY tab_des";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  ";
            t.Rows.InsertAt(x, 0);
            cmbArtTgr.DataSource = t;
            cmbArtTgr.DisplayMember = "tab_des";
            cmbArtTgr.ValueMember = "tab_cod";
            cmbArtTgr.SelectedValue = "";

            p = "TabReparti";
            s = "SELECT * FROM " + p + " WHERE tab_ann=0 ORDER BY tab_des";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbArtRep.DataSource = t;
            cmbArtRep.DisplayMember = "tab_des";
            cmbArtRep.ValueMember = "tab_cod";
            cmbArtRep.SelectedValue = "";

            p = "TabListini";
            s = "SELECT * FROM TabListini WHERE tab_ann=0 AND tab_tip<>'O'";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "";
            t.Rows.InsertAt(x, 0);
            cmbLivLis.DataSource = t;
            cmbLivLis.DisplayMember = "tab_des";
            cmbLivLis.ValueMember = "tab_cod";
            SelectDefaultListinoCassa();

            p = "AnaFornitori";
            s = "SELECT for_cod, for_des FROM AnaFornitori WHERE for_ann=0 ORDER BY for_des";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["for_cod"] = "";
            x["for_des"] = "";
            t.Rows.InsertAt(x, 0);
            cmbLiaFor.DataSource = t;
            cmbLiaFor.DisplayMember = "for_des";
            cmbLiaFor.ValueMember = "for_cod";
            cmbLiaFor.SelectedValue = "";

            if (File.Exists(_strFilVoci))
            {
                using (StreamReader sr = new StreamReader(_strFilVoci))
                {
                    string sRig = "";

                    while ((sRig = sr.ReadLine()) != null)
                    {
                        if (sRig.Length > 4)
                        {
                            string sVal = sRig.Substring(4).Trim();

                            if (sRig.StartsWith("MEM:T", StringComparison.OrdinalIgnoreCase))
                                chkMem.Checked = true;

                            if (sRig.StartsWith("AUT:", StringComparison.OrdinalIgnoreCase))
                            {
                                string sAut = sRig.Substring(4).Trim();
                                if (chkRicercaAuto != null)
                                    chkRicercaAuto.Checked = (sAut.Equals("True", StringComparison.OrdinalIgnoreCase) || sAut.Equals("1") || sAut.Equals("T", StringComparison.OrdinalIgnoreCase));
                            }

                            if (chkMem.Checked)
                            {
                                if (sRig.Substring(0, 3) == "IVA")
                                    cmbArtIva.SelectedValue = sVal;
                                else if (sRig.Substring(0, 3) == "UMI")
                                    cmbArtUmi.SelectedValue = sVal;
                                else if (sRig.Substring(0, 3) == "TGR")
                                    cmbArtTgr.SelectedValue = sVal;
                                else if (sRig.Substring(0, 3) == "REP")
                                    cmbArtRep.SelectedValue = sVal;
                                else if (sRig.Substring(0, 3) == "FOR")
                                    cmbLiaFor.SelectedValue = sVal;
                                else if (sRig.Substring(0, 3) == "LIV")
                                    cmbLivLis.SelectedValue = sVal;
                                else if (sRig.Substring(0, 3) == "RIC")
                                    txtRicarico.Text = sVal;
                            }
                        }
                    }
                }
            }

            if (cmbLivLis.SelectedValue == null || string.IsNullOrEmpty(cmbLivLis.SelectedValue.ToString()))
            {
                SelectDefaultListinoCassa();
            }
        }

        private void SelectDefaultListinoCassa()
        {
            if (cmbLivLis == null || cmbLivLis.DataSource == null) return;
            try
            {
                DataTable dt = cmbLivLis.DataSource as DataTable;
                if (dt != null)
                {
                    // 1. Priorità massima: cerca per descrizione "LISTINO CASSA" o contenente "CASSA"
                    foreach (DataRow row in dt.Rows)
                    {
                        string des = row["tab_des"] != null ? row["tab_des"].ToString().Trim() : "";
                        if (des.IndexOf("CASSA", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            cmbLivLis.SelectedValue = row[cmbLivLis.ValueMember];
                            return;
                        }
                    }

                    // 2. Cerca per tipo 'C' (Cassa)
                    if (dt.Columns.Contains("tab_tip"))
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            string tip = row["tab_tip"] != null ? row["tab_tip"].ToString().Trim() : "";
                            if (tip.Equals("C", StringComparison.OrdinalIgnoreCase))
                            {
                                cmbLivLis.SelectedValue = row[cmbLivLis.ValueMember];
                                return;
                            }
                        }
                    }

                    // 3. Fallback: primo listino valido
                    foreach (DataRow row in dt.Rows)
                    {
                        string cod = row[cmbLivLis.ValueMember] != null ? row[cmbLivLis.ValueMember].ToString().Trim() : "";
                        if (!string.IsNullOrEmpty(cod))
                        {
                            cmbLivLis.SelectedValue = row[cmbLivLis.ValueMember];
                            return;
                        }
                    }
                }
            }
            catch { }
        }

        #region Ricerca Web EAN & Autocompilazione
        private class WebProductInfo
        {
            public bool Found { get; set; }
            public string Name { get; set; }
            public string Brand { get; set; }
            public string Quantity { get; set; }
            public decimal Content { get; set; }
            public string Unit { get; set; }
            public string Grammatura { get; set; }
            public string IvaCode { get; set; }
            public decimal Price { get; set; }
            public string ImageUrl { get; set; }
            public string Category { get; set; }

            public WebProductInfo()
            {
                Found = false;
                Name = "";
                Brand = "";
                Quantity = "";
                Content = 1;
                Unit = "PZ";
                Grammatura = "PZ";
                IvaCode = "010";
                Price = 0;
                ImageUrl = "";
                Category = "";
            }
        }

        public bool CercaArticoloLocale(string eanInput, out string outCodArt, out string outDesArt)
        {
            outCodArt = "";
            outDesArt = "";
            if (string.IsNullOrEmpty(eanInput)) return false;
            string ean = eanInput.Trim();
            if (ean.Length < 3) return false;

            if (string.IsNullOrEmpty(_strConSql))
                _strConSql = _clsFun.ConSql("");

            try
            {
                List<string> variants = new List<string> { ean };
                if (ean.Length == 12) variants.Add("0" + ean);
                if (ean.StartsWith("0") && ean.Length > 1) variants.Add(ean.TrimStart('0'));
                if (ean.Length < 13 && _clsFun.Numerico(ean)) variants.Add(ean.PadLeft(13, '0'));

                string inClause = string.Join(",", variants.Distinct().Select(v => "'" + v.Replace("'", "''") + "'"));

                // 1. Cerca in AnaBarcode con anagrafica collegata
                string sql = @"
                    SELECT TOP 1 b.ean_art, a.art_des, a.art_sta 
                    FROM AnaBarcode b 
                    LEFT JOIN AnaArticoli a ON b.ean_art = a.art_cod 
                    WHERE b.ean_ean IN (" + inClause + @") 
                      AND (b.ean_ann = 0 OR b.ean_ann IS NULL)
                    ORDER BY CASE WHEN b.ean_ean = '" + ean.Replace("'", "''") + @"' THEN 0 ELSE 1 END,
                             CASE WHEN a.art_sta = 'A' THEN 0 ELSE 1 END";

                DataTable dt = _clsFun.FillTabSql("AnaBarcode", sql, false, _strConSql);
                if (dt != null && dt.Rows.Count > 0)
                {
                    outCodArt = dt.Rows[0]["ean_art"] != null ? dt.Rows[0]["ean_art"].ToString().Trim() : "";
                    outDesArt = dt.Rows[0]["art_des"] != null ? dt.Rows[0]["art_des"].ToString().Trim() : "";
                    if (!string.IsNullOrEmpty(outCodArt)) return true;
                }

                // 2. Cerca direttamente in AnaArticoli (se barcode coincide col codice articolo)
                sql = @"
                    SELECT TOP 1 art_cod, art_des 
                    FROM AnaArticoli 
                    WHERE art_cod IN (" + inClause + @")
                    ORDER BY CASE WHEN art_sta = 'A' THEN 0 ELSE 1 END";
                dt = _clsFun.FillTabSql("AnaArticoli", sql, false, _strConSql);
                if (dt != null && dt.Rows.Count > 0)
                {
                    outCodArt = dt.Rows[0]["art_cod"] != null ? dt.Rows[0]["art_cod"].ToString().Trim() : "";
                    outDesArt = dt.Rows[0]["art_des"] != null ? dt.Rows[0]["art_des"].ToString().Trim() : "";
                    if (!string.IsNullOrEmpty(outCodArt)) return true;
                }

                // 3. Cerca in GesLisAcquisto (lia_arf)
                sql = @"
                    SELECT TOP 1 l.lia_art, a.art_des 
                    FROM GesLisAcquisto l 
                    LEFT JOIN AnaArticoli a ON l.lia_art = a.art_cod 
                    WHERE l.lia_arf IN (" + inClause + @") 
                      AND (l.lia_ann = 0 OR l.lia_ann IS NULL)
                    ORDER BY CASE WHEN a.art_sta = 'A' THEN 0 ELSE 1 END";
                dt = _clsFun.FillTabSql("GesLisAcquisto", sql, false, _strConSql);
                if (dt != null && dt.Rows.Count > 0)
                {
                    outCodArt = dt.Rows[0]["lia_art"] != null ? dt.Rows[0]["lia_art"].ToString().Trim() : "";
                    outDesArt = dt.Rows[0]["art_des"] != null ? dt.Rows[0]["art_des"].ToString().Trim() : "";
                    if (!string.IsNullOrEmpty(outCodArt)) return true;
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArtNuovo.CercaArticoloLocale", ex.Message);
            }
            return false;
        }

        private bool VerificaBarcodeEsistente(string ean, bool isAutoSearch)
        {
            string codArt = "";
            string desArt = "";
            if (CercaArticoloLocale(ean, out codArt, out desArt))
            {
                DialogResult dr = MessageBox.Show(this, 
                    "Il codice a barre " + ean + " è già presente nel database per l'articolo:\n\n" +
                    "Codice: " + codArt + "\n" +
                    "Descrizione: " + desArt + "\n\n" +
                    "Vuoi aprire la scheda anagrafica dell'articolo esistente?", 
                    "Articolo già registrato", 
                    MessageBoxButtons.YesNo, 
                    MessageBoxIcon.Information);

                if (dr == DialogResult.Yes)
                {
                    _strRes = "art_cod:" + codArt + ";";
                    Esci();
                    return true;
                }
                else
                {
                    lblWebStatus.Text = "ℹ️ Barcode già associato ad art. " + codArt + (desArt != "" ? " (" + desArt + ")" : "");
                    lblWebStatus.ForeColor = Color.FromArgb(234, 88, 12);
                    return true;
                }
            }
            return false;
        }

        private void btnCercaWeb_Click(object sender, EventArgs e)
        {
            string ean = txtEanEan.Text.Trim();
            _strLastEanWebSearched = ean;
            CercaWeb(ean);
        }

        private void CercaWeb(string strEanInput)
        {
            if (string.IsNullOrEmpty(strEanInput) || strEanInput.Trim().Length < 4)
            {
                lblWebStatus.Text = "⚠️ Inserisci un codice a barre EAN valido";
                lblWebStatus.ForeColor = Color.FromArgb(220, 38, 38);
                return;
            }

            string strEanClean = strEanInput.Trim();

            // Controlla prima se il barcode è già registrato nel database aziendale
            if (VerificaBarcodeEsistente(strEanClean, true))
            {
                return;
            }

            // Se è un barcode bilancia a peso o prezzo variabile (prefisso 2 standard GS1 / cassa interna)
            if (strEanClean.StartsWith("2") && (strEanClean.Length == 12 || strEanClean.Length == 13))
            {
                lblWebStatus.Text = "⚖️ Barcode bilancia / peso variabile rilevato (codice interno)";
                lblWebStatus.ForeColor = Color.FromArgb(16, 185, 129);
                lblPicCaption.Text = "Articolo Bilancia";

                if (!chkMem.Checked || cmbArtUmi.SelectedValue == null || string.IsNullOrEmpty(cmbArtUmi.SelectedValue.ToString()))
                {
                    SelectComboValueOrText(cmbArtUmi, "KG", new string[] { "KG", "CHILO", "PESO" });
                }
                if (!chkMem.Checked || cmbArtTgr.SelectedValue == null || string.IsNullOrEmpty(cmbArtTgr.SelectedValue.ToString()))
                {
                    SelectComboValueOrText(cmbArtTgr, "GR", new string[] { "GR", "GRAMMI" });
                }
                if (!chkMem.Checked || cmbArtRep.SelectedValue == null || string.IsNullOrEmpty(cmbArtRep.SelectedValue.ToString()))
                {
                    SelectComboValueOrText(cmbArtRep, "003", new string[] { "gastronomia", "salumi", "formaggi", "freschi", "macelleria", "ortofrutta" });
                }
                return;
            }

            lblWebStatus.Text = "⏳ Ricerca prodotto su Internet in corso...";
            lblWebStatus.ForeColor = Color.FromArgb(37, 99, 235);
            lblPicCaption.Text = "Caricamento...";
            btnCercaWeb.Enabled = false;

            ThreadPool.QueueUserWorkItem(state =>
            {
                WebProductInfo info = null;
                try
                {
                    info = ExecuteWebLookup(strEanClean);
                }
                catch (Exception ex)
                {
                    _clsFun.ErrorLog("frmAnaArtNuovo.CercaWeb.Async", ex.Message);
                }

                if (this.IsDisposed || !this.IsHandleCreated) return;

                this.Invoke(new MethodInvoker(() =>
                {
                    try
                    {
                        btnCercaWeb.Enabled = true;

                        // Previene race condition se l'utente nel frattempo ha modificato o azzerato il barcode
                        if (!string.Equals(txtEanEan.Text.Trim(), strEanClean, StringComparison.OrdinalIgnoreCase))
                            return;

                        if (info != null && info.Found)
                        {
                            lblWebStatus.Text = "✅ Prodotto trovato: " + info.Name;
                        lblWebStatus.ForeColor = Color.FromArgb(22, 163, 74);

                        if (!string.IsNullOrEmpty(info.Name))
                        {
                            txtArtDes.Text = info.Name;
                            txtArtDes.BackColor = Color.FromArgb(240, 253, 244);
                        }

                        if (info.Content > 0)
                        {
                            txtArtNet.Text = info.Content.ToString(System.Globalization.CultureInfo.InvariantCulture);
                            txtArtNet.BackColor = Color.FromArgb(240, 253, 244);
                        }

                        if (!chkMem.Checked || cmbArtIva.SelectedValue == null || string.IsNullOrEmpty(cmbArtIva.SelectedValue.ToString()))
                        {
                            SelectComboValueOrText(cmbArtIva, info.IvaCode, new string[] { "10", "4", "22" });
                        }

                        if (!chkMem.Checked || cmbArtUmi.SelectedValue == null || string.IsNullOrEmpty(cmbArtUmi.SelectedValue.ToString()))
                        {
                            SelectComboValueOrText(cmbArtUmi, info.Unit, new string[] { info.Unit });
                        }

                        if (!chkMem.Checked || cmbArtTgr.SelectedValue == null || string.IsNullOrEmpty(cmbArtTgr.SelectedValue.ToString()))
                        {
                            SelectComboValueOrText(cmbArtTgr, info.Grammatura, new string[] { info.Grammatura });
                        }

                        if (!chkMem.Checked || cmbArtRep.SelectedValue == null || string.IsNullOrEmpty(cmbArtRep.SelectedValue.ToString()))
                        {
                            if (!string.IsNullOrEmpty(info.Category) || !string.IsNullOrEmpty(info.Name))
                            {
                                SelectRepartoByCategory(info.Category, info.Name);
                            }
                        }

                        // Calcolo prezzo: se è presente costo e % ricarico calcola, altrimenti se web ha prezzo propone quello
                        decimal dCurCost = 0;
                        decimal.TryParse((txtLiaPrc.Text ?? "").Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dCurCost);
                        decimal dCurRic = 0;
                        decimal.TryParse((txtRicarico.Text ?? "").Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dCurRic);

                        if (dCurCost > 0 && dCurRic > 0)
                        {
                            CalcolaPrezzoDaRicarico();
                        }
                        else if (info.Price > 0)
                        {
                            txtLivPrv.Text = info.Price.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);
                            txtLivPrv.BackColor = Color.FromArgb(240, 253, 244);
                        }

                        if (!string.IsNullOrEmpty(info.ImageUrl))
                        {
                            LoadWebImageToPictureBox(info.ImageUrl);
                        }
                        else
                        {
                            _strLastWebImageUrl = "";
                            if (picWebProduct.Image != null)
                            {
                                try { picWebProduct.Image.Dispose(); } catch { }
                                picWebProduct.Image = null;
                            }
                            lblPicCaption.Text = "Nessuna Foto";
                        }
                    }
                    else
                    {
                        _strLastWebImageUrl = "";
                        if (picWebProduct.Image != null)
                        {
                            try { picWebProduct.Image.Dispose(); } catch { }
                            picWebProduct.Image = null;
                        }
                        lblWebStatus.Text = "⚠️ Nessun prodotto trovato online per " + strEanClean;
                        lblWebStatus.ForeColor = Color.FromArgb(217, 119, 6);
                        lblPicCaption.Text = "Nessuna Foto";
                    }
                }
                catch (Exception ex)
                {
                    _clsFun.ErrorLog("frmAnaArtNuovo.CercaWeb.Ui", ex.Message);
                }
                finally
                {
                    if (btnCercaWeb != null) btnCercaWeb.Enabled = true;
                }
                }));
            });
        }

        private void LoadWebImageToPictureBox(string url)
        {
            if (string.IsNullOrEmpty(url)) return;
            _strLastWebImageUrl = url;

            ThreadPool.QueueUserWorkItem(state =>
            {
                try
                {
                    System.Net.ServicePointManager.SecurityProtocol = (System.Net.SecurityProtocolType)12288 | (System.Net.SecurityProtocolType)3072 | (System.Net.SecurityProtocolType)768 | System.Net.SecurityProtocolType.Tls;
                    System.Net.ServicePointManager.ServerCertificateValidationCallback = (senderCert, cert, chain, sslPolicyErrors) => true;

                    System.Net.HttpWebRequest req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
                    req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36";
                    req.Headers.Add("Accept-Language", "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7");
                    req.AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate;
                    req.Timeout = 8000;
                    req.ReadWriteTimeout = 8000;

                    using (System.Net.HttpWebResponse resp = (System.Net.HttpWebResponse)req.GetResponse())
                    using (System.IO.Stream stream = resp.GetResponseStream())
                    using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
                    {
                        stream.CopyTo(ms);
                        ms.Position = 0;
                        using (Image downloadedImg = Image.FromStream(ms))
                        {
                            Bitmap bmp = new Bitmap(downloadedImg);

                            if (this.IsDisposed || !this.IsHandleCreated)
                            {
                                bmp.Dispose();
                                return;
                            }

                            this.Invoke(new MethodInvoker(() =>
                            {
                                if (_strLastWebImageUrl != url)
                                {
                                    bmp.Dispose();
                                    return;
                                }

                                if (picWebProduct.Image != null)
                                {
                                    try { picWebProduct.Image.Dispose(); } catch { }
                                }
                                picWebProduct.Image = bmp;
                                lblPicCaption.Text = "Foto da Web";
                            }));
                        }
                    }
                }
                catch
                {
                    if (this.IsDisposed || !this.IsHandleCreated) return;
                    this.Invoke(new MethodInvoker(() =>
                    {
                        lblPicCaption.Text = "Foto non disp.";
                    }));
                }
            });
        }

        private void ApplyHighContrastTheme()
        {
            this.BackColor = Color.FromArgb(248, 250, 252);

            if (menuStrip1 != null)
            {
                menuStrip1.BackColor = Color.FromArgb(15, 23, 42);
                if (esciToolStripMenuItem != null)
                {
                    esciToolStripMenuItem.ForeColor = Color.White;
                    esciToolStripMenuItem.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
                }
            }

            if (panelHeader != null)
            {
                panelHeader.BackColor = Color.FromArgb(15, 23, 42);
                if (lblHeaderTitle != null)
                {
                    lblHeaderTitle.ForeColor = Color.FromArgb(250, 204, 21); // Vivid Contrast Yellow
                    lblHeaderTitle.Font = new Font("Segoe UI", 14F, FontStyle.Bold);
                    lblHeaderTitle.Text = "🛍️ INSERIMENTO ARTICOLO VELOCE";
                }
                if (lblHeaderSub != null)
                {
                    lblHeaderSub.ForeColor = Color.White;
                    lblHeaderSub.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
                    lblHeaderSub.Text = "Ricerca automatica Web EAN & compilazione veloce informazioni prodotto da Internet";
                }
            }

            foreach (Control c in this.Controls)
            {
                if (c is Label && c != lblHeaderTitle && c != lblHeaderSub && c != lblWebStatus && c != lblPicCaption)
                {
                    c.ForeColor = Color.FromArgb(15, 23, 42);
                    c.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
                }
            }

            if (panelFooter != null)
            {
                panelFooter.BackColor = Color.FromArgb(15, 23, 42);
            }

            if (btnSalvaNuovo != null)
            {
                btnSalvaNuovo.FlatStyle = FlatStyle.Flat;
                btnSalvaNuovo.FlatAppearance.BorderSize = 0;
                btnSalvaNuovo.BackColor = Color.FromArgb(16, 185, 129); // Emerald Green
                btnSalvaNuovo.ForeColor = Color.White;
                btnSalvaNuovo.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
                btnSalvaNuovo.Text = "✔ Salva e Continua (F10)";
            }

            if (btnOk != null)
            {
                btnOk.FlatStyle = FlatStyle.Flat;
                btnOk.FlatAppearance.BorderSize = 0;
                btnOk.BackColor = Color.FromArgb(37, 99, 235); // Royal Blue
                btnOk.ForeColor = Color.White;
                btnOk.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
                btnOk.Text = "💾 Salva e Chiudi";
            }

            if (btnEsci != null)
            {
                btnEsci.FlatStyle = FlatStyle.Flat;
                btnEsci.FlatAppearance.BorderSize = 0;
                btnEsci.BackColor = Color.FromArgb(220, 38, 38); // Crimson Red
                btnEsci.ForeColor = Color.White;
                btnEsci.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
                btnEsci.Text = "❌ Annulla (Esc)";
            }

            if (btnCercaWeb != null)
            {
                btnCercaWeb.FlatStyle = FlatStyle.Flat;
                btnCercaWeb.FlatAppearance.BorderSize = 0;
                btnCercaWeb.BackColor = Color.FromArgb(2, 132, 199); // Cyan Blue
                btnCercaWeb.ForeColor = Color.White;
                btnCercaWeb.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            }

            try
            {
                ToolTip toolTip = new ToolTip();
                toolTip.AutoPopDelay = 5000;
                toolTip.InitialDelay = 150;
                toolTip.ReshowDelay = 50;
                toolTip.ShowAlways = true;

                if (btnSalvaNuovo != null) toolTip.SetToolTip(btnSalvaNuovo, "Salva l'articolo corrente e passa al successivo (F10)");
                if (btnOk != null) toolTip.SetToolTip(btnOk, "Salva l'articolo corrente e torna all'Anagrafica Articolo");
                if (btnEsci != null) toolTip.SetToolTip(btnEsci, "Annulla l'inserimento ed esci immediatamente senza salvare (ESC)");
                if (btnCercaWeb != null) toolTip.SetToolTip(btnCercaWeb, "Cerca dati prodotto e prezzo consigliato su Internet (F2)");
                if (chkRicercaAuto != null) toolTip.SetToolTip(chkRicercaAuto, "Se attivo cerca automaticamente le informazioni del prodotto online durante la digitazione o scansione del barcode. Se disattivato, la ricerca web si avvia solo premendo Cerca Web (F2).");
                if (txtArtArf != null) toolTip.SetToolTip(txtArtArf, "Codice articolo del fornitore (se vuoto viene usato il codice articolo interno)");
                if (txtRicarico != null) toolTip.SetToolTip(txtRicarico, "Percentuale di ricarico sul costo: calcola automaticamente il prezzo ivato");
                if (txtLiaPrc != null) toolTip.SetToolTip(txtLiaPrc, "Prezzo di costo del fornitore");
                if (txtLivPrv != null) toolTip.SetToolTip(txtLivPrv, "Prezzo di vendita al pubblico (IVA inclusa)");
            }
            catch { }
        }

        private WebProductInfo FinalizeProductInfo(WebProductInfo info, string ean)
        {
            if (info == null || !info.Found || string.IsNullOrEmpty(info.Name)) return null;

            if (string.IsNullOrEmpty(info.ImageUrl))
            {
                info.ImageUrl = SearchProductImageFallback(ean, info.Name);
            }
            return info;
        }

        private string SearchProductImageFallback(string ean, string productName)
        {
            try
            {
                // 1. Yahoo Images search by EAN
                if (!string.IsNullOrEmpty(ean))
                {
                    string urlEan = "https://it.images.search.yahoo.com/search/images?p=" + ean;
                    string htmlEan = FetchUrlContentWithUserAgent(urlEan);
                    if (!string.IsNullOrEmpty(htmlEan))
                    {
                        string imgUrl = ExtractImageUrlFromHtml(htmlEan);
                        if (!string.IsNullOrEmpty(imgUrl)) return imgUrl;
                    }
                }

                // 2. Yahoo Images search by product Name
                if (!string.IsNullOrEmpty(productName))
                {
                    string cleanTitle = Regex.Replace(productName, @"\s+", " ").Trim();
                    string urlTitle = "https://it.images.search.yahoo.com/search/images?p=" + System.Uri.EscapeDataString(cleanTitle);
                    string htmlTitle = FetchUrlContentWithUserAgent(urlTitle);
                    if (!string.IsNullOrEmpty(htmlTitle))
                    {
                        string imgUrl = ExtractImageUrlFromHtml(htmlTitle);
                        if (!string.IsNullOrEmpty(imgUrl)) return imgUrl;
                    }
                }

                // 3. Bing Images search by EAN / Name
                string q = !string.IsNullOrEmpty(ean) ? ean : productName;
                if (!string.IsNullOrEmpty(q))
                {
                    string urlBing = "https://www.bing.com/images/search?q=" + System.Uri.EscapeDataString(q);
                    string htmlBing = FetchUrlContentWithUserAgent(urlBing);
                    if (!string.IsNullOrEmpty(htmlBing))
                    {
                        Match mBing = Regex.Match(htmlBing, @"murl&quot;:&quot;(https?://[^&""]+)&quot;", RegexOptions.IgnoreCase);
                        if (mBing.Success) return System.Net.WebUtility.HtmlDecode(mBing.Groups[1].Value);

                        Match mBingSrc = Regex.Match(htmlBing, @"src=""(https://tse\d\.mm\.bing\.net/th\?[^""]+)""", RegexOptions.IgnoreCase);
                        if (mBingSrc.Success) return System.Net.WebUtility.HtmlDecode(mBingSrc.Groups[1].Value);
                    }
                }
            }
            catch { }
            return "";
        }

        private string ExtractImageUrlFromHtml(string html)
        {
            try
            {
                // Match Yahoo Images data-src or src or imgurl
                Match mImgUrl = Regex.Match(html, @"imgurl=(https?%3A%2F%2F[^&]+)", RegexOptions.IgnoreCase);
                if (mImgUrl.Success)
                {
                    string decoded = System.Uri.UnescapeDataString(mImgUrl.Groups[1].Value);
                    if (decoded.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                        return decoded;
                }

                Match mData = Regex.Match(html, @"data-src=""(https://[^""]+\.(?:jpg|jpeg|png|webp)[^""]*)""", RegexOptions.IgnoreCase);
                if (mData.Success)
                    return System.Net.WebUtility.HtmlDecode(mData.Groups[1].Value);

                Match mSrc = Regex.Match(html, @"<img[^>]+src=""(https://tse\d\.mm\.bing\.net/th\?[^""]+)""", RegexOptions.IgnoreCase);
                if (mSrc.Success)
                    return System.Net.WebUtility.HtmlDecode(mSrc.Groups[1].Value);
            }
            catch { }
            return "";
        }

        private WebProductInfo ExecuteWebLookup(string ean)
        {
            WebProductInfo info = null;

            if (string.IsNullOrEmpty(ean) || ean.Trim().Length < 4)
                return null;

            ean = ean.Trim();

            try
            {
                // Abilita TLS 1.3, TLS 1.2, TLS 1.1 e ignora errori certificati intermedi
                System.Net.ServicePointManager.SecurityProtocol = (System.Net.SecurityProtocolType)12288 | (System.Net.SecurityProtocolType)3072 | (System.Net.SecurityProtocolType)768 | System.Net.SecurityProtocolType.Tls;
                System.Net.ServicePointManager.ServerCertificateValidationCallback = (senderCert, cert, chain, sslPolicyErrors) => true;

                // 1. Query Open Food Facts v2 (API rapida con selezione mirata dei campi)
                string fields = "product_name,product_name_it,generic_name,generic_name_it,brands,quantity,product_quantity,product_quantity_unit,net_weight_value,net_weight_unit,serving_size,image_front_url,image_url,categories,status";
                string urlOffV2 = "https://world.openfoodfacts.org/api/v2/product/" + ean + ".json?fields=" + fields;
                string json = FetchUrlContent(urlOffV2);
                info = ParseOpenFoodFactsJson(json, ean);
                if (info != null && info.Found && !string.IsNullOrEmpty(info.Name) && IsValidProductTitle(info.Name, ean))
                    return FinalizeProductInfo(info, ean);

                // Variante EAN normalizzata (UPC 12 cifre <-> EAN 13 cifre con zero)
                string altEan = (ean.Length == 13 && ean.StartsWith("0")) ? ean.Substring(1) : (ean.Length == 12 ? "0" + ean : "");
                if (!string.IsNullOrEmpty(altEan))
                {
                    string urlOffAlt = "https://world.openfoodfacts.org/api/v2/product/" + altEan + ".json?fields=" + fields;
                    json = FetchUrlContent(urlOffAlt);
                    info = ParseOpenFoodFactsJson(json, ean);
                    if (info != null && info.Found && !string.IsNullOrEmpty(info.Name) && IsValidProductTitle(info.Name, ean))
                        return FinalizeProductInfo(info, ean);
                }

                // 2. Query DuckDuckGo HTML (Motore di ricerca retail pulito senza captcha / consent)
                info = SearchDuckDuckGoWeb(ean);
                if (info != null && info.Found && !string.IsNullOrEmpty(info.Name) && IsValidProductTitle(info.Name, ean))
                    return FinalizeProductInfo(info, ean);

                // 3. Query Open Food Facts IT v0 fallback
                string urlItV0 = "https://it.openfoodfacts.org/api/v0/product/" + ean + ".json";
                json = FetchUrlContent(urlItV0);
                info = ParseOpenFoodFactsJson(json, ean);
                if (info != null && info.Found && !string.IsNullOrEmpty(info.Name) && IsValidProductTitle(info.Name, ean))
                    return FinalizeProductInfo(info, ean);

                // 4. Query Yahoo Search IT (Motore primario non bloccato ad altissima precisione EAN retail italiano)
                info = SearchYahooWeb(ean);
                if (info != null && info.Found && !string.IsNullOrEmpty(info.Name) && IsValidProductTitle(info.Name, ean))
                    return FinalizeProductInfo(info, ean);

                // 5. Query Brave Search Engine (Motore secondario indipendente non bloccato)
                info = SearchBraveWeb(ean);
                if (info != null && info.Found && !string.IsNullOrEmpty(info.Name) && IsValidProductTitle(info.Name, ean))
                    return FinalizeProductInfo(info, ean);

                // 6. Query Open Beauty Facts IT & World (Cosmetici, Igiene persona, Saponi, Shampoo, Dopocera)
                string urlBeautyIt = "https://it.openbeautyfacts.org/api/v0/product/" + ean + ".json";
                json = FetchUrlContent(urlBeautyIt);
                info = ParseOpenFoodFactsJson(json, ean);
                if (info != null && info.Found && !string.IsNullOrEmpty(info.Name) && IsValidProductTitle(info.Name, ean))
                    return FinalizeProductInfo(info, ean);

                string urlBeautyWorld = "https://world.openbeautyfacts.org/api/v0/product/" + ean + ".json";
                json = FetchUrlContent(urlBeautyWorld);
                info = ParseOpenFoodFactsJson(json, ean);
                if (info != null && info.Found && !string.IsNullOrEmpty(info.Name) && IsValidProductTitle(info.Name, ean))
                    return FinalizeProductInfo(info, ean);

                // 7. Query Open Products Facts (Casalinghi, Non-Food, Cancelleria)
                string urlProdWorld = "https://world.openproductsfacts.org/api/v0/product/" + ean + ".json";
                json = FetchUrlContent(urlProdWorld);
                info = ParseOpenFoodFactsJson(json, ean);
                if (info != null && info.Found && !string.IsNullOrEmpty(info.Name) && IsValidProductTitle(info.Name, ean))
                    return FinalizeProductInfo(info, ean);

                // 8. Query Open Pet Food Facts (Cibo per Animali)
                string urlPetWorld = "https://world.openpetfoodfacts.org/api/v0/product/" + ean + ".json";
                json = FetchUrlContent(urlPetWorld);
                info = ParseOpenFoodFactsJson(json, ean);
                if (info != null && info.Found && !string.IsNullOrEmpty(info.Name) && IsValidProductTitle(info.Name, ean))
                    return FinalizeProductInfo(info, ean);

                // 9. Query UPCItemDB API (Database mondiale barcode EAN/UPC)
                string urlUpc = "https://api.upcitemdb.com/prod/trial/lookup?upc=" + ean;
                json = FetchUrlContent(urlUpc);
                info = ParseUpcItemDbJson(json, ean);
                if (info != null && info.Found && !string.IsNullOrEmpty(info.Name) && IsValidProductTitle(info.Name, ean))
                    return FinalizeProductInfo(info, ean);

                // 10. Query Yahoo Supermercati & Drogherie Italiane
                info = SearchYahooSupermarketsWeb(ean);
                if (info != null && info.Found && !string.IsNullOrEmpty(info.Name) && IsValidProductTitle(info.Name, ean))
                    return FinalizeProductInfo(info, ean);

                // 11. Fallback Motori Ricerca Ausiliari
                info = SearchBingWeb(ean);
                if (info != null && info.Found && !string.IsNullOrEmpty(info.Name) && IsValidProductTitle(info.Name, ean))
                    return FinalizeProductInfo(info, ean);

                info = SearchGoogleWeb(ean);
                if (info != null && info.Found && !string.IsNullOrEmpty(info.Name) && IsValidProductTitle(info.Name, ean))
                    return FinalizeProductInfo(info, ean);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArtNuovo.ExecuteWebLookup", ex.Message);
            }

            return null;
        }

        private WebProductInfo SearchDuckDuckGoWeb(string ean)
        {
            WebProductInfo info = new WebProductInfo();
            try
            {
                string url = "https://html.duckduckgo.com/html/?q=" + ean;
                string html = FetchUrlContentWithUserAgent(url);
                if (string.IsNullOrEmpty(html)) return info;

                MatchCollection matches = Regex.Matches(html, @"<a class=""result__a""[^>]*>([\s\S]*?)</a>", RegexOptions.IgnoreCase);
                foreach (Match m in matches)
                {
                    string raw = Regex.Replace(m.Groups[1].Value, @"<[^>]+>", "").Trim();
                    raw = System.Net.WebUtility.HtmlDecode(raw);
                    string clean = CleanWebProductTitle(raw, ean);
                    if (IsValidProductTitle(clean, ean))
                    {
                        info.Found = true;
                        ParseWeightAndGrammatura(clean, "", 0, "", 0, "", "", info);
                        info.Name = FormatFinalProductTitle(clean, info, ean);
                        info.IvaCode = DeduceIvaCode(info.Name, raw);
                        return info;
                    }
                }

                MatchCollection snippetMatches = Regex.Matches(html, @"<a class=""result__snippet""[^>]*>([\s\S]*?)</a>", RegexOptions.IgnoreCase);
                foreach (Match m in snippetMatches)
                {
                    string raw = Regex.Replace(m.Groups[1].Value, @"<[^>]+>", "").Trim();
                    raw = System.Net.WebUtility.HtmlDecode(raw);
                    string clean = CleanWebProductTitle(raw, ean);
                    if (IsValidProductTitle(clean, ean))
                    {
                        info.Found = true;
                        ParseWeightAndGrammatura(clean, "", 0, "", 0, "", "", info);
                        info.Name = FormatFinalProductTitle(clean, info, ean);
                        info.IvaCode = DeduceIvaCode(info.Name, raw);
                        return info;
                    }
                }
            }
            catch { }
            return info;
        }

        private WebProductInfo SearchYahooWeb(string ean)
        {
            WebProductInfo info = new WebProductInfo();
            try
            {
                string url = "https://it.search.yahoo.com/search?p=" + ean;
                string html = FetchUrlContentWithUserAgent(url);
                if (string.IsNullOrEmpty(html)) return info;

                // 1. Estrazione ad alta priorità da aria-label sui link di risultato
                MatchCollection ariaMatches = Regex.Matches(html, @"<a[^>]*aria-label=""([^""]+)""[^>]*>", RegexOptions.IgnoreCase);
                foreach (Match m in ariaMatches)
                {
                    string raw = System.Net.WebUtility.HtmlDecode(m.Groups[1].Value).Trim();
                    string clean = CleanWebProductTitle(raw, ean);
                    if (IsValidProductTitle(clean, ean))
                    {
                        info.Found = true;
                        ParseWeightAndGrammatura(clean, "", 0, "", 0, "", "", info);
                        info.Name = FormatFinalProductTitle(clean, info, ean);
                        info.IvaCode = DeduceIvaCode(info.Name, raw);
                        return info;
                    }
                }

                // 2. Estrazione da tag <h3>
                MatchCollection h3Matches = Regex.Matches(html, @"<h3[^>]*>([\s\S]*?)</h3>", RegexOptions.IgnoreCase);
                foreach (Match m in h3Matches)
                {
                    string raw = Regex.Replace(m.Groups[1].Value, @"<[^>]+>", "").Trim();
                    raw = System.Net.WebUtility.HtmlDecode(raw);
                    // Rimuovi eventuale breadcrumb iniziale
                    raw = Regex.Replace(raw, @"^.*?(›|»|»|\.\.\.)\s*", "");
                    string clean = CleanWebProductTitle(raw, ean);
                    if (IsValidProductTitle(clean, ean))
                    {
                        info.Found = true;
                        ParseWeightAndGrammatura(clean, "", 0, "", 0, "", "", info);
                        info.Name = FormatFinalProductTitle(clean, info, ean);
                        info.IvaCode = DeduceIvaCode(info.Name, raw);
                        return info;
                    }
                }
            }
            catch { }
            return info;
        }

        private WebProductInfo SearchYahooSupermarketsWeb(string ean)
        {
            WebProductInfo info = new WebProductInfo();
            try
            {
                string url = "https://it.search.yahoo.com/search?p=site:dm-drogeriemarkt.it+OR+site:carrefour.it+OR+site:conad.it+OR+site:esselunga.it+OR+site:cliviaprofumi.eu+OR+site:cicalia.com+" + ean;
                string html = FetchUrlContentWithUserAgent(url);
                if (string.IsNullOrEmpty(html)) return info;

                MatchCollection ariaMatches = Regex.Matches(html, @"<a[^>]*aria-label=""([^""]+)""[^>]*>", RegexOptions.IgnoreCase);
                foreach (Match m in ariaMatches)
                {
                    string raw = System.Net.WebUtility.HtmlDecode(m.Groups[1].Value).Trim();
                    string clean = CleanWebProductTitle(raw, ean);
                    if (IsValidProductTitle(clean, ean))
                    {
                        info.Found = true;
                        ParseWeightAndGrammatura(clean, "", 0, "", 0, "", "", info);
                        info.Name = FormatFinalProductTitle(clean, info, ean);
                        info.IvaCode = DeduceIvaCode(info.Name, raw);
                        return info;
                    }
                }
            }
            catch { }
            return info;
        }

        private WebProductInfo SearchBraveWeb(string ean)
        {
            WebProductInfo info = new WebProductInfo();
            try
            {
                string url = "https://search.brave.com/search?q=" + ean;
                string html = FetchUrlContentWithUserAgent(url);
                if (string.IsNullOrEmpty(html)) return info;

                MatchCollection titleMatches = Regex.Matches(html, @"<div class=""title[^""]*""[^>]*>(.*?)</div>", RegexOptions.IgnoreCase);
                if (titleMatches.Count == 0)
                {
                    titleMatches = Regex.Matches(html, @"<span class=""snippet-title""[^>]*>(.*?)</span>", RegexOptions.IgnoreCase);
                }

                foreach (Match m in titleMatches)
                {
                    string raw = Regex.Replace(m.Groups[1].Value, @"<[^>]+>", "").Trim();
                    raw = System.Net.WebUtility.HtmlDecode(raw);
                    string clean = CleanWebProductTitle(raw, ean);
                    if (IsValidProductTitle(clean, ean))
                    {
                        info.Found = true;
                        ParseWeightAndGrammatura(clean, "", 0, "", 0, "", "", info);
                        info.Name = FormatFinalProductTitle(clean, info, ean);
                        info.IvaCode = DeduceIvaCode(info.Name, raw);
                        return info;
                    }
                }
            }
            catch { }
            return info;
        }

        private WebProductInfo SearchGoogleWeb(string ean)
        {
            WebProductInfo info = new WebProductInfo();
            try
            {
                string url = "https://www.google.it/search?q=" + ean + "+prodotto+prezzo+spesa&hl=it&gbv=1";
                string html = FetchUrlContentWithUserAgent(url);

                if (!string.IsNullOrEmpty(html) && !html.Contains("302 Moved") && !html.Contains("Consent"))
                {
                    MatchCollection matches = Regex.Matches(html, @"<h3[^>]*>(.*?)</h3>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    foreach (Match m in matches)
                    {
                        string rawText = Regex.Replace(m.Groups[1].Value, @"<[^>]+>", "").Trim();
                        rawText = System.Net.WebUtility.HtmlDecode(rawText);
                        string cleanName = CleanWebProductTitle(rawText, ean);

                        if (IsValidProductTitle(cleanName, ean))
                        {
                            info.Found = true;
                            ParseWeightAndGrammatura(cleanName, "", 0, "", 0, "", "", info);
                            info.Name = FormatFinalProductTitle(cleanName, info, ean);
                            info.IvaCode = DeduceIvaCode(info.Name, rawText);
                            return info;
                        }
                    }
                }
            }
            catch { }
            return info;
        }

        private WebProductInfo SearchBingWeb(string ean)
        {
            WebProductInfo info = new WebProductInfo();
            try
            {
                string url = "https://www.bing.com/search?q=" + ean + "+prodotto+spesa";
                string html = FetchUrlContentWithUserAgent(url);
                if (!string.IsNullOrEmpty(html))
                {
                    MatchCollection matches = Regex.Matches(html, @"<h2[^>]*><a[^>]*>(.*?)</a></h2>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    foreach (Match m in matches)
                    {
                        string rawText = Regex.Replace(m.Groups[1].Value, @"<[^>]+>", "").Trim();
                        rawText = System.Net.WebUtility.HtmlDecode(rawText);
                        string cleanName = CleanWebProductTitle(rawText, ean);

                        if (IsValidProductTitle(cleanName, ean))
                        {
                            info.Found = true;
                            ParseWeightAndGrammatura(cleanName, "", 0, "", 0, "", "", info);
                            info.Name = FormatFinalProductTitle(cleanName, info, ean);
                            info.IvaCode = DeduceIvaCode(info.Name, rawText);
                            return info;
                        }
                    }
                }
            }
            catch { }
            return info;
        }

        private bool IsValidProductTitle(string title, string ean)
        {
            if (string.IsNullOrEmpty(title)) return false;
            string t = title.Trim().ToUpper();

            // 1. Lunghezza minima e massima
            if (t.Length < 3 || t.Length > 100) return false;

            // 2. Non deve essere solo il codice EAN o caratteri non alfabetici
            if (!string.IsNullOrEmpty(ean) && t == ean) return false;
            int letterCount = 0;
            foreach (char c in t)
            {
                if (char.IsLetter(c)) letterCount++;
            }
            if (letterCount < 3) return false;

            // 3. Rifiuta frammenti URL o percorsi file
            if (t.Contains("HTTP://") || t.Contains("HTTPS://") || t.Contains("WWW.") || t.Contains(".COM/") || t.Contains(".IT/") || t.Contains(".HTML") || t.Contains(".PHP"))
                return false;

            // 4. Blacklist parole rumorose, AI hallucination, pagine errore, cookie/privacy e query di ricerca generiche
            string[] blacklisted = new string[] {
                "ACCUEIL", "BENVENUTO", "BENVENUTI", "HOMEPAGE", "HOME PAGE", "INDEX OF", "IMPOSTAZIONI", "SETTINGS",
                "IDENTIFY OWNER", "SEARCH ANY NUMBER", "CONFRONTA PREZZI", "MIGLIOR PREZZO", "TROVAPREZZI",
                "CHATGPT", "CHAT GPT", "OPENAI", "PROMPT", "AI RESULT", "AI SEARCH", "GENERATED BY", "BOT",
                "GOOGLE", "ACCETTA", "ACCETTI", "PRIVACY", "COOKIE", "COOKIES", "CONSENT", "SEARCH", "CERCA CON", "CERCA SU",
                "PRIMA DI PROSEGUIRE", "TERMINI DI SERVIZIO", "POLITICA DI PRIVACY", "DISCLAIMER", "NOTIFICHE", "CONSENTI",
                "TRUFFA", "CHI CHIAMA", "NUMERO VERDE", "TELEFONO", "CELLULARE", "CALL CENTER", "PREFISSO", "RECLAMI",
                "FORUM", "WIKIPEDIA", "YOUTUBE", "FACEBOOK", "INSTAGRAM", "TIKTOK", "TWITTER", "REDDIT", "PINTEREST",
                "LOGIN", "ACCEDI", "SIGN IN", "SIGN UP", "PASSWORD", "ACCOUNT", "REGISTRATI", "LOG IN", "REGISTRAZIONE",
                "DOWNLOAD", "SCARICA", "PDF", "MANUALE", "DOCUMENT", "FILE EXTENSION",
                "RISULTATI PER", "RISULTATO PER", "RESULTS FOR", "NOT FOUND", "NON TROVATO", "NESSUN RISULTATO", "NESSUNA CORRISPONDENZA",
                "PAGE NOT FOUND", "404", "ERROR", "ERRORE", "BAD REQUEST", "ACCESS DENIED", "FORBIDDEN",
                "RECENSIONI", "OPINIONI", "VALUTAZIONE", "FEEDBACK", "COMMENTI",
                "CARRELLO", "CHECKOUT", "VAI AL CARRELLO", "SPEDIZIONE GRATUITA", "AGGIUNGI AL CARRELLO", "AGGIUNGI AL",
                "OFFERTE VOLANTINO", "ACQUISTA ONLINE", "COMPRA ORA", "PREZZO MIGLIORE"
            };

            foreach (string bad in blacklisted)
            {
                if (t.Contains(bad)) return false;
            }

            return true;
        }

        private string CleanWebProductTitle(string rawTitle, string ean)
        {
            if (string.IsNullOrEmpty(rawTitle)) return "";

            string text = rawTitle;
            if (!string.IsNullOrEmpty(ean))
            {
                text = text.Replace(ean, "");
            }
            text = text.Trim();

            // Rimuovi prefissi breadcrumb (es: "www.dm-drogeriemarkt.it › p › dRESTASE ...")
            text = Regex.Replace(text, @"^.*?(›|»|»|\.\.\.)\s*", "");

            string[] noiseTokens = new string[] {
                "DM ITALIA", "DM-DROGERIEMARKT.IT", "CLIVIA PROFUMI", "CLIVIAPROFUMI.EU", "BELLISSIMA-SHOP.NET", "BELLISSIMA-SHOP",
                "PROFUMERIAMUNNO.COM", "PROFUMERIAMUNNO", "MELONI STORE", "PIÙME", "PIUME", "GALIANI STORE",
                "AMAZON.IT", "AMAZON", "EBAY.IT", "EBAY", "COOP ONLINE", "COOP", "CARREFOUR.IT", "CARREFOUR", 
                "ESSELUNGA A CASA", "ESSELUNGA", "CONAD.IT", "CONAD", "TROVAPREZZI.IT", "TROVAPREZZI",
                "BENNET.COM", "BENNET", "TIGROS.IT", "TIGROS", "DESPAR", "EUROSPIN", "CRAI SPESA ONLINE", "CRAI",
                "PAM PANORAMA", "CICALIA", "DECÒ", "TODIS", "MD SPESA ONLINE", "MD DISCOUT", "IPER",
                "PREZZO", "ACQUISTA ORA", "ACQUISTA", "COMPRA ORA", "COMPRA", "OFFERTA SPECIALE", "OFFERTA", 
                "SPESA ONLINE", ".IT", ".COM", "SUPERMERCATO", "IPERMERCATO",
                "BELL-ITALIA", "BELL ITALIA", "SANTANNA.IT", "CARRYGO", "MYBEVERAGES", "EVERLI", "SUPERMERCATO24",
                "PREZZI E OFFERTE", "SCHEDA PRODOTTO", "VOLANTINO"
            };

            string[] parts = text.Split(new string[] { " - ", " | ", " : ", " – ", " — ", " :: " }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                string p = CleanWhitespaceAndPunctuation(part.Trim());
                bool containsNoise = false;
                foreach (string noise in noiseTokens)
                {
                    if (p.ToUpper().Contains(noise))
                    {
                        containsNoise = true;
                        break;
                    }
                }

                if (!containsNoise && IsValidProductTitle(p, ean))
                {
                    return p;
                }
            }

            if (parts.Length > 0)
            {
                string first = CleanWhitespaceAndPunctuation(parts[0].Trim());
                if (IsValidProductTitle(first, ean))
                    return first;
            }

            string cleanedText = CleanWhitespaceAndPunctuation(text);
            return IsValidProductTitle(cleanedText, ean) ? cleanedText : "";
        }

        private string CleanWhitespaceAndPunctuation(string str)
        {
            if (string.IsNullOrEmpty(str)) return "";
            string res = str.Trim();
            res = res.Trim('-', '|', ':', '•', '/', '\\', '_', '.', ',', ';', ' ', '\t', '\r', '\n');
            res = Regex.Replace(res, @"\s+", " ");
            return res;
        }

        private string FetchUrlContentWithUserAgent(string url)
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol = (System.Net.SecurityProtocolType)12288 | (System.Net.SecurityProtocolType)3072 | (System.Net.SecurityProtocolType)768 | System.Net.SecurityProtocolType.Tls;
                System.Net.ServicePointManager.ServerCertificateValidationCallback = (senderCert, cert, chain, sslPolicyErrors) => true;

                System.Net.HttpWebRequest req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
                req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36";
                req.Headers.Add("Accept-Language", "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7");
                req.AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate;
                req.Timeout = 6000;
                req.ReadWriteTimeout = 6000;
                req.CookieContainer = new System.Net.CookieContainer();
                req.AllowAutoRedirect = true;
                req.MaximumAutomaticRedirections = 5;

                using (System.Net.HttpWebResponse resp = (System.Net.HttpWebResponse)req.GetResponse())
                using (System.IO.Stream stream = resp.GetResponseStream())
                using (System.IO.StreamReader sr = new System.IO.StreamReader(stream, System.Text.Encoding.UTF8))
                {
                    return sr.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArtNuovo.FetchUrlContentWithUserAgent", ex.Message);
                return "";
            }
        }

        private string FetchUrlContent(string url)
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol = (System.Net.SecurityProtocolType)12288 | (System.Net.SecurityProtocolType)3072 | (System.Net.SecurityProtocolType)768 | System.Net.SecurityProtocolType.Tls;
                System.Net.ServicePointManager.ServerCertificateValidationCallback = (senderCert, cert, chain, sslPolicyErrors) => true;

                System.Net.HttpWebRequest req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
                req.UserAgent = "APOffice Retail Manager/1.0 (contact@apsistemi.it)";
                req.Headers.Add("Accept-Language", "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7");
                req.AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate;
                req.Timeout = 6000;
                req.ReadWriteTimeout = 6000;
                req.CookieContainer = new System.Net.CookieContainer();
                req.AllowAutoRedirect = true;
                req.MaximumAutomaticRedirections = 5;

                using (System.Net.HttpWebResponse resp = (System.Net.HttpWebResponse)req.GetResponse())
                using (System.IO.Stream stream = resp.GetResponseStream())
                using (System.IO.StreamReader sr = new System.IO.StreamReader(stream, Encoding.UTF8))
                {
                    return sr.ReadToEnd();
                }
            }
            catch
            {
                return "";
            }
        }

        private WebProductInfo ParseOpenFoodFactsJson(string json, string ean)
        {
            WebProductInfo info = new WebProductInfo();
            if (string.IsNullOrEmpty(json))
                return info;

            string status = ExtractJsonRaw(json, "status");
            if (status != "1" && !json.Contains("\"product\":"))
                return info;

            // 1. Estrai Nome Prodotto (privilegia versione italiana)
            string name = ExtractJsonString(json, "product_name_it");
            if (string.IsNullOrEmpty(name)) name = ExtractJsonString(json, "product_name");
            if (string.IsNullOrEmpty(name)) name = ExtractJsonString(json, "generic_name_it");
            if (string.IsNullOrEmpty(name)) name = ExtractJsonString(json, "abbreviated_product_name");
            if (string.IsNullOrEmpty(name)) name = ExtractJsonString(json, "generic_name");
            if (string.IsNullOrEmpty(name)) name = ExtractJsonString(json, "product_name_en");

            // 2. Estrai Marca
            string brand = ExtractJsonString(json, "brands");
            if (string.IsNullOrEmpty(brand)) brand = ExtractJsonString(json, "brand_owner");
            if (string.IsNullOrEmpty(brand)) brand = ExtractJsonString(json, "brands_tags");

            // 3. Estrai Quantità / Peso Netto
            string qtyStr = ExtractJsonString(json, "quantity");
            decimal prodQty = ExtractJsonDecimal(json, "product_quantity");
            string prodQtyUnit = ExtractJsonString(json, "product_quantity_unit");
            decimal netWeight = ExtractJsonDecimal(json, "net_weight_value");
            string netWeightUnit = ExtractJsonString(json, "net_weight_unit");
            string servingSize = ExtractJsonString(json, "serving_size");

            // 4. Estrai Categoria e Immagine
            string imgUrl = ExtractJsonString(json, "image_front_url");
            if (string.IsNullOrEmpty(imgUrl)) imgUrl = ExtractJsonString(json, "image_url");
            if (string.IsNullOrEmpty(imgUrl)) imgUrl = ExtractJsonString(json, "image_small_url");
            if (string.IsNullOrEmpty(imgUrl)) imgUrl = ExtractJsonString(json, "image_front_small_url");
            if (string.IsNullOrEmpty(imgUrl)) imgUrl = ExtractJsonString(json, "image_nutrition_url");
            string catStr = ExtractJsonString(json, "categories");

            if (!string.IsNullOrEmpty(name))
            {
                name = CleanWhitespaceAndPunctuation(name);
                brand = CleanWhitespaceAndPunctuation(brand);

                // Unione pulita di Marca e Nome
                string combinedTitle = BuildCleanBrandAndName(brand, name);

                // Estrazione avanzata di grammatura e peso netto
                ParseWeightAndGrammatura(combinedTitle, qtyStr, prodQty, prodQtyUnit, netWeight, netWeightUnit, servingSize, info);

                // Formatta il titolo finale con il suffisso del peso se non già presente
                string finalTitle = FormatFinalProductTitle(combinedTitle, info, ean);

                if (IsValidProductTitle(finalTitle, ean))
                {
                    info.Found = true;
                    info.Name = finalTitle.ToUpper();
                    info.Brand = brand;
                    info.ImageUrl = imgUrl;
                    info.Category = catStr;
                    info.IvaCode = DeduceIvaCode(info.Name, catStr);
                }
            }

            return info;
        }

        private WebProductInfo ParseUpcItemDbJson(string json, string ean)
        {
            WebProductInfo info = new WebProductInfo();
            if (string.IsNullOrEmpty(json) || !json.Contains("\"items\":"))
                return info;

            string title = ExtractJsonString(json, "title");
            string brand = ExtractJsonString(json, "brand");
            string category = ExtractJsonString(json, "category");
            string imgUrl = ExtractJsonString(json, "images");
            if (string.IsNullOrEmpty(imgUrl)) imgUrl = ExtractJsonString(json, "image");

            if (!string.IsNullOrEmpty(title))
            {
                title = CleanWhitespaceAndPunctuation(title);
                brand = CleanWhitespaceAndPunctuation(brand);

                string combinedTitle = BuildCleanBrandAndName(brand, title);
                ParseWeightAndGrammatura(combinedTitle, "", 0, "", 0, "", "", info);
                string finalTitle = FormatFinalProductTitle(combinedTitle, info, ean);

                if (IsValidProductTitle(finalTitle, ean))
                {
                    info.Found = true;
                    info.Name = finalTitle.ToUpper();
                    info.Brand = brand;
                    info.ImageUrl = imgUrl;
                    info.Category = category;
                    info.IvaCode = DeduceIvaCode(info.Name, category);
                }
            }

            return info;
        }

        private string BuildCleanBrandAndName(string brand, string name)
        {
            if (string.IsNullOrEmpty(name)) return CleanWhitespaceAndPunctuation(brand).ToUpper();
            if (string.IsNullOrEmpty(brand)) return CleanWhitespaceAndPunctuation(name).ToUpper();

            string cBrand = CleanWhitespaceAndPunctuation(brand).ToUpper();
            string cName = CleanWhitespaceAndPunctuation(name).ToUpper();

            // Evita duplicazione della marca se è già inclusa all'inizio o all'interno del nome
            if (cName.StartsWith(cBrand + " ") || cName.EndsWith(" " + cBrand) || cName.Contains(" " + cBrand + " ") || cName == cBrand)
            {
                return cName;
            }

            return (cBrand + " " + cName).Trim();
        }

        private void ParseWeightAndGrammatura(string text, string qtyStr, decimal prodQty, string prodQtyUnit, decimal netWeight, string netWeightUnit, string servingSize, WebProductInfo info)
        {
            info.Content = 1;
            info.Unit = "PZ";
            info.Grammatura = "PZ";

            // 1. Controlla valori numerici strutturati dell'API se presenti
            if (prodQty > 0 && !string.IsNullOrEmpty(prodQtyUnit))
            {
                AssignWeight(prodQty, prodQtyUnit, info);
                return;
            }
            if (netWeight > 0 && !string.IsNullOrEmpty(netWeightUnit))
            {
                AssignWeight(netWeight, netWeightUnit, info);
                return;
            }

            // 2. Combina stringhe di quantità e testo descrizione
            string fullSearch = (qtyStr + " " + servingSize + " " + text).Trim();
            if (string.IsNullOrEmpty(fullSearch)) return;

            // Pattern confezione multipla: es. "6x33cl", "6 x 33 cl", "4 x 100 g", "3 x 200g", "2 x 1,5 L", "8x200ml"
            Match mMulti = Regex.Match(fullSearch, @"(\d+)\s*[xX*]\s*(\d+([\.,]\d+)?)\s*(g|gr|grammi|kg|l|lt|litri|ml|cl|cc|pz)\b", RegexOptions.IgnoreCase);
            if (mMulti.Success)
            {
                int packs = int.Parse(mMulti.Groups[1].Value);
                decimal singleVal;
                string numStr = mMulti.Groups[2].Value.Replace(',', '.');
                if (decimal.TryParse(numStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out singleVal))
                {
                    string unit = mMulti.Groups[4].Value.ToUpper();
                    AssignWeight(singleVal, unit, info);
                    info.Quantity = packs + "X" + singleVal.ToString(System.Globalization.CultureInfo.InvariantCulture) + info.Grammatura;
                    return;
                }
            }

            // Pattern singola quantità/peso: es. "500 g", "500g", "1.5 l", "1,5lt", "750 ml", "33 cl", "1 kg", "100 pz"
            Match mSingle = Regex.Match(fullSearch, @"(\d+([\.,]\d+)?)\s*(g|gr|grammi|kg|l|lt|litri|ml|cl|cc|pz|capsule|cialde|filtri|bustine|pezzi)\b", RegexOptions.IgnoreCase);
            if (mSingle.Success)
            {
                decimal val;
                string numStr = mSingle.Groups[1].Value.Replace(',', '.');
                if (decimal.TryParse(numStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out val))
                {
                    string unit = mSingle.Groups[3].Value.ToUpper();
                    AssignWeight(val, unit, info);
                    info.Quantity = val.ToString(System.Globalization.CultureInfo.InvariantCulture) + info.Grammatura;
                    return;
                }
            }
        }

        private void AssignWeight(decimal val, string rawUnit, WebProductInfo info)
        {
            string u = rawUnit.Trim().ToUpper();
            info.Content = val;
            info.Unit = "PZ";

            if (u == "G" || u == "GR" || u == "GRAMMI" || u == "GRAMS" || u == "G.")
            {
                info.Grammatura = "GR";
            }
            else if (u == "KG" || u == "KILOGRAMMI" || u == "KG.")
            {
                info.Grammatura = "KG";
            }
            else if (u == "L" || u == "LT" || u == "LITRI" || u == "LITRE" || u == "L.")
            {
                info.Grammatura = "LT";
            }
            else if (u == "ML" || u == "MILLILITRI" || u == "ML.")
            {
                info.Grammatura = "ML";
            }
            else if (u == "CL" || u == "CENTILITRI" || u == "CL.")
            {
                info.Grammatura = "CL";
            }
            else if (u == "CC")
            {
                info.Grammatura = "CC";
            }
            else
            {
                info.Grammatura = "PZ";
            }
        }

        private string FormatFinalProductTitle(string name, WebProductInfo info, string ean)
        {
            if (string.IsNullOrEmpty(name)) return "";

            string title = CleanWebProductTitle(name, ean);
            if (string.IsNullOrEmpty(title)) return "";

            title = title.ToUpper();

            // Aggiungi la grammatura al termine del nome se non già presente nel testo
            if (info != null && info.Content > 0 && info.Grammatura != "PZ")
            {
                string weightToken1 = info.Content.ToString(System.Globalization.CultureInfo.InvariantCulture) + info.Grammatura;
                string weightToken2 = info.Content.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " + info.Grammatura;
                string weightToken3 = info.Content.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + info.Grammatura;
                string numOnly = info.Content.ToString(System.Globalization.CultureInfo.InvariantCulture);

                bool alreadyHasWeight = title.Contains(weightToken1) || title.Contains(weightToken2) || title.Contains(weightToken3) ||
                                        (title.Contains(numOnly) && (title.Contains(" " + info.Grammatura) || title.Contains(info.Grammatura + " ") || title.EndsWith(info.Grammatura)));

                if (!alreadyHasWeight)
                {
                    if (!string.IsNullOrEmpty(info.Quantity))
                    {
                        title = (title + " " + info.Quantity).Trim();
                    }
                    else
                    {
                        title = (title + " " + weightToken1).Trim();
                    }
                }
            }

            return CleanWhitespaceAndPunctuation(title);
        }

        private string ExtractJsonString(string json, string key)
        {
            if (string.IsNullOrEmpty(json)) return "";
            try
            {
                // Match stringa: "key": "val"
                string pattern = "\"" + Regex.Escape(key) + "\":\\s*\"((?:\\\\\"|[^\"])*)\"";
                Match m = Regex.Match(json, pattern);
                if (m.Success)
                {
                    return UnescapeJsonString(m.Groups[1].Value).Trim();
                }

                // Match primo elemento array: "key": ["val1", "val2"]
                string arrPattern = "\"" + Regex.Escape(key) + "\":\\s*\\[\\s*\"((?:\\\\\"|[^\"])*)\"";
                Match mArr = Regex.Match(json, arrPattern);
                if (mArr.Success)
                {
                    return UnescapeJsonString(mArr.Groups[1].Value).Trim();
                }
            }
            catch { }
            return "";
        }

        private decimal ExtractJsonDecimal(string json, string key)
        {
            if (string.IsNullOrEmpty(json)) return 0;
            try
            {
                string pattern = "\"" + Regex.Escape(key) + "\":\\s*\"?(-?[0-9]+(?:\\.[0-9]+)?)\"?";
                Match m = Regex.Match(json, pattern);
                if (m.Success)
                {
                    decimal d;
                    if (decimal.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out d))
                    {
                        return d;
                    }
                }
            }
            catch { }
            return 0;
        }

        private string ExtractJsonRaw(string json, string key)
        {
            if (string.IsNullOrEmpty(json)) return "";
            try
            {
                string pattern = "\"" + Regex.Escape(key) + "\":\\s*([^,\r\n}\\]]+)";
                Match m = Regex.Match(json, pattern);
                if (m.Success)
                {
                    return m.Groups[1].Value.Trim().Trim('"');
                }
            }
            catch { }
            return "";
        }

        private string UnescapeJsonString(string val)
        {
            if (string.IsNullOrEmpty(val)) return "";
            try
            {
                // Decodifica sequenze unicode \uXXXX
                val = Regex.Replace(val, @"\\u(?<val>[a-fA-F0-9]{4})", m =>
                {
                    int code = int.Parse(m.Groups["val"].Value, System.Globalization.NumberStyles.HexNumber);
                    return ((char)code).ToString();
                });

                val = val.Replace("\\\"", "\"")
                         .Replace("\\\\", "\\")
                         .Replace("\\/", "/")
                         .Replace("\\n", " ")
                         .Replace("\\r", " ")
                         .Replace("\\t", " ");
            }
            catch { }
            return val;
        }

        private string DeduceIvaCode(string name, string json)
        {
            string lowerStr = (name + " " + json).ToLower();

            // 4% IVA items (pane fresco, latte fresco, frutta fresca, burro, libri)
            if (lowerStr.Contains("pane fresco") || lowerStr.Contains("latte fresco") || lowerStr.Contains("frutta fresca") || lowerStr.Contains("burro"))
                return "004";

            // 22% IVA items (cosmetici, depilazione, dopocera, epilazione, oli corpo, igiene, detersivi, cura casa, cura persona, profumi, alcolici, vino, birra, liquori, casalinghi, non food)
            if (lowerStr.Contains("dopocera") || lowerStr.Contains("epilazion") || lowerStr.Contains("depilator") || lowerStr.Contains("depilazion") || lowerStr.Contains("cera ") || lowerStr.Contains("ceretta") || lowerStr.Contains("olio corpo") ||
                lowerStr.Contains("shampoo") || lowerStr.Contains("bagnoschiuma") || lowerStr.Contains("sapone") || lowerStr.Contains("detergente") || lowerStr.Contains("detersivo") || lowerStr.Contains("dentifricio") || lowerStr.Contains("candeggina") || lowerStr.Contains("ammorbidente") || lowerStr.Contains("deodorante") || lowerStr.Contains("crema") || lowerStr.Contains("profum") || lowerStr.Contains("cosmetic") || lowerStr.Contains("cosmesi") || lowerStr.Contains("trucco") || lowerStr.Contains("fazzoletti") || lowerStr.Contains("carta igienica") ||
                lowerStr.Contains("birra") || lowerStr.Contains("vino") || lowerStr.Contains("liquor") || lowerStr.Contains("amaro") || lowerStr.Contains("gin ") || lowerStr.Contains("vodka") || lowerStr.Contains("whisky") || lowerStr.Contains("rum ") || lowerStr.Contains("alcol"))
            {
                return "022";
            }

            // 10% IVA items (prodotti da scaffale: pasta, sugo, pesto, conserve, caffe, biscotti, farina, olio, carne, riso, formaggio, salumi, bevande analcoliche)
            return "010";
        }

        private void SelectComboValueOrText(ComboBox cmb, string targetVal, string[] keywords)
        {
            if (cmb == null || cmb.DataSource == null) return;

            try
            {
                DataTable dt = cmb.DataSource as DataTable;
                if (dt != null)
                {
                    targetVal = (targetVal ?? "").Trim();

                    // 1. Ricerca per codice ValueMember
                    if (!string.IsNullOrEmpty(targetVal))
                    {
                        // 1a. Match esatto
                        foreach (DataRow row in dt.Rows)
                        {
                            string cod = row[cmb.ValueMember] != null ? row[cmb.ValueMember].ToString().Trim() : "";
                            if (!string.IsNullOrEmpty(cod) && cod.Equals(targetVal, StringComparison.OrdinalIgnoreCase))
                            {
                                cmb.SelectedValue = row[cmb.ValueMember];
                                return;
                            }
                        }

                        // 1b. Match normalizzato (senza zeri iniziali, es. "022" == "22", "004" == "04")
                        string normTarget = targetVal.TrimStart('0');
                        if (string.IsNullOrEmpty(normTarget)) normTarget = "0";

                        foreach (DataRow row in dt.Rows)
                        {
                            string cod = row[cmb.ValueMember] != null ? row[cmb.ValueMember].ToString().Trim() : "";
                            if (string.IsNullOrEmpty(cod)) continue; // MAI confrontare con la riga vuota!

                            string normCod = cod.TrimStart('0');
                            if (string.IsNullOrEmpty(normCod)) normCod = "0";

                            if (normCod.Equals(normTarget, StringComparison.OrdinalIgnoreCase))
                            {
                                cmb.SelectedValue = row[cmb.ValueMember];
                                return;
                            }
                        }

                        // 1c. Match EndsWith / StartsWith (solo su codici non vuoti!)
                        foreach (DataRow row in dt.Rows)
                        {
                            string cod = row[cmb.ValueMember] != null ? row[cmb.ValueMember].ToString().Trim() : "";
                            if (string.IsNullOrEmpty(cod)) continue; // Evita il bug targetVal.EndsWith("") == true!

                            if (cod.EndsWith(targetVal, StringComparison.OrdinalIgnoreCase) || targetVal.EndsWith(cod, StringComparison.OrdinalIgnoreCase))
                            {
                                cmb.SelectedValue = row[cmb.ValueMember];
                                return;
                            }
                        }
                    }

                    // 2. Ricerca per DisplayMember text keywords (es. "Alimentari", "Igiene", "PZ", "KG")
                    if (keywords != null && keywords.Length > 0)
                    {
                        foreach (string kw in keywords)
                        {
                            if (string.IsNullOrEmpty(kw) || kw.Trim().Length == 0) continue;
                            string cleanKw = kw.Trim();

                            foreach (DataRow row in dt.Rows)
                            {
                                string cod = row[cmb.ValueMember] != null ? row[cmb.ValueMember].ToString().Trim() : "";
                                if (string.IsNullOrEmpty(cod)) continue; // Ignora la riga "Non definito"

                                string des = row[cmb.DisplayMember] != null ? row[cmb.DisplayMember].ToString().Trim() : "";
                                if (des.IndexOf(cleanKw, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    cmb.SelectedValue = row[cmb.ValueMember];
                                    return;
                                }
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void SelectRepartoByCategory(string category, string productName)
        {
            string fullStr = (category + " " + productName).ToLower();

            // 1. Igiene, Cosmetica, Cura Persona, Depilazione, Profumeria, Detersivi, Non Food
            if (fullStr.Contains("dopocera") || fullStr.Contains("epilazion") || fullStr.Contains("depil") || fullStr.Contains("cera") || fullStr.Contains("ceretta") || fullStr.Contains("cosmetic") || fullStr.Contains("cosmesi") || fullStr.Contains("profum") ||
                fullStr.Contains("igiene") || fullStr.Contains("detersiv") || fullStr.Contains("detergent") || fullStr.Contains("sapone") || fullStr.Contains("shampoo") || fullStr.Contains("bagnoschiuma") || fullStr.Contains("crema") || fullStr.Contains("cura casa") || fullStr.Contains("cura persona") || fullStr.Contains("bazar") || fullStr.Contains("non food"))
            {
                SelectComboValueOrText(cmbArtRep, "004", new string[] { "igiene", "cosmesi", "cosmetica", "profumeria", "cura persona", "detersivi", "non food", "casa", "bazar", "generale" });
                return;
            }

            // 2. Bevande, Acqua, Bibite, Vini, Birre, Liquori
            if (fullStr.Contains("bevand") || fullStr.Contains("acqua") || fullStr.Contains("bibit") || fullStr.Contains("birra") || fullStr.Contains("vino") || fullStr.Contains("succo") || fullStr.Contains("liquor") || fullStr.Contains("amaro") || fullStr.Contains("spumante") || fullStr.Contains("whisky"))
            {
                SelectComboValueOrText(cmbArtRep, "001", new string[] { "bevande", "liquori", "acqua", "vini", "bibite" });
                return;
            }

            // 3. Freschi, Latticini, Salumi, Formaggi, Gastronomia, Yogurt, Burro
            if (fullStr.Contains("fresc") || fullStr.Contains("latte") || fullStr.Contains("formagg") || fullStr.Contains("salum") || fullStr.Contains("yogurt") || fullStr.Contains("burro") || fullStr.Contains("mozzarella") || fullStr.Contains("prosciutto"))
            {
                SelectComboValueOrText(cmbArtRep, "003", new string[] { "freschi", "latticini", "gastronomia", "salumi", "formaggi" });
                return;
            }

            // 4. Ortofrutta
            if (fullStr.Contains("ortofrutta") || fullStr.Contains("verdura") || fullStr.Contains("frutta"))
            {
                SelectComboValueOrText(cmbArtRep, "005", new string[] { "ortofrutta", "frutta", "verdura" });
                return;
            }

            // 5. Macelleria, Carni
            if (fullStr.Contains("macelleria") || fullStr.Contains("carne ") || fullStr.Contains("carni ") || fullStr.Contains("bovino") || fullStr.Contains("suino") || fullStr.Contains("pollo") || fullStr.Contains("tacchino"))
            {
                SelectComboValueOrText(cmbArtRep, "006", new string[] { "macelleria", "carne", "carni" });
                return;
            }

            // 6. Default: Alimentari / Scaffale / Drogheria
            SelectComboValueOrText(cmbArtRep, "002", new string[] { "alimentari", "drogheria", "pasta", "generale", "scaffale" });
        }
        #endregion

        #region Calcolo Automatico Prezzo da Ricarico & Aliquota IVA
        private decimal GetSelectedIvaAliquota()
        {
            try
            {
                if (cmbArtIva != null && cmbArtIva.SelectedItem != null)
                {
                    DataRowView drv = cmbArtIva.SelectedItem as DataRowView;
                    if (drv != null && drv.Row != null)
                    {
                        if (drv.Row.Table.Columns.Contains("tab_ali") && drv.Row["tab_ali"] != DBNull.Value)
                        {
                            decimal ali = 0;
                            if (decimal.TryParse(drv.Row["tab_ali"].ToString().Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out ali))
                            {
                                return ali;
                            }
                        }
                    }
                }

                if (cmbArtIva != null)
                {
                    string txt = cmbArtIva.Text;
                    Match m = Regex.Match(txt, @"(\d+(?:[\.,]\d+)?)");
                    if (m.Success)
                    {
                        decimal ali = 0;
                        if (decimal.TryParse(m.Groups[1].Value.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out ali))
                        {
                            return ali;
                        }
                    }
                }
            }
            catch { }
            return 22m;
        }

        private bool _isCalculatingPrice = false;

        private decimal CalcolaCoefficienteRicarico(decimal inputVal)
        {
            if (inputVal <= 0m) return 1.0m;
            // Se l'utente inserisce un valore >= 5.0 (es. 20, 30, 45), è espresso in percentuale (%)
            if (inputVal >= 5.0m)
                return 1.0m + (inputVal / 100.0m);
            // Se l'utente inserisce un valore tra 1.0 e 5.0 (es. 1.20, 1.30, 1.45), è il coefficiente diretto
            if (inputVal >= 1.0m)
                return inputVal;
            // Se l'utente inserisce un valore tra 0.01 e 0.99 (es. 0.30 per indicare +30%)
            return 1.0m + inputVal;
        }

        private void CalcolaPrezzoDaRicarico()
        {
            if (_isCalculatingPrice) return;
            try
            {
                _isCalculatingPrice = true;

                decimal dCost = _clsFun.Txt2Dec(txtLiaPrc != null ? txtLiaPrc.Text : "0");
                decimal dPercInput = _clsFun.Txt2Dec(txtRicarico != null ? txtRicarico.Text : "0");
                decimal dAli = GetSelectedIvaAliquota();

                if (dCost > 0 && dPercInput > 0)
                {
                    decimal dK = CalcolaCoefficienteRicarico(dPercInput);
                    // Formula reversibile standard: Prezzo Ivato = Costo * K * (1 + Aliquota/100)
                    // Esempio: Costo 1.00, Iva 10%, Ricarico 30% (K=1.30) -> 1.00 * 1.30 * 1.10 = 1.43
                    decimal dPrezzoIvato = dCost * dK * (1m + (dAli / 100m));
                    decimal dPrezzoFinale = Math.Round(dPrezzoIvato, 2, MidpointRounding.AwayFromZero);

                    if (txtLivPrv != null)
                    {
                        txtLivPrv.Text = dPrezzoFinale.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);
                        txtLivPrv.BackColor = Color.FromArgb(240, 253, 244);
                    }
                }
            }
            catch { }
            finally
            {
                _isCalculatingPrice = false;
            }
        }

        private void txtLiaPrc_TextChanged(object sender, EventArgs e)
        {
            CalcolaPrezzoDaRicarico();
        }

        private void txtRicarico_TextChanged(object sender, EventArgs e)
        {
            CalcolaPrezzoDaRicarico();
        }

        private void cmbArtIva_SelectionChangeCommitted(object sender, EventArgs e)
        {
            CalcolaPrezzoDaRicarico();
        }
        #endregion

        private void txtArt_KeyDown(object sender, KeyEventArgs e)
        {
            TextBox txt = (TextBox)sender;

            if (e.KeyCode == Keys.Return)
            {
                if (txt.Name == "txtArtDes")
                    txtArtNet.Select();
                else if (txt.Name == "txtArtNet")
                {
                    if (txtArtArf != null)
                        txtArtArf.Select();
                    else
                        txtLiaPrc.Select();
                }
                else if (txt.Name == "txtArtArf")
                    txtLiaPrc.Select();
                else if (txt.Name == "txtLiaPrc")
                {
                    if (txtRicarico != null)
                        txtRicarico.Select();
                    else
                        txtLivPrv.Select();
                }
                else if (txt.Name == "txtRicarico")
                    txtLivPrv.Select();
                else if (txt.Name == "txtLivPrv")
                    btnSalvaNuovo.Select();
                else if (txt.Name == "txtEanEan")
                {
                    string ean = txtEanEan.Text.Trim();
                    if (_tmrEanInput != null) _tmrEanInput.Stop();

                    if (chkRicercaAuto != null && chkRicercaAuto.Checked)
                    {
                        if (ean.Length >= 6)
                        {
                            _strLastEanWebSearched = ean;
                            CercaWeb(ean);
                        }
                    }
                    else
                    {
                        if (ean.Length >= 4)
                        {
                            if (VerificaBarcodeEsistente(ean, false))
                                return;
                        }
                    }
                    txtArtDes.Select();
                }
            }
        }

        private void txtArt_Validated(object sender, EventArgs e)
        {
            TextBox txt = (TextBox)sender;

            if (txt.Name == "txtArtNet" || txt.Name == "txtLiaPrc" || txt.Name == "txtLivPrv" || txt.Name == "txtRicarico")
            {
                txt.Text = txt.Text.Replace(",", ".");
                if (string.IsNullOrWhiteSpace(txt.Text))
                {
                    txt.Text = (txt.Name == "txtArtNet") ? "1" : "0";
                }

                if (txt.Text != "0" && txt.Name == "txtLiaPrc" && (cmbLiaFor.SelectedValue == null || cmbLiaFor.SelectedValue.ToString() == ""))
                {
                    // Fornitore opzionale ma avvisabile
                }
                else if (txt.Text != "0" && txt.Name == "txtLivPrv" && (cmbLivLis.SelectedValue == null || cmbLivLis.SelectedValue.ToString() == ""))
                {
                    MessageBox.Show("Listino di vendita non definito!", "CONTROLLO VALORI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    cmbLivLis.Select();
                }
                else if (!_clsFun.Numerico(txt.Text, "0123456789,."))
                {
                    MessageBox.Show("Richiesti valori numerici!", "CONTROLLO VALORI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txt.Text = (txt.Name == "txtArtNet") ? "1" : "0";
                    txt.Select();
                }

                if (txt.Name == "txtLiaPrc" || txt.Name == "txtRicarico")
                {
                    CalcolaPrezzoDaRicarico();
                }
            }
            else if (txt.Name == "txtEanEan")
            {
                string ean = txtEanEan.Text.Trim();
                if (chkRicercaAuto != null && chkRicercaAuto.Checked)
                {
                    if (ean.Length >= 6 && ean != _strLastEanWebSearched)
                    {
                        if (_tmrEanInput != null) _tmrEanInput.Stop();
                        _strLastEanWebSearched = ean;
                        CercaWeb(ean);
                    }
                }
                else
                {
                    if (ean.Length >= 4)
                    {
                        VerificaBarcodeEsistente(ean, false);
                    }
                }
            }
        }

        private void txtEanEan_TextChanged(object sender, EventArgs e)
        {
            string strEan = txtEanEan.Text.Trim();
            if (chkRicercaAuto == null || !chkRicercaAuto.Checked)
            {
                if (_tmrEanInput != null) _tmrEanInput.Stop();
                return;
            }

            if (strEan.Length >= 6 && strEan != _strLastEanWebSearched)
            {
                if (strEan.Length == 8 || strEan.Length == 12 || strEan.Length == 13 || strEan.Length == 14)
                {
                    if (_tmrEanInput != null) _tmrEanInput.Stop();
                    _strLastEanWebSearched = strEan;
                    CercaWeb(strEan);
                }
                else
                {
                    if (_tmrEanInput != null)
                    {
                        _tmrEanInput.Stop();
                        _tmrEanInput.Start();
                    }
                }
            }
        }

        private Boolean CtrlEan()
        {
            Boolean b = false;

            string sEan = txtEanEan.Text.Trim();

            if (string.IsNullOrEmpty(sEan))
                return false;

            if (!_clsFun.Numerico(sEan))
            {
                MessageBox.Show("Barcode non corretto!", "CONTROLLO BARCODE", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            b = true;
            if (!(sEan.Substring(0, 1) == "2" && sEan.Length == 13 && sEan.Substring(8, 5) == "00000"))
            {
                double dVal = 0;
                if (double.TryParse(sEan, out dVal) && dVal > 799999)
                {
                    string sMod = sEan.Substring(0, sEan.Length - 1) + new clsCtrlCodici().FindMod10Digit(sEan.Substring(0, sEan.Length - 1));
                    if (sMod != sEan)
                    {
                        b = false;
                        MessageBox.Show("Barcode non corretto! (" + sMod + ")", "CONTROLLO BARCODE", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }

            if (b)
            {
                if (VerificaBarcodeEsistente(sEan, false))
                {
                    return false;
                }
            }

            return b;
        }

        private void btnSalvaNuovo_Click(object sender, EventArgs e)
        {
            string sMsg = "";

            if (string.IsNullOrEmpty(txtArtDes.Text.Trim()))
                sMsg += "Descrizione non definita " + _clsDef.CRLF;
            if (string.IsNullOrEmpty(txtEanEan.Text.Trim()))
                sMsg += "Barcode non definito " + _clsDef.CRLF;
            else if (!_clsFun.Numerico(txtEanEan.Text.Trim()))
                sMsg += "Il codice a barre deve contenere solo caratteri numerici " + _clsDef.CRLF;

            if (sMsg != "")
                MessageBox.Show(sMsg, "DATI MANCANTI O NON VALIDI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                SalvaVoci();
                string sCreatedCod = SalvaArticoloDb();
                if (!string.IsNullOrEmpty(sCreatedCod))
                {
                    _strRes = "art_cod:" + sCreatedCod + ";";
                    MessageBox.Show("Articolo '" + txtArtDes.Text + "' (Cod. " + sCreatedCod + ") salvato con successo!", "SALVATAGGIO OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ResetFormForNextInsertion();
                }
            }
        }

        private void ResetFormForNextInsertion()
        {
            _strLastEanWebSearched = "";
            _strLastWebImageUrl = "";
            txtEanEan.Text = "";
            txtArtDes.Text = "";
            if (txtArtArf != null) txtArtArf.Text = "";
            txtLiaPrc.Text = "0";
            txtLivPrv.Text = "0";
            txtArtCod.Text = "NEW";
            if (picWebProduct.Image != null)
            {
                try { picWebProduct.Image.Dispose(); } catch { }
                picWebProduct.Image = null;
            }
            UpdateWebStatusLabelPrompt();
            lblPicCaption.Text = "Foto da Web";

            txtArtDes.BackColor = Color.White;
            txtArtNet.BackColor = Color.White;
            txtLivPrv.BackColor = Color.White;

            if (!chkMem.Checked)
            {
                txtArtNet.Text = "1";
                if (txtRicarico != null) txtRicarico.Text = "0";
                cmbArtIva.SelectedValue = "";
                cmbArtUmi.SelectedValue = "";
                cmbArtTgr.SelectedValue = "";
                cmbArtRep.SelectedValue = "";
                cmbLiaFor.SelectedValue = "";
                SelectDefaultListinoCassa();
            }
            else
            {
                // Se chkMem.Checked == true:
                // MANTIENE i valori impostati (IVA, UM, Grammatura, Reparto, Fornitore, Listino, % Ricarico, Contenuto/Peso),
                // azzerando solo Nome, Cod. Art. Fornitore, Costo e Prezzo per il nuovo articolo.
            }

            txtEanEan.Select();
            txtEanEan.Focus();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            // 1. Controlla se ci sono dati minimi per tentare il salvataggio
            bool isMinimallyFilled = !string.IsNullOrEmpty(txtEanEan.Text.Trim()) ||
                                     !string.IsNullOrEmpty(txtArtDes.Text.Trim());

            if (!isMinimallyFilled)
            {
                // Schermata vuota o senza dati essenziali: esce direttamente
                Esci();
                return;
            }

            // 2. Controllo dei campi obbligatori
            string sMsg = "";

            if (string.IsNullOrEmpty(txtArtDes.Text.Trim()))
                sMsg += "Descrizione non definita " + _clsDef.CRLF;
            if (string.IsNullOrEmpty(txtEanEan.Text.Trim()))
                sMsg += "Barcode non definito " + _clsDef.CRLF;
            else if (!_clsFun.Numerico(txtEanEan.Text.Trim()))
                sMsg += "Il codice a barre deve contenere solo caratteri numerici " + _clsDef.CRLF;

            if (sMsg != "")
            {
                // Prodotto non completo: avvisa l'utente e chiede se vuole uscire senza salvare
                DialogResult res = MessageBox.Show(
                    "Impossibile salvare l'articolo. I seguenti dati risultano incompleti:\n\n" + sMsg + "\nUscire senza salvare l'articolo?",
                    "PRODOTTO NON COMPLETO",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);

                if (res == DialogResult.Yes)
                {
                    Esci();
                }
                return;
            }

            // 3. Se completo, salva le voci preferite, salva l'articolo ed esce
            SalvaVoci();
            string sCreatedCod = SalvaArticoloDb();
            if (!string.IsNullOrEmpty(sCreatedCod))
            {
                _strRes = "art_cod:" + sCreatedCod + ";";
                Esci();
            }
        }

        private void Salva()
        {
            string sCreatedArtCod = SalvaArticoloDb();

            if (string.IsNullOrEmpty(sCreatedArtCod))
                sCreatedArtCod = txtArtCod.Text;

            _strRes = "art_cod:" + sCreatedArtCod + ";";
        }

        private bool EsisteArticoloCodice(string cod)
        {
            if (string.IsNullOrEmpty(cod)) return false;
            try
            {
                DataTable t = _clsFun.FillTabSql("AnaArticoli", "SELECT COUNT(*) FROM AnaArticoli WHERE art_cod='" + cod.Replace("'", "''") + "'", false, _strConSql);
                if (t != null && t.Rows.Count > 0 && t.Rows[0][0] != DBNull.Value)
                {
                    int cnt = 0;
                    if (int.TryParse(t.Rows[0][0].ToString(), out cnt) && cnt > 0)
                        return true;
                }
            }
            catch { }
            return false;
        }

        private string SalvaArticoloDb()
        {
            try
            {
                if (string.IsNullOrEmpty(_strConSql))
                    _strConSql = _clsFun.ConSql("");

                string sEan = txtEanEan.Text.Trim();
                if (string.IsNullOrEmpty(sEan)) return "";

                // 1. Controllo rigoroso duplicati: se il barcode è già associato a un articolo esistente
                string existingArt = "";
                string existingDes = "";
                if (CercaArticoloLocale(sEan, out existingArt, out existingDes))
                {
                    MessageBox.Show(
                        "Impossibile salvare il nuovo articolo!\n\nIl codice a barre " + sEan + " risulta già registrato per l'articolo:\n" +
                        "Codice: " + existingArt + "\n" +
                        "Descrizione: " + existingDes + "\n\n" +
                        "Modifica il codice a barre prima di salvare o apri la scheda dell'articolo esistente.",
                        "BARCODE GIÀ ESISTENTE",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return "";
                }

                // 2. Generate new unique article code con controllo anti-collisione
                string sNewArtCod = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaArticoli, 7, _strConSql);
                if (string.IsNullOrEmpty(sNewArtCod))
                    sNewArtCod = "0000001";

                int safetyTries = 0;
                while (EsisteArticoloCodice(sNewArtCod) && safetyTries < 500)
                {
                    sNewArtCod = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaArticoli, 7, _strConSql);
                    safetyTries++;
                }

                bool isBilancia = (sEan.Length == 12 || sEan.Length == 13) && sEan.StartsWith("2");

                string sDes = txtArtDes.Text.Trim();
                if (sDes.Length > 50)
                    sDes = sDes.Substring(0, 50);

                string sDeb = sDes.Length > 20 ? sDes.Substring(0, 20) : sDes;

                string sIva = cmbArtIva.SelectedValue != null ? cmbArtIva.SelectedValue.ToString() : "022";
                if (string.IsNullOrEmpty(sIva)) sIva = "022";

                string sUmi = cmbArtUmi.SelectedValue != null ? cmbArtUmi.SelectedValue.ToString() : (isBilancia ? "KG" : "PZ");
                if (string.IsNullOrEmpty(sUmi)) sUmi = isBilancia ? "KG" : "PZ";

                string sTgr = cmbArtTgr.SelectedValue != null ? cmbArtTgr.SelectedValue.ToString() : "GR";
                if (string.IsNullOrEmpty(sTgr)) sTgr = "GR";

                // Merceologie fisse 999 (Generico) - non richieste all'operatore
                EnsureEcr999();
                string sEc1 = "999";
                string sEc2 = "999";
                string sEc3 = "999";

                string sRep = cmbArtRep.SelectedValue != null ? cmbArtRep.SelectedValue.ToString() : "";
                if (string.IsNullOrEmpty(sRep)) sRep = "01";

                // Format content/weight
                decimal dNet = _clsFun.Txt2Dec(txtArtNet != null ? txtArtNet.Text : "1");
                if (dNet <= 0m) dNet = 1m;

                // Format selling price
                decimal dPrv = _clsFun.Txt2Dec(txtLivPrv != null ? txtLivPrv.Text : "0");

                // Format cost price
                decimal dPrc = _clsFun.Txt2Dec(txtLiaPrc != null ? txtLiaPrc.Text : "0");

                string sFor = cmbLiaFor.SelectedValue != null ? cmbLiaFor.SelectedValue.ToString() : "";
                string sLis = cmbLivLis.SelectedValue != null ? cmbLivLis.SelectedValue.ToString() : "01";
                if (string.IsNullOrEmpty(sLis)) sLis = "01";

                // Codice articolo fornitore: se non inserito mantiene il codice articolo dell'inserimento
                string sArf = (txtArtArf != null && !string.IsNullOrEmpty(txtArtArf.Text.Trim())) ? txtArtArf.Text.Trim() : sNewArtCod;

                // Download / Save photo into C:\ApProject\Temp\Img\<sNewArtCod>.jpg
                string sImgDir = @"C:\ApProject\Temp\Img";
                if (!Directory.Exists(sImgDir))
                {
                    try { Directory.CreateDirectory(sImgDir); } catch { }
                }

                string sImgFilePath = Path.Combine(sImgDir, sNewArtCod + ".jpg");
                string sFinalImgDb = "";

                try
                {
                    if (picWebProduct != null && picWebProduct.Image != null)
                    {
                        using (Bitmap bmp = new Bitmap(picWebProduct.Image))
                        {
                            bmp.Save(sImgFilePath, System.Drawing.Imaging.ImageFormat.Jpeg);
                        }
                    }
                    else if (!string.IsNullOrEmpty(_strLastWebImageUrl))
                    {
                        System.Net.ServicePointManager.SecurityProtocol = (System.Net.SecurityProtocolType)3072 | (System.Net.SecurityProtocolType)768 | System.Net.SecurityProtocolType.Tls;
                        using (System.Net.WebClient wc = new System.Net.WebClient())
                        {
                            wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36");
                            wc.DownloadFile(_strLastWebImageUrl, sImgFilePath);
                        }
                    }

                    if (File.Exists(sImgFilePath))
                    {
                        sFinalImgDb = sImgFilePath;
                    }
                }
                catch (Exception ex)
                {
                    _clsFun.ErrorLog("frmAnaArtNuovo.SaveImg", ex.Message);
                }

                // 3. INSERT INTO AnaArticoli via SqlInsertRow (guarantees schema compatibility & default values)
                DataTable tabArt = _clsFun.FillTabSql("AnaArticoli", "SELECT * FROM AnaArticoli WHERE art_cod='9999999999999'", false, _strConSql);
                DataRow xArt = tabArt.NewRow();
                if (tabArt.Columns.Contains("art_cod")) xArt["art_cod"] = sNewArtCod;
                if (tabArt.Columns.Contains("art_des")) xArt["art_des"] = sDes;
                if (tabArt.Columns.Contains("art_deb")) xArt["art_deb"] = sDeb;
                if (tabArt.Columns.Contains("art_sta")) xArt["art_sta"] = "A";
                if (tabArt.Columns.Contains("art_iva")) xArt["art_iva"] = sIva;
                if (tabArt.Columns.Contains("art_eqp")) xArt["art_eqp"] = "";
                if (tabArt.Columns.Contains("art_umi")) xArt["art_umi"] = sUmi;
                if (tabArt.Columns.Contains("art_tgr")) xArt["art_tgr"] = sTgr;
                if (tabArt.Columns.Contains("art_pxc")) xArt["art_pxc"] = 0;
                if (tabArt.Columns.Contains("art_pne")) xArt["art_pne"] = dNet;
                if (tabArt.Columns.Contains("art_rep")) xArt["art_rep"] = sRep;
                if (tabArt.Columns.Contains("art_ec1")) xArt["art_ec1"] = sEc1;
                if (tabArt.Columns.Contains("art_ec2")) xArt["art_ec2"] = sEc2;
                if (tabArt.Columns.Contains("art_ec3")) xArt["art_ec3"] = sEc3;
                if (tabArt.Columns.Contains("art_ori")) xArt["art_ori"] = "";
                if (tabArt.Columns.Contains("art_cal")) xArt["art_cal"] = "";
                if (tabArt.Columns.Contains("art_reb")) xArt["art_reb"] = "";
                if (tabArt.Columns.Contains("art_cat")) xArt["art_cat"] = "";
                if (tabArt.Columns.Contains("art_tra")) xArt["art_tra"] = "";
                if (tabArt.Columns.Contains("art_gsc")) xArt["art_gsc"] = 0;
                if (tabArt.Columns.Contains("art_bil")) xArt["art_bil"] = isBilancia;
                if (tabArt.Columns.Contains("art_web")) xArt["art_web"] = false;
                if (tabArt.Columns.Contains("art_cel")) xArt["art_cel"] = false;
                if (tabArt.Columns.Contains("art_plu")) xArt["art_plu"] = "";
                if (tabArt.Columns.Contains("art_sfr")) xArt["art_sfr"] = 0;
                if (tabArt.Columns.Contains("art_tar")) xArt["art_tar"] = 0;
                if (tabArt.Columns.Contains("art_gia")) xArt["art_gia"] = 0;
                if (tabArt.Columns.Contains("art_eti")) xArt["art_eti"] = "";
                if (tabArt.Columns.Contains("art_mar")) xArt["art_mar"] = "";
                if (tabArt.Columns.Contains("art_tas")) xArt["art_tas"] = 0;
                if (tabArt.Columns.Contains("art_img")) xArt["art_img"] = sFinalImgDb;
                if (tabArt.Columns.Contains("art_bpz")) xArt["art_bpz"] = false;
                if (tabArt.Columns.Contains("art_dtm"))
                {
                    if (tabArt.Columns["art_dtm"].DataType == typeof(DateTime))
                        xArt["art_dtm"] = DateTime.Now;
                    else
                        xArt["art_dtm"] = DateTime.Now.ToString("dd/MM/yyyy");
                }
                if (tabArt.Columns.Contains("art_dti"))
                {
                    if (tabArt.Columns["art_dti"].DataType == typeof(DateTime))
                        xArt["art_dti"] = DateTime.Today;
                    else
                        xArt["art_dti"] = DateTime.Today.ToString("dd/MM/yyyy");
                }
                if (tabArt.Columns.Contains("art_cos")) xArt["art_cos"] = dPrc;
                if (tabArt.Columns.Contains("art_prv")) xArt["art_prv"] = dPrv;

                string sqlArt = _clsFun.SqlInsertRow("AnaArticoli", tabArt, xArt);
                bool okArt = _clsFun.SqlWrite(sqlArt, _strConSql);
                if (!okArt)
                {
                    _clsFun.ErrorLog("frmAnaArtNuovo.SalvaArticoloDb", "Fallita INSERT AnaArticoli: " + sqlArt);
                    MessageBox.Show("Errore durante il salvataggio dell'anagrafica articolo (Cod. " + sNewArtCod + ").\nOperazione annullata.", "ERRORE SALVATAGGIO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return "";
                }

                if (!string.IsNullOrEmpty(sFinalImgDb))
                {
                    try
                    {
                        string sqlImgUpd = "UPDATE AnaArticoli SET art_img='" + sFinalImgDb.Replace("'", "''") + "' WHERE art_cod='" + sNewArtCod + "'";
                        _clsFun.SqlWrite(sqlImgUpd, _strConSql);
                    }
                    catch { }
                }

                // 4. INSERT INTO AnaBarcode (pulizia preventiva di barcode storici o annullati per evitare violazione PK)
                try
                {
                    string sqlCleanBar = "DELETE FROM AnaBarcode WHERE ean_ean='" + sEan.Replace("'", "''") + "'";
                    _clsFun.SqlWrite(sqlCleanBar, _strConSql);
                }
                catch { }

                DataTable tabEan = _clsFun.FillTabSql("AnaBarcode", "SELECT * FROM AnaBarcode WHERE ean_art='9999999999999'", false, _strConSql);
                DataRow xEan = tabEan.NewRow();
                if (tabEan.Columns.Contains("ean_art")) xEan["ean_art"] = sNewArtCod;
                if (tabEan.Columns.Contains("ean_ean")) xEan["ean_ean"] = sEan;
                if (tabEan.Columns.Contains("ean_qta")) xEan["ean_qta"] = 1;
                if (tabEan.Columns.Contains("ean_prv")) xEan["ean_prv"] = 0m;
                if (tabEan.Columns.Contains("ean_dti")) xEan["ean_dti"] = DateTime.Today;
                if (tabEan.Columns.Contains("ean_dtm")) xEan["ean_dtm"] = DateTime.Today;
                if (tabEan.Columns.Contains("ean_ann")) xEan["ean_ann"] = false;
                if (tabEan.Columns.Contains("ean_bil")) xEan["ean_bil"] = isBilancia;
                if (tabEan.Columns.Contains("ean_ecp")) xEan["ean_ecp"] = false;
                string sqlBar = _clsFun.SqlInsertRow("AnaBarcode", tabEan, xEan);
                _clsFun.SqlWrite(sqlBar, _strConSql);

                // 5. INSERT / UPDATE GesLisVendita (Selling price listino)
                if (!string.IsNullOrEmpty(sLis))
                {
                    string sSqlCheckLiv = "SELECT COUNT(*) FROM GesLisVendita WHERE liv_art='" + sNewArtCod + "' AND liv_lis='" + sLis + "'";
                    DataTable tLivExist = _clsFun.FillTabSql("GesLisVendita", sSqlCheckLiv, false, _strConSql);
                    int nLivExist = 0;
                    if (tLivExist != null && tLivExist.Rows.Count > 0 && tLivExist.Rows[0][0] != DBNull.Value)
                    {
                        int.TryParse(tLivExist.Rows[0][0].ToString(), out nLivExist);
                    }

                    if (nLivExist > 0)
                    {
                        string sqlUpdLiv = "UPDATE GesLisVendita SET " +
                                           "liv_prv=" + dPrv.ToString(System.Globalization.CultureInfo.InvariantCulture) + ", " +
                                           "liv_dti='" + DateTime.Today.ToString("yyyyMMdd") + "', " +
                                           "liv_ann=0, " +
                                           "liv_sta='A' " +
                                           "WHERE liv_art='" + sNewArtCod + "' AND liv_lis='" + sLis + "'";
                        _clsFun.SqlWrite(sqlUpdLiv, _strConSql);
                    }
                    else
                    {
                        DataTable tabLiv = _clsFun.FillTabSql("GesLisVendita", "SELECT * FROM GesLisVendita WHERE 1=0", false, _strConSql);
                        DataRow xLiv = tabLiv.NewRow();
                        if (tabLiv.Columns.Contains("liv_art")) xLiv["liv_art"] = sNewArtCod;
                        if (tabLiv.Columns.Contains("liv_lis")) xLiv["liv_lis"] = sLis;
                        if (tabLiv.Columns.Contains("liv_prv")) xLiv["liv_prv"] = dPrv;
                        if (tabLiv.Columns.Contains("liv_dti")) xLiv["liv_dti"] = DateTime.Today;
                        if (tabLiv.Columns.Contains("liv_dtf")) xLiv["liv_dtf"] = _clsDef.DAYOUT;
                        if (tabLiv.Columns.Contains("liv_ann")) xLiv["liv_ann"] = false;
                        if (tabLiv.Columns.Contains("liv_sta")) xLiv["liv_sta"] = "A";
                        if (tabLiv.Columns.Contains("liv_day")) xLiv["liv_day"] = DateTime.Now;
                        string sqlLis = _clsFun.SqlInsertRow("GesLisVendita", tabLiv, xLiv);
                        _clsFun.SqlWrite(sqlLis, _strConSql);
                    }
                }

                // 6. FORTIFICAZIONE GesLisAcquisto (Cost price listino fornitore)
                if (string.IsNullOrEmpty(sFor) && dPrc > 0)
                {
                    try
                    {
                        DataTable tDefFor = _clsFun.FillTabSql("AnaFornitori", "SELECT TOP 1 for_cod FROM AnaFornitori WHERE for_ann=0 ORDER BY for_cod", false, _strConSql);
                        if (tDefFor != null && tDefFor.Rows.Count > 0 && tDefFor.Rows[0]["for_cod"] != DBNull.Value)
                        {
                            sFor = tDefFor.Rows[0]["for_cod"].ToString().Trim();
                        }
                    }
                    catch { }
                }

                if (!string.IsNullOrEmpty(sFor))
                {
                    string sSqlCheckLia = "SELECT COUNT(*) FROM GesLisAcquisto WHERE lia_art='" + sNewArtCod + "' AND lia_for='" + sFor + "'";
                    DataTable tLiaExist = _clsFun.FillTabSql("GesLisAcquisto", sSqlCheckLia, false, _strConSql);
                    int nLiaExist = 0;
                    if (tLiaExist != null && tLiaExist.Rows.Count > 0 && tLiaExist.Rows[0][0] != DBNull.Value)
                    {
                        int.TryParse(tLiaExist.Rows[0][0].ToString(), out nLiaExist);
                    }

                    if (nLiaExist > 0)
                    {
                        string sqlUpdLia = "UPDATE GesLisAcquisto SET " +
                                           "lia_tip='L', " +
                                           "lia_cos=" + dPrc.ToString(System.Globalization.CultureInfo.InvariantCulture) + ", " +
                                           "lia_prv=" + dPrv.ToString(System.Globalization.CultureInfo.InvariantCulture) + ", " +
                                           "lia_arf='" + sArf.Replace("'", "''") + "', " +
                                           "lia_dti='" + DateTime.Today.ToString("yyyyMMdd") + "', " +
                                           "lia_stf='A', " +
                                           "lia_ann=0, " +
                                           "lia_pxc=1, " +
                                           "lia_cxp=0, " +
                                           "lia_day=GETDATE() " +
                                           "WHERE lia_art='" + sNewArtCod + "' AND lia_for='" + sFor + "'";
                        _clsFun.SqlWrite(sqlUpdLia, _strConSql);
                    }
                    else
                    {
                        DataTable tabLia = _clsFun.FillTabSql("GesLisAcquisto", "SELECT * FROM GesLisAcquisto WHERE 1=0", false, _strConSql);
                        DataRow xLia = tabLia.NewRow();
                        if (tabLia.Columns.Contains("lia_art")) xLia["lia_art"] = sNewArtCod;
                        if (tabLia.Columns.Contains("lia_for")) xLia["lia_for"] = sFor;
                        if (tabLia.Columns.Contains("lia_arf")) xLia["lia_arf"] = sArf;
                        if (tabLia.Columns.Contains("lia_cos")) xLia["lia_cos"] = dPrc;
                        if (tabLia.Columns.Contains("lia_prv")) xLia["lia_prv"] = dPrv;
                        if (tabLia.Columns.Contains("lia_dti")) xLia["lia_dti"] = DateTime.Today;
                        if (tabLia.Columns.Contains("lia_dtf")) xLia["lia_dtf"] = _clsDef.DAYOUT;
                        if (tabLia.Columns.Contains("lia_ann")) xLia["lia_ann"] = false;
                        if (tabLia.Columns.Contains("lia_tip")) xLia["lia_tip"] = "L";
                        if (tabLia.Columns.Contains("lia_stf")) xLia["lia_stf"] = "A";
                        if (tabLia.Columns.Contains("lia_pxc")) xLia["lia_pxc"] = 1;
                        if (tabLia.Columns.Contains("lia_cxp")) xLia["lia_cxp"] = 0;
                        if (tabLia.Columns.Contains("lia_day")) xLia["lia_day"] = DateTime.Now;
                        string sqlFor = _clsFun.SqlInsertRow("GesLisAcquisto", tabLia, xLia);
                        _clsFun.SqlWrite(sqlFor, _strConSql);
                    }
                }

                // 7. Register Variazioni to send changes to ECR cash registers and POS scales
                _clsVar.Variazioni(sNewArtCod, sqlArt, "Inserimento articolo veloce", _clsDef.VARALL);

                txtArtCod.Text = sNewArtCod;
                return sNewArtCod;
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArtNuovo.SalvaArticoloDb", ex.Message);
                MessageBox.Show("Errore durante il salvataggio dell'articolo: " + ex.Message, "ERRORE SALVATAGGIO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "";
            }
        }

        private void EnsureEcr999()
        {
            try
            {
                string s1 = "SELECT tab_cod FROM TabEcrLv1 WHERE tab_cod='999'";
                DataTable t1 = _clsFun.FillTabSql("TabEcrLv1", s1, false, _strConSql);
                if (t1 == null || t1.Rows.Count == 0)
                {
                    _clsFun.SqlWrite("INSERT INTO TabEcrLv1 (tab_cod, tab_des) VALUES ('999', 'GENERICO')", _strConSql);
                }

                string s2 = "SELECT tab_cod FROM TabEcrLv2 WHERE tab_lv1='999' AND tab_cod='999'";
                DataTable t2 = _clsFun.FillTabSql("TabEcrLv2", s2, false, _strConSql);
                if (t2 == null || t2.Rows.Count == 0)
                {
                    _clsFun.SqlWrite("INSERT INTO TabEcrLv2 (tab_lv1, tab_cod, tab_des) VALUES ('999', '999', 'GENERICO')", _strConSql);
                }

                string s3 = "SELECT tab_cod FROM TabEcrLv3 WHERE tab_lv1='999' AND tab_lv2='999' AND tab_cod='999'";
                DataTable t3 = _clsFun.FillTabSql("TabEcrLv3", s3, false, _strConSql);
                if (t3 == null || t3.Rows.Count == 0)
                {
                    _clsFun.SqlWrite("INSERT INTO TabEcrLv3 (tab_lv1, tab_lv2, tab_cod, tab_des) VALUES ('999', '999', '999', 'GENERICO')", _strConSql);
                }
            }
            catch { }
        }

        private void SalvaVoci()
        {
            try
            {
                string s = "";
                string sFil = _strFilVoci;

                string sDir = Path.GetDirectoryName(sFil);
                if (!Directory.Exists(sDir))
                    Directory.CreateDirectory(sDir);

                using (StreamWriter sw = new StreamWriter(sFil, false))
                {
                    s += "MEM:" + (chkMem != null ? chkMem.Checked.ToString() : "False") + _clsDef.CRLF;
                    s += "AUT:" + (chkRicercaAuto != null ? chkRicercaAuto.Checked.ToString() : "True") + _clsDef.CRLF;
                    s += "IVA:" + (cmbArtIva != null && cmbArtIva.SelectedValue != null ? cmbArtIva.SelectedValue.ToString() : "") + _clsDef.CRLF;
                    s += "UMI:" + (cmbArtUmi != null && cmbArtUmi.SelectedValue != null ? cmbArtUmi.SelectedValue.ToString() : "") + _clsDef.CRLF;
                    s += "TGR:" + (cmbArtTgr != null && cmbArtTgr.SelectedValue != null ? cmbArtTgr.SelectedValue.ToString() : "") + _clsDef.CRLF;
                    s += "REP:" + (cmbArtRep != null && cmbArtRep.SelectedValue != null ? cmbArtRep.SelectedValue.ToString() : "") + _clsDef.CRLF;

                    if (cmbLiaFor != null && cmbLiaFor.SelectedValue != null && cmbLiaFor.SelectedValue.ToString() != "")
                        s += "FOR:" + cmbLiaFor.SelectedValue.ToString() + _clsDef.CRLF;

                    s += "LIV:" + (cmbLivLis != null && cmbLivLis.SelectedValue != null ? cmbLivLis.SelectedValue.ToString() : "") + _clsDef.CRLF;
                    s += "RIC:" + (txtRicarico != null ? txtRicarico.Text.Trim() : "0") + _clsDef.CRLF;

                    sw.Write(s + _clsDef.CRLF);
                    sw.Flush();
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaArtNuovo.SalvaVoci", ex.Message);
            }
        }
    }
}
