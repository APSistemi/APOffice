using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace APOffice
{
    public partial class frmUtyImportXls : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();
        clsVariazioni _clsVar = new clsVariazioni();

        private string _strConSql = "";
        private DataTable _tImport = null;
        private bool _bolCancel = false;

        public frmUtyImportXls()
        {
            InitializeComponent();
            _strConSql = _clsFun.ConSql("");

            // Proactive Migration: ensure MapSkip and FixVal exist in SQL Server
            try { _clsFun.SqlWrite("ALTER TABLE XlsMapTestata ADD MapSkip INT", _strConSql); } catch { }
            try { _clsFun.SqlWrite("ALTER TABLE XlsMapDettaglio ADD FixVal VARCHAR(255)", _strConSql); } catch { }

            new clsGesGraph().SetGraph(this, 0);

            // Wizard initialization
            tabWizard.ItemSize = new Size(0, 1);
            tabWizard.SizeMode = TabSizeMode.Fixed;

            dgvImport.DataBindingComplete += dgvImport_DataBindingComplete;
        }

        private void frmUtyImportXls_Load(object sender, EventArgs e)
        {
            try { new clsDbMigration().CheckDb(_strConSql); } catch { }
            FillMappings();
            SetGrid();
            UpdateWizardUI();
        }

        private void UpdateWizardUI()
        {
            int step = tabWizard.SelectedIndex + 1;
            lblStepInfo.Text = "Passaggio " + step + " di 4";

            btnBack.Enabled = (step > 1);
            btnNext.Enabled = true;
            btnNext.Visible = (step < 4);

            // Override contrast for title label (Fix for dark-on-dark issues)
            lblTitle.ForeColor = Color.FromArgb(44, 62, 80);

            if (step == 1) lblTitle.Text = "Importazione Dati Excel - Selezione File";
            if (step == 2)
            {
                lblTitle.Text = "Importazione Dati Excel - Mappatura";
                lblFileSelected.Text = "File in uso: " + Path.GetFileName(txtFilePath.Text);
            }
            if (step == 3) lblTitle.Text = "Importazione Dati Excel - Anteprima";
            if (step == 4)
            {
                lblTitle.Text = "Importazione Dati Excel - Elaborazione";
                lblSummary.Text = "Tutto pronto per l'importazione.\nVerifica i dati nell'anteprima e procedi.\n\nRighe da elaborare: " + (dgvImport.Rows.Count);
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (tabWizard.SelectedIndex == 0 && string.IsNullOrEmpty(txtFilePath.Text))
            {
                MessageBox.Show("Selezionare prima un file Excel.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (tabWizard.SelectedIndex == 1 && (cmbMapping.SelectedValue == null || cmbMapping.SelectedValue.ToString() == ""))
            {
                MessageBox.Show("Selezionare una mappatura.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (tabWizard.SelectedIndex == 1) // Moving to Preview
            {
                btnNext.Visible = false;
                btnNext.Enabled = false;
                btnBack.Enabled = false;
                Application.DoEvents();

                bool success = ProcessFile(txtFilePath.Text);
                if (!success)
                {
                    UpdateWizardUI();
                    return;
                }
            }

            if (tabWizard.SelectedIndex < tabWizard.TabCount - 1)
            {
                tabWizard.SelectedIndex++;
                UpdateWizardUI();
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            if (tabWizard.SelectedIndex > 0)
            {
                tabWizard.SelectedIndex--;
                UpdateWizardUI();
            }
        }

        private void FillMappings()
        {
            try
            {
                string s = "SELECT MapId as trk_cod, MapDes as trk_des FROM XlsMapTestata ORDER BY MapDes";
                DataTable t = _clsFun.FillTabSql("XlsMapTestata", s, false, _strConSql);

                if (t != null)
                {
                    if (t.Columns.Contains("trk_des")) t.Columns["trk_des"].MaxLength = -1;
                    if (t.Columns.Contains("trk_cod")) t.Columns["trk_cod"].MaxLength = -1;

                    DataRow x = t.NewRow();
                    x["trk_cod"] = "";
                    x["trk_des"] = "  Seleziona Mappatura...";
                    t.Rows.InsertAt(x, 0);

                    cmbMapping.ValueMember = "trk_cod";
                    cmbMapping.DisplayMember = "trk_des";
                    cmbMapping.DataSource = t;
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "FillMappings");
            }
        }

        private void btnNewMapping_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFilePath.Text))
            {
                MessageBox.Show("Selezionare prima un file Excel.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrEmpty(txtFilePath.Text)) return;
            frmUtyImportXlsMap f = new frmUtyImportXlsMap(txtFilePath.Text, _strConSql);
            if (f.ShowDialog() == DialogResult.OK)
            {
                FillMappings();
                if (!string.IsNullOrEmpty(f._strSelectedMap))
                    cmbMapping.SelectedValue = f._strSelectedMap;
            }
        }

        private void btnEditMapping_Click(object sender, EventArgs e)
        {
            if (cmbMapping.SelectedValue == null || cmbMapping.SelectedValue.ToString() == "")
            {
                MessageBox.Show("Selezionare una mappatura da modificare.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string sId = cmbMapping.SelectedValue.ToString();
            frmUtyImportXlsMap f = new frmUtyImportXlsMap(txtFilePath.Text, _strConSql, sId);
            if (f.ShowDialog() == DialogResult.OK)
            {
                FillMappings();
                cmbMapping.SelectedValue = sId;
            }
        }

        private void SetGrid()
        {
            dgvImport.AutoGenerateColumns = false;
            dgvImport.AllowUserToAddRows = false;
            dgvImport.ReadOnly = true;

            dgvImport.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "art_cod", Name = "art_cod", HeaderText = "Codice", Width = 80 });
            dgvImport.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "art_ean", Name = "art_ean", HeaderText = "Barcode", Width = 115 });
            dgvImport.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "plu", Name = "plu", HeaderText = "PLU", Width = 55, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
            dgvImport.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "reb", Name = "reb", HeaderText = "Rep. Bil.", Width = 65, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
            dgvImport.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "art_des", Name = "art_des", HeaderText = "Descrizione", Width = 250 });
            dgvImport.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "art_for", Name = "art_for", HeaderText = "Fornitore", Width = 80 });
            dgvImport.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "art_arf", Name = "art_arf", HeaderText = "Cod. Art. Forn.", Width = 100 });
            dgvImport.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "old_cos", Name = "old_cos", HeaderText = "Costo Att.", Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvImport.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "new_cos", Name = "new_cos", HeaderText = "Costo Nuo.", Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvImport.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "old_prv", Name = "old_prv", HeaderText = "Prezzo Att.", Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvImport.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "new_prv", Name = "new_prv", HeaderText = "Prezzo Nuo.", Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvImport.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "diff_cos", Name = "diff_cos", HeaderText = "Diff Costo", Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } });
            dgvImport.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "diff_prp", Name = "diff_prp", HeaderText = "Diff. %", Width = 60, DefaultCellStyle = new DataGridViewCellStyle { Format = "N1", Alignment = DataGridViewContentAlignment.MiddleRight } });
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            PerformExit();
        }

        private void PerformExit()
        {
            if (progressBar1.Visible)
            {
                if (MessageBox.Show("Procedura in corso. Interrompere l'elaborazione?", "Attenzione", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    _bolCancel = true;
                }
                return;
            }
            this.Close();
        }

        private string Format3Digits(string sInput, string sDefault = "000")
        {
            if (string.IsNullOrEmpty(sInput)) return sDefault;
            string sClean = sInput.Trim();
            if (int.TryParse(sClean, out int val))
                return val.ToString("000");
            return sClean.PadLeft(3, '0');
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "File Excel (*.xls;*.xlsx)|*.xls;*.xlsx|File CSV (*.csv)|*.csv";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = ofd.FileName;
                lblFileStatus.Text = "File caricato: " + Path.GetFileName(ofd.FileName);
                lblFileStatus.ForeColor = Color.Green;

                // Avanzamento automatico al tab della mappatura
                if (tabWizard.SelectedIndex < tabWizard.TabCount - 1)
                {
                    tabWizard.SelectedIndex++;
                    UpdateWizardUI();
                }
            }
        }

        private DataRow GetArticleInfo(string sArt)
        {
            try
            {
                string sSql = "SELECT A.art_cod, A.art_des, A.art_cos, " +
                              "(SELECT TOP 1 V.liv_prv FROM GesLisVendita V WHERE V.liv_art = A.art_cod AND V.liv_ann = 0) AS art_prv " +
                              "FROM AnaArticoli A WHERE A.art_cod = '" + sArt.Replace("'", "''") + "'";
                DataTable t = _clsFun.FillTabSql("ArtInfo", sSql, false, _strConSql);
                if (t != null && t.Rows.Count > 0) return t.Rows[0];
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "GetArticleInfo");
            }
            return null;
        }

        private bool ProcessFile(string strPath)
        {
            if (cmbMapping.SelectedValue == null || cmbMapping.SelectedValue.ToString() == "")
            {
                MessageBox.Show("Selezionare una mappatura prima di procedere.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            try
            {
                _bolCancel = false;
                progressBar1.Visible = true;
                string sId = cmbMapping.SelectedValue.ToString();
                DataTable tMap = _clsFun.FillTabSql("Map", "SELECT * FROM XlsMapDettaglio WHERE MapId='" + sId + "'", false, _strConSql);
                DataTable tTes = _clsFun.FillTabSql("Tes", "SELECT MapSkip, MapFor FROM XlsMapTestata WHERE MapId='" + sId + "'", false, _strConSql);
                int iSkip = (tTes.Rows.Count > 0) ? Convert.ToInt32(tTes.Rows[0]["MapSkip"]) : 0;
                string sDefFor = (tTes.Rows.Count > 0) ? tTes.Rows[0]["MapFor"].ToString() : "";

                DataTable tXls = ReadExcel(strPath);

                if (tXls == null) return false;

                lblFileStatus.Text = "Caricamento indici di ricerca DB... attendere.";
                Application.DoEvents();

                // 1. Pre-cache EAN -> art_cod from AnaBarcode
                Dictionary<string, string> dictEanDb = new Dictionary<string, string>();
                DataTable tAllEan = _clsFun.FillTabSql("AllEan", "SELECT ean_ean, ean_art FROM AnaBarcode WHERE ean_ann=0", false, _strConSql);
                if (tAllEan != null)
                {
                    foreach (DataRow rEan in tAllEan.Rows)
                    {
                        string kEan = rEan["ean_ean"].ToString().Trim();
                        if (!string.IsNullOrEmpty(kEan) && !dictEanDb.ContainsKey(kEan))
                            dictEanDb[kEan] = rEan["ean_art"].ToString().Trim();
                    }
                }

                // 2. Pre-cache AnaArticoli (art_cod, art_des, art_cos, art_reb, art_for, liv_prv)
                HashSet<string> setCodDb = new HashSet<string>();
                Dictionary<string, string> dictDesDb = new Dictionary<string, string>();
                Dictionary<string, string> dictArtDesInfo = new Dictionary<string, string>();
                Dictionary<string, decimal> dictArtCosInfo = new Dictionary<string, decimal>();
                Dictionary<string, decimal> dictArtPrvInfo = new Dictionary<string, decimal>();
                Dictionary<string, string> dictArtRebDb = new Dictionary<string, string>();
                Dictionary<string, string> dictArtForDb = new Dictionary<string, string>();

                DataTable tAllArt = _clsFun.FillTabSql("AllArt", "SELECT a.art_cod, a.art_des, a.art_cos, a.art_reb, v.liv_prv, l.lia_for FROM AnaArticoli a LEFT JOIN GesLisVendita v ON a.art_cod = v.liv_art AND v.liv_ann=0 LEFT JOIN GesLisAcquisto l ON a.art_cod = l.lia_art AND l.lia_ann=0", false, _strConSql);
                if (tAllArt != null)
                {
                    foreach (DataRow rArt in tAllArt.Rows)
                    {
                        string sCod = rArt["art_cod"].ToString().Trim();
                        string sDes = rArt["art_des"].ToString().Trim().ToUpper();
                        if (!string.IsNullOrEmpty(sCod))
                        {
                            setCodDb.Add(sCod);
                            if (!string.IsNullOrEmpty(sDes) && sDes.Length >= 3 && !dictDesDb.ContainsKey(sDes))
                                dictDesDb[sDes] = sCod;

                            if (!dictArtDesInfo.ContainsKey(sCod))
                            {
                                dictArtDesInfo[sCod] = rArt["art_des"].ToString();
                                dictArtCosInfo[sCod] = (rArt["art_cos"] != DBNull.Value) ? _clsFun.Txt2Dec(rArt["art_cos"].ToString()) : 0m;
                                dictArtPrvInfo[sCod] = (rArt["liv_prv"] != DBNull.Value) ? _clsFun.Txt2Dec(rArt["liv_prv"].ToString()) : 0m;
                                dictArtRebDb[sCod] = (rArt.Table.Columns.Contains("art_reb") && rArt["art_reb"] != DBNull.Value) ? rArt["art_reb"].ToString().Trim() : "";
                                dictArtForDb[sCod] = (rArt.Table.Columns.Contains("lia_for") && rArt["lia_for"] != DBNull.Value) ? rArt["lia_for"].ToString().Trim() : "";
                            }
                        }
                    }
                }

                // 3. Pre-cache GesLisAcquisto (lia_for + '|' + lia_arf -> lia_art)
                Dictionary<string, string> dictArfDb = new Dictionary<string, string>();
                DataTable tAllLia = _clsFun.FillTabSql("AllLia", "SELECT lia_for, lia_arf, lia_art FROM GesLisAcquisto WHERE lia_ann=0 ORDER BY lia_dti DESC", false, _strConSql);
                if (tAllLia != null)
                {
                    foreach (DataRow rLia in tAllLia.Rows)
                    {
                        string sFor = rLia["lia_for"].ToString().Trim();
                        string sArf = rLia["lia_arf"].ToString().Trim();
                        string sArt = rLia["lia_art"].ToString().Trim();
                        if (!string.IsNullOrEmpty(sArf))
                        {
                            string k1 = sFor + "|" + sArf;
                            string k2 = "|" + sArf;
                            if (!dictArfDb.ContainsKey(k1)) dictArfDb[k1] = sArt;
                            if (!dictArfDb.ContainsKey(k2)) dictArfDb[k2] = sArt;
                        }
                    }
                }
                
                // 4. Pre-cache Scale Articles (art_plu + '|' + art_reb -> art_cod, and art_plu -> art_cod)
                Dictionary<string, string> dictPluDb = new Dictionary<string, string>();
                DataTable tAllPlu = _clsFun.FillTabSql("AllPlu", "SELECT art_cod, art_plu, art_reb FROM AnaArticoli WHERE art_sta='A' AND art_bil=1 AND art_plu IS NOT NULL AND art_plu <> '' ORDER BY art_dti DESC", false, _strConSql);
                if (tAllPlu != null)
                {
                    foreach (DataRow rPlu in tAllPlu.Rows)
                    {
                        string sPluVal = rPlu["art_plu"].ToString().Trim();
                        string sRebVal = rPlu["art_reb"] != null ? rPlu["art_reb"].ToString().Trim() : "";
                        string sArtVal = rPlu["art_cod"].ToString().Trim();
                        if (!string.IsNullOrEmpty(sPluVal))
                        {
                            if (int.TryParse(sPluVal, out int iPlu))
                            {
                                string kPluReb = iPlu.ToString() + "|" + sRebVal;
                                string kPluOnly = iPlu.ToString();
                                if (!dictPluDb.ContainsKey(kPluReb)) dictPluDb[kPluReb] = sArtVal;
                                if (!dictPluDb.ContainsKey(kPluOnly)) dictPluDb[kPluOnly] = sArtVal;
                            }
                            string rawKey = sPluVal + "|" + sRebVal;
                            if (!dictPluDb.ContainsKey(rawKey)) dictPluDb[rawKey] = sArtVal;
                            if (!dictPluDb.ContainsKey(sPluVal)) dictPluDb[sPluVal] = sArtVal;
                        }
                    }
                }

                _tImport = CreateTargetTable();

                int totalRows = tXls.Rows.Count;
                progressBar1.Value = 0;
                progressBar1.Maximum = totalRows * 2;

                Dictionary<string, string> dictEanBatch = new Dictionary<string, string>();
                Dictionary<string, string> dictArfBatch = new Dictionary<string, string>();
                Dictionary<string, string> dictCodBatch = new Dictionary<string, string>();
                Dictionary<string, string> dictDesBatch = new Dictionary<string, string>();

                int currentRowPass1 = 0;

                // PASS 1: Scan file to pre-resolve exact duplicates (by exact supplier code, mapped code or exact description)
                foreach (DataRow rXls in tXls.Rows)
                {
                    currentRowPass1++;
                    if (currentRowPass1 % 50 == 0)
                    {
                        progressBar1.Value = currentRowPass1;
                        lblFileStatus.Text = "Analisi e allineamento righe: " + currentRowPass1 + " di " + totalRows;
                        Application.DoEvents();
                        if (_bolCancel) break;
                    }

                    if (currentRowPass1 <= iSkip) continue;

                    string sEan = GetMappedValue(rXls, tMap, "art_ean").Trim();
                    string sDes = GetMappedValue(rXls, tMap, "art_des").Trim();
                    string sDesKey = sDes.ToUpper();
                    string sMappedCod = GetMappedValue(rXls, tMap, "art_cod").Trim();
                    if (IsPlaceholderCode(sMappedCod)) sMappedCod = "";

                    string sFor = GetMappedValue(rXls, tMap, "art_for").Trim();
                    string sRowFor = (sFor != "") ? sFor : sDefFor;
                    string sArf = GetMappedValue(rXls, tMap, "art_cod_for").Trim();
                    string sArfKey = sRowFor + "|" + sArf;

                    string sArt = "";

                    // Standard commercial EAN
                    if (!string.IsNullOrEmpty(sEan) && !sEan.StartsWith("2") && dictEanDb.ContainsKey(sEan))
                    {
                        sArt = dictEanDb[sEan];
                    }

                    // Explicit internal code
                    if (string.IsNullOrEmpty(sArt) && !string.IsNullOrEmpty(sMappedCod) && setCodDb.Contains(sMappedCod))
                    {
                        sArt = sMappedCod;
                    }

                    // Supplier Article Code
                    if (string.IsNullOrEmpty(sArt) && !string.IsNullOrEmpty(sArf))
                    {
                        if (dictArfDb.ContainsKey(sArfKey)) sArt = dictArfDb[sArfKey];
                        else if (dictArfDb.ContainsKey("|" + sArf)) sArt = dictArfDb["|" + sArf];
                    }

                    // Exact Description
                    if (string.IsNullOrEmpty(sArt) && !string.IsNullOrEmpty(sDesKey) && sDesKey.Length >= 3 && dictDesDb.ContainsKey(sDesKey))
                    {
                        sArt = dictDesDb[sDesKey];
                    }

                    if (!string.IsNullOrEmpty(sArt))
                    {
                        if (!string.IsNullOrEmpty(sEan) && !sEan.StartsWith("2") && !dictEanBatch.ContainsKey(sEan)) dictEanBatch[sEan] = sArt;
                        if (!string.IsNullOrEmpty(sMappedCod) && !dictCodBatch.ContainsKey(sMappedCod)) dictCodBatch[sMappedCod] = sArt;
                        if (!string.IsNullOrEmpty(sArf) && !dictArfBatch.ContainsKey(sArfKey)) dictArfBatch[sArfKey] = sArt;
                        if (!string.IsNullOrEmpty(sDesKey) && sDesKey.Length >= 3 && !dictDesBatch.ContainsKey(sDesKey)) dictDesBatch[sDesKey] = sArt;
                    }
                }

                int currentRowPass2 = 0;

                // PASS 2: Populate _tImport rows in memory
                foreach (DataRow rXls in tXls.Rows)
                {
                    currentRowPass2++;

                    if (currentRowPass2 % 50 == 0)
                    {
                        progressBar1.Value = totalRows + currentRowPass2;
                        lblFileStatus.Text = "Preparazione anteprima riga: " + currentRowPass2 + " di " + totalRows;
                        Application.DoEvents();
                        if (_bolCancel) break;
                    }

                    if (currentRowPass2 <= iSkip) continue;

                    string sEan = GetMappedValue(rXls, tMap, "art_ean").Trim();
                    string sDes = GetMappedValue(rXls, tMap, "art_des").Trim();
                    string sMappedCod = GetMappedValue(rXls, tMap, "art_cod").Trim();
                    string sArf = GetMappedValue(rXls, tMap, "art_cod_for").Trim();

                    // Non ignorare le righe se almeno un campo identificativo (EAN, Descrizione, Codice, Cod. Fornitore) è presente
                    if (string.IsNullOrEmpty(sEan) && string.IsNullOrEmpty(sDes) && string.IsNullOrEmpty(sMappedCod) && string.IsNullOrEmpty(sArf))
                        continue;

                    DataRow r = _tImport.NewRow();
                    r["art_ean"] = sEan;
                    r["art_des"] = sDes;
                    string sDesKey = sDes.ToUpper();

                    r["new_cos"] = _clsFun.Txt2Dec(GetMappedValue(rXls, tMap, "art_cos"));
                    r["new_prv"] = _clsFun.Txt2Dec(GetMappedValue(rXls, tMap, "art_pre"));

                    r["iva"] = ResolveIvaCode(GetMappedValue(rXls, tMap, "art_iva"));
                    string sFor = GetMappedValue(rXls, tMap, "art_for").Trim();
                    string sRowFor = (sFor != "") ? sFor : sDefFor;
                    r["art_for"] = sRowFor;
                    r["art_arf"] = sArf;
                    string sArfKey = sRowFor + "|" + sArf;

                    r["pxc"] = GetMappedValue(rXls, tMap, "art_pxc");
                    r["gr"] = GetMappedValue(rXls, tMap, "art_gr");
                    r["peso"] = GetMappedValue(rXls, tMap, "art_ps");
                    r["reparto"] = ResolveRepCode(GetMappedValue(rXls, tMap, "art_rep"));
                    string sRawEc1 = GetMappedValue(rXls, tMap, "art_ec1");
                    if (string.IsNullOrEmpty(sRawEc1)) sRawEc1 = GetMappedValue(rXls, tMap, "art_fam");
                    string sEc1 = ResolveEc1Code(sRawEc1);
                    string sEc2 = ResolveEc2Code(GetMappedValue(rXls, tMap, "art_ec2"), sEc1);
                    string sEc3 = ResolveEc3Code(GetMappedValue(rXls, tMap, "art_ec3"), sEc1, sEc2);
                    r["art_ec1"] = sEc1;
                    r["art_ec2"] = sEc2;
                    r["art_ec3"] = sEc3;
                    r["famiglia"] = sEc1;
                    r["eti"] = GetMappedValue(rXls, tMap, "art_eti");
                    // Parametri Bilancia
                    string sPlu = ResolvePlu(GetMappedValue(rXls, tMap, "art_plu"));
                    if (string.IsNullOrEmpty(sPlu))
                    {
                        if (int.TryParse(sArf, out int iArfPlu) && iArfPlu > 0 && iArfPlu <= 99999)
                        {
                            sPlu = iArfPlu.ToString("D4");
                        }
                    }
                    r["plu"] = sPlu;
                    r["reb"] = ResolveRebCode(GetMappedValue(rXls, tMap, "art_reb"));
                    r["ori"] = ResolveOriCode(GetMappedValue(rXls, tMap, "art_ori"));
                    r["cal"] = ResolveCalCode(GetMappedValue(rXls, tMap, "art_cal"));
                    r["cat_bil"] = ResolveCatBilCode(GetMappedValue(rXls, tMap, "art_cat_bil"));
                    r["tas"] = ResolveTas(GetMappedValue(rXls, tMap, "art_tas"));
                    if (string.IsNullOrEmpty(r["tas"].ToString()) && !string.IsNullOrEmpty(sPlu) && int.TryParse(sPlu, out int iPluTas))
                    {
                        r["tas"] = iPluTas.ToString();
                    }
                    r["tra"] = ResolveTraCode(GetMappedValue(rXls, tMap, "art_tra"));
                    r["gsc"] = GetMappedValue(rXls, tMap, "art_gsc");
                    r["bpz"] = ResolveBpz(GetMappedValue(rXls, tMap, "art_bpz"));
                    r["tar"] = GetMappedValue(rXls, tMap, "art_tar");

                    string sBilStr = GetMappedValue(rXls, tMap, "art_bil").Trim();
                    string sRebStr = r["reb"].ToString().Trim();
                    bool bIsBil = (!string.IsNullOrEmpty(sPlu)) ||
                                  (!string.IsNullOrEmpty(sRebStr)) ||
                                  (!string.IsNullOrEmpty(sEan) && sEan.StartsWith("2")) ||
                                  sBilStr == "1" || sBilStr.Equals("S", StringComparison.OrdinalIgnoreCase) ||
                                  sBilStr.Equals("SI", StringComparison.OrdinalIgnoreCase) ||
                                  sBilStr.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
                                  sBilStr.Equals("YES", StringComparison.OrdinalIgnoreCase);
                    r["bil"] = bIsBil ? "1" : "0";

                    // Generazione Barcode Bilancia EAN-13 (prefisso 29 o 2) se abilitato e mancante o non bilancia
                    string sGenEanBilOpt = GetConfigOption(tMap, "gen_ean_bil", "0");
                    bool bGenEanBil = (sGenEanBilOpt == "1" || sGenEanBilOpt.Equals("true", StringComparison.OrdinalIgnoreCase));
                    string sPrfEanBil = GetConfigOption(tMap, "prf_ean_bil", "2");

                    string sGenEan = "";
                    if (bGenEanBil && bIsBil && !string.IsNullOrEmpty(sPlu))
                    {
                        if (string.IsNullOrEmpty(sEan) || !sEan.StartsWith("2"))
                        {
                            sGenEan = GenerateScaleEan13(sPlu, sPrfEanBil);
                        }
                    }
                    r["ean_bil_gen"] = sGenEan;

                    if (!string.IsNullOrEmpty(sGenEan))
                    {
                        if (string.IsNullOrEmpty(sEan) || !sEan.StartsWith("2"))
                        {
                            sEan = sGenEan;
                            r["art_ean"] = sGenEan;
                        }
                    }

                    // Risoluzione Articolo Intelligente e Sicura:
                    // 1. Explicit internal code
                    string sArt = "";
                    if (!string.IsNullOrEmpty(sMappedCod) && !IsPlaceholderCode(sMappedCod) && setCodDb.Contains(sMappedCod))
                    {
                        sArt = sMappedCod;
                    }

                    // 2. Commercial EAN standard (escluso prefisso bilancia 2/29)
                    if (string.IsNullOrEmpty(sArt) && !string.IsNullOrEmpty(sEan) && !sEan.StartsWith("2"))
                    {
                        if (dictEanBatch.ContainsKey(sEan)) sArt = dictEanBatch[sEan];
                        else if (dictEanDb.ContainsKey(sEan)) sArt = dictEanDb[sEan];
                    }

                    // 3. Codice Articolo Fornitore (lia_arf) per quel fornitore
                    if (string.IsNullOrEmpty(sArt) && !string.IsNullOrEmpty(sArf))
                    {
                        if (dictArfBatch.ContainsKey(sArfKey)) sArt = dictArfBatch[sArfKey];
                        else if (dictArfDb.ContainsKey(sArfKey)) sArt = dictArfDb[sArfKey];
                    }

                    // 4. Descrizione Esatta
                    if (string.IsNullOrEmpty(sArt) && !string.IsNullOrEmpty(sDesKey) && sDesKey.Length >= 3)
                    {
                        if (dictDesBatch.ContainsKey(sDesKey)) sArt = dictDesBatch[sDesKey];
                        else if (dictDesDb.ContainsKey(sDesKey)) sArt = dictDesDb[sDesKey];
                    }

                    // 5. Descrizione Semantica / Somiglianza (Fuzzy matching)
                    if (string.IsNullOrEmpty(sArt) && !string.IsNullOrEmpty(sDesKey) && sDesKey.Length >= 3)
                    {
                        foreach (var kvp in dictArtDesInfo)
                        {
                            if (AreDescriptionsSimilar(kvp.Value, sDes))
                            {
                                sArt = kvp.Key;
                                break;
                            }
                        }
                    }

                    // 6. Scale PLU / Scale EAN Candidate Check con validazione di coerenza (anti-falsi positivi)
                    if (string.IsNullOrEmpty(sArt) && bIsBil && (!string.IsNullOrEmpty(sPlu) || !string.IsNullOrEmpty(sGenEan)))
                    {
                        string cand = "";
                        int.TryParse(sPlu, out int iPlu);
                        string k1 = (iPlu > 0 ? iPlu.ToString() : sPlu) + "|" + sRebStr;
                        string k2 = (iPlu > 0 ? iPlu.ToString() : sPlu);

                        if (!string.IsNullOrEmpty(sRebStr) && dictPluDb.ContainsKey(k1)) cand = dictPluDb[k1];
                        else if (dictPluDb.ContainsKey(k2)) cand = dictPluDb[k2];
                        else if (!string.IsNullOrEmpty(sGenEan) && dictEanDb.ContainsKey(sGenEan)) cand = dictEanDb[sGenEan];

                        if (!string.IsNullOrEmpty(cand) && dictArtDesInfo.ContainsKey(cand))
                        {
                            string candDes = dictArtDesInfo[cand];
                            string candReb = dictArtRebDb.ContainsKey(cand) ? dictArtRebDb[cand] : "";
                            string candFor = dictArtForDb.ContainsKey(cand) ? dictArtForDb[cand] : "";

                            // Accetta il candidato solo se la descrizione è congruente, o se appartiene allo stesso reparto bilancia e fornitore
                            if (AreDescriptionsSimilar(candDes, sDes) || (candReb == sRebStr && !string.IsNullOrEmpty(sRebStr) && (candFor == sRowFor || string.IsNullOrEmpty(candFor))))
                            {
                                sArt = cand;
                            }
                        }
                    }

                    // Indicizza l'articolo risolto per righe successive identiche (solo se stessa descrizione o fornitore)
                    if (!string.IsNullOrEmpty(sArt))
                    {
                        if (!string.IsNullOrEmpty(sEan) && !sEan.StartsWith("2") && !dictEanBatch.ContainsKey(sEan)) dictEanBatch[sEan] = sArt;
                        if (!string.IsNullOrEmpty(sMappedCod) && !IsPlaceholderCode(sMappedCod) && !dictCodBatch.ContainsKey(sMappedCod)) dictCodBatch[sMappedCod] = sArt;
                        if (!string.IsNullOrEmpty(sArf) && !dictArfBatch.ContainsKey(sArfKey)) dictArfBatch[sArfKey] = sArt;
                        if (!string.IsNullOrEmpty(sDesKey) && sDesKey.Length >= 3 && !dictDesBatch.ContainsKey(sDesKey)) dictDesBatch[sDesKey] = sArt;
                    }

                    decimal dNewCos = Convert.ToDecimal(r["new_cos"]);
                    decimal dNewPrv = Convert.ToDecimal(r["new_prv"]);

                    if (!string.IsNullOrEmpty(sArt))
                    {
                        r["art_cod"] = sArt;
                        r["is_new"] = false;
                        if (dictArtDesInfo.ContainsKey(sArt))
                        {
                            r["old_des"] = dictArtDesInfo[sArt];
                            r["old_cos"] = dictArtCosInfo[sArt];
                            r["old_prv"] = dictArtPrvInfo[sArt];
                        }
                        else
                        {
                            r["old_des"] = "";
                            r["old_cos"] = 0m;
                            r["old_prv"] = 0m;
                        }

                        decimal dOldCos = Convert.ToDecimal(r["old_cos"]);
                        decimal dOldPrv = Convert.ToDecimal(r["old_prv"]);

                        r["diff_cos"] = dNewCos - dOldCos;
                        if (dOldPrv > 0)
                            r["diff_prp"] = Math.Round(((dNewPrv - dOldPrv) / dOldPrv) * 100, 2);
                    }
                    else
                    {
                        // Nuova referenza: se art_cod non è mappato, imposta [NUOVO]
                        string sCodeExcel = (!string.IsNullOrEmpty(sMappedCod) && !IsPlaceholderCode(sMappedCod)) ? sMappedCod : "[NUOVO]";
                        r["art_cod"] = sCodeExcel;
                        r["is_new"] = true;
                        r["old_cos"] = 0m;
                        r["old_prv"] = 0m;
                        r["old_des"] = "";
                        r["diff_cos"] = dNewCos;
                    }

                    _tImport.Rows.Add(r);
                }

                if (_bolCancel)
                {
                    progressBar1.Visible = false;
                    lblFileStatus.Text = "Elaborazione annullata dall'utente.";
                    return false;
                }

                lblFileStatus.Text = "Completato. Righe elaborate: " + _tImport.Rows.Count;
                BindingSource bs = new BindingSource();
                bs.DataSource = _tImport;
                dgvImport.DataSource = bs;
                FormatGrid();
                progressBar1.Visible = false;
                return true;
            }
            catch (Exception ex)
            {
                progressBar1.Visible = false;
                MessageBox.Show("Errore durante l'elaborazione del file: " + ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void FormatGrid()
        {
            if (dgvImport.Columns.Count == 0) return;
            if (dgvImport.Columns.Contains("art_cod")) dgvImport.Columns["art_cod"].HeaderText = "Codice";
            if (dgvImport.Columns.Contains("art_ean")) dgvImport.Columns["art_ean"].HeaderText = "Barcode";
            if (dgvImport.Columns.Contains("plu")) { dgvImport.Columns["plu"].HeaderText = "PLU"; dgvImport.Columns["plu"].Width = 55; dgvImport.Columns["plu"].Visible = true; }
            if (dgvImport.Columns.Contains("reb")) { dgvImport.Columns["reb"].HeaderText = "Rep. Bil."; dgvImport.Columns["reb"].Width = 65; dgvImport.Columns["reb"].Visible = true; }
            if (dgvImport.Columns.Contains("art_des")) dgvImport.Columns["art_des"].HeaderText = "Descrizione";
            if (dgvImport.Columns.Contains("art_for")) dgvImport.Columns["art_for"].HeaderText = "Fornitore";
            if (dgvImport.Columns.Contains("art_arf")) dgvImport.Columns["art_arf"].HeaderText = "Cod. Art. Forn.";
            if (dgvImport.Columns.Contains("old_cos")) dgvImport.Columns["old_cos"].HeaderText = "Costo Att.";
            if (dgvImport.Columns.Contains("new_cos")) dgvImport.Columns["new_cos"].HeaderText = "Costo Nuo.";
            if (dgvImport.Columns.Contains("old_prv")) dgvImport.Columns["old_prv"].HeaderText = "Prezzo Att.";
            if (dgvImport.Columns.Contains("new_prv")) dgvImport.Columns["new_prv"].HeaderText = "Prezzo Nuo.";
            if (dgvImport.Columns.Contains("diff_cos")) dgvImport.Columns["diff_cos"].HeaderText = "Diff Costo";
            if (dgvImport.Columns.Contains("diff_prp")) dgvImport.Columns["diff_prp"].HeaderText = "Diff. %";

            if (dgvImport.Columns.Contains("old_cos")) dgvImport.Columns["old_cos"].DefaultCellStyle.Format = "N2";
            if (dgvImport.Columns.Contains("new_cos")) dgvImport.Columns["new_cos"].DefaultCellStyle.Format = "N2";
            if (dgvImport.Columns.Contains("old_prv")) dgvImport.Columns["old_prv"].DefaultCellStyle.Format = "N2";
            if (dgvImport.Columns.Contains("new_prv")) dgvImport.Columns["new_prv"].DefaultCellStyle.Format = "N2";
            if (dgvImport.Columns.Contains("diff_cos")) dgvImport.Columns["diff_cos"].DefaultCellStyle.Format = "N2";
            if (dgvImport.Columns.Contains("diff_prp")) dgvImport.Columns["diff_prp"].DefaultCellStyle.Format = "N2";

            if (dgvImport.Columns.Contains("art_des")) dgvImport.Columns["art_des"].Width = 250;

            // Nascondi colonne tecniche non necessarie in anteprima
            if (dgvImport.Columns.Contains("iva")) dgvImport.Columns["iva"].Visible = false;
            if (dgvImport.Columns.Contains("reparto")) { dgvImport.Columns["reparto"].HeaderText = "Reparto"; dgvImport.Columns["reparto"].Width = 60; dgvImport.Columns["reparto"].Visible = true; }
            if (dgvImport.Columns.Contains("art_ec1")) { dgvImport.Columns["art_ec1"].HeaderText = "Merc. 1"; dgvImport.Columns["art_ec1"].Width = 60; dgvImport.Columns["art_ec1"].Visible = true; }
            if (dgvImport.Columns.Contains("art_ec2")) { dgvImport.Columns["art_ec2"].HeaderText = "Merc. 2"; dgvImport.Columns["art_ec2"].Width = 60; dgvImport.Columns["art_ec2"].Visible = true; }
            if (dgvImport.Columns.Contains("art_ec3")) { dgvImport.Columns["art_ec3"].HeaderText = "Merc. 3"; dgvImport.Columns["art_ec3"].Width = 60; dgvImport.Columns["art_ec3"].Visible = true; }
            if (dgvImport.Columns.Contains("famiglia")) dgvImport.Columns["famiglia"].Visible = false;
            if (dgvImport.Columns.Contains("eti")) dgvImport.Columns["eti"].Visible = false;
            if (dgvImport.Columns.Contains("ori")) dgvImport.Columns["ori"].Visible = false;
            if (dgvImport.Columns.Contains("cal")) dgvImport.Columns["cal"].Visible = false;
            if (dgvImport.Columns.Contains("cat_bil")) dgvImport.Columns["cat_bil"].Visible = false;
            if (dgvImport.Columns.Contains("tas")) dgvImport.Columns["tas"].Visible = false;
            if (dgvImport.Columns.Contains("tra")) dgvImport.Columns["tra"].Visible = false;
            if (dgvImport.Columns.Contains("gsc")) dgvImport.Columns["gsc"].Visible = false;
            if (dgvImport.Columns.Contains("bpz")) dgvImport.Columns["bpz"].Visible = false;
            if (dgvImport.Columns.Contains("tar")) dgvImport.Columns["tar"].Visible = false;
            if (dgvImport.Columns.Contains("bil")) dgvImport.Columns["bil"].Visible = false;
            if (dgvImport.Columns.Contains("ean_bil_gen")) dgvImport.Columns["ean_bil_gen"].Visible = false;
            if (dgvImport.Columns.Contains("is_new")) dgvImport.Columns["is_new"].Visible = false;
        }

        private string GetConfigOption(DataTable tMap, string fldName, string defVal = "")
        {
            if (tMap == null) return defVal;
            DataRow[] rows = tMap.Select("FldName='" + fldName + "'");
            if (rows != null && rows.Length > 0)
            {
                if (tMap.Columns.Contains("FixVal") && !DBNull.Value.Equals(rows[0]["FixVal"]))
                {
                    string v = rows[0]["FixVal"].ToString().Trim();
                    if (!string.IsNullOrEmpty(v)) return v;
                }
                if (tMap.Columns.Contains("ColPos") && !DBNull.Value.Equals(rows[0]["ColPos"]))
                {
                    string v = rows[0]["ColPos"].ToString().Trim();
                    if (!string.IsNullOrEmpty(v) && v != "0") return v;
                }
            }
            return defVal;
        }

        private string GetMappedValue(DataRow rXls, DataTable tMap, string fldName)
        {
            DataRow[] rows = tMap.Select("FldName='" + fldName + "'");
            if (rows.Length == 0 && (fldName == "art_fam" || fldName == "art_ec1"))
                rows = tMap.Select("FldName='art_fam' OR FldName='art_ec1'");
            if (rows.Length == 0 && (fldName == "art_cat" || fldName == "art_cat_bil"))
                rows = tMap.Select("FldName='art_cat' OR FldName='art_cat_bil'");

            if (rows.Length > 0)
            {
                int col = Convert.ToInt32(rows[0]["ColPos"]) - 1; // 1-based to 0-based
                string fixVal = tMap.Columns.Contains("FixVal") && !DBNull.Value.Equals(rows[0]["FixVal"]) ? rows[0]["FixVal"].ToString().Trim() : "";

                if (col >= 0 && col < rXls.Table.Columns.Count)
                {
                    object val = rXls[col];
                    if (val != null && val != DBNull.Value)
                    {
                        string sVal = "";
                        if (val is double || val is float || val is decimal || val is int || val is long)
                        {
                            double d = Convert.ToDouble(val);
                            if (fldName == "art_ean" || fldName == "art_plu")
                                sVal = d.ToString("0");
                            else if (fldName == "art_rep")
                            {
                                int iR = Convert.ToInt32(d);
                                sVal = iR > 0 ? iR.ToString("D3") : "001";
                            }
                            else if (fldName == "art_reb")
                            {
                                int iR = Convert.ToInt32(d);
                                sVal = iR.ToString();
                            }
                            else if (fldName == "art_pxc" || fldName == "art_fam" || fldName == "art_ec1" || fldName == "art_ec2" || fldName == "art_ec3" || fldName == "art_iva" || fldName == "art_for" || fldName == "art_cod_for" || fldName == "art_tas" || fldName == "art_gsc" || fldName == "art_ori" || fldName == "art_cal" || fldName == "art_cat_bil")
                                sVal = d.ToString("0");
                            else
                                sVal = d.ToString("0.00");
                        }
                        else
                        {
                            sVal = val.ToString().Trim();
                        }

                        if (!string.IsNullOrEmpty(sVal))
                            return sVal;
                    }
                }

                if (!string.IsNullOrEmpty(fixVal))
                    return fixVal;
            }
            return "";
        }

        private DataTable ReadExcel(string strPath)
        {
            if (string.IsNullOrEmpty(strPath) || !File.Exists(strPath))
                throw new FileNotFoundException("File non trovato o percorso non valido:\n" + strPath);

            string ext = Path.GetExtension(strPath).ToLower();

            if (ext == ".csv" || ext == ".txt")
            {
                return ReadCsvToDataTable(strPath);
            }

            // Per file .xls e .xlsx: crea una copia temporanea in Temp per evitare l'errore di file bloccato da Excel
            string tempFile = Path.Combine(Path.GetTempPath(), "ap_imp_" + Guid.NewGuid().ToString("N") + ext);
            try
            {
                File.Copy(strPath, tempFile, true);
            }
            catch
            {
                tempFile = strPath;
            }

            try
            {
                string[] providers;
                string[] extProps;

                if (ext == ".xls")
                {
                    providers = new string[] { "Microsoft.Jet.OLEDB.4.0", "Microsoft.ACE.OLEDB.12.0", "Microsoft.ACE.OLEDB.16.0" };
                    extProps = new string[] { "'Excel 8.0;HDR=No;IMEX=1;'", "'Excel 8.0;HDR=No;IMEX=1;'", "'Excel 8.0;HDR=No;IMEX=1;'" };
                }
                else
                {
                    providers = new string[] { "Microsoft.ACE.OLEDB.12.0", "Microsoft.ACE.OLEDB.16.0", "Microsoft.Jet.OLEDB.4.0" };
                    extProps = new string[] { "'Excel 12.0 Xml;HDR=No;IMEX=1;'", "'Excel 12.0 Xml;HDR=No;IMEX=1;'", "'Excel 8.0;HDR=No;IMEX=1;'" };
                }

                Exception lastEx = null;
                for (int i = 0; i < providers.Length; i++)
                {
                    try
                    {
                        string connStr = string.Format("Provider={0};Data Source=\"{1}\";Extended Properties={2}", providers[i], tempFile, extProps[i]);
                        using (OleDbConnection conn = new OleDbConnection(connStr))
                        {
                            conn.Open();
                            DataTable schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            if (schema == null || schema.Rows.Count == 0) continue;

                            string sheetName = schema.Rows[0]["TABLE_NAME"].ToString();
                            using (OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + sheetName + "]", conn))
                            {
                                using (OleDbDataAdapter da = new OleDbDataAdapter(cmd))
                                {
                                    DataTable dt = new DataTable();
                                    da.Fill(dt);
                                    return dt;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        lastEx = ex;
                    }
                }

                throw lastEx ?? new Exception("Impossibile leggere il file Excel con i driver OLEDB installati.");
            }
            finally
            {
                if (tempFile != strPath && File.Exists(tempFile))
                {
                    try { File.Delete(tempFile); } catch { }
                }
            }
        }

        private DataTable ReadCsvToDataTable(string filePath)
        {
            DataTable dt = new DataTable();
            string[] lines = File.ReadAllLines(filePath, Encoding.Default);
            if (lines.Length == 0) return dt;

            char sep = ';';
            if (lines[0].Contains("\t")) sep = '\t';
            else if (lines[0].Contains(";") && !lines[0].Contains(",")) sep = ';';
            else if (lines[0].Contains(",") && !lines[0].Contains(";")) sep = ',';

            string[] firstRow = lines[0].Split(sep);
            for (int i = 0; i < firstRow.Length; i++)
            {
                dt.Columns.Add("F" + (i + 1), typeof(string));
            }

            foreach (string line in lines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                string[] parts = line.Split(sep);
                DataRow r = dt.NewRow();
                for (int i = 0; i < Math.Min(parts.Length, dt.Columns.Count); i++)
                {
                    r[i] = parts[i].Trim('"', '\'', ' ');
                }
                dt.Rows.Add(r);
            }
            return dt;
        }

        private DataTable CreateTargetTable()
        {
            DataTable t = new DataTable();
            t.Columns.Add("art_cod", typeof(string));
            t.Columns.Add("art_ean", typeof(string));
            t.Columns.Add("art_des", typeof(string));
            t.Columns.Add("old_des", typeof(string));
            t.Columns.Add("old_cos", typeof(decimal));
            t.Columns.Add("new_cos", typeof(decimal));
            t.Columns.Add("old_prv", typeof(decimal));
            t.Columns.Add("new_prv", typeof(decimal));
            t.Columns.Add("diff_cos", typeof(decimal));
            t.Columns.Add("diff_prp", typeof(decimal));
            t.Columns.Add("iva", typeof(string));
            t.Columns.Add("art_for", typeof(string));
            t.Columns.Add("art_arf", typeof(string));
            t.Columns.Add("pxc", typeof(string));
            t.Columns.Add("gr", typeof(string));
            t.Columns.Add("peso", typeof(string));
            t.Columns.Add("reparto", typeof(string));
            t.Columns.Add("famiglia", typeof(string));
            t.Columns.Add("art_ec1", typeof(string));
            t.Columns.Add("art_ec2", typeof(string));
            t.Columns.Add("art_ec3", typeof(string));
            t.Columns.Add("bil", typeof(string));
            t.Columns.Add("eti", typeof(string));
            t.Columns.Add("plu", typeof(string));
            t.Columns.Add("reb", typeof(string));
            t.Columns.Add("ori", typeof(string));
            t.Columns.Add("cal", typeof(string));
            t.Columns.Add("cat_bil", typeof(string));
            t.Columns.Add("tas", typeof(string));
            t.Columns.Add("tra", typeof(string));
            t.Columns.Add("gsc", typeof(string));
            t.Columns.Add("bpz", typeof(bool));
            t.Columns.Add("tar", typeof(string));
            t.Columns.Add("ean_bil_gen", typeof(string));
            t.Columns.Add("is_new", typeof(bool));
            return t;
        }

        private void dgvImport_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow r in dgvImport.Rows)
            {
                if (r.Cells["new_cos"].Value == DBNull.Value) continue;

                bool isNew = false;
                if (dgvImport.Columns.Contains("is_new") && r.Cells["is_new"].Value != null && r.Cells["is_new"].Value != DBNull.Value)
                {
                    isNew = Convert.ToBoolean(r.Cells["is_new"].Value);
                }

                if (isNew)
                {
                    // Articolo non trovato (Nuova referenza) -> BLU
                    r.DefaultCellStyle.ForeColor = Color.Blue;
                }
                else
                {
                    if (r.Cells["old_cos"].Value != DBNull.Value)
                    {
                        decimal dOldCos = Convert.ToDecimal(r.Cells["old_cos"].Value);
                        decimal dNewCos = Convert.ToDecimal(r.Cells["new_cos"].Value);

                        if (dNewCos > dOldCos && dOldCos > 0)
                            r.DefaultCellStyle.ForeColor = Color.Red; // Increase
                        else if (dNewCos < dOldCos && dOldCos > 0)
                            r.DefaultCellStyle.ForeColor = Color.Green; // Decrease
                    }
                }
            }
        }

        private void btnAcquire_Click(object sender, EventArgs e)
        {
            if (_tImport == null || _tImport.Rows.Count == 0)
            {
                MessageBox.Show("Nessun dato da importare.", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show("Procedere con l'acquisizione delle variazioni e l'allineamento anagrafiche?\nL'operazione su grandi volumi può richiedere qualche minuto.", "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                SqlConnection cn = new SqlConnection(_strConSql);
                SqlTransaction tr = null;
                try
                {
                    btnAcquire.Enabled = false;
                    btnBack.Enabled = false;
                    btnNext.Enabled = false;
                    progressBar1.Value = 0;
                    progressBar1.Maximum = _tImport.Rows.Count;
                    progressBar1.Visible = true;

                    string sIdMap = cmbMapping.SelectedValue.ToString();
                    DataTable tTesMapping = _clsFun.FillTabSql("Tes", "SELECT MapFor FROM XlsMapTestata WHERE MapId='" + sIdMap + "'", false, _strConSql);
                    string sDefFor = (tTesMapping.Rows.Count > 0 && tTesMapping.Rows[0]["MapFor"] != DBNull.Value) ? tTesMapping.Rows[0]["MapFor"].ToString() : "";

                    lblFileStatus.Text = "Predisposizione cache... attendere.";
                    Application.DoEvents();

                    // 0. Cache EAN, Supplier Article Code, and Articles for faster lookup
                    Dictionary<string, string> dictEan = new Dictionary<string, string>();
                    DataTable tAllEan = _clsFun.FillTabSql("AllEan", "SELECT ean_ean, ean_art FROM AnaBarcode", false, _strConSql);
                    foreach (DataRow rEan in tAllEan.Rows)
                    {
                        string sKey = rEan["ean_ean"].ToString().Trim();
                        if (!dictEan.ContainsKey(sKey)) dictEan.Add(sKey, rEan["ean_art"].ToString().Trim());
                    }

                    Dictionary<string, string> dictArfToArt = new Dictionary<string, string>();
                    Dictionary<string, string> dictDesToArt = new Dictionary<string, string>();
                    DataTable tArtSchema = _clsFun.FillTabSql("AnaArticoli", "SELECT * FROM AnaArticoli WHERE 1=0", false, _strConSql);
                    DataTable tEanSchema = _clsFun.FillTabSql("AnaBarcode", "SELECT * FROM AnaBarcode WHERE 1=0", false, _strConSql);
                    DataTable tLiaSchema = _clsFun.FillTabSql("GesLisAcquisto", "SELECT * FROM GesLisAcquisto WHERE 1=0", false, _strConSql);
                    DataTable tLivSchema = _clsFun.FillTabSql("GesLisVendita", "SELECT * FROM GesLisVendita WHERE 1=0", false, _strConSql);
                    foreach (DataRow rImp in _tImport.Rows)
                    {
                        bool isImpNew = rImp.Table.Columns.Contains("is_new") && rImp["is_new"] != DBNull.Value && Convert.ToBoolean(rImp["is_new"]);
                        if (!isImpNew && rImp["art_cod"] != DBNull.Value && !string.IsNullOrEmpty(rImp["art_cod"].ToString().Trim()))
                        {
                                                        string sC = rImp["art_cod"].ToString().Trim();
                            if (!IsPlaceholderCode(sC))
                            {
                                string sD = rImp["art_des"].ToString().Trim().ToUpper();
                                if (!string.IsNullOrEmpty(sD) && sD.Length >= 3 && !dictDesToArt.ContainsKey(sD))
                                {
                                    dictDesToArt[sD] = sC;
                                }
                            }
                        }
                    }

                    cn.Open();
                    tr = cn.BeginTransaction();
                    List<string> logEntries = new List<string>();
                    logEntries.Add("===============================================================================");
                    logEntries.Add("APOFFICE - REPORT IMPORTAZIONE DATI EXCEL");
                    logEntries.Add("Data/Ora: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                    logEntries.Add("File Excel: " + txtFilePath.Text);
                    logEntries.Add("Opzione Nuove Referenze: " + (chkInsertNew.Checked ? "ABILITATA" : "DISABILITATA"));
                    logEntries.Add("Opzione Stato Attivo ('A'): " + (chkSetActiveStatus.Checked ? "ABILITATA" : "DISABILITATA"));
                    logEntries.Add("===============================================================================\n");

                    int countInsertedNew = 0;
                    int countUpdatedExisting = 0;
                    int countSkipped = 0;
                    int countErrors = 0;
                    int currentRowAcq = 0;
                    _bolCancel = false;

                    foreach (DataRow r in _tImport.Rows)
                    {
                        if (_bolCancel) break;

                        currentRowAcq++;
                        if (currentRowAcq % 500 == 0)
                        {
                            progressBar1.Value = currentRowAcq;
                            lblFileStatus.Text = "Acquisizione riga " + currentRowAcq + " di " + _tImport.Rows.Count;
                            Application.DoEvents();
                        }

                                                string sEan = r["art_ean"].ToString().Trim();
                        string sGenEan = r.Table.Columns.Contains("ean_bil_gen") && r["ean_bil_gen"] != null ? r["ean_bil_gen"].ToString().Trim() : "";
                        string sPlu = (r.Table.Columns.Contains("plu") && r["plu"] != null) ? r["plu"].ToString().Trim() : "";
                        string sReb = (r.Table.Columns.Contains("reb") && r["reb"] != null) ? r["reb"].ToString().Trim() : "";
                        string sArt = "";
                        string sRowFor = (r["art_for"].ToString() != "") ? r["art_for"].ToString() : sDefFor;
                        string sArf = r["art_arf"].ToString().Trim();
                        string sArfKey = sRowFor + "|" + sArf;
                        string sDesKey = r["art_des"].ToString().Trim().ToUpper();
                        string sExplicitCod = (r["art_cod"] != null && r["art_cod"] != DBNull.Value) ? r["art_cod"].ToString().Trim() : "";
                        if (IsPlaceholderCode(sExplicitCod)) sExplicitCod = "";

                        bool isNewRow = r.Table.Columns.Contains("is_new") && r["is_new"] != DBNull.Value && Convert.ToBoolean(r["is_new"]);

                        string sBilStr = (r["bil"] != null) ? r["bil"].ToString().Trim() : "";
                        bool bIsBilancia = (!string.IsNullOrEmpty(sPlu)) ||
                                           (!string.IsNullOrEmpty(sReb)) ||
                                           (!string.IsNullOrEmpty(sEan) && sEan.StartsWith("2")) ||
                                           (!string.IsNullOrEmpty(sGenEan) && sGenEan.StartsWith("2")) ||
                                           sBilStr == "1" || sBilStr.Equals("S", StringComparison.OrdinalIgnoreCase) ||
                                           sBilStr.Equals("SI", StringComparison.OrdinalIgnoreCase) ||
                                           sBilStr.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
                                           sBilStr.Equals("YES", StringComparison.OrdinalIgnoreCase);

                        tr.Save("RowStart");
                        try
                        {
                            // 1. Resolve Article Code (sArt)
                            // If preview matched an existing article in DB, use it
                            if (!isNewRow && !string.IsNullOrEmpty(sExplicitCod) && !IsPlaceholderCode(sExplicitCod) && CheckArticleExists(sExplicitCod, cn, tr))
                            {
                                sArt = sExplicitCod;
                            }
                            else if (!isNewRow)
                            {
                                // Match existing reference from cache / preview
                                if (!string.IsNullOrEmpty(sArf) && dictArfToArt.ContainsKey(sArfKey))
                                {
                                    sArt = dictArfToArt[sArfKey];
                                }
                                else if (!string.IsNullOrEmpty(sDesKey) && dictDesToArt.ContainsKey(sDesKey))
                                {
                                    sArt = dictDesToArt[sDesKey];
                                }
                                else if (!string.IsNullOrEmpty(sEan) && !sEan.StartsWith("2") && dictEan.ContainsKey(sEan))
                                {
                                    sArt = dictEan[sEan];
                                }
                                else if (!string.IsNullOrEmpty(sEan) && !sEan.StartsWith("2"))
                                {
                                    sArt = FindArticleByBarcode(sEan, cn, tr);
                                }
                                else if (!string.IsNullOrEmpty(sArf))
                                {
                                    sArt = FindArticleBySupplierCode(sRowFor, sArf, cn, tr);
                                }
                            }

                            decimal dNewCos = Convert.ToDecimal(r["new_cos"]);
                            decimal dNewPrv = Convert.ToDecimal(r["new_prv"]);

                            // 2. If existing article found -> UPDATE AND CONSOLIDATE
                            if (!string.IsNullOrEmpty(sArt))
                            {
                                countUpdatedExisting++;

                                UpdateExistingArticle(sArt, r, sRowFor, dNewCos, dNewPrv,
                                    tArtSchema, tEanSchema, tLiaSchema, tLivSchema,
                                    dictEan, dictArfToArt, dictDesToArt, cn, tr);

                                logEntries.Add("[AGGIORNATO/CONSOLIDATO] Codice: " + sArt + " | Barcode: " + sEan + " | " + r["art_des"].ToString());
                            }
                            // 3. Otherwise -> INSERT NEW ARTICLE (or consolidate if unique code collision happens)
                            else if (chkInsertNew.Checked)
                            {
                                string sNewArt = "";
                                if (!string.IsNullOrEmpty(sExplicitCod) && !IsPlaceholderCode(sExplicitCod))
                                {
                                    if (CheckArticleExists(sExplicitCod, cn, tr))
                                    {
                                        // Code already exists in DB! Switch immediately to update and consolidate!
                                        countUpdatedExisting++;
                                        UpdateExistingArticle(sExplicitCod, r, sRowFor, dNewCos, dNewPrv,
                                            tArtSchema, tEanSchema, tLiaSchema, tLivSchema,
                                            dictEan, dictArfToArt, dictDesToArt, cn, tr);

                                        logEntries.Add("[CONSOLIDATO] Codice esistente: " + sExplicitCod + " | Barcode: " + sEan + " | " + r["art_des"].ToString());
                                        continue;
                                    }
                                    sNewArt = sExplicitCod;
                                }
                                else
                                {
                                    sNewArt = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaArticoli, 7, _strConSql);
                                    while (CheckArticleExists(sNewArt, cn, tr))
                                    {
                                        sNewArt = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaArticoli, 7, _strConSql);
                                    }
                                }

                                countInsertedNew++;

                                string sDes = r["art_des"].ToString().Trim();
                                string sDeb = (sDes.Length > 20) ? sDes.Substring(0, 20) : sDes;
                                string sIva = (r["iva"] != null && r["iva"].ToString().Trim() != "") ? ResolveIvaCode(r["iva"].ToString().Trim()) : ResolveIvaCode("");

                                string sPxc = (r["pxc"] != null && r["pxc"].ToString().Trim() != "") ? r["pxc"].ToString().Replace(",", ".") : "1";
                                if (sPxc.Contains(".")) sPxc = sPxc.Split('.')[0];

                                string sPne = (r["peso"] != null && r["peso"].ToString().Trim() != "") ? r["peso"].ToString().Replace(",", ".") : "0";

                                string sRep = (r["reparto"] != null && r["reparto"].ToString().Trim() != "") ? ResolveRepCode(r["reparto"].ToString().Trim()) : ResolveRepCode("");
                                string sEc1 = (r["art_ec1"] != null && r["art_ec1"].ToString().Trim() != "") ? Format3Digits(r["art_ec1"].ToString().Trim(), "") : ((r["famiglia"] != null) ? Format3Digits(r["famiglia"].ToString().Trim(), "") : "");
                                string sEc2 = (r["art_ec2"] != null) ? Format3Digits(r["art_ec2"].ToString().Trim(), "") : "";
                                string sEc3 = (r["art_ec3"] != null) ? Format3Digits(r["art_ec3"].ToString().Trim(), "") : "";
                                string sGr = (r["gr"] != null) ? r["gr"].ToString().Trim() : "";

                                string sOri = (r.Table.Columns.Contains("ori") && r["ori"] != null) ? r["ori"].ToString().Trim() : "";
                                string sCal = (r.Table.Columns.Contains("cal") && r["cal"] != null) ? r["cal"].ToString().Trim() : "";
                                string sCatBil = (r.Table.Columns.Contains("cat_bil") && r["cat_bil"] != null) ? r["cat_bil"].ToString().Trim() : "";
                                string sTas = (r.Table.Columns.Contains("tas") && r["tas"] != null && !string.IsNullOrEmpty(r["tas"].ToString().Trim())) ? r["tas"].ToString().Trim() : sGr;
                                string sTra = (r.Table.Columns.Contains("tra") && r["tra"] != null) ? r["tra"].ToString().Trim() : "";
                                string sGsc = (r.Table.Columns.Contains("gsc") && r["gsc"] != null) ? r["gsc"].ToString().Trim() : "0";
                                bool bIsBpz = (r.Table.Columns.Contains("bpz") && r["bpz"] != DBNull.Value) ? Convert.ToBoolean(r["bpz"]) : false;
                                string sTar = (r.Table.Columns.Contains("tar") && r["tar"] != null) ? r["tar"].ToString().Trim() : "0";

                                string sEtiMap = (r["eti"] != null && !string.IsNullOrEmpty(r["eti"].ToString().Trim())) ? Format3Digits(r["eti"].ToString().Trim(), "043") : "043";

                                // Build new AnaArticoli DataRow
                                DataRow rNewArt = tArtSchema.NewRow();
                                foreach (DataColumn col in tArtSchema.Columns)
                                {
                                    if (col.DataType == typeof(string)) rNewArt[col] = "";
                                    else if (col.DataType == typeof(decimal) || col.DataType == typeof(double) || col.DataType == typeof(float)) rNewArt[col] = 0m;
                                    else if (col.DataType == typeof(int) || col.DataType == typeof(short) || col.DataType == typeof(long) || col.DataType == typeof(byte)) rNewArt[col] = 0;
                                    else if (col.DataType == typeof(DateTime)) rNewArt[col] = DateTime.Today;
                                    else if (col.DataType == typeof(bool)) rNewArt[col] = false;
                                }

                                if (tArtSchema.Columns.Contains("art_cod")) rNewArt["art_cod"] = TruncateField(tArtSchema, "art_cod", sNewArt);
                                if (tArtSchema.Columns.Contains("art_des")) rNewArt["art_des"] = TruncateField(tArtSchema, "art_des", sDes);
                                if (tArtSchema.Columns.Contains("art_deb")) rNewArt["art_deb"] = TruncateField(tArtSchema, "art_deb", sDeb);
                                if (tArtSchema.Columns.Contains("art_cos")) rNewArt["art_cos"] = dNewCos;
                                if (tArtSchema.Columns.Contains("art_sta")) rNewArt["art_sta"] = "A";
                                if (tArtSchema.Columns.Contains("art_iva")) rNewArt["art_iva"] = TruncateField(tArtSchema, "art_iva", sIva);
                                if (tArtSchema.Columns.Contains("art_umi")) rNewArt["art_umi"] = "NR";
                                if (tArtSchema.Columns.Contains("art_umc")) rNewArt["art_umc"] = "NR";
                                if (tArtSchema.Columns.Contains("art_tgr")) rNewArt["art_tgr"] = "PZ";
                                if (tArtSchema.Columns.Contains("art_pxc")) rNewArt["art_pxc"] = _clsFun.Txt2Dec(sPxc);
                                if (tArtSchema.Columns.Contains("art_pne")) rNewArt["art_pne"] = _clsFun.Txt2Dec(sPne);
                                if (tArtSchema.Columns.Contains("art_tas")) rNewArt["art_tas"] = TruncateField(tArtSchema, "art_tas", ResolveTas(sTas));
                                if (tArtSchema.Columns.Contains("art_rep")) rNewArt["art_rep"] = TruncateField(tArtSchema, "art_rep", sRep);
                                if (tArtSchema.Columns.Contains("art_cat")) rNewArt["art_cat"] = TruncateField(tArtSchema, "art_cat", sCatBil);
                                if (tArtSchema.Columns.Contains("art_ec1")) rNewArt["art_ec1"] = TruncateField(tArtSchema, "art_ec1", sEc1);
                                if (tArtSchema.Columns.Contains("art_ec2")) rNewArt["art_ec2"] = TruncateField(tArtSchema, "art_ec2", sEc2);
                                if (tArtSchema.Columns.Contains("art_ec3")) rNewArt["art_ec3"] = TruncateField(tArtSchema, "art_ec3", sEc3);
                                if (tArtSchema.Columns.Contains("art_eti")) rNewArt["art_eti"] = TruncateField(tArtSchema, "art_eti", sEtiMap);
                                if (tArtSchema.Columns.Contains("art_bil")) rNewArt["art_bil"] = bIsBilancia;
                                if (tArtSchema.Columns.Contains("art_plu")) rNewArt["art_plu"] = TruncateField(tArtSchema, "art_plu", sPlu);
                                if (tArtSchema.Columns.Contains("art_reb")) rNewArt["art_reb"] = TruncateField(tArtSchema, "art_reb", sReb);
                                if (tArtSchema.Columns.Contains("art_ori")) rNewArt["art_ori"] = TruncateField(tArtSchema, "art_ori", sOri);
                                if (tArtSchema.Columns.Contains("art_cal")) rNewArt["art_cal"] = TruncateField(tArtSchema, "art_cal", sCal);
                                if (tArtSchema.Columns.Contains("art_tra")) rNewArt["art_tra"] = TruncateField(tArtSchema, "art_tra", sTra);
                                if (tArtSchema.Columns.Contains("art_gsc")) rNewArt["art_gsc"] = _clsFun.Txt2Dec(sGsc);
                                if (tArtSchema.Columns.Contains("art_bpz")) rNewArt["art_bpz"] = bIsBpz;
                                if (tArtSchema.Columns.Contains("art_tar")) rNewArt["art_tar"] = _clsFun.Txt2Dec(sTar);
                                if (tArtSchema.Columns.Contains("art_dti"))
                                {
                                    if (tArtSchema.Columns["art_dti"].DataType == typeof(DateTime)) rNewArt["art_dti"] = DateTime.Today;
                                    else rNewArt["art_dti"] = DateTime.Today.ToString("dd/MM/yyyy");
                                }
                                if (tArtSchema.Columns.Contains("art_dtm"))
                                {
                                    if (tArtSchema.Columns["art_dtm"].DataType == typeof(DateTime)) rNewArt["art_dtm"] = DateTime.Today;
                                    else rNewArt["art_dtm"] = DateTime.Today.ToString("dd/MM/yyyy");
                                }

                                try
                                {
                                    string sSqlInsArt = _clsFun.SqlInsertRow("AnaArticoli", tArtSchema, rNewArt);
                                    using (SqlCommand cm = new SqlCommand(sSqlInsArt, cn, tr))
                                    {
                                        cm.ExecuteNonQuery();
                                    }
                                }
                                catch (SqlException sqlex) when (sqlex.Number == 2601 || sqlex.Number == 2627)
                                {
                                    // Primary key conflict! Consolidate with existing record
                                    countUpdatedExisting++;
                                    UpdateExistingArticle(sNewArt, r, sRowFor, dNewCos, dNewPrv,
                                        tArtSchema, tEanSchema, tLiaSchema, tLivSchema,
                                        dictEan, dictArfToArt, dictDesToArt, cn, tr);

                                    logEntries.Add("[CONSOLIDATO] Articolo duplicato: " + sNewArt + " | Barcode: " + sEan + " | " + r["art_des"].ToString());
                                    continue;
                                }

                                // GesLisAcquisto
                                decimal dPxc = 1m;
                                if (r["pxc"] != null && !string.IsNullOrEmpty(r["pxc"].ToString().Trim()))
                                {
                                    dPxc = _clsFun.Txt2Dec(r["pxc"].ToString());
                                    if (dPxc <= 0) dPxc = 1m;
                                }
                                EnsureLiaRecord(sNewArt, sRowFor, sArf, dNewCos, dPxc, tLiaSchema, cn, tr);

                                // AnaBarcode
                                string sGenEanIns = r.Table.Columns.Contains("ean_bil_gen") && r["ean_bil_gen"] != null ? r["ean_bil_gen"].ToString().Trim() : "";
                                bool bBpzIns = r.Table.Columns.Contains("bpz") && r["bpz"] != DBNull.Value ? Convert.ToBoolean(r["bpz"]) : false;

                                if (!string.IsNullOrEmpty(sEan))
                                {
                                    bool bEanInsIsBil = bIsBilancia && sEan.StartsWith("2");
                                    EnsureBarcodeRecord(sEan, sNewArt, dNewPrv, bEanInsIsBil, bBpzIns, tEanSchema, cn, tr);
                                    dictEan[sEan] = sNewArt;
                                }
                                if (!string.IsNullOrEmpty(sGenEanIns) && sGenEanIns != sEan)
                                {
                                    EnsureBarcodeRecord(sGenEanIns, sNewArt, dNewPrv, true, bBpzIns, tEanSchema, cn, tr);
                                    dictEan[sGenEanIns] = sNewArt;
                                }

                                // GesLisVendita
                                if (dNewPrv > 0)
                                {
                                    EnsureLivRecord(sNewArt, dNewPrv, tLivSchema, cn, tr);
                                }

                                // Cache lookup
                                if (!string.IsNullOrEmpty(sArf)) dictArfToArt[sArfKey] = sNewArt;
                                if (!string.IsNullOrEmpty(sDesKey) && sDesKey.Length >= 3) dictDesToArt[sDesKey] = sNewArt;

                                // Variations
                                if (chkSendToCasse.Checked) AddVariation(sNewArt, "POS", "NEW ART XLS", cn, tr);
                                if (chkSendToStampa.Checked) AddVariation(sNewArt, "ETI", "NEW ART XLS", cn, tr);
                                if (bIsBilancia) AddVariation(sNewArt, "BIL", "NEW ART XLS", cn, tr);

                                logEntries.Add("[INSERITO] Nuovo Articolo: " + sNewArt + " | Barcode: " + sEan + " | " + r["art_des"].ToString());
                            }
                            else
                            {
                                countSkipped++;
                                logEntries.Add("[NON INSERITO - NUOVE REFERENZE DISABILITATE] Barcode: " + sEan + " | Descrizione: " + r["art_des"].ToString());
                            }
                        }
                        catch (SqlException sqlex)
                        {
                            if (sqlex.Number == 2601 || sqlex.Number == 2627)
                            {
                                tr.Rollback("RowStart");
                                // Try fallback consolidation
                                                                string sFallbackArt = !string.IsNullOrEmpty(sEan) ? FindArticleByBarcode(sEan, cn, tr) : "";
                                if (string.IsNullOrEmpty(sFallbackArt) && !string.IsNullOrEmpty(sGenEan))
                                    sFallbackArt = FindArticleByBarcode(sGenEan, cn, tr);
                                if (string.IsNullOrEmpty(sFallbackArt) && !string.IsNullOrEmpty(sPlu))
                                    sFallbackArt = FindArticleByPlu(sPlu, sReb, cn, tr);
                                if (string.IsNullOrEmpty(sFallbackArt) && !string.IsNullOrEmpty(sArf))
                                    sFallbackArt = FindArticleBySupplierCode(sRowFor, sArf, cn, tr);

                                if (!string.IsNullOrEmpty(sFallbackArt))
                                {
                                    try
                                    {
                                        decimal dNewCos = Convert.ToDecimal(r["new_cos"]);
                                        decimal dNewPrv = Convert.ToDecimal(r["new_prv"]);
                                        UpdateExistingArticle(sFallbackArt, r, sRowFor, dNewCos, dNewPrv,
                                            tArtSchema, tEanSchema, tLiaSchema, tLivSchema,
                                            dictEan, dictArfToArt, dictDesToArt, cn, tr);
                                        countUpdatedExisting++;
                                        logEntries.Add("[CONSOLIDATO DA ECCEZIONE] Articolo: " + sFallbackArt + " | Barcode: " + sEan);
                                        continue;
                                    }
                                    catch { }
                                }

                                countSkipped++;
                                logEntries.Add("[CONSOLIDAMENTO NON RIUSCITO] Barcode: " + sEan + " | Descrizione: " + r["art_des"].ToString());
                                continue;
                            }
                            countErrors++;
                            logEntries.Add("[ERRORE DATABASE] Barcode: [" + sEan + "] " + r["art_des"].ToString() + " -> " + sqlex.Message);
                            throw new Exception("Prodotto: [" + sEan + "] " + r["art_des"].ToString() + "\nErrore Database: " + sqlex.Message, sqlex);
                        }
                        catch (Exception innerEx)
                        {
                            tr.Rollback("RowStart");
                            countErrors++;
                            logEntries.Add("[ERRORE GENERICO] Barcode: [" + sEan + "] " + r["art_des"].ToString() + " -> " + innerEx.Message);
                            throw new Exception("Prodotto: [" + sEan + "] " + r["art_des"].ToString() + "\nErrore: " + innerEx.Message, innerEx);
                        }
                    }

                    if (_bolCancel)
                    {
                        if (tr != null) tr.Rollback();
                        MessageBox.Show("Acquisizione interrotta. Nessuna modifica è stata salvata.", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        if (tr != null) tr.Commit();

                        // Write Log File
                        string logFolder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
                        if (!System.IO.Directory.Exists(logFolder)) System.IO.Directory.CreateDirectory(logFolder);
                        string logFile = System.IO.Path.Combine(logFolder, "ImportXls_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt");

                        logEntries.Add("\n===============================================================================");
                        logEntries.Add("RIEPILOGO FINALE:");
                        logEntries.Add("Totale righe elaborate: " + _tImport.Rows.Count);
                        logEntries.Add("Articoli aggiornati / consolidati: " + countUpdatedExisting);
                        logEntries.Add("Nuovi articoli creati: " + countInsertedNew);
                        logEntries.Add("Prodotti non inseriti/scartati: " + countSkipped);
                        logEntries.Add("Errori riscontrati: " + countErrors);
                        logEntries.Add("===============================================================================");

                        System.IO.File.WriteAllLines(logFile, logEntries);

                        string sMsgResult = "Importazione ed Allineamento completati con successo!\n\n" +
                                            "• Articoli aggiornati / consolidati: " + countUpdatedExisting + "\n" +
                                            "• Nuovi articoli creati: " + countInsertedNew + "\n" +
                                            "• Prodotti non inseriti/scartati: " + countSkipped + "\n\n" +
                                            "Il file di report dettagliato è stato salvato in:\n" + logFile;

                        MessageBox.Show(sMsgResult, "Fine Importazione", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.Close();
                    }
                }
                catch (Exception ex)
                {
                    if (tr != null) tr.Rollback();
                    MessageBox.Show("Errore irreversibile durante l'acquisizione:\n" + ex.Message + "\n\nL'operazione è stata annullata per garantire l'integrità dei dati.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (cn.State == ConnectionState.Open) cn.Close();
                    cn.Dispose();
                    btnAcquire.Enabled = true;
                    btnBack.Enabled = true;
                    btnNext.Enabled = true;
                    progressBar1.Visible = false;
                    _bolCancel = false;
                }
            }
        }

        private void AddVariation(string sArt, string sTip, string sOri, SqlConnection cn, SqlTransaction tr)
        {
            // Transaction-aware UPSERT logic to handle duplicate key constraints gracefully
            string sCheck = "SELECT COUNT(*) FROM GesVariazioni WHERE var_art=@art AND var_tip=@tip AND var_num=@num";
            using (SqlCommand cm = new SqlCommand(sCheck, cn, tr))
            {
                cm.Parameters.AddWithValue("@art", sArt);
                cm.Parameters.AddWithValue("@tip", sTip);
                cm.Parameters.AddWithValue("@num", _clsDef.COD03Z);
                int count = (int)cm.ExecuteScalar();
                if (count == 0)
                {
                    // New variation
                    string sIns = "INSERT INTO GesVariazioni (var_tip, var_art, var_num, var_dti, var_dtv, var_ori, var_inv, var_off) VALUES (@tip, @art, @num, GETDATE(), GETDATE(), @ori, @inv, '')";
                    using (SqlCommand cmIns = new SqlCommand(sIns, cn, tr))
                    {
                        cmIns.Parameters.AddWithValue("@tip", sTip);
                        cmIns.Parameters.AddWithValue("@art", sArt);
                        cmIns.Parameters.AddWithValue("@num", _clsDef.COD03Z);
                        cmIns.Parameters.AddWithValue("@ori", sOri);
                        cmIns.Parameters.AddWithValue("@inv", _clsDef.DIVDAD);
                        cmIns.ExecuteNonQuery();
                    }
                }
                else
                {
                    // Update existing variation (Reset sent flag and update timestamp)
                    string sUpd = "UPDATE GesVariazioni SET var_dti=GETDATE(), var_dtv=GETDATE(), var_ori=@ori, var_inv=@inv WHERE var_tip=@tip AND var_art=@art AND var_num=@num";
                    using (SqlCommand cmUpd = new SqlCommand(sUpd, cn, tr))
                    {
                        cmUpd.Parameters.AddWithValue("@tip", sTip);
                        cmUpd.Parameters.AddWithValue("@art", sArt);
                        cmUpd.Parameters.AddWithValue("@num", _clsDef.COD03Z);
                        cmUpd.Parameters.AddWithValue("@ori", sOri);
                        cmUpd.Parameters.AddWithValue("@inv", _clsDef.DIVDAD);
                        cmUpd.ExecuteNonQuery();
                    }
                }
            }
        }

        private void UpdateSqlFieldTx(ref string sSql, ref bool bFirst, string fldName, object val, List<SqlParameter> list, DataTable schemaTable = null)
        {
            if (val != null && val != DBNull.Value && !string.IsNullOrEmpty(val.ToString()))
            {
                string sVal = val.ToString().Trim();
                // Assicura che i campi numerici decimali vengano convertiti correttamente con punto per SQL
                if (fldName == "art_pne" || fldName == "old_cos" || fldName == "new_cos" || fldName == "art_tar" || fldName == "art_gsc")
                    sVal = sVal.Replace(",", ".");

                // Per i campi INTERI / CODICI, se presente un decimale (es. .00), lo tronchiamo per evitare errori di conversione SQL
                if (fldName == "art_pxc" || fldName == "art_rep" || fldName == "art_cat" || fldName == "art_iva" || fldName == "art_tas" || fldName == "art_reb" || fldName == "art_ec1" || fldName == "art_ec2" || fldName == "art_ec3")
                {
                    sVal = sVal.Replace(",", ".");
                    if (sVal.Contains(".")) sVal = sVal.Split('.')[0];
                }

                if (fldName == "art_ec1" || fldName == "art_ec2" || fldName == "art_ec3")
                {
                    if (int.TryParse(sVal, out int iEc) && iEc >= 0)
                        sVal = iEc.ToString("D3");
                }
                else if (fldName == "art_rep")
                {
                    if (int.TryParse(sVal, out int iRepUpd) && iRepUpd > 0)
                        sVal = iRepUpd.ToString("D3");
                    else
                        sVal = "001";
                }
                else if (fldName == "art_reb")
                {
                    if (int.TryParse(sVal, out int iRebUpd) && iRebUpd > 0)
                        sVal = iRebUpd.ToString();
                }
                else if (fldName == "art_tas")
                {
                    if (int.TryParse(sVal, out int iTasUpd) && iTasUpd >= 0)
                        sVal = iTasUpd.ToString();
                }
                else if (fldName == "art_plu")
                {
                    if (int.TryParse(sVal, out int iPluUpd) && iPluUpd > 0)
                        sVal = iPluUpd.ToString("D4");
                }

                if (schemaTable != null && schemaTable.Columns.Contains(fldName))
                {
                    int maxLen = schemaTable.Columns[fldName].MaxLength;
                    if (maxLen > 0 && sVal.Length > maxLen)
                    {
                        sVal = sVal.Substring(0, maxLen);
                    }
                }

                // Avoid duplicate parameter and column in UPDATE statement
                SqlParameter existingParam = list.Find(p => p.ParameterName.Equals("@" + fldName, StringComparison.OrdinalIgnoreCase));
                if (existingParam != null)
                {
                    existingParam.Value = sVal;
                }
                else
                {
                    if (!bFirst) sSql += ", ";
                    sSql += fldName + "=@" + fldName;
                    list.Add(new SqlParameter("@" + fldName, sVal));
                    bFirst = false;
                }
            }
        }

        private string TruncateField(DataTable schemaTable, string colName, string val)
        {
            if (string.IsNullOrEmpty(val)) return "";
            if (schemaTable != null && schemaTable.Columns.Contains(colName))
            {
                int maxLen = schemaTable.Columns[colName].MaxLength;
                if (maxLen > 0 && val.Length > maxLen)
                {
                    return val.Substring(0, maxLen);
                }
            }
            return val;
        }

        private struct IvaItem
        {
            public string Codice;
            public decimal Aliquota;
            public string Descrizione;
        }

        private List<IvaItem> _listTabIva = new List<IvaItem>();

        private void LoadTabIva()
        {
            _listTabIva.Clear();
            try
            {
                string sSql = "SELECT tab_cod, tab_des, tab_ali FROM TabIva WHERE tab_ann=0 ORDER BY tab_ali DESC";
                DataTable t = _clsFun.FillTabSql("TabIva", sSql, false, _strConSql);
                if (t != null)
                {
                    foreach (DataRow r in t.Rows)
                    {
                        IvaItem item = new IvaItem();
                        item.Codice = r["tab_cod"] != DBNull.Value ? r["tab_cod"].ToString().Trim() : "";
                        item.Descrizione = r["tab_des"] != DBNull.Value ? r["tab_des"].ToString().Trim().ToUpper() : "";
                        item.Aliquota = r["tab_ali"] != DBNull.Value ? Convert.ToDecimal(r["tab_ali"]) : 0m;
                        _listTabIva.Add(item);
                    }
                }
            }
            catch { }
        }

        private string ResolveIvaCode(string sRawIva)
        {
            if (_listTabIva.Count == 0) LoadTabIva();

            if (string.IsNullOrEmpty(sRawIva))
            {
                var default22 = _listTabIva.FirstOrDefault(x => Math.Abs(x.Aliquota - 22m) < 0.01m);
                if (!string.IsNullOrEmpty(default22.Codice)) return Format3Digits(default22.Codice, "022");
                if (_listTabIva.Count > 0) return Format3Digits(_listTabIva[0].Codice, "022");
                return "022";
            }

            string sClean = sRawIva.Trim();

            string sNumOnly = System.Text.RegularExpressions.Regex.Replace(sClean, @"[^\d,\.]", "").Trim();
            if (!string.IsNullOrEmpty(sNumOnly))
            {
                sNumOnly = sNumOnly.Replace(",", ".");
                if (decimal.TryParse(sNumOnly, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal dParsedAli))
                {
                    if (dParsedAli > 0m && dParsedAli < 1m) dParsedAli = dParsedAli * 100m;

                    var matchAli = _listTabIva.FirstOrDefault(x => Math.Abs(x.Aliquota - dParsedAli) < 0.01m);
                    if (!string.IsNullOrEmpty(matchAli.Codice))
                    {
                        return Format3Digits(matchAli.Codice, "022");
                    }
                }
            }

            foreach (var item in _listTabIva)
            {
                if (string.Equals(item.Codice, sClean, StringComparison.OrdinalIgnoreCase))
                {
                    return Format3Digits(item.Codice, "022");
                }
            }

            string sUpper = sClean.ToUpper();
            foreach (var item in _listTabIva)
            {
                if (!string.IsNullOrEmpty(item.Descrizione) && item.Descrizione.Contains(sUpper))
                {
                    return Format3Digits(item.Codice, "022");
                }
            }

            var fb22 = _listTabIva.FirstOrDefault(x => Math.Abs(x.Aliquota - 22m) < 0.01m);
            if (!string.IsNullOrEmpty(fb22.Codice)) return Format3Digits(fb22.Codice, "022");
            if (_listTabIva.Count > 0) return Format3Digits(_listTabIva[0].Codice, "022");

            return Format3Digits(sClean, "022");
        }

        private struct RepItem
        {
            public string Codice;
            public string Descrizione;
        }

        private List<RepItem> _listTabRep = new List<RepItem>();

        private void LoadTabReparti()
        {
            _listTabRep.Clear();
            try
            {
                string sSql = "SELECT tab_cod, tab_des FROM TabReparti WHERE tab_ann=0 ORDER BY tab_cod";
                DataTable t = _clsFun.FillTabSql("TabReparti", sSql, false, _strConSql);
                if (t != null)
                {
                    foreach (DataRow r in t.Rows)
                    {
                        RepItem item = new RepItem();
                        item.Codice = r["tab_cod"] != DBNull.Value ? r["tab_cod"].ToString().Trim() : "";
                        item.Descrizione = r["tab_des"] != DBNull.Value ? r["tab_des"].ToString().Trim().ToUpper() : "";
                        if (!string.IsNullOrEmpty(item.Codice))
                        {
                            _listTabRep.Add(item);
                        }
                    }
                }
            }
            catch { }
        }

        private string ResolveRepCode(string sRawRep)
        {
            if (_listTabRep.Count == 0) LoadTabReparti();

            if (string.IsNullOrEmpty(sRawRep))
            {
                if (_listTabRep.Count > 0)
                {
                    return Format3Digits(_listTabRep[0].Codice.Trim(), "001");
                }
                return "001";
            }

            string sClean = sRawRep.Trim();
            string sPadded3 = Format3Digits(sClean, "001");
            string sPadded2 = sClean;

            if (int.TryParse(sClean, out int iVal) && iVal >= 0)
            {
                sPadded2 = iVal.ToString("D2");
            }

            foreach (var item in _listTabRep)
            {
                string c = item.Codice.Trim();
                if (string.Equals(c, sPadded3, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(c, sPadded2, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(c, sClean, StringComparison.OrdinalIgnoreCase))
                {
                    return Format3Digits(c, "001");
                }
            }

            string sUpper = sClean.ToUpper();
            foreach (var item in _listTabRep)
            {
                if (string.Equals(item.Descrizione, sUpper, StringComparison.OrdinalIgnoreCase))
                {
                    return Format3Digits(item.Codice.Trim(), "001");
                }
            }

            if (sUpper.Length >= 3)
            {
                foreach (var item in _listTabRep)
                {
                    if (!string.IsNullOrEmpty(item.Descrizione) && (item.Descrizione.Contains(sUpper) || sUpper.Contains(item.Descrizione)))
                    {
                        return Format3Digits(item.Codice.Trim(), "001");
                    }
                }
            }

            if (int.TryParse(sClean, out int iFallback) && iFallback >= 0)
                return iFallback.ToString("D3");

            if (_listTabRep.Count > 0)
            {
                return Format3Digits(_listTabRep[0].Codice.Trim(), "001");
            }

            return Format3Digits(sClean, "001");
        }

        private struct LookupItem
        {
            public string Codice;
            public string Descrizione;
        }
        private List<LookupItem> _listTabReb = new List<LookupItem>();
        private List<LookupItem> _listTabOri = new List<LookupItem>();
        private List<LookupItem> _listTabCal = new List<LookupItem>();
        private List<LookupItem> _listTabCatBil = new List<LookupItem>();

        private void LoadTabBilanceLookup()
        {
            try
            {
                _listTabReb.Clear();
                DataTable tReb = null;
                try { tReb = _clsFun.FillTabSql("TabRepBilance", "SELECT tab_cod, tab_des FROM TabRepBilance ORDER BY tab_cod", false, _strConSql); } catch { }
                if (tReb == null || tReb.Rows.Count == 0)
                {
                    try { tReb = _clsFun.FillTabSql("TabRepBilance", "SELECT tab_cod, tab_des FROM TabTabelle WHERE tab_tip='REB' ORDER BY tab_cod", false, _strConSql); } catch { }
                }
                if (tReb != null)
                {
                    foreach (DataRow r in tReb.Rows)
                    {
                        if (tReb.Columns.Contains("tab_cod") && tReb.Columns.Contains("tab_des"))
                            _listTabReb.Add(new LookupItem { Codice = r["tab_cod"] != DBNull.Value ? r["tab_cod"].ToString().Trim() : "", Descrizione = r["tab_des"] != DBNull.Value ? r["tab_des"].ToString().Trim().ToUpper() : "" });
                    }
                }

                _listTabOri.Clear();
                DataTable tOri = null;
                try { tOri = _clsFun.FillTabSql("TabOrigine", "SELECT tab_cod, tab_des FROM TabOrigine ORDER BY tab_cod", false, _strConSql); } catch { }
                if (tOri == null || tOri.Rows.Count == 0)
                {
                    try { tOri = _clsFun.FillTabSql("TabOrigine", "SELECT tab_cod, tab_des FROM TabTabelle WHERE tab_tip='ORI' ORDER BY tab_cod", false, _strConSql); } catch { }
                }
                if (tOri != null)
                {
                    foreach (DataRow r in tOri.Rows)
                    {
                        if (tOri.Columns.Contains("tab_cod") && tOri.Columns.Contains("tab_des"))
                            _listTabOri.Add(new LookupItem { Codice = r["tab_cod"] != DBNull.Value ? r["tab_cod"].ToString().Trim() : "", Descrizione = r["tab_des"] != DBNull.Value ? r["tab_des"].ToString().Trim().ToUpper() : "" });
                    }
                }

                _listTabCal.Clear();
                DataTable tCal = null;
                try { tCal = _clsFun.FillTabSql("TabCalibro", "SELECT tab_cod, tab_des FROM TabCalibro ORDER BY tab_cod", false, _strConSql); } catch { }
                if (tCal == null || tCal.Rows.Count == 0)
                {
                    try { tCal = _clsFun.FillTabSql("TabCalibro", "SELECT tab_cod, tab_des FROM TabTabelle WHERE tab_tip='CAL' ORDER BY tab_cod", false, _strConSql); } catch { }
                }
                if (tCal != null)
                {
                    foreach (DataRow r in tCal.Rows)
                    {
                        if (tCal.Columns.Contains("tab_cod") && tCal.Columns.Contains("tab_des"))
                            _listTabCal.Add(new LookupItem { Codice = r["tab_cod"] != DBNull.Value ? r["tab_cod"].ToString().Trim() : "", Descrizione = r["tab_des"] != DBNull.Value ? r["tab_des"].ToString().Trim().ToUpper() : "" });
                    }
                }

                _listTabCatBil.Clear();
                DataTable tCat = null;
                try { tCat = _clsFun.FillTabSql("TabCategoria", "SELECT tab_cod, tab_des FROM TabCategoria ORDER BY tab_cod", false, _strConSql); } catch { }
                if (tCat == null || tCat.Rows.Count == 0)
                {
                    try { tCat = _clsFun.FillTabSql("TabCategoria", "SELECT tab_cod, tab_des FROM TabTabelle WHERE tab_tip='CAT' ORDER BY tab_cod", false, _strConSql); } catch { }
                }
                if (tCat != null)
                {
                    foreach (DataRow r in tCat.Rows)
                    {
                        if (tCat.Columns.Contains("tab_cod") && tCat.Columns.Contains("tab_des"))
                            _listTabCatBil.Add(new LookupItem { Codice = r["tab_cod"] != DBNull.Value ? r["tab_cod"].ToString().Trim() : "", Descrizione = r["tab_des"] != DBNull.Value ? r["tab_des"].ToString().Trim().ToUpper() : "" });
                    }
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "LoadTabBilanceLookup");
            }
        }

        private string ResolveLookupCode(string sRaw, List<LookupItem> list)
        {
            if (string.IsNullOrEmpty(sRaw)) return "";
            string sClean = sRaw.Trim();
            string sUpper = sClean.ToUpper();

            foreach (var item in list)
            {
                if (string.Equals(item.Codice, sClean, StringComparison.OrdinalIgnoreCase))
                    return item.Codice;
            }

            foreach (var item in list)
            {
                if (string.Equals(item.Descrizione, sUpper, StringComparison.OrdinalIgnoreCase))
                    return item.Codice;
            }

            if (sUpper.Length >= 2)
            {
                foreach (var item in list)
                {
                    if (!string.IsNullOrEmpty(item.Descrizione) && (item.Descrizione.Contains(sUpper) || sUpper.Contains(item.Descrizione)))
                        return item.Codice;
                }
            }

            return sClean;
        }

        private struct EcrItem
        {
            public string Codice;
            public string Descrizione;
            public string Lv1;
            public string Lv2;
        }

        private List<EcrItem> _listTabEc1 = new List<EcrItem>();
        private List<EcrItem> _listTabEc2 = new List<EcrItem>();
        private List<EcrItem> _listTabEc3 = new List<EcrItem>();

        private void LoadTabEcr()
        {
            _listTabEc1.Clear();
            _listTabEc2.Clear();
            _listTabEc3.Clear();
            try
            {
                DataTable t1 = _clsFun.FillTabSql("TabEcrLv1", "SELECT tab_cod, tab_des FROM TabEcrLv1 ORDER BY tab_cod", false, _strConSql);
                if (t1 != null)
                {
                    foreach (DataRow r in t1.Rows)
                    {
                        _listTabEc1.Add(new EcrItem { Codice = r["tab_cod"] != null ? r["tab_cod"].ToString().Trim() : "", Descrizione = r["tab_des"] != null ? r["tab_des"].ToString().Trim().ToUpper() : "" });
                    }
                }

                DataTable t2 = _clsFun.FillTabSql("TabEcrLv2", "SELECT tab_cod, tab_des, tab_lv1 FROM TabEcrLv2 ORDER BY tab_cod", false, _strConSql);
                if (t2 != null)
                {
                    foreach (DataRow r in t2.Rows)
                    {
                        _listTabEc2.Add(new EcrItem { Codice = r["tab_cod"] != null ? r["tab_cod"].ToString().Trim() : "", Descrizione = r["tab_des"] != null ? r["tab_des"].ToString().Trim().ToUpper() : "", Lv1 = (t2.Columns.Contains("tab_lv1") && r["tab_lv1"] != DBNull.Value) ? r["tab_lv1"].ToString().Trim() : "" });
                    }
                }

                DataTable t3 = _clsFun.FillTabSql("TabEcrLv3", "SELECT tab_cod, tab_des, tab_lv1, tab_lv2 FROM TabEcrLv3 ORDER BY tab_cod", false, _strConSql);
                if (t3 != null)
                {
                    foreach (DataRow r in t3.Rows)
                    {
                        _listTabEc3.Add(new EcrItem { Codice = r["tab_cod"] != null ? r["tab_cod"].ToString().Trim() : "", Descrizione = r["tab_des"] != null ? r["tab_des"].ToString().Trim().ToUpper() : "", Lv1 = (t3.Columns.Contains("tab_lv1") && r["tab_lv1"] != DBNull.Value) ? r["tab_lv1"].ToString().Trim() : "", Lv2 = (t3.Columns.Contains("tab_lv2") && r["tab_lv2"] != DBNull.Value) ? r["tab_lv2"].ToString().Trim() : "" });
                    }
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "LoadTabEcr");
            }
        }

        private string ResolveEc1Code(string sRaw)
        {
            if (string.IsNullOrEmpty(sRaw)) return "";
            if (_listTabEc1.Count == 0) LoadTabEcr();

            string sClean = sRaw.Trim();
            string sPad3 = Format3Digits(sClean, "");

            // 1. Direct code match
            foreach (var item in _listTabEc1)
            {
                if (string.Equals(item.Codice, sPad3, StringComparison.OrdinalIgnoreCase) || string.Equals(item.Codice, sClean, StringComparison.OrdinalIgnoreCase))
                    return Format3Digits(item.Codice, "000");
            }

            // 2. Description match
            string sUpper = sClean.ToUpper();
            foreach (var item in _listTabEc1)
            {
                if (string.Equals(item.Descrizione, sUpper, StringComparison.OrdinalIgnoreCase))
                    return Format3Digits(item.Codice, "000");
            }
            foreach (var item in _listTabEc1)
            {
                if (!string.IsNullOrEmpty(item.Descrizione) && (item.Descrizione.Contains(sUpper) || sUpper.Contains(item.Descrizione)))
                    return Format3Digits(item.Codice, "000");
            }

            if (int.TryParse(sClean, out int iVal)) return iVal.ToString("D3");
            return sClean.Length <= 3 ? sClean : sClean.Substring(0, 3);
        }

        private string ResolveEc2Code(string sRaw, string sLv1 = "")
        {
            if (string.IsNullOrEmpty(sRaw)) return "";
            if (_listTabEc2.Count == 0) LoadTabEcr();

            string sClean = sRaw.Trim();
            string sPad3 = Format3Digits(sClean, "");
            string sLv1Pad3 = Format3Digits(sLv1, "");

            // 1. Match code filtered by Lv1
            foreach (var item in _listTabEc2)
            {
                if (!string.IsNullOrEmpty(sLv1) && !string.IsNullOrEmpty(item.Lv1) && !string.Equals(Format3Digits(item.Lv1, ""), sLv1Pad3, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (string.Equals(item.Codice, sPad3, StringComparison.OrdinalIgnoreCase) || string.Equals(item.Codice, sClean, StringComparison.OrdinalIgnoreCase))
                    return Format3Digits(item.Codice, "000");
            }

            // 2. Match description filtered by Lv1
            string sUpper = sClean.ToUpper();
            foreach (var item in _listTabEc2)
            {
                if (!string.IsNullOrEmpty(sLv1) && !string.IsNullOrEmpty(item.Lv1) && !string.Equals(Format3Digits(item.Lv1, ""), sLv1Pad3, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (string.Equals(item.Descrizione, sUpper, StringComparison.OrdinalIgnoreCase) || (!string.IsNullOrEmpty(item.Descrizione) && (item.Descrizione.Contains(sUpper) || sUpper.Contains(item.Descrizione))))
                    return Format3Digits(item.Codice, "000");
            }

            // 3. Fallback without Lv1 filter
            foreach (var item in _listTabEc2)
            {
                if (string.Equals(item.Codice, sPad3, StringComparison.OrdinalIgnoreCase) || string.Equals(item.Codice, sClean, StringComparison.OrdinalIgnoreCase))
                    return Format3Digits(item.Codice, "000");
            }

            if (int.TryParse(sClean, out int iVal)) return iVal.ToString("D3");
            return sClean.Length <= 3 ? sClean : sClean.Substring(0, 3);
        }

        private string ResolveEc3Code(string sRaw, string sLv1 = "", string sLv2 = "")
        {
            if (string.IsNullOrEmpty(sRaw)) return "";
            if (_listTabEc3.Count == 0) LoadTabEcr();

            string sClean = sRaw.Trim();
            string sPad3 = Format3Digits(sClean, "");
            string sLv1Pad3 = Format3Digits(sLv1, "");
            string sLv2Pad3 = Format3Digits(sLv2, "");

            // 1. Match code filtered by Lv1 and Lv2
            foreach (var item in _listTabEc3)
            {
                if (!string.IsNullOrEmpty(sLv1) && !string.IsNullOrEmpty(item.Lv1) && !string.Equals(Format3Digits(item.Lv1, ""), sLv1Pad3, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!string.IsNullOrEmpty(sLv2) && !string.IsNullOrEmpty(item.Lv2) && !string.Equals(Format3Digits(item.Lv2, ""), sLv2Pad3, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (string.Equals(item.Codice, sPad3, StringComparison.OrdinalIgnoreCase) || string.Equals(item.Codice, sClean, StringComparison.OrdinalIgnoreCase))
                    return Format3Digits(item.Codice, "000");
            }

            // 2. Match description
            string sUpper = sClean.ToUpper();
            foreach (var item in _listTabEc3)
            {
                if (!string.IsNullOrEmpty(sLv1) && !string.IsNullOrEmpty(item.Lv1) && !string.Equals(Format3Digits(item.Lv1, ""), sLv1Pad3, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!string.IsNullOrEmpty(sLv2) && !string.IsNullOrEmpty(item.Lv2) && !string.Equals(Format3Digits(item.Lv2, ""), sLv2Pad3, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (string.Equals(item.Descrizione, sUpper, StringComparison.OrdinalIgnoreCase) || (!string.IsNullOrEmpty(item.Descrizione) && (item.Descrizione.Contains(sUpper) || sUpper.Contains(item.Descrizione))))
                    return Format3Digits(item.Codice, "000");
            }

            if (int.TryParse(sClean, out int iVal)) return iVal.ToString("D3");
            return sClean.Length <= 3 ? sClean : sClean.Substring(0, 3);
        }

        public static bool AreDescriptionsSimilar(string des1, string des2)
        {
            if (string.IsNullOrWhiteSpace(des1) || string.IsNullOrWhiteSpace(des2)) return false;
            string s1 = des1.Trim().ToUpper();
            string s2 = des2.Trim().ToUpper();
            if (s1 == s2) return true;
            if (s1.StartsWith(s2) || s2.StartsWith(s1)) return true;

            HashSet<string> stopWords = new HashSet<string> { 
                "CON", "PER", "DEL", "DEI", "DEGLI", "DELLE", "DELLA", "ALLA", "ALLE", "ALLO", "AGLI", 
                "CAT", "CAT.", "ORIGINE", "ORIG.", "CONF", "CONF.", "SFUSO", "SFUSA", "KG", "PZ", "GR", 
                "GR.", "THE", "AND", "EST", "SUD", "NORD", "DOC", "DOP", "IGP", "BIO", "VASO", "VASI", "CEST", "CESTINO" 
            };

            char[] seps = new char[] { ' ', ',', '.', ';', ':', '-', '/', '(', ')', '[', ']', '+', '*', '\\', '%' };
            string[] words1 = s1.Split(seps, StringSplitOptions.RemoveEmptyEntries);
            string[] words2 = s2.Split(seps, StringSplitOptions.RemoveEmptyEntries);

            foreach (string w1 in words1)
            {
                if (w1.Length >= 3 && !stopWords.Contains(w1))
                {
                    foreach (string w2 in words2)
                    {
                        if (w2.Length >= 3 && !stopWords.Contains(w2))
                        {
                            if (w1 == w2 || (w1.Length >= 4 && w2.Length >= 4 && (w1.StartsWith(w2) || w2.StartsWith(w1))))
                                return true;
                        }
                    }
                }
            }
            return false;
        }

        private string ResolveRebCode(string sRaw)
        {
            if (string.IsNullOrEmpty(sRaw)) return "";
            string sClean = sRaw.Trim();
            string sUpper = sClean.ToUpper();
            if (sUpper.Contains("ORTOFRUTTA") || sUpper.Contains("FRUTTA") || sUpper.Contains("VERDURA") || sUpper == "2" || sUpper == "02") return "2";
            if (sUpper.Contains("GASTRONOMIA") || sUpper == "1" || sUpper == "01") return "1";
            if (sUpper.Contains("MACELLARIA") || sUpper.Contains("CARNE") || sUpper == "4" || sUpper == "04") return "4";
            if (sUpper.Contains("PANETTERIA") || sUpper.Contains("PANE") || sUpper == "5" || sUpper == "05") return "5";
            if (sUpper.Contains("PESCHERIA") || sUpper.Contains("PESCE") || sUpper == "3" || sUpper == "03") return "3";

            if (_listTabReb.Count == 0) LoadTabBilanceLookup();
            string res = ResolveLookupCode(sRaw, _listTabReb);
            if (int.TryParse(res, out int iReb) && iReb > 0) return iReb.ToString();
            return res.Length > 1 ? res.Substring(0, 1) : res;
        }

        private string ResolveOriCode(string sRaw)
        {
            if (string.IsNullOrEmpty(sRaw)) return "";
            if (_listTabOri.Count == 0) LoadTabBilanceLookup();
            string sClean = sRaw.Trim();
            string sUpper = sClean.ToUpper();

            string res = ResolveLookupCode(sClean, _listTabOri);
            if (!string.IsNullOrEmpty(res) && res != sClean) return res;

            if (sUpper.Contains("ITALIA") || sUpper == "IT" || sUpper.StartsWith("IT ")) return "001";
            if (sUpper.Contains("OLANDA") || sUpper.Contains("PAESI BASSI") || sUpper.Contains("NL")) return "NLD";
            if (sUpper.Contains("SPAGNA") || sUpper.Contains("ES")) return "ESP";
            if (sUpper.Contains("SUDAFRICA") || sUpper.Contains("ZA")) return "ZAF";
            if (sUpper.Contains("EGITTO") || sUpper.Contains("EG")) return "EGY";
            if (sUpper.Contains("ECUADOR") || sUpper.Contains("EC")) return "ECU";
            if (sUpper.Contains("BRASILE") || sUpper.Contains("BR")) return "BRA";
            if (sUpper.Contains("PERU") || sUpper.Contains("PERÙ") || sUpper.Contains("PE")) return "PER";
            if (sUpper.Contains("COSTA RICA") || sUpper.Contains("CR")) return "CRI";
            if (sUpper.Contains("CILE") || sUpper.Contains("CL")) return "CHL";
            if (sUpper.Contains("FRANCIA") || sUpper.Contains("FR")) return "FRA";
            if (sUpper.Contains("GRECIA") || sUpper.Contains("GR")) return "GRC";
            if (sUpper.Contains("GERMANIA") || sUpper.Contains("DE")) return "DEU";
            if (sUpper.Contains("ARGENTINA") || sUpper.Contains("AR")) return "ARG";
            if (sUpper.Contains("NUOVA ZELANDA") || sUpper.Contains("NZ")) return "NZL";
            if (sUpper.Contains("POLONIA") || sUpper.Contains("PL")) return "POL";
            if (sUpper.Contains("TURCHIA") || sUpper.Contains("TR")) return "TUR";
            if (sUpper.Contains("MAROCCO") || sUpper.Contains("MA")) return "MAR";
            if (sUpper.Contains("MESSICO") || sUpper.Contains("MX")) return "MEX";
            if (sUpper.Contains("USA") || sUpper.Contains("STATI UNITI")) return "USA";

            return sClean.Length > 3 ? sClean.Substring(0, 3) : sClean;
        }

        private string ResolveCalCode(string sRaw)
        {
            if (string.IsNullOrEmpty(sRaw)) return "";
            string sClean = sRaw.Trim();
            if (int.TryParse(sClean, out int iCal) && iCal > 0) return iCal.ToString("D3");
            if (_listTabCal.Count == 0) LoadTabBilanceLookup();
            return ResolveLookupCode(sRaw, _listTabCal);
        }

        private string ResolveCatBilCode(string sRaw)
        {
            if (string.IsNullOrEmpty(sRaw)) return "";
            string sClean = sRaw.Trim();
            string sUpper = sClean.ToUpper();
            if (sUpper.Contains("III") || sUpper.Contains("TERZA") || sUpper == "3") return "3";
            if (sUpper.Contains("II") || sUpper.Contains("SECONDA") || sUpper == "2") return "2";
            if (sUpper.Contains("I") || sUpper.Contains("PRIMA") || sUpper == "1") return "1";

            if (_listTabCatBil.Count == 0) LoadTabBilanceLookup();
            return ResolveLookupCode(sRaw, _listTabCatBil);
        }

        private string ResolveTas(string sRaw)
        {
            if (string.IsNullOrEmpty(sRaw)) return "";
            string sClean = sRaw.Trim();
            if (int.TryParse(sClean, out int iTas) && iTas >= 0)
            {
                return iTas.ToString();
            }
            return sClean.Length > 4 ? sClean.Substring(0, 4) : sClean;
        }

        private string ResolvePlu(string sRaw)
        {
            if (string.IsNullOrEmpty(sRaw)) return "";
            string sClean = sRaw.Trim();
            if (int.TryParse(sClean, out int iPlu) && iPlu > 0)
            {
                return iPlu.ToString("D4");
            }
            return sClean.Length > 6 ? sClean.Substring(0, 6) : sClean;
        }

        private string ResolveTraCode(string sRaw)
        {
            if (string.IsNullOrEmpty(sRaw)) return "";
            string s = sRaw.Trim().ToUpper();
            if (s == "1" || s == "S" || s == "SI" || s == "TRUE" || s == "YES" || s == "LOTTO") return "S";
            return "";
        }

        private bool ResolveBpz(string sRaw)
        {
            if (string.IsNullOrEmpty(sRaw)) return false;
            string s = sRaw.Trim().ToUpper();
            if (s == "1" || s == "S" || s == "SI" || s == "TRUE" || s == "YES" || s == "PZ" || s == "PEZZO" || s == "A PEZZO") return true;
            return false;
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PerformExit();
        }

        public static string GenerateScaleEan13(string sPluRaw, string sPrefix = "2")
        {
            if (string.IsNullOrEmpty(sPluRaw)) return "";
            if (string.IsNullOrEmpty(sPrefix)) sPrefix = "2";

            string sNumPlu = System.Text.RegularExpressions.Regex.Replace(sPluRaw.Trim(), @"[^\d]", "");
            if (string.IsNullOrEmpty(sNumPlu) || !int.TryParse(sNumPlu, out int iPlu) || iPlu <= 0)
                return "";

            // EAN-13: 13 cifre totali = Prefisso + PLU (zeri a sinistra) + "00000" (5 zeri per peso/prezzo) + 1 Cifra Controllo Modulo 10
            // Base12 = Prefisso.Length + PluLen + 5 = 12 -> PluLen = 7 - Prefisso.Length
            int itemLen = 7 - sPrefix.Length;
            if (itemLen < 1) itemLen = 4;

            string sPaddedPlu = iPlu.ToString().PadLeft(itemLen, '0');
            if (sPaddedPlu.Length > itemLen)
                sPaddedPlu = sPaddedPlu.Substring(sPaddedPlu.Length - itemLen);

            string sBase12 = sPrefix + sPaddedPlu + "00000";
            if (sBase12.Length > 12) sBase12 = sBase12.Substring(0, 12);
            else if (sBase12.Length < 12) sBase12 = sBase12.PadRight(12, '0');

            string checkDigit = new clsCtrlCodici().FindMod10Digit(sBase12);
            return sBase12 + checkDigit;
        }

        private bool CheckArticleExists(string sArt, SqlConnection cn, SqlTransaction tr)
        {
            if (string.IsNullOrEmpty(sArt)) return false;
            try
            {
                using (SqlCommand cm = new SqlCommand("SELECT COUNT(*) FROM AnaArticoli WHERE art_cod=@cod", cn, tr))
                {
                    cm.Parameters.AddWithValue("@cod", sArt);
                    return Convert.ToInt32(cm.ExecuteScalar()) > 0;
                }
            }
            catch { return false; }
        }

        private string FindArticleByBarcode(string sEan, SqlConnection cn, SqlTransaction tr)
        {
            if (string.IsNullOrEmpty(sEan)) return "";
            try
            {
                using (SqlCommand cm = new SqlCommand("SELECT TOP 1 ean_art FROM AnaBarcode WHERE ean_ean=@ean", cn, tr))
                {
                    cm.Parameters.AddWithValue("@ean", sEan.Trim());
                    object res = cm.ExecuteScalar();
                    return (res != null && res != DBNull.Value) ? res.ToString().Trim() : "";
                }
            }
            catch { return ""; }
        }

        private bool IsPlaceholderCode(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return true;
            string s = code.Trim().ToUpper();
            return s == "[NUOVO]" || s == "NUOVO" || s == "[NEW]" || s == "NEW" || s == "N" || s == "[N]" || s == "NUOVA" || s == "[NUOVA]" || s == "AUTO" || s == "[AUTO]";
        }

        private string FindArticleBySupplierCode(string sFor, string sArf, SqlConnection cn, SqlTransaction tr)
        {
            if (string.IsNullOrEmpty(sArf)) return "";
            try
            {
                string sql = "SELECT TOP 1 lia_art FROM GesLisAcquisto WHERE lia_arf=@arf ";
                if (!string.IsNullOrEmpty(sFor)) sql += "AND (lia_for=@for OR lia_for='') ";
                sql += "AND lia_ann=0 ORDER BY lia_dti DESC";
                using (SqlCommand cm = new SqlCommand(sql, cn, tr))
                {
                    cm.Parameters.AddWithValue("@arf", sArf.Trim());
                    if (!string.IsNullOrEmpty(sFor)) cm.Parameters.AddWithValue("@for", sFor.Trim());
                    object res = cm.ExecuteScalar();
                    return (res != null && res != DBNull.Value) ? res.ToString().Trim() : "";
                }
            }
            catch { return ""; }
        }

        private string FindArticleByPlu(string sPlu, string sReb, SqlConnection cn, SqlTransaction tr)
        {
            if (string.IsNullOrWhiteSpace(sPlu)) return "";
            string sCleanPlu = sPlu.Trim();
            try
            {
                string sql = @"
                    SELECT TOP 1 art_cod 
                    FROM AnaArticoli 
                    WHERE (art_plu = @plu OR (ISNUMERIC(art_plu) = 1 AND ISNUMERIC(@plu) = 1 AND CAST(art_plu AS int) = CAST(@plu AS int))) 
                      AND (@reb = '' OR art_reb = @reb OR art_reb IS NULL OR art_reb = '')
                      AND art_sta = 'A' AND art_bil = 1 
                    ORDER BY art_dti DESC";
                using (SqlCommand cm = new SqlCommand(sql, cn, tr))
                {
                    cm.Parameters.AddWithValue("@plu", sCleanPlu);
                    cm.Parameters.AddWithValue("@reb", sReb != null ? sReb.Trim() : "");
                    object res = cm.ExecuteScalar();
                    return (res != null && res != DBNull.Value) ? res.ToString().Trim() : "";
                }
            }
            catch { return ""; }
        }

        private void EnsureBarcodeRecord(string sEan, string sArt, decimal dPrv, bool bIsBil, bool bApezz, DataTable tEanSchema, SqlConnection cn, SqlTransaction tr)
        {
            if (string.IsNullOrEmpty(sEan) || string.IsNullOrEmpty(sArt) || tEanSchema == null) return;
            string sCleanEan = sEan.Trim();
            if (string.IsNullOrEmpty(sCleanEan)) return;

            try
            {
                string sCheck = "SELECT COUNT(*) FROM AnaBarcode WHERE ean_ean=@ean";
                int count = 0;
                using (SqlCommand cm = new SqlCommand(sCheck, cn, tr))
                {
                    cm.Parameters.AddWithValue("@ean", sCleanEan);
                    count = Convert.ToInt32(cm.ExecuteScalar());
                }

                if (count == 0)
                {
                    try
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

                        if (tEanSchema.Columns.Contains("ean_ean")) rNewEan["ean_ean"] = TruncateField(tEanSchema, "ean_ean", sCleanEan);
                        if (tEanSchema.Columns.Contains("ean_art")) rNewEan["ean_art"] = TruncateField(tEanSchema, "ean_art", sArt);
                        if (tEanSchema.Columns.Contains("ean_qta")) rNewEan["ean_qta"] = 1;
                        if (tEanSchema.Columns.Contains("ean_dti"))
                        {
                            if (tEanSchema.Columns["ean_dti"].DataType == typeof(DateTime)) rNewEan["ean_dti"] = DateTime.Today;
                            else rNewEan["ean_dti"] = DateTime.Today.ToString("dd/MM/yyyy");
                        }
                        if (tEanSchema.Columns.Contains("ean_dtm"))
                        {
                            if (tEanSchema.Columns["ean_dtm"].DataType == typeof(DateTime)) rNewEan["ean_dtm"] = DateTime.Today;
                            else rNewEan["ean_dtm"] = DateTime.Today.ToString("dd/MM/yyyy");
                        }
                        if (tEanSchema.Columns.Contains("ean_ann")) rNewEan["ean_ann"] = 0;
                        if (tEanSchema.Columns.Contains("ean_prv")) rNewEan["ean_prv"] = dPrv;
                        if (tEanSchema.Columns.Contains("ean_bil")) rNewEan["ean_bil"] = bIsBil;
                        if (tEanSchema.Columns.Contains("ean_ecp")) rNewEan["ean_ecp"] = false;

                        string sSqlInsBar = _clsFun.SqlInsertRow("AnaBarcode", tEanSchema, rNewEan);
                        using (SqlCommand cmIns = new SqlCommand(sSqlInsBar, cn, tr))
                        {
                            cmIns.ExecuteNonQuery();
                        }
                        return;
                    }
                    catch (SqlException ex) when (ex.Number == 2601 || ex.Number == 2627)
                    {
                        // Conflict on duplicate key -> fallback to update below
                    }
                }

                string sUpd = "UPDATE AnaBarcode SET ean_art=@art, ean_prv=@prv, ean_dtm=GETDATE(), ean_bil=@bil, ean_ecp=@ecp WHERE ean_ean=@ean";
                using (SqlCommand cmUpd = new SqlCommand(sUpd, cn, tr))
                {
                    cmUpd.Parameters.AddWithValue("@art", sArt);
                    cmUpd.Parameters.AddWithValue("@prv", dPrv);
                    cmUpd.Parameters.AddWithValue("@bil", bIsBil);
                    cmUpd.Parameters.AddWithValue("@ecp", false);
                    cmUpd.Parameters.AddWithValue("@ean", sCleanEan);
                    cmUpd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "EnsureBarcodeRecord");
            }
        }

        private void EnsureLiaRecord(string sArt, string sFor, string sArf, decimal dCos, decimal dPxc, DataTable tLiaSchema, SqlConnection cn, SqlTransaction tr)
        {
            if (string.IsNullOrEmpty(sArt) || string.IsNullOrEmpty(sFor) || tLiaSchema == null) return;
            string sCleanArf = !string.IsNullOrEmpty(sArf) ? sArf.Trim() : sArt;
            try
            {
                string sCheck = "SELECT COUNT(*) FROM GesLisAcquisto WHERE lia_art=@art AND lia_for=@for AND lia_ann=0";
                int count = 0;
                using (SqlCommand cm = new SqlCommand(sCheck, cn, tr))
                {
                    cm.Parameters.AddWithValue("@art", sArt);
                    cm.Parameters.AddWithValue("@for", sFor);
                    count = Convert.ToInt32(cm.ExecuteScalar());
                }

                if (count == 0)
                {
                    try
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

                        if (tLiaSchema.Columns.Contains("lia_art")) rNewLia["lia_art"] = TruncateField(tLiaSchema, "lia_art", sArt);
                        if (tLiaSchema.Columns.Contains("lia_for")) rNewLia["lia_for"] = TruncateField(tLiaSchema, "lia_for", sFor);
                        if (tLiaSchema.Columns.Contains("lia_arf")) rNewLia["lia_arf"] = TruncateField(tLiaSchema, "lia_arf", sCleanArf);
                        if (tLiaSchema.Columns.Contains("lia_cos")) rNewLia["lia_cos"] = dCos;
                        if (tLiaSchema.Columns.Contains("lia_pxc")) rNewLia["lia_pxc"] = dPxc;
                        if (tLiaSchema.Columns.Contains("lia_cxp")) rNewLia["lia_cxp"] = 1;
                        if (tLiaSchema.Columns.Contains("lia_day"))
                        {
                            if (tLiaSchema.Columns["lia_day"].DataType == typeof(DateTime)) rNewLia["lia_day"] = DateTime.Today;
                            else rNewLia["lia_day"] = DateTime.Today.ToString("dd/MM/yyyy");
                        }
                        if (tLiaSchema.Columns.Contains("lia_dti"))
                        {
                            if (tLiaSchema.Columns["lia_dti"].DataType == typeof(DateTime)) rNewLia["lia_dti"] = DateTime.Today;
                            else rNewLia["lia_dti"] = DateTime.Today.ToString("dd/MM/yyyy");
                        }
                        if (tLiaSchema.Columns.Contains("lia_dtf"))
                        {
                            if (tLiaSchema.Columns["lia_dtf"].DataType == typeof(DateTime)) rNewLia["lia_dtf"] = _clsDef.DAYOUT;
                            else rNewLia["lia_dtf"] = _clsDef.DAYOUT.ToString("dd/MM/yyyy");
                        }
                        if (tLiaSchema.Columns.Contains("lia_tip")) rNewLia["lia_tip"] = "M";
                        if (tLiaSchema.Columns.Contains("lia_ann")) rNewLia["lia_ann"] = 0;

                        string sSqlInsLia = _clsFun.SqlInsertRow("GesLisAcquisto", tLiaSchema, rNewLia);
                        using (SqlCommand cmIns = new SqlCommand(sSqlInsLia, cn, tr))
                        {
                            cmIns.ExecuteNonQuery();
                        }
                        return;
                    }
                    catch (SqlException ex) when (ex.Number == 2601 || ex.Number == 2627)
                    {
                        // Fallback to update below
                    }
                }

                string sSqlUpdLia = "UPDATE GesLisAcquisto SET lia_cos=@cos, lia_arf=@arf, lia_pxc=@pxc, lia_cxp=1, lia_day=GETDATE(), lia_dtf=" + _clsFun.DaySql(_clsDef.DAYOUT) + " WHERE lia_art=@art AND lia_for=@for AND lia_ann=0";
                using (SqlCommand cmUpdLia = new SqlCommand(sSqlUpdLia, cn, tr))
                {
                    cmUpdLia.Parameters.AddWithValue("@cos", dCos);
                    cmUpdLia.Parameters.AddWithValue("@arf", sCleanArf);
                    cmUpdLia.Parameters.AddWithValue("@pxc", dPxc);
                    cmUpdLia.Parameters.AddWithValue("@art", sArt);
                    cmUpdLia.Parameters.AddWithValue("@for", sFor);
                    cmUpdLia.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "EnsureLiaRecord");
            }
        }

        private void EnsureLivRecord(string sArt, decimal dPrv, DataTable tLivSchema, SqlConnection cn, SqlTransaction tr)
        {
            if (string.IsNullOrEmpty(sArt) || dPrv <= 0 || tLivSchema == null) return;
            try
            {
                string sCheck = "SELECT COUNT(*) FROM GesLisVendita WHERE liv_art=@art AND liv_lis=@lis AND liv_ann=0";
                int count = 0;
                using (SqlCommand cm = new SqlCommand(sCheck, cn, tr))
                {
                    cm.Parameters.AddWithValue("@art", sArt);
                    cm.Parameters.AddWithValue("@lis", _clsDef.LISPOS);
                    count = Convert.ToInt32(cm.ExecuteScalar());
                }

                if (count == 0)
                {
                    try
                    {
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
                        if (tLivSchema.Columns.Contains("liv_art")) rNewLiv["liv_art"] = TruncateField(tLivSchema, "liv_art", sArt);
                        if (tLivSchema.Columns.Contains("liv_qta")) rNewLiv["liv_qta"] = 1;
                        if (tLivSchema.Columns.Contains("liv_prv")) rNewLiv["liv_prv"] = dPrv;
                        if (tLivSchema.Columns.Contains("liv_dti"))
                        {
                            if (tLivSchema.Columns["liv_dti"].DataType == typeof(DateTime)) rNewLiv["liv_dti"] = DateTime.Today;
                            else rNewLiv["liv_dti"] = DateTime.Today.ToString("dd/MM/yyyy");
                        }
                        if (tLivSchema.Columns.Contains("liv_dtf"))
                        {
                            if (tLivSchema.Columns["liv_dtf"].DataType == typeof(DateTime)) rNewLiv["liv_dtf"] = _clsDef.DAYOUT;
                            else rNewLiv["liv_dtf"] = _clsDef.DAYOUT.ToString("dd/MM/yyyy");
                        }
                        if (tLivSchema.Columns.Contains("liv_sta")) rNewLiv["liv_sta"] = "A";
                        if (tLivSchema.Columns.Contains("liv_ann")) rNewLiv["liv_ann"] = 0;

                        string sSqlInsLiv = _clsFun.SqlInsertRow("GesLisVendita", tLivSchema, rNewLiv);
                        using (SqlCommand cmIns = new SqlCommand(sSqlInsLiv, cn, tr))
                        {
                            cmIns.ExecuteNonQuery();
                        }
                        return;
                    }
                    catch (SqlException ex) when (ex.Number == 2601 || ex.Number == 2627)
                    {
                        // Fallback to update below
                    }
                }

                string sSqlUpdLiv = "UPDATE GesLisVendita SET liv_prv=@prv, liv_sta='A', liv_dtf=" + _clsFun.DaySql(_clsDef.DAYOUT) + " WHERE liv_art=@art AND liv_lis=@lis AND liv_ann=0";
                using (SqlCommand cmUpd = new SqlCommand(sSqlUpdLiv, cn, tr))
                {
                    cmUpd.Parameters.AddWithValue("@prv", dPrv);
                    cmUpd.Parameters.AddWithValue("@art", sArt);
                    cmUpd.Parameters.AddWithValue("@lis", _clsDef.LISPOS);
                    cmUpd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "EnsureLivRecord");
            }
        }
        private void UpdateExistingArticle(string sArt, DataRow r, string sRowFor, decimal dNewCos, decimal dNewPrv,
            DataTable tArtSchema, DataTable tEanSchema, DataTable tLiaSchema, DataTable tLivSchema,
            Dictionary<string, string> dictEan, Dictionary<string, string> dictArfToArt, Dictionary<string, string> dictDesToArt,
            SqlConnection cn, SqlTransaction tr)
        {
            if (string.IsNullOrEmpty(sArt)) return;

            string sArf = r["art_arf"] != null ? r["art_arf"].ToString().Trim() : "";
            string sArfKey = sRowFor + "|" + sArf;
            string sDesKey = r["art_des"] != null ? r["art_des"].ToString().Trim().ToUpper() : "";

            if (!string.IsNullOrEmpty(sArf) && !dictArfToArt.ContainsKey(sArfKey))
                dictArfToArt[sArfKey] = sArt;
            if (!string.IsNullOrEmpty(sDesKey) && sDesKey.Length >= 3 && !dictDesToArt.ContainsKey(sDesKey))
                dictDesToArt[sDesKey] = sArt;

                        string sPluUpd = (r.Table.Columns.Contains("plu") && r["plu"] != null) ? r["plu"].ToString().Trim() : "";
            string sRebUpd = (r.Table.Columns.Contains("reb") && r["reb"] != null) ? r["reb"].ToString().Trim() : "";
            string sEanUpd = r["art_ean"] != null ? r["art_ean"].ToString().Trim() : "";
            string sGenEanUpd = r.Table.Columns.Contains("ean_bil_gen") && r["ean_bil_gen"] != null ? r["ean_bil_gen"].ToString().Trim() : "";
            string sBilStrUpd = r["bil"] != null ? r["bil"].ToString().Trim() : "";
            bool bIsBilUpd = (!string.IsNullOrEmpty(sPluUpd)) ||
                             (!string.IsNullOrEmpty(sRebUpd)) ||
                             (!string.IsNullOrEmpty(sEanUpd) && sEanUpd.StartsWith("2")) ||
                             (!string.IsNullOrEmpty(sGenEanUpd) && sGenEanUpd.StartsWith("2")) ||
                             sBilStrUpd == "1" || sBilStrUpd.Equals("S", StringComparison.OrdinalIgnoreCase) ||
                             sBilStrUpd.Equals("SI", StringComparison.OrdinalIgnoreCase) ||
                             sBilStrUpd.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
                             sBilStrUpd.Equals("YES", StringComparison.OrdinalIgnoreCase);
            bool bBpzUpd = r.Table.Columns.Contains("bpz") && r["bpz"] != DBNull.Value ? Convert.ToBoolean(r["bpz"]) : false;

            if (!string.IsNullOrEmpty(sEanUpd))
            {
                bool bEanUpdIsBil = bIsBilUpd || sEanUpd.StartsWith("2");
                EnsureBarcodeRecord(sEanUpd, sArt, dNewPrv, bEanUpdIsBil, bBpzUpd, tEanSchema, cn, tr);
                dictEan[sEanUpd] = sArt;
            }
            if (!string.IsNullOrEmpty(sGenEanUpd) && sGenEanUpd != sEanUpd)
            {
                EnsureBarcodeRecord(sGenEanUpd, sArt, dNewPrv, true, bBpzUpd, tEanSchema, cn, tr);
                dictEan[sGenEanUpd] = sArt;
            }

            // 1. Update AnaArticoli
            string sSqlArt = "UPDATE AnaArticoli SET ";
            List<SqlParameter> listParams = new List<SqlParameter>();
            bool bFirst = true;

            string sDesUpd = r["art_des"] != null ? r["art_des"].ToString().Trim() : "";
            if (!string.IsNullOrEmpty(sDesUpd))
            {
                string sDebUpd = sDesUpd.Length > 20 ? sDesUpd.Substring(0, 20) : sDesUpd;
                sSqlArt += "art_des=@des, art_deb=@deb";
                listParams.Add(new SqlParameter("@des", sDesUpd));
                listParams.Add(new SqlParameter("@deb", sDebUpd));
                bFirst = false;
            }
            else
            {
                sSqlArt += "art_deb = ISNULL(NULLIF(art_deb, ''), LEFT(art_des, 20))";
                bFirst = false;
            }

            if (chkSetActiveStatus.Checked)
            {
                if (!bFirst) sSqlArt += ", ";
                sSqlArt += "art_sta='A'";
                bFirst = false;
            }

            if (!bFirst) sSqlArt += ", ";
            sSqlArt += "art_umc = ISNULL(NULLIF(art_umc, ''), 'NR'), art_umi = ISNULL(NULLIF(art_umi, ''), 'NR'), art_tgr = ISNULL(NULLIF(art_tgr, ''), 'PZ')";
            bFirst = false;

            if (dNewCos > 0)
            {
                if (!bFirst) sSqlArt += ", ";
                sSqlArt += "art_cos=@cos";
                bFirst = false;
                listParams.Add(new SqlParameter("@cos", dNewCos));
            }

            UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_pxc", r["pxc"], listParams, tArtSchema);
            UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_pne", r["peso"], listParams, tArtSchema);

            // Tara / Gruppo / Tasto (art_tas)
            object valTas = (r.Table.Columns.Contains("tas") && r["tas"] != null && !string.IsNullOrEmpty(r["tas"].ToString().Trim()))
                ? r["tas"]
                : ((r.Table.Columns.Contains("gr") && r["gr"] != null && !string.IsNullOrEmpty(r["gr"].ToString().Trim())) ? r["gr"] : null);
            if (valTas != null)
            {
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_tas", valTas, listParams, tArtSchema);
            }

            if (r["reparto"] != null && !string.IsNullOrEmpty(r["reparto"].ToString().Trim()))
            {
                string sUpdRep = ResolveRepCode(r["reparto"].ToString().Trim());
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_rep", sUpdRep, listParams, tArtSchema);
            }

            // Categorie merceologiche
            if (r["art_ec1"] != null && !string.IsNullOrEmpty(r["art_ec1"].ToString().Trim()))
            {
                string sEc1Upd = Format3Digits(r["art_ec1"].ToString().Trim(), "");
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_ec1", sEc1Upd, listParams, tArtSchema);
            }
            else if (r["famiglia"] != null && !string.IsNullOrEmpty(r["famiglia"].ToString().Trim()))
            {
                string sFamUpd = Format3Digits(r["famiglia"].ToString().Trim(), "");
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_ec1", sFamUpd, listParams, tArtSchema);
            }

            if (r["art_ec2"] != null && !string.IsNullOrEmpty(r["art_ec2"].ToString().Trim()))
            {
                string sEc2Upd = Format3Digits(r["art_ec2"].ToString().Trim(), "");
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_ec2", sEc2Upd, listParams, tArtSchema);
            }
            if (r["art_ec3"] != null && !string.IsNullOrEmpty(r["art_ec3"].ToString().Trim()))
            {
                string sEc3Upd = Format3Digits(r["art_ec3"].ToString().Trim(), "");
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_ec3", sEc3Upd, listParams, tArtSchema);
            }

            // Categoria bilancia (art_cat)
            object valCat = (r.Table.Columns.Contains("cat_bil") && r["cat_bil"] != null && !string.IsNullOrEmpty(r["cat_bil"].ToString().Trim()))
                ? r["cat_bil"]
                : null;
            if (valCat != null && !string.IsNullOrEmpty(valCat.ToString().Trim()))
            {
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_cat", valCat, listParams, tArtSchema);
            }

            if (r["iva"] != null && !string.IsNullOrEmpty(r["iva"].ToString().Trim()))
            {
                string sUpdIva = ResolveIvaCode(r["iva"].ToString().Trim());
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_iva", sUpdIva, listParams, tArtSchema);
            }
            if (r["bil"] != null && !string.IsNullOrEmpty(r["bil"].ToString().Trim()))
            {
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_bil", bIsBilUpd, listParams, tArtSchema);
            }
            if (r["eti"] != null && !string.IsNullOrEmpty(r["eti"].ToString().Trim()))
            {
                string sEtiUpd = Format3Digits(r["eti"].ToString().Trim(), "043");
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_eti", sEtiUpd, listParams, tArtSchema);
            }

            // Parametri Bilancia updates
            if (r.Table.Columns.Contains("plu") && r["plu"] != null && !string.IsNullOrEmpty(r["plu"].ToString().Trim()))
            {
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_plu", r["plu"], listParams, tArtSchema);
            }
            if (r.Table.Columns.Contains("reb") && r["reb"] != null && !string.IsNullOrEmpty(r["reb"].ToString().Trim()))
            {
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_reb", r["reb"], listParams, tArtSchema);
            }
            if (r.Table.Columns.Contains("ori") && r["ori"] != null && !string.IsNullOrEmpty(r["ori"].ToString().Trim()))
            {
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_ori", r["ori"], listParams, tArtSchema);
            }
            if (r.Table.Columns.Contains("cal") && r["cal"] != null && !string.IsNullOrEmpty(r["cal"].ToString().Trim()))
            {
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_cal", r["cal"], listParams, tArtSchema);
            }
            if (r.Table.Columns.Contains("tra") && r["tra"] != null && !string.IsNullOrEmpty(r["tra"].ToString().Trim()))
            {
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_tra", r["tra"], listParams, tArtSchema);
            }
            if (r.Table.Columns.Contains("gsc") && r["gsc"] != null && !string.IsNullOrEmpty(r["gsc"].ToString().Trim()))
            {
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_gsc", _clsFun.Txt2Dec(r["gsc"].ToString()), listParams, tArtSchema);
            }
            if (r.Table.Columns.Contains("bpz") && r["bpz"] != DBNull.Value)
            {
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_bpz", Convert.ToBoolean(r["bpz"]), listParams, tArtSchema);
            }
            if (r.Table.Columns.Contains("tar") && r["tar"] != null && !string.IsNullOrEmpty(r["tar"].ToString().Trim()))
            {
                UpdateSqlFieldTx(ref sSqlArt, ref bFirst, "art_tar", _clsFun.Txt2Dec(r["tar"].ToString()), listParams, tArtSchema);
            }

            if (!bFirst)
            {
                sSqlArt += " WHERE art_cod=@art";
                listParams.Add(new SqlParameter("@art", sArt));
                using (SqlCommand cm = new SqlCommand(sSqlArt, cn, tr))
                {
                    foreach (var pNum in listParams) cm.Parameters.Add(pNum);
                    cm.ExecuteNonQuery();
                }
            }

            // 2. Update GesLisAcquisto
            decimal dPxc = 1m;
            if (r["pxc"] != null && !string.IsNullOrEmpty(r["pxc"].ToString().Trim()))
            {
                dPxc = _clsFun.Txt2Dec(r["pxc"].ToString());
                if (dPxc <= 0) dPxc = 1m;
            }
            EnsureLiaRecord(sArt, sRowFor, sArf, dNewCos, dPxc, tLiaSchema, cn, tr);

            // 3. Update GesLisVendita
            if (dNewPrv > 0)
            {
                EnsureLivRecord(sArt, dNewPrv, tLivSchema, cn, tr);
            }

            // 4. Variations
            if (chkSendToCasse.Checked) AddVariation(sArt, "POS", "UPDATE ART XLS", cn, tr);
            if (chkSendToStampa.Checked) AddVariation(sArt, "ETI", "UPDATE ART XLS", cn, tr);
            if (bIsBilUpd) AddVariation(sArt, "BIL", "UPDATE ART XLS", cn, tr);
        }
    }
}
