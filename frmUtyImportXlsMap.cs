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

namespace APOffice
{
    public partial class frmUtyImportXlsMap : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public string _strFilePath = "";
        public string _strConSqlMdb = "";
        public string _strSelectedMap = "";

        private DataTable _tPreview = null;
        private string _strMapId = "";

        public class ColumnComboItem
        {
            public string Text { get; set; }
            public int Value { get; set; }
            public override string ToString() { return Text; }
        }

        public frmUtyImportXlsMap(string filePath, string conSql, string mapId = "")
        {
            InitializeComponent();
            _strFilePath = filePath;
            _strMapId = mapId;
            _strConSqlMdb = conSql;
        }

        private void frmUtyImportXlsMap_Load(object sender, EventArgs e)
        {
            InitDatabase();
            LoadSuppliers();
            LoadExcelPreview();
            PopulateColumnSelectors();

            if (!string.IsNullOrEmpty(_strMapId))
                LoadExistingMapping();
        }

        private void InitDatabase()
        {
            try
            {
                string sqlTes = "CREATE TABLE XlsMapTestata (MapId VARCHAR(50) PRIMARY KEY, MapDes VARCHAR(255), MapFor VARCHAR(50), MapSkip INT)";
                _clsFun.SqlWrite(sqlTes, _strConSqlMdb);

                string sql = "CREATE TABLE XlsMapDettaglio (MapId VARCHAR(50), FldName VARCHAR(50), ColPos INT, FixVal VARCHAR(255))";
                _clsFun.SqlWrite(sql, _strConSqlMdb);
            }
            catch
            {
                // Schema Migration for MapSkip and FixVal
                try { _clsFun.SqlWrite("ALTER TABLE XlsMapTestata ADD MapSkip INT", _strConSqlMdb); } catch { }
                try { _clsFun.SqlWrite("ALTER TABLE XlsMapDettaglio ADD FixVal VARCHAR(255)", _strConSqlMdb); } catch { }
            }
        }

        private void LoadSuppliers()
        {
            try
            {
                string sStrConSql = _clsFun.ConSql("");
                string s = "SELECT for_cod, for_des FROM AnaFornitori WHERE for_ann=0 ORDER BY for_des";
                DataTable t = _clsFun.FillTabSql("for", s, false, sStrConSql);

                DataRow r = t.NewRow();
                r["for_cod"] = "";
                r["for_des"] = "-- Nessun Fornitore --";
                t.Rows.InsertAt(r, 0);

                cmbFor.DataSource = t;
                cmbFor.DisplayMember = "for_des";
                cmbFor.ValueMember = "for_cod";
            }
            catch (Exception ex) { MessageBox.Show("Errore caricamento fornitori: " + ex.Message); }
        }

        private void LoadExcelPreview()
        {
            try
            {
                if (string.IsNullOrEmpty(_strFilePath) || !File.Exists(_strFilePath))
                {
                    MessageBox.Show("Nessun file selezionato o file non trovato:\n" + _strFilePath,
                        "FILE NON TROVATO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                _tPreview = ReadExcelOrCsv(_strFilePath);
                if (_tPreview != null && _tPreview.Rows.Count > 10)
                {
                    DataTable tTop = _tPreview.Clone();
                    for (int i = 0; i < Math.Min(10, _tPreview.Rows.Count); i++)
                        tTop.ImportRow(_tPreview.Rows[i]);

                    dgvPreview.DataSource = tTop;
                }
                else
                {
                    dgvPreview.DataSource = _tPreview;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore durante il caricamento dell'anteprima Excel/CSV:\n\n" + ex.Message,
                    "ERRORE CARICAMENTO ANTEPRIMA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private DataTable ReadExcelOrCsv(string strPath)
        {
            if (string.IsNullOrEmpty(strPath) || !File.Exists(strPath))
                throw new FileNotFoundException("File non trovato o percorso non valido:\n" + strPath);

            string ext = Path.GetExtension(strPath).ToLower();

            if (ext == ".csv" || ext == ".txt")
            {
                return ReadCsvToDataTable(strPath);
            }

            // Per file .xls e .xlsx: crea una copia temporanea per evitare l'errore di file bloccato da Excel
            string tempFile = Path.Combine(Path.GetTempPath(), "ap_map_" + Guid.NewGuid().ToString("N") + ext);
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
                            using (OleDbCommand cmd = new OleDbCommand("SELECT TOP 10 * FROM [" + sheetName + "]", conn))
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

        private void PopulateColumnSelectors()
        {
            if (_tPreview == null) return;

            var columns = new List<ColumnComboItem>();
            columns.Add(new ColumnComboItem { Text = "-- Non mappato --", Value = 0 });
            for (int i = 0; i < _tPreview.Columns.Count; i++)
            {
                string val = GetPreviewValue(i);
                columns.Add(new ColumnComboItem { Text = "Col " + (i + 1) + " (" + val + ")", Value = i + 1 });
            }

            SetupCombo(cmbEan, columns);
            SetupCombo(cmbDes, columns);
            SetupCombo(cmbCos, columns);
            SetupCombo(cmbPrv, columns);
            SetupCombo(cmbIva, columns);
            SetupCombo(cmbCodFor, columns);
            SetupCombo(cmbPxc, columns);
            SetupCombo(cmbGr, columns);
            SetupCombo(cmbPs, columns);
            SetupCombo(cmbRep, columns);
            SetupCombo(cmbEc1, columns);
            SetupCombo(cmbEc2, columns);
            SetupCombo(cmbEc3, columns);
            SetupCombo(cmbArtFor, columns);
            SetupCombo(cmbEti, columns);

            // Bilancia Column Combos
            SetupCombo(cmbReb, columns);
            SetupCombo(cmbOri, columns);
            SetupCombo(cmbCal, columns);
            SetupCombo(cmbCat, columns);
            SetupCombo(cmbPlu, columns);
            SetupCombo(cmbTas, columns);
            SetupCombo(cmbBil, columns);
            SetupCombo(cmbTra, columns);
            SetupCombo(cmbGsc, columns);
            SetupCombo(cmbBpz, columns);
            SetupCombo(cmbTar, columns);

            LoadFixedValueDropdowns();
        }

        private void LoadFixedValueDropdowns()
        {
            try
            {
                string sConSql = !string.IsNullOrEmpty(_strConSqlMdb) ? _strConSqlMdb : _clsFun.ConSql("");

                // 1. Reparto Cassa (TabReparti come in Anagrafica Articolo)
                BindFixedCombo(cmbFixRep, "TabReparti", new string[] {
                    "SELECT tab_cod, tab_des FROM TabReparti WHERE tab_ann=0 ORDER BY tab_cod",
                    "SELECT tab_cod, tab_des FROM TabReparti ORDER BY tab_cod",
                    "SELECT tab_cod, tab_des FROM TabTabelle WHERE tab_tip='REP' AND tab_ann=0 ORDER BY tab_cod",
                    "SELECT tab_cod, tab_des FROM TabTabelle WHERE tab_tip='REP' ORDER BY tab_cod"
                }, "-- Valore Fisso Reparto --", sConSql);

                // 2. Tipo Etichetta (TabEtichette come in Anagrafica Articolo)
                BindFixedCombo(cmbFixEti, "TabEtichette", new string[] {
                    "SELECT tab_cod, tab_des FROM TabEtichette WHERE tab_ann=0 ORDER BY tab_cod",
                    "SELECT tab_cod, tab_des FROM TabEtichette ORDER BY tab_cod",
                    "SELECT fmt_cod, fmt_des FROM TabFormati WHERE fmt_ann=0 ORDER BY fmt_cod",
                    "SELECT fmt_cod, fmt_des FROM TabFormati ORDER BY fmt_cod"
                }, "-- Valore Fisso Etichetta --", sConSql);

                // 3. Merceologia 1 (TabEcrLv1 come in Anagrafica Articolo)
                BindFixedCombo(cmbFixEc1, "TabEcrLv1", new string[] {
                    "SELECT tab_cod, tab_des FROM TabEcrLv1 ORDER BY tab_cod",
                    "SELECT tab_cod, tab_des FROM TabTabelle WHERE tab_tip='EC1' AND tab_ann=0 ORDER BY tab_cod",
                    "SELECT tab_cod, tab_des FROM TabTabelle WHERE tab_tip='EC1' ORDER BY tab_cod"
                }, "-- Valore Fisso Merc. 1 --", sConSql);

                cmbFixEc1.SelectedIndexChanged -= cmbFixEc1_SelectedIndexChanged;
                cmbFixEc1.SelectedIndexChanged += cmbFixEc1_SelectedIndexChanged;
                cmbFixEc2.SelectedIndexChanged -= cmbFixEc2_SelectedIndexChanged;
                cmbFixEc2.SelectedIndexChanged += cmbFixEc2_SelectedIndexChanged;

                ReloadEc2Combo();

                // 6. IVA (TabIva come in Anagrafica Articolo)
                BindFixedCombo(cmbFixIva, "TabIva", new string[] {
                    "SELECT tab_cod, tab_des + ' (' + CAST(tab_ali AS VARCHAR) + '%)' as tab_des FROM TabIva ORDER BY tab_cod",
                    "SELECT tab_cod, tab_des FROM TabIva ORDER BY tab_cod"
                }, "-- Valore Fisso IVA --", sConSql);

                // --- TABELLE LOOKUP PARAMETRI BILANCIA ---

                // 7. Reparto Bilancia (TabRepBilance)
                BindFixedCombo(cmbFixReb, "TabRepBilance", new string[] {
                    "SELECT tab_cod, tab_des FROM TabRepBilance ORDER BY tab_cod",
                    "SELECT * FROM TabRepBilance",
                    "SELECT tab_cod, tab_des FROM TabTabelle WHERE tab_tip='REB' ORDER BY tab_cod"
                }, "-- Valore Fisso Rep. Bilancia --", sConSql);

                // 8. Origine (TabOrigine)
                BindFixedCombo(cmbFixOri, "TabOrigine", new string[] {
                    "SELECT tab_cod, tab_des FROM TabOrigine ORDER BY tab_cod",
                    "SELECT * FROM TabOrigine",
                    "SELECT tab_cod, tab_des FROM TabTabelle WHERE tab_tip='ORI' ORDER BY tab_cod"
                }, "-- Valore Fisso Origine --", sConSql);

                // 9. Calibro (TabCalibro)
                BindFixedCombo(cmbFixCal, "TabCalibro", new string[] {
                    "SELECT tab_cod, tab_des FROM TabCalibro ORDER BY tab_cod",
                    "SELECT * FROM TabCalibro",
                    "SELECT tab_cod, tab_des FROM TabTabelle WHERE tab_tip='CAL' ORDER BY tab_cod"
                }, "-- Valore Fisso Calibro --", sConSql);

                // 10. Categoria (TabCategoria)
                BindFixedCombo(cmbFixCat, "TabCategoria", new string[] {
                    "SELECT tab_cod, tab_des FROM TabCategoria ORDER BY tab_cod",
                    "SELECT * FROM TabCategoria",
                    "SELECT tab_cod, tab_des FROM TabTabelle WHERE tab_tip='CAT' ORDER BY tab_cod"
                }, "-- Valore Fisso Categoria --", sConSql);

                // 11. Valori Fissi Standard Bilancia (Attivo, Lotto, A Pezzo)
                DataTable tFixBil = new DataTable();
                tFixBil.Columns.Add("tab_cod", typeof(string));
                tFixBil.Columns.Add("tab_des", typeof(string));
                tFixBil.Rows.Add("", "-- Valore Fisso Attivo --");
                tFixBil.Rows.Add("1", "Sì (Attivo in Bilancia)");
                tFixBil.Rows.Add("0", "No (Disattivo)");
                cmbFixBil.DataSource = tFixBil; cmbFixBil.DisplayMember = "tab_des"; cmbFixBil.ValueMember = "tab_cod";

                DataTable tFixTra = new DataTable();
                tFixTra.Columns.Add("tab_cod", typeof(string));
                tFixTra.Columns.Add("tab_des", typeof(string));
                tFixTra.Rows.Add("", "-- Valore Fisso Lotto --");
                tFixTra.Rows.Add("S", "Sì (Tracciabilità/Lotto)");
                tFixTra.Rows.Add("N", "No (Nessun Lotto)");
                cmbFixTra.DataSource = tFixTra; cmbFixTra.DisplayMember = "tab_des"; cmbFixTra.ValueMember = "tab_cod";

                DataTable tFixBpz = new DataTable();
                tFixBpz.Columns.Add("tab_cod", typeof(string));
                tFixBpz.Columns.Add("tab_des", typeof(string));
                tFixBpz.Rows.Add("", "-- Valore Fisso Tipo --");
                tFixBpz.Rows.Add("1", "A Pezzo (PZ)");
                tFixBpz.Rows.Add("0", "A Peso (KG)");
                cmbFixBpz.DataSource = tFixBpz; cmbFixBpz.DisplayMember = "tab_des"; cmbFixBpz.ValueMember = "tab_cod";

                // 12. Prefisso Barcode Bilancia
                cmbPrfEanBil.Items.Clear();
                cmbPrfEanBil.Items.AddRange(new object[] { "2", "20", "22", "24", "28", "29" });
                cmbPrfEanBil.SelectedIndex = 0; // "2"
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "LoadFixedValueDropdowns");
            }
        }

        private void BindFixedCombo(ComboBox cmb, string tabName, string[] queries, string defaultText, string sConSql)
        {
            if (cmb == null) return;
            try
            {
                DataTable tRaw = null;
                if (queries != null)
                {
                    foreach (string q in queries)
                    {
                        if (string.IsNullOrEmpty(q)) continue;
                        try
                        {
                            tRaw = _clsFun.FillTabSql(tabName, q, false, sConSql);
                            if (tRaw != null && tRaw.Rows.Count > 0)
                                break;
                        }
                        catch { }
                    }
                }

                DataTable tClean = new DataTable();
                tClean.Columns.Add("tab_cod", typeof(string));
                tClean.Columns.Add("tab_des", typeof(string));

                // Always add the default prompt row at index 0
                tClean.Rows.Add("", defaultText);

                if (tRaw != null)
                {
                    string colCod = "tab_cod";
                    if (!tRaw.Columns.Contains(colCod))
                    {
                        if (tRaw.Columns.Contains("fmt_cod")) colCod = "fmt_cod";
                        else if (tRaw.Columns.Count > 0) colCod = tRaw.Columns[0].ColumnName;
                    }

                    string colDes = "tab_des";
                    if (!tRaw.Columns.Contains(colDes))
                    {
                        if (tRaw.Columns.Contains("fmt_des")) colDes = "fmt_des";
                        else if (tRaw.Columns.Count > 1) colDes = tRaw.Columns[1].ColumnName;
                    }

                    foreach (DataRow r in tRaw.Rows)
                    {
                        string cod = (colCod != null && tRaw.Columns.Contains(colCod) && r[colCod] != DBNull.Value) ? r[colCod].ToString().Trim() : "";
                        string des = (colDes != null && tRaw.Columns.Contains(colDes) && r[colDes] != DBNull.Value) ? r[colDes].ToString().Trim() : "";
                        if (string.IsNullOrEmpty(cod) && string.IsNullOrEmpty(des)) continue;

                        string display = !string.IsNullOrEmpty(cod) && !string.IsNullOrEmpty(des) && !des.StartsWith(cod) ? (cod + " - " + des) : (!string.IsNullOrEmpty(des) ? des : cod);
                        tClean.Rows.Add(cod, display);
                    }
                }

                cmb.DataSource = tClean;
                cmb.DisplayMember = "tab_des";
                cmb.ValueMember = "tab_cod";
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "BindFixedCombo_" + tabName);
            }
        }

        private void SetupCombo(ComboBox cmb, object source)
        {
            if (cmb == null) return;
            cmb.DataSource = new BindingSource(source, null);
            cmb.DisplayMember = "Text";
            cmb.ValueMember = "Value";
        }

        private string GetPreviewValue(int index)
        {
            if (_tPreview.Rows.Count > 0 && !DBNull.Value.Equals(_tPreview.Rows[0][index]))
            {
                string s = _tPreview.Rows[0][index].ToString().Trim();
                return s.Length > 20 ? s.Substring(0, 20) + "..." : s;
            }
            return "Vuota";
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtMapName.Text))
            {
                MessageBox.Show("Inserire un nome per la mappatura.");
                return;
            }

            try
            {
                string sId = string.IsNullOrEmpty(_strMapId) ? DateTime.Now.ToString("yyyyMMddHHmmss") : _strMapId;
                string sName = txtMapName.Text.Trim();
                string sFor = cmbFor.SelectedValue != null ? cmbFor.SelectedValue.ToString() : "";

                // Se in modifica, eliminiamo i vecchi dati prima di reinserire
                if (!string.IsNullOrEmpty(_strMapId))
                {
                    _clsFun.SqlWrite("DELETE FROM XlsMapTestata WHERE MapId='" + _strMapId + "'", _strConSqlMdb);
                    _clsFun.SqlWrite("DELETE FROM XlsMapDettaglio WHERE MapId='" + _strMapId + "'", _strConSqlMdb);
                }

                // 1. Salvataggio Testata
                string sSqlTes = "INSERT INTO XlsMapTestata (MapId, MapDes, MapFor, MapSkip) VALUES ('" + sId + "', '" + sName.Replace("'", "''") + "', '" + sFor + "', " + numSkip.Value + ")";
                _clsFun.SqlWrite(sSqlTes, _strConSqlMdb);

                // 2. Salvataggio Dettagli Base
                SaveDetail(sId, "art_ean", cmbEan);
                SaveDetail(sId, "art_des", cmbDes);
                SaveDetail(sId, "art_cos", cmbCos);
                SaveDetail(sId, "art_pre", cmbPrv);
                SaveDetail(sId, "art_iva", cmbIva, cmbFixIva);
                SaveDetail(sId, "art_cod_for", cmbCodFor);
                SaveDetail(sId, "art_pxc", cmbPxc);
                SaveDetail(sId, "art_gr", cmbGr, txtFixGr);
                SaveDetail(sId, "art_ps", cmbPs, txtFixPs);
                SaveDetail(sId, "art_rep", cmbRep, cmbFixRep);
                SaveDetail(sId, "art_ec1", cmbEc1, cmbFixEc1);
                SaveDetail(sId, "art_ec2", cmbEc2, cmbFixEc2);
                SaveDetail(sId, "art_ec3", cmbEc3, cmbFixEc3);
                SaveDetail(sId, "art_for", cmbArtFor);
                SaveDetail(sId, "art_eti", cmbEti, cmbFixEti);

                // 3. Salvataggio Dettagli Bilancia
                SaveDetail(sId, "art_reb", cmbReb, cmbFixReb);
                SaveDetail(sId, "art_ori", cmbOri, cmbFixOri);
                SaveDetail(sId, "art_cal", cmbCal, cmbFixCal);
                SaveDetail(sId, "art_cat_bil", cmbCat, cmbFixCat);
                SaveDetail(sId, "art_plu", cmbPlu, txtFixPlu);
                SaveDetail(sId, "art_tas", cmbTas, txtFixTas);
                SaveDetail(sId, "art_bil", cmbBil, cmbFixBil);
                SaveDetail(sId, "art_tra", cmbTra, cmbFixTra);
                SaveDetail(sId, "art_gsc", cmbGsc, txtFixGsc);
                SaveDetail(sId, "art_bpz", cmbBpz, cmbFixBpz);
                SaveDetail(sId, "art_tar", cmbTar, txtFixTar);

                // 4. Salvataggio Opzioni Barcode Bilancia
                string sqlGenEan = "INSERT INTO XlsMapDettaglio (MapId, FldName, ColPos, FixVal) VALUES ('" + sId + "', 'gen_ean_bil', 0, '" + (chkGenEanBil.Checked ? "1" : "0") + "')";
                _clsFun.SqlWrite(sqlGenEan, _strConSqlMdb);
                string sPrf = cmbPrfEanBil.SelectedItem != null ? cmbPrfEanBil.SelectedItem.ToString() : (cmbPrfEanBil.Text != "" ? cmbPrfEanBil.Text : "2");
                string sqlPrfEan = "INSERT INTO XlsMapDettaglio (MapId, FldName, ColPos, FixVal) VALUES ('" + sId + "', 'prf_ean_bil', 0, '" + sPrf.Replace("'", "''") + "')";
                _clsFun.SqlWrite(sqlPrfEan, _strConSqlMdb);

                _strSelectedMap = sId;
                MessageBox.Show("Mappatura salvata con successo!");
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore salvataggio: " + ex.Message);
            }
        }

        private void SaveDetail(string mapId, string fldName, ComboBox cmbCol, Control ctrlFix = null)
        {
            int val = 0;
            if (cmbCol != null)
            {
                if (cmbCol.SelectedValue != null && int.TryParse(cmbCol.SelectedValue.ToString(), out int parsedVal))
                    val = parsedVal;
                else if (cmbCol.SelectedItem is ColumnComboItem cci)
                    val = cci.Value;
            }

            string fixVal = "";
            if (ctrlFix is ComboBox cmbFix)
            {
                if (cmbFix.SelectedValue != null)
                    fixVal = cmbFix.SelectedValue.ToString().Trim();
                else if (cmbFix.SelectedItem is DataRowView drv && drv.Row.Table.Columns.Contains("tab_cod"))
                    fixVal = drv["tab_cod"].ToString().Trim();
            }
            else if (ctrlFix is TextBox txtFix)
            {
                fixVal = txtFix.Text.Trim();
            }

            if (val > 0 || !string.IsNullOrEmpty(fixVal))
            {
                string sql = "INSERT INTO XlsMapDettaglio (MapId, FldName, ColPos, FixVal) VALUES ('" + mapId + "', '" + fldName + "', " + val + ", '" + fixVal.Replace("'", "''") + "')";
                _clsFun.SqlWrite(sql, _strConSqlMdb);
            }
        }

        private void cmbFixEc1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReloadEc2Combo();
        }

        private void cmbFixEc2_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReloadEc3Combo();
        }

        private void ReloadEc2Combo(string targetVal = "")
        {
            try
            {
                string sConSql = !string.IsNullOrEmpty(_strConSqlMdb) ? _strConSqlMdb : _clsFun.ConSql("");
                string sLv1 = cmbFixEc1.SelectedValue != null ? cmbFixEc1.SelectedValue.ToString().Trim() : "";

                string q = "SELECT tab_cod, tab_des FROM TabEcrLv2 WHERE 1=1 ";
                if (!string.IsNullOrEmpty(sLv1))
                {
                    q += "AND (tab_lv1='" + sLv1.Replace("'", "''") + "' OR tab_lv1='" + sLv1.PadLeft(3, '0') + "') ";
                }
                q += "ORDER BY tab_cod";

                BindFixedCombo(cmbFixEc2, "TabEcrLv2", new string[] { q, "SELECT tab_cod, tab_des FROM TabEcrLv2 ORDER BY tab_cod" }, "-- Valore Fisso Merc. 2 --", sConSql);
                if (!string.IsNullOrEmpty(targetVal))
                {
                    SetComboSelectedValue(cmbFixEc2, targetVal);
                }
                ReloadEc3Combo();
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "ReloadEc2Combo");
            }
        }

        private void ReloadEc3Combo(string targetVal = "")
        {
            try
            {
                string sConSql = !string.IsNullOrEmpty(_strConSqlMdb) ? _strConSqlMdb : _clsFun.ConSql("");
                string sLv1 = cmbFixEc1.SelectedValue != null ? cmbFixEc1.SelectedValue.ToString().Trim() : "";
                string sLv2 = cmbFixEc2.SelectedValue != null ? cmbFixEc2.SelectedValue.ToString().Trim() : "";

                string q = "SELECT tab_cod, tab_des FROM TabEcrLv3 WHERE 1=1 ";
                if (!string.IsNullOrEmpty(sLv1))
                {
                    q += "AND (tab_lv1='" + sLv1.Replace("'", "''") + "' OR tab_lv1='" + sLv1.PadLeft(3, '0') + "') ";
                }
                if (!string.IsNullOrEmpty(sLv2))
                {
                    q += "AND (tab_lv2='" + sLv2.Replace("'", "''") + "' OR tab_lv2='" + sLv2.PadLeft(3, '0') + "') ";
                }
                q += "ORDER BY tab_cod";

                BindFixedCombo(cmbFixEc3, "TabEcrLv3", new string[] { q, "SELECT tab_cod, tab_des FROM TabEcrLv3 ORDER BY tab_cod" }, "-- Valore Fisso Merc. 3 --", sConSql);
                if (!string.IsNullOrEmpty(targetVal))
                {
                    SetComboSelectedValue(cmbFixEc3, targetVal);
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "ReloadEc3Combo");
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SetComboSelectedValue(ComboBox cmb, object value)
        {
            if (cmb == null || value == null) return;
            try
            {
                string targetStr = value.ToString().Trim();
                if (string.IsNullOrEmpty(targetStr)) return;

                // 1. Se associato a DataTable (Valori Fissi: Reparto, IVA, Merceologie, ecc.)
                if (cmb.DataSource is DataTable dt)
                {
                    cmb.SelectedValue = targetStr;
                    if (cmb.SelectedValue != null && string.Equals(cmb.SelectedValue.ToString().Trim(), targetStr, StringComparison.OrdinalIgnoreCase))
                        return;

                    string pad3 = targetStr.PadLeft(3, '0');
                    cmb.SelectedValue = pad3;
                    if (cmb.SelectedValue != null && string.Equals(cmb.SelectedValue.ToString().Trim(), pad3, StringComparison.OrdinalIgnoreCase))
                        return;

                    string colCod = dt.Columns.Contains("tab_cod") ? "tab_cod" : (dt.Columns.Contains("fmt_cod") ? "fmt_cod" : (dt.Columns.Count > 0 ? dt.Columns[0].ColumnName : ""));
                    if (!string.IsNullOrEmpty(colCod))
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            string cod = dt.Rows[i][colCod].ToString().Trim();
                            if (string.Equals(cod, targetStr, StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(cod, pad3, StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(cod.TrimStart('0'), targetStr.TrimStart('0'), StringComparison.OrdinalIgnoreCase))
                            {
                                cmb.SelectedIndex = i;
                                return;
                            }
                        }
                    }
                    return;
                }

                // 2. Se è un combo di mappatura colonna Excel (ColumnComboItem)
                if (int.TryParse(targetStr, out int targetVal))
                {
                    cmb.SelectedValue = targetVal;
                    if (cmb.SelectedValue != null && cmb.SelectedValue.Equals(targetVal))
                        return;

                    for (int i = 0; i < cmb.Items.Count; i++)
                    {
                        if (cmb.Items[i] is ColumnComboItem cci && cci.Value == targetVal)
                        {
                            cmb.SelectedIndex = i;
                            return;
                        }
                    }
                }
                else
                {
                    cmb.SelectedValue = targetStr;
                }
            }
            catch { }
        }

        private void LoadExistingMapping()
        {
            try
            {
                DataTable tTes = _clsFun.FillTabSql("Tes", "SELECT * FROM XlsMapTestata WHERE MapId='" + _strMapId + "'", false, _strConSqlMdb);
                if (tTes.Rows.Count > 0)
                {
                    txtMapName.Text = tTes.Rows[0]["MapDes"].ToString();
                    if (tTes.Rows[0]["MapFor"] != DBNull.Value)
                        SetComboSelectedValue(cmbFor, tTes.Rows[0]["MapFor"].ToString());
                    if (tTes.Columns.Contains("MapSkip") && tTes.Rows[0]["MapSkip"] != DBNull.Value)
                        numSkip.Value = Convert.ToInt32(tTes.Rows[0]["MapSkip"]);
                }

                string savedFixEc1 = "";
                string savedFixEc2 = "";
                string savedFixEc3 = "";

                DataTable tDet = _clsFun.FillTabSql("Det", "SELECT * FROM XlsMapDettaglio WHERE MapId='" + _strMapId + "'", false, _strConSqlMdb);
                foreach (DataRow r in tDet.Rows)
                {
                    string fld = r["FldName"].ToString();
                    int pos = Convert.ToInt32(r["ColPos"]);
                    string fixVal = tDet.Columns.Contains("FixVal") && !DBNull.Value.Equals(r["FixVal"]) ? r["FixVal"].ToString() : "";

                    if (fld == "art_ean") SetComboSelectedValue(cmbEan, pos);
                    else if (fld == "art_des") SetComboSelectedValue(cmbDes, pos);
                    else if (fld == "art_cos") SetComboSelectedValue(cmbCos, pos);
                    else if (fld == "art_pre") SetComboSelectedValue(cmbPrv, pos);
                    else if (fld == "art_iva") { SetComboSelectedValue(cmbIva, pos); if (!string.IsNullOrEmpty(fixVal)) SetComboSelectedValue(cmbFixIva, fixVal); }
                    else if (fld == "art_cod" || fld == "art_cod_for") SetComboSelectedValue(cmbCodFor, pos);
                    else if (fld == "art_pxc") SetComboSelectedValue(cmbPxc, pos);
                    else if (fld == "art_gr") { SetComboSelectedValue(cmbGr, pos); if (!string.IsNullOrEmpty(fixVal)) txtFixGr.Text = fixVal; }
                    else if (fld == "art_ps") { SetComboSelectedValue(cmbPs, pos); if (!string.IsNullOrEmpty(fixVal)) txtFixPs.Text = fixVal; }
                    else if (fld == "art_rep") { SetComboSelectedValue(cmbRep, pos); if (!string.IsNullOrEmpty(fixVal)) SetComboSelectedValue(cmbFixRep, fixVal); }
                    else if (fld == "art_fam" || fld == "art_ec1") { SetComboSelectedValue(cmbEc1, pos); if (!string.IsNullOrEmpty(fixVal)) savedFixEc1 = fixVal; }
                    else if (fld == "art_ec2") { SetComboSelectedValue(cmbEc2, pos); if (!string.IsNullOrEmpty(fixVal)) savedFixEc2 = fixVal; }
                    else if (fld == "art_ec3") { SetComboSelectedValue(cmbEc3, pos); if (!string.IsNullOrEmpty(fixVal)) savedFixEc3 = fixVal; }
                    else if (fld == "art_for") SetComboSelectedValue(cmbArtFor, pos);
                    else if (fld == "art_eti") { SetComboSelectedValue(cmbEti, pos); if (!string.IsNullOrEmpty(fixVal)) SetComboSelectedValue(cmbFixEti, fixVal); }
                    
                    // Bilancia mappings
                    else if (fld == "art_reb") { SetComboSelectedValue(cmbReb, pos); if (!string.IsNullOrEmpty(fixVal)) SetComboSelectedValue(cmbFixReb, fixVal); }
                    else if (fld == "art_ori") { SetComboSelectedValue(cmbOri, pos); if (!string.IsNullOrEmpty(fixVal)) SetComboSelectedValue(cmbFixOri, fixVal); }
                    else if (fld == "art_cal") { SetComboSelectedValue(cmbCal, pos); if (!string.IsNullOrEmpty(fixVal)) SetComboSelectedValue(cmbFixCal, fixVal); }
                    else if (fld == "art_cat_bil" || (fld == "art_cat" && cmbCat.SelectedValue == null)) { SetComboSelectedValue(cmbCat, pos); if (!string.IsNullOrEmpty(fixVal)) SetComboSelectedValue(cmbFixCat, fixVal); }
                    else if (fld == "art_plu") { SetComboSelectedValue(cmbPlu, pos); if (!string.IsNullOrEmpty(fixVal)) txtFixPlu.Text = fixVal; }
                    else if (fld == "art_tas") { SetComboSelectedValue(cmbTas, pos); if (!string.IsNullOrEmpty(fixVal)) txtFixTas.Text = fixVal; }
                    else if (fld == "art_bil") { SetComboSelectedValue(cmbBil, pos); if (!string.IsNullOrEmpty(fixVal)) SetComboSelectedValue(cmbFixBil, fixVal); }
                    else if (fld == "art_tra") { SetComboSelectedValue(cmbTra, pos); if (!string.IsNullOrEmpty(fixVal)) SetComboSelectedValue(cmbFixTra, fixVal); }
                    else if (fld == "art_gsc") { SetComboSelectedValue(cmbGsc, pos); if (!string.IsNullOrEmpty(fixVal)) txtFixGsc.Text = fixVal; }
                    else if (fld == "art_bpz") { SetComboSelectedValue(cmbBpz, pos); if (!string.IsNullOrEmpty(fixVal)) SetComboSelectedValue(cmbFixBpz, fixVal); }
                    else if (fld == "art_tar") { SetComboSelectedValue(cmbTar, pos); if (!string.IsNullOrEmpty(fixVal)) txtFixTar.Text = fixVal; }
                    else if (fld == "gen_ean_bil") { chkGenEanBil.Checked = (pos == 1 || fixVal == "1" || fixVal.Equals("true", StringComparison.OrdinalIgnoreCase)); }
                    else if (fld == "prf_ean_bil") { if (!string.IsNullOrEmpty(fixVal)) { if (!cmbPrfEanBil.Items.Contains(fixVal)) cmbPrfEanBil.Items.Add(fixVal); cmbPrfEanBil.SelectedItem = fixVal; } }
                }

                if (!string.IsNullOrEmpty(savedFixEc1))
                {
                    SetComboSelectedValue(cmbFixEc1, savedFixEc1);
                    ReloadEc2Combo(savedFixEc2);
                    if (!string.IsNullOrEmpty(savedFixEc3))
                    {
                        ReloadEc3Combo(savedFixEc3);
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Errore caricamento mappatura: " + ex.Message); }
        }
    }
}
