using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    public partial class frmSeekDocs : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        clsResize _form_resize;

        private const string DOCTIPPES = "K";

        private const string TABTABDOC = "TabDocumenti";
        private const string TABTABDTP = "TabDocTipo";
        private const string TABTABCAU = "TabMovCausali";

        private const string TABGESFAT = "GesFatTestate";
        private const string TABGESMOT = "GesMovTestate";

        public string _strImpDoc = "";

        private string _strConSql = "";

        private DataSet _dasGen = new DataSet();

        private string _strFrmPar = "";
        private string _strIni11VenditaTouch = "";

        private string _strPar026FattFromCassa = "";

        public frmSeekDocs()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmGesSeekDocs_Load(object sender, EventArgs e)
        {
            ApplyModernUi();

            if(_clsFun.ParGet(clsDefine.enuParametri.Par033FattureElettroniche,_strConSql) != "S")
                fattureElettronicheToolStripMenuItem.Visible = false;

            _strConSql = _clsFun.ConSql("");

            for (int i = DateTime.Today.Year - 4; i < DateTime.Today.Year + 1; i++)
                cmbYea.Items.Add(i.ToString());
            cmbYea.SelectedIndex = 4;

            _strIni11VenditaTouch = _clsFun.FileIni("R", clsDefine.enuIni.Ini11VenditaTouch, "");

            importFattureDaScontriniToolStripMenuItem.Visible = false;
            _strPar026FattFromCassa = _clsFun.ParGet(clsDefine.enuParametri.Par026FattFromCassa, _strConSql);
            if(_strPar026FattFromCassa != "")
                importFattureDaScontriniToolStripMenuItem.Visible = true;

            SetDgv1();

            FillTab();

            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmSeekDocs_dgv1");
        }

        private void frmSeekDocs_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmSeekDocs_dgv1");
        }

        private void ApplyModernUi()
        {
            try
            {
                if (menuStrip1 != null)
                {
                    menuStrip1.Renderer = clsUiIcons.GetModernMenuRenderer();
                    if (esciToolStripMenuItem != null)
                        esciToolStripMenuItem.Image = clsUiIcons.GetIcon("exit", 16);
                    if (lottiToolStripMenuItem != null)
                        lottiToolStripMenuItem.Image = clsUiIcons.GetIcon("boxes", 16);
                    if (elenchiToolStripMenuItem != null)
                        elenchiToolStripMenuItem.Image = clsUiIcons.GetIcon("list", 16);
                    if (excelDocumentiToolStripMenuItem != null)
                        excelDocumentiToolStripMenuItem.Image = clsUiIcons.GetIcon("excel", 16);
                    if (excelElencoToolStripMenuItem != null)
                        excelElencoToolStripMenuItem.Image = clsUiIcons.GetIcon("excel", 16);
                    if (excelElencoCsvToolStripMenuItem != null)
                        excelElencoCsvToolStripMenuItem.Image = clsUiIcons.GetIcon("excel", 16);
                    if (excelElencoDettaglioToolStripMenuItem != null)
                        excelElencoDettaglioToolStripMenuItem.Image = clsUiIcons.GetIcon("excel", 16);
                    if (importFattureDaScontriniToolStripMenuItem != null)
                        importFattureDaScontriniToolStripMenuItem.Image = clsUiIcons.GetIcon("receipt", 16);
                    if (fattureElettronicheToolStripMenuItem != null)
                        fattureElettronicheToolStripMenuItem.Image = clsUiIcons.GetIcon("invoice", 16);
                }

                this.BackColor = Color.FromArgb(243, 244, 246);

                if (btnSeek != null)
                {
                    clsUiIcons.StyleStatButton(btnSeek, "Estrai", "search", 20,
                        Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254),
                        Color.FromArgb(96, 165, 250), Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));
                }

                if (btnNew != null)
                {
                    clsUiIcons.StyleStatButton(btnNew, "Nuovo", "plus", 18,
                        Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208),
                        Color.FromArgb(52, 211, 153), Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));
                }

                if (btnDocRaggruppa != null)
                {
                    clsUiIcons.StyleStatButton(btnDocRaggruppa, "Raggruppamento documenti per fattura differita", "sync", 16,
                        Color.FromArgb(245, 243, 255), Color.FromArgb(221, 214, 254),
                        Color.FromArgb(167, 139, 250), Color.FromArgb(76, 29, 149), Color.FromArgb(124, 58, 237));
                }

                if (dgv1 != null)
                {
                    clsUiIcons.StyleDataGridView(dgv1);
                    dgv1.CellFormatting += dgv1_CellFormatting;
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmSeekDocs.ApplyModernUi", ex.Message);
            }
        }

        private void dgv1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            try
            {
                if (dgv1.Columns.Contains("Annullato") && dgv1.Rows[e.RowIndex].Cells["Annullato"].Value is bool isAnn && isAnn)
                {
                    e.CellStyle.ForeColor = Color.FromArgb(148, 163, 184);
                }

                string colName = dgv1.Columns[e.ColumnIndex].Name;
                if (colName == "Importo" && e.Value != null)
                {
                    if (decimal.TryParse(e.Value.ToString(), out decimal valDec))
                    {
                        e.Value = valDec.ToString("N2");
                        e.FormattingApplied = true;
                    }
                }
                else if (colName == "Data" && e.Value != null)
                {
                    if (DateTime.TryParse(e.Value.ToString(), out DateTime valDt))
                    {
                        e.Value = valDt.ToString("dd/MM/yyyy");
                        e.FormattingApplied = true;
                    }
                }
            }
            catch { }
        }

        private void frmSeekDocs_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void Esci()
        {
            if (cmbMovFat.SelectedValue != null && cmbCauTpd.SelectedValue != null)
            {
                string sTip = cmbMovFat.SelectedValue.ToString();
                string sTpd = cmbCauTpd.SelectedValue.ToString();

                string s = sTip + "-" + sTpd;
                _clsQry.ParForm(this, "W", s);
            }

            this.Close();
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

            //dgv1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //dgv1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            //s += "fat_yfa AS DocYea, ";
            //s += "fat_nfa AS DocNum, ";
            //s += "fat_ndo AS DocNdo, ";
            //s += "fat_ddo AS DocDdo, ";
            //s += "fat_cfo AS DocCfo, ";
            //s += "AnaClienti.cli_des AS DocCfd, ";
            //s += "fat_ann AS DocAnn ";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocYea";
            cTbc.Name = "Anno";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocTdo";
            cTbc.Name = "T";
            cTbc.Width = 23;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocFat";
            cTbc.Name = "Fatturato";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.Visible = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocNeg";
            cTbc.Name = "Negozio";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocCfd";
            cTbc.Name = "Cliente/Fornitore";
            cTbc.Width = 220;
            cTbc.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocNdo";
            cTbc.Name = "Documento";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocDdo";
            cTbc.Name = "Data";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "MovImp";
            cTbc.Name = "Importo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(decimal);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format="#,##0.00";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocTpg";
            cTbc.Name = "Pagamento";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocStd";
            cTbc.Name = "Stato";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocNo1";
            cTbc.Name = "Note";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.DataPropertyName = "DocAnn";
            cCbc.Name = "Annullato";
            cCbc.Width = 40;
            cCbc.ValueType = typeof(string);
            cCbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "FatEle";
            cTbc.Name = "Fattura Elettronica";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocNum";
            cTbc.Name = "Movimento";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocUbi";
            cTbc.Name = "Negozio/Sede";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);
        }

        private void FillTab()
        {
            DataRow x;
            string s = "";
            DataTable t = new DataTable();

            s = "SELECT * FROM TabDocumenti";
            t = _clsFun.FillTabSql(TABTABDOC, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbMovFat.DataSource = t;
            cmbMovFat.DisplayMember = "tab_des";
            cmbMovFat.ValueMember = "tab_cod";
            cmbMovFat.SelectedValue = TipDoc("TIP", _strFrmPar);

            s = "SELECT * FROM TabDocTipo";
            t = _clsFun.FillTabSql(TABTABDTP, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            if (_dasGen.Tables.IndexOf(t.TableName) >= 0)
                _dasGen.Tables.Remove(t.TableName);
            _dasGen.Tables.Add(t);

            s = "SELECT * FROM TabMovCausali WHERE tab_ann=0";
            t = _clsFun.FillTabSql(TABTABCAU, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            if (_dasGen.Tables.IndexOf(t.TableName) >= 0)
                _dasGen.Tables.Remove(t.TableName);
            _dasGen.Tables.Add(t);

            if (_strFrmPar.Length > 3)
            {
                FillDocTpd();
                FillDati(cmbMovFat.SelectedValue.ToString(), cmbCauTpd.SelectedValue.ToString());
                dgv1.Select();
            }
        }

        private string TipDoc(string strTip, string strPar)
        {
            string sRes = "";
            if(strPar != "")
            {
                string[] a = strPar.Split('-');
                if (strTip == "TIP" && a.Length > 0)
                    sRes = a[0];
                if (strTip == "TPD" && a.Length > 1)
                    sRes = a[1];
            }
            return sRes;
        }

        private void cmbYea_SelectedIndexChanged(object sender, EventArgs e)
        {
            Boolean b = false;
            if (cmbMovFat.SelectedValue != null && cmbCauTpd.SelectedValue != null)
            {
                FillDati(cmbMovFat.SelectedValue.ToString(), cmbCauTpd.SelectedValue.ToString());
                dgv1.Select();
            }
        }

        private void cmbDocDoc_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillDocTpd();
        }

        private void FillDocTpd()
        {
            DataTable t = new DataTable();
            if (cmbMovFat.SelectedIndex == 1)
            {
                lblDocTip.Text = "Tipo documento";
                t = _dasGen.Tables[TABTABDTP];
            }
            else
            { 
                lblDocTip.Text = "Causale";
                t = _dasGen.Tables[TABTABCAU];
            }

            cmbCauTpd.DataSource = t;
            cmbCauTpd.DisplayMember = "tab_des";
            cmbCauTpd.ValueMember = "tab_cod";
            cmbCauTpd.SelectedValue = TipDoc("TPD", _strFrmPar);

            //FillDati(cmbMovFat.SelectedValue.ToString(), cmbCauTpd.SelectedValue.ToString());
            if (cmbMovFat.SelectedValue != null && cmbMovFat.SelectedValue.ToString() != "" && cmbCauTpd.SelectedValue != null && cmbCauTpd.SelectedValue.ToString() != "")
                FillDati(cmbMovFat.SelectedValue.ToString(), cmbCauTpd.SelectedValue.ToString());
            else
                dgv1.DataSource = null;
        }

        private void cmbCauTpd_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillDati(cmbMovFat.SelectedValue.ToString(), cmbCauTpd.SelectedValue.ToString());
            dgv1.Select();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            Boolean b = false;
            if (cmbMovFat.SelectedValue.ToString() == "")
                MessageBox.Show("Tipo documento non definito!");
            else if (cmbCauTpd.SelectedValue == null || cmbCauTpd.SelectedValue.ToString() == "")
                MessageBox.Show(lblDocTip.Text + " non definito!");

            else
                b = true;
            if(b)
            {
                string sSuf = "";
                string sTik = "";
                string sTip = cmbMovFat.SelectedValue.ToString() + cmbCauTpd.SelectedValue.ToString();

                DataRow[] j = ((DataTable)cmbMovFat.DataSource).Select("tab_cod='" + cmbMovFat.SelectedValue.ToString()+"'");
                if(j.Length > 0)
                {
                    string s = (string)j[0]["tab_suf"];
                    string[] aa = s.Split(',');
                    if (aa.Length > 0)
                        sSuf = aa[0];
                }
                j = ((DataTable)cmbCauTpd.DataSource).Select("tab_cod='" + cmbCauTpd.SelectedValue.ToString() + "'");
                if (j.Length > 0)
                {
                    if(((string)j[0]["tab_tip"]).Contains(DOCTIPPES))
                        sTik = DOCTIPPES;
                }

                string[] a = _strIni11VenditaTouch.Split(';');

                if(a.Length > 0 && a[0] == "S" && sTik == DOCTIPPES)
                {
                    string sPorta = a[1];
                    string sBilTime = a[3];

                    frmGesDocTouch f = new frmGesDocTouch();
                    f._strMovFat = cmbMovFat.SelectedValue.ToString();
                    f._strDocSuf = sSuf;
                    f._strCauTpd = cmbCauTpd.SelectedValue.ToString();
                    f._strMftYea = cmbYea.Text;
                    f._strMftNum = _clsDef.CODNEW;
                    f._strOpenPorta = sPorta;
                    f._strBilTime = sBilTime;
                    f.ShowDialog();
                }
                else
                {
                    frmGesDocumento f = new frmGesDocumento();
                    f._strMovFat = cmbMovFat.SelectedValue.ToString();
                    f._strDocSuf = sSuf;
                    f._strCauTpd = cmbCauTpd.SelectedValue.ToString();
                    f._strMftYea = cmbYea.Text;
                    f._strMftNum = _clsDef.CODNEW;
                    f.ShowDialog();
                }

                FillDati(cmbMovFat.SelectedValue.ToString(), cmbCauTpd.SelectedValue.ToString());
            }
        }

        private void btnSeek_Click(object sender, EventArgs e)
        {
            Boolean b = false;
            if (cmbMovFat.SelectedValue.ToString() == "")
                MessageBox.Show("Tipo documento non definito!");
            else if (cmbCauTpd.SelectedValue.ToString() == "")
                MessageBox.Show(lblDocTip.Text + " non definita!");
            else
                b = true;
            if (b)
            {
                FillDati(cmbMovFat.SelectedValue.ToString(), cmbCauTpd.SelectedValue.ToString());
            }
        }

        private void FillDati(string strMovFat, string strCauTdo)
        {
            string s = "";
            string sCfo = "";
            //string sTip = "";
            DataRow[] j;
            decimal d = 0;
            
            dgv1.Columns["Fatturato"].Visible = false;

            if (strMovFat == "F")
            {

                j = _dasGen.Tables[TABTABDTP].Select("tab_cod='" + strCauTdo + "'");
                if (!DBNull.Value.Equals(j[0]["tab_cfo"]))
                {
                    sCfo = (string)j[0]["tab_cfo"];
                    //sTip = (string)j[0]["tab_tip"];
                }

                if(sCfo == _clsDef.TIPCLI)
                {
                    s = "SELECT ";
                    s += "fat_ubi AS DocUbi, ";
                    s += "fat_yfa AS DocYea, ";
                    s += "fat_nfa AS DocNum, ";
                    s += "fat_tdo AS DocTdo, ";
                    s += "fat_ndo AS DocNdo, ";
                    s += "fat_ddo AS DocDdo, ";
                    s += "fat_neg AS DocNeg, ";
                    s += "fat_cfo AS DocCfo, ";
                    s += "fat_ele AS FatEle, ";
                    s += "AnaClienti.cli_des AS DocCfd, ";
                    s += "TabStato.tab_des AS DocStd, ";
                    s += "TabPagamenti.tab_des AS DocTpg, ";
                    s += "R.mov_imp + R.MovIva AS MovImp, ";
                    s += "fat_no1 AS DocNo1, ";
                    s += "fat_ann AS DocAnn ";
                    s += "FROM (((GesFatTestate ";
                    s += "LEFT OUTER JOIN AnaClienti ON GesFatTestate.fat_cfo = AnaClienti.cli_cod) ";
                    s += "LEFT OUTER JOIN TabStato ON GesFatTestate.fat_sta = TabStato.tab_cod) ";
                    s += "LEFT OUTER JOIN TabPagamenti ON GesFatTestate.fat_tpg = TabPagamenti.tab_cod) ";
                    s += "LEFT JOIN (";

                    s += "SELECT ";
                    s += "mov_nfa, ";
                    s += "SUM(mov_imp) AS mov_imp, ";
                    s += "SUM(CASE WHEN tab_ali IS NOT NULL AND tab_ali > 0 AND mov_imp > 0 THEN mov_imp * tab_ali / 100 ELSE 0 END) AS MovIva ";
                    s += "FROM GesMovimenti ";
                    s += "LEFT OUTER JOIN TabIva ON GesMovimenti.mov_iva = TabIva.tab_cod ";
                    s += "WHERE (mov_ann = 0 AND mov_yfa='" + cmbYea.Text + "') ";
                    s += "GROUP BY mov_yfa, mov_nfa ";

                    s += ") R ON GesFatTestate.fat_nfa = R.mov_nfa ";

                    s += "WHERE ";
                    s += "fat_yfa ='" + cmbYea.Text + "' AND ";
                    s += "fat_tpd='" + strCauTdo + "' ";
                    //s += "fat_tip='" + sTip + "' ";
                    s += "ORDER BY fat_ddo, fat_ndo";
                }
                else 
                {
                    s = "SELECT ";
                    s += "fat_ubi AS DocUbi, ";
                    s += "fat_yfa AS DocYea, ";
                    s += "fat_nfa AS DocNum, ";
                    s += "fat_tdo AS DocTdo, ";
                    s += "fat_ndo AS DocNdo, ";
                    s += "fat_ddo AS DocDdo, ";
                    s += "fat_neg AS DocNeg, ";
                    s += "fat_cfo AS DocCfo, ";
                    s += "AnaFornitori.for_des AS DocCfd, ";
                    s += "TabStato.tab_des AS DocStd, ";
                    s += "TabPagamenti.tab_des AS DocTpg, ";
                    s += "R.mov_imp + R.MovIva AS MovImp, ";
                    s += "fat_no1 AS DocNo1, ";
                    s += "fat_ann AS DocAnn ";
                    s += "FROM (((GesFatTestate ";
                    s += "LEFT OUTER JOIN AnaFornitori ON GesFatTestate.fat_cfo = AnaFornitori.for_cod) ";
                    s += "LEFT OUTER JOIN TabStato ON GesFatTestate.fat_sta = TabStato.tab_cod) ";
                    s += "LEFT OUTER JOIN TabPagamenti ON GesFatTestate.fat_tpg = TabPagamenti.tab_cod) ";
                    s += "LEFT JOIN (";

                    s += "SELECT ";
                    s += "mov_nfa, ";
                    s += "SUM(mov_imp) AS mov_imp, ";
                    s += "SUM(CASE WHEN tab_ali IS NOT NULL AND tab_ali > 0 AND mov_imp > 0 THEN mov_imp * tab_ali / 100 ELSE 0 END) AS MovIva ";
                    s += "FROM GesMovimenti ";
                    s += "LEFT OUTER JOIN TabIva ON GesMovimenti.mov_iva = TabIva.tab_cod ";
                    s += "WHERE (mov_ann = 0 AND mov_yfa='" + cmbYea.Text + "') ";
                    s += "GROUP BY mov_yfa, mov_nfa ";

                    s += ") R ON GesFatTestate.fat_nfa = R.mov_nfa ";

                    s += "WHERE ";
                    s += "fat_yfa ='" + cmbYea.Text + "' AND ";
                    s += "fat_tpd='" + strCauTdo + "' ";
                    //s += "fat_tip='" + sTip + "' ";
                    s += "ORDER BY fat_ddo, fat_ndo";
                }
            }
            else
            {
                if (strCauTdo != "")
                {

                    j = _dasGen.Tables[TABTABCAU].Select("tab_cod='" + strCauTdo + "'");
                    sCfo = (string)j[0]["tab_cfo"];

                    //if (((string)j[0]["tab_tip"]).Contains(DOCTIPPES))
                    //    sTip = DOCTIPPES;

                    if (sCfo == _clsDef.TIPCLI)
                    {
                        dgv1.Columns["Fatturato"].Visible = true;

                        s = "SELECT ";
                        s += "mot_ubi AS DocUbi, ";
                        s += "mot_ymo AS DocYea, ";
                        s += "mot_nmo AS DocNum, ";
                        s += "CASE WHEN mot_nfa IS NULL OR  mot_nfa = '' THEN '' ELSE 'S' END AS DocFat, ";
                        s += "mot_ndo AS DocNdo, ";
                        s += "mot_ddo AS DocDdo, ";
                        s += "mot_cfo AS DocCfo, ";
                        s += "mot_neg AS DocNeg, ";
                        s += "AnaClienti.cli_des AS DocCfd, ";
                        s += "R.mov_imp + R.MovIva AS MovImp, ";
                        s += "mot_no1 AS DocNo1, ";
                        s += "TabStato.tab_des AS DocStd, ";
                        s += "TabPagamenti.tab_des AS DocTpg, ";
                        s += "mot_ann AS DocAnn ";
                        s += "FROM GesMovTestate ";
                        s += "LEFT OUTER JOIN AnaClienti ON GesMovTestate.mot_cfo = AnaClienti.cli_cod ";
                        s += "LEFT OUTER JOIN TabStato ON GesMovTestate.mot_sta = TabStato.tab_cod ";
                        s += "LEFT OUTER JOIN TabPagamenti ON GesMovTestate.mot_tpg = TabPagamenti.tab_cod ";

                        s += "LEFT JOIN (";

                        s += "SELECT ";
                        s += "mov_nmo, ";
                        s += "SUM(mov_imp) AS mov_imp, ";
                        s += "SUM(CASE WHEN tab_ali IS NOT NULL AND tab_ali > 0 AND mov_imp > 0 THEN mov_imp * tab_ali / 100 ELSE 0 END) AS MovIva ";
                        s += "FROM GesMovimenti ";
                        s += "LEFT OUTER JOIN TabIva ON GesMovimenti.mov_iva = TabIva.tab_cod ";
                        s += "WHERE (mov_ann = 0 AND mov_ymo='" + cmbYea.Text + "') ";
                        s += "GROUP BY mov_ymo, mov_nmo ";

                        s += ") R ON GesMovTestate.mot_nmo = R.mov_nmo ";


                        s += "WHERE ";
                        s += "mot_ymo >='" + cmbYea.Text + "' AND ";
                        s += "mot_cau='" + strCauTdo + "' ";
                        //s += "mot_tip='" + sTip + "' ";
                        s += "ORDER BY mot_yfa, mot_nfa";
                    }
                    else
                    {
                        s = "SELECT ";
                        s += "mot_ubi AS DocUbi, ";
                        s += "mot_ymo AS DocYea, ";
                        s += "mot_nmo AS DocNum, ";
                        s += "mot_ndo AS DocNdo, ";
                        s += "mot_ddo AS DocDdo, ";
                        s += "mot_cfo AS DocCfo, ";
                        s += "mot_neg AS DocNeg, ";
                        if(sCfo != "")
                            s += "AnaFornitori.for_des AS DocCfd, ";
                        s += "R.mov_imp + R.MovIva AS MovImp, ";
                        s += "mot_ann AS DocAnn ";
                        s += "FROM GesMovTestate ";
                        if (sCfo != "")
                            s += "LEFT OUTER JOIN AnaFornitori ON GesMovTestate.mot_cfo = AnaFornitori.for_cod ";

                        s += "LEFT JOIN (";

                        s += "SELECT ";
                        s += "mov_nmo, ";
                        s += "SUM(mov_imp) AS mov_imp, ";
                        s += "SUM(CASE WHEN tab_ali IS NOT NULL AND tab_ali > 0 AND mov_imp > 0 THEN mov_imp * tab_ali / 100 ELSE 0 END) AS MovIva ";
                        s += "FROM GesMovimenti ";
                        s += "LEFT OUTER JOIN TabIva ON GesMovimenti.mov_iva = TabIva.tab_cod ";
                        s += "WHERE (mov_ann = 0 AND mov_ymo='" + cmbYea.Text + "') ";
                        s += "GROUP BY mov_ymo, mov_nmo ";

                        s += ") R ON GesMovTestate.mot_nmo = R.mov_nmo ";

                        s += "WHERE ";
                        s += "mot_ymo >='" + cmbYea.Text + "' AND ";
                        s += "mot_cau='" + strCauTdo + "' ";
                        //s += "mot_tip='" + sTip + "' ";
                        s += "ORDER BY mot_yfa, mot_nfa";
                    }
                }
            }

            DataTable t = _clsFun.FillTabSql("DOC", s, false, _strConSql);

            dgv1.DataSource = t;

            d = 0;
            if (t.Columns.Contains("MovImp"))
            {
                foreach (DataRow y in t.Rows)
                {
                    if (!(Boolean)y["DocAnn"] && (string)y["DocNdo"] != "")
                    {
                        if (strMovFat == "F" && (string)y["DocTdo"] == "NA")
                            d -= (decimal)y["MovImp"];
                        else
                            d += (decimal)y["MovImp"];
                    }
                }
            }

            if (t.Rows.Count > 0)
            {
                int i = dgv1.Rows.Count - 1;
                if (i >= 0)
                    dgv1.CurrentCell = dgv1[3, i];
            }

            lblTotVal.Text = d.ToString("###,##0.00");
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                if (_strImpDoc == "S")
                {
                    _strImpDoc = cmbMovFat.SelectedValue.ToString() + ";";
                    _strImpDoc += cmbCauTpd.SelectedValue.ToString() + ";";
                    _strImpDoc += (string)x["DocYea"] + ";";
                    _strImpDoc += (string)x["DocNum"] + ";";
                    _strImpDoc += (string)x["DocUbi"] + ";";
                    Esci();
                }
                else
                {
                    string sSuf = "";
                    string sTik = "";
                    string sTip = cmbMovFat.SelectedValue.ToString() + cmbCauTpd.SelectedValue.ToString();
                    
                    DataRow[] j = ((DataTable)cmbMovFat.DataSource).Select("tab_cod='" + cmbMovFat.SelectedValue.ToString() + "'");
                    if (j.Length > 0)
                    {
                        //sSuf = ((string)j[0]["tab_suf"]).Trim();

                        string s = (string)j[0]["tab_suf"];
                        string[] aa = s.Split(',');
                        if (aa.Length > 0)
                            sSuf = aa[0];
                    }

                    j = ((DataTable)cmbCauTpd.DataSource).Select("tab_cod='" + cmbCauTpd.SelectedValue.ToString() + "'");
                    if (j.Length > 0)
                    {
                        if(((string)j[0]["tab_tip"]).Contains(DOCTIPPES))
                            sTik = DOCTIPPES;
                    }
                    string[] a = _strIni11VenditaTouch.Split(';');

                    if (a.Length > 0 && a[0] == "S" && sTik == DOCTIPPES)
                    {
                        string sPorta = a[1];
                        string sBilTime = a[3];

                        frmGesDocTouch f = new frmGesDocTouch();
                        f._strMovFat = cmbMovFat.SelectedValue.ToString();
                        f._strCauTpd = cmbCauTpd.SelectedValue.ToString();
                        f._strMftYea = (string)x["DocYea"];
                        f._strMftNum = (string)x["DocNum"];
                        f._strMftUbi = (string)x["DocUbi"];
                        f._strDocSuf = sSuf;
                        f._strOpenPorta = sPorta;
                        f._strBilTime = sBilTime;

                        f.ShowDialog();
                        if (f._strReturn != "")
                        {
                            if (f._strCauTpd != cmbCauTpd.SelectedValue.ToString())
                                x.Delete();
                            else if (f._strReturn == "CANC")
                                x.Delete();
                            else
                            {
                                string[] aa = f._strReturn.Split('|');
                                //if (cmbMovFat.SelectedValue.ToString() == "F" && cmbCauTpd.SelectedValue.ToString() == "FV")
                                if (aa.Length >= 3)
                                {
                                    //x["DocStd"] = aa[0];

                                    if (cmbMovFat.SelectedValue.ToString() == "F" && cmbCauTpd.SelectedValue.ToString() == "FV")
                                        x["DocStd"] = aa[0];
                                    if(x.Table.Columns.Contains("DocCfd"))
                                        x["DocCfd"] = aa[1];
                                    x["DocNdo"] = aa[2];
                                }
                            }
                        }
                    }
                    else
                    {
                        frmGesDocumento f = new frmGesDocumento();
                        f._strMovFat = cmbMovFat.SelectedValue.ToString();
                        f._strCauTpd = cmbCauTpd.SelectedValue.ToString();
                        f._strMftYea = (string)x["DocYea"];
                        f._strMftNum = (string)x["DocNum"];
                        if (!DBNull.Value.Equals(x["DocUbi"]))
                            f._strMftUbi = Convert.ToString(x["DocUbi"]);
                        f._strDocSuf = sSuf;
                        f.ShowDialog();
                        if (f._strReturn != "")
                        {
                            if (f._strReturn == "CANC")
                                FillDati(cmbMovFat.SelectedValue.ToString(), cmbCauTpd.SelectedValue.ToString());
                            else
                            {
                                string[] aa = f._strReturn.Split('|');
                                if (cmbMovFat.SelectedValue.ToString() == "F") // && cmbCauTpd.SelectedValue.ToString() == "FV"
                                {
                                    x["DocStd"] = aa[0];
                                    x["DocCfd"] = aa[1];
                                    x["DocNdo"] = aa[2];
                                    x["DocNo1"] = aa[4];
                                }
                            }
                        }
                    }
                }
            }
        }

        private void btnDocRaggruppa_Click(object sender, EventArgs e)
        {
            new frmGesDocRaggruppa().ShowDialog();
            FillDati(cmbMovFat.SelectedValue.ToString(), cmbCauTpd.SelectedValue.ToString());
        }

        private void txtSeek_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                string s = txtSeek.Text.Trim();

                if (s != "" && dgv1.RowCount > 0)
                {
                    Boolean b = false;
                    int iCur = dgv1.CurrentCell.RowIndex;
                    int i = 0;

                    foreach (DataGridViewRow row in dgv1.Rows)
                    {
                        i++;

                        if (i > iCur+1)
                        {
                            if (row.Cells["Cliente/Fornitore"].Value.ToString().ToLower().Contains(s.ToLower()))
                            {
                                dgv1.Rows[row.Index].Selected = true;
                                dgv1.FirstDisplayedScrollingRowIndex = row.Index;
                                dgv1.CurrentCell = dgv1.Rows[row.Index].Cells[0];
                                b = true;
                                break;
                            }
                        }
                    }
                    if(!b)
                        dgv1.CurrentCell = dgv1.Rows[0].Cells[0];
                }
            }
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //if ((DataTable)dgv1.DataSource == null)
            //    MessageBox.Show("Dati non estratti!", "ELENCO DOCUMENTI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //else
            //{
            //    frmUtyPeriodo f = new frmUtyPeriodo();
            //    //f._dayIni = new DateTime(Convert.ToInt16(cmbYea.Text), 1, 1);
            //    //f._dayFin = new DateTime(Convert.ToInt16(cmbYea.Text), DateTime.Today.Month, DateTime.Today.Day);
            //    f._dayIni = DateTime.Today;
            //    f._dayFin = DateTime.Today;
            //    f.ShowDialog();
            //    if (f._strRes == "S")
            //    {
            //        FillExcel(f._dayIni, f._dayFin);
            //    }
            //}
        }

        private void FillExcel(DateTime dayIni, DateTime dayFin)
        {
            string s = "DocDdo >= " + _clsFun.DayMdb(dayIni) + " AND DocDdo <= " + _clsFun.DayMdb(dayFin);

            DataTable t = (DataTable)dgv1.DataSource;

            DataView v = new DataView(t, s, "DocNdo, DocDdo", DataViewRowState.CurrentRows);

            if (v.Count == 0)
                MessageBox.Show("Dati non trovati!", "ELENCO DOCUMENTI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                DataRow[] j;
                t = v.Table.Clone();

                decimal d = 0;

                //if (!(Boolean)r["DocAnn"] && (string)r["DocNdo"] != "")
                //{
                //    if (((string)r["DocTdo"]).Substring(0, 1) == "F" && (string)r["DocTdo"] == "NA")
                //        d -= (decimal)r["MovImp"];
                //    else
                //        d += (decimal)r["MovImp"];
                //    t.ImportRow(r.Row);
                //}

                foreach (DataRowView r in v)
                {
                    if (!(Boolean)r["DocAnn"] && (string)r["DocNdo"] != "")
                    {
                        if (cmbMovFat.SelectedValue.ToString().Substring(0, 1) == "F" && (string)r["DocTdo"] == "NA")
                            d -= (decimal)r["MovImp"];
                        else
                            d += (decimal)r["MovImp"];
                        t.ImportRow(r.Row);
                    }
                }

                string sTit = "SITUAZIONE " + cmbMovFat.Text + "/" + cmbCauTpd.Text + " DAL " + dayIni.ToString("dd/MM/yyyy") + " AL " + dayFin.ToString("dd/MM/yyyy");
                string sFil = "DocElenco.xls";

                string sFld = "";
                //s += "DocNum, Num, 30, StringLiteral;";
                sFld += "DocTdo, Tipo, 50, StringLiteral;";
                sFld += "DocNdo, Numero, 50, StringLiteral;";
                sFld += "DocDdo, Data, 50, DateTime;";

                if (t.Columns.Contains("DocCfd"))
                    sFld += "DocCfd, Cli/for, 180, StringLiteral;";

                sFld += "MovImp, Importo, 50, Decimal;";

                if (t.Columns.Contains("DocTpg"))
                    sFld += "DocTpg, Pagamento, 120, StringLiteral;";

                if (t.Columns.Contains("DocStd"))
                    sFld += "DocStd, Stato, 80, StringLiteral;";

                if (t.Columns.Contains("DocNo1"))
                    sFld += "DocNo1, Note, 100, StringLiteral;";

                //sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd_HHmm") + "_Documenti.xls";

                //DataSet ds = new DataSet();
                //DataTable tTit = (new clsExcel()).FillTitStr(s, "TabTit");
                //ds.Tables.Add(tTit);
                //ds.Tables.Add(t);
                //(new clsExcel()).exportToExcel(ds, sFil, sTit, "3|Totale  " + d.ToString("#,###,##0.00"), true);
                //ds.Tables.Remove(t);
                //ds.Clear();
                //ds.Dispose();

                string sFoo = ",,,'Totali'," + d.ToString("#0.00").Replace(",", ".") + ",,,";

                (new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);
            }
        }

        private void csvToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //ElencoCsv();
        }

        private void ElencoCsv()
        {
            Boolean b = false;
            DateTime dIni = new DateTime();
            DateTime dFin = new DateTime();

            if ((DataTable)dgv1.DataSource == null)
                MessageBox.Show("Dati non estratti!", "ELENCO DOCUMENTI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                frmUtyPeriodo f = new frmUtyPeriodo();
                f._dayIni = new DateTime(Convert.ToInt16(cmbYea.Text), 1, 1);
                f._dayFin = new DateTime(Convert.ToInt16(cmbYea.Text), DateTime.Today.Month, DateTime.Today.Day);
                f.ShowDialog();
                if (f._strRes == "S")
                {
                    b = true;
                    dIni = f._dayIni;
                    dFin = f._dayFin;
                }
            }

            if (b)
            {
                string s = "DocDdo >= " + _clsFun.DayMdb(dIni) + " AND DocDdo <= " + _clsFun.DayMdb(dFin);

                DataTable t = (DataTable)dgv1.DataSource;

                DataView v = new DataView(t, s, "DocNdo, DocDdo", DataViewRowState.CurrentRows);

                DataRow[] j;
                t = v.Table.Clone();

                s = "";
                decimal d = 0;

                foreach (DataRowView r in v)
                {
                    if (!(Boolean)r["DocAnn"])
                    {
                        d += (decimal)r["MovImp"];
                        t.ImportRow(r.Row);
                    }
                }

                string sTit = "SITUAZIONE " + cmbMovFat.Text + "/" + cmbCauTpd.Text + " DAL " + dIni.ToString("dd/MM/yyyy") + " AL " + dFin.ToString("dd/MM/yyyy");
                string sFil = "";

                string sFld = "";
                //s += "DocNum, Num, 30, StringLiteral;";
                sFld += "DocNdo, Numero, 50, StringLiteral;";
                sFld += "DocDdo, Data, 50, StringLiteral;";
                sFld += "DocCfd, Cli/for, 180, StringLiteral;";
                sFld += "MovImp, Importo, 50, Decimal2;";
                sFld += "DocTpg, Pagamento, 120, StringLiteral;";
                sFld += "DocStd, Stato, 80, StringLiteral;";
                sFld += "DocNo1, Note, 100, StringLiteral;";

                //sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd_HHmm") + "_Documenti.xls";

                DataSet ds = new DataSet();

                DataTable tTit = (new clsExcel()).FillTitStr(s, "TabTit");

                ds.Tables.Add(tTit);
                ds.Tables.Add(t);

                //(new clsExcel()).exportToExcel(ds, sFil, sTit, "3|Totale  " + d.ToString("#,###,##0.00"), true);
                (new clsExcel()).exportToCsv1(sFld, t, sFil, sTit, "");

                ds.Tables.Remove(t);
                ds.Clear();
                ds.Dispose();
            }
        }

        private void lottiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new frmSeekDocLotti().ShowDialog();
        }

        private void importFattureDaScontriniToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmGesDocScoFatture f = new frmGesDocScoFatture();
            f._strPar026FattFromCassa = _strPar026FattFromCassa;
            f.ShowDialog();
            if (cmbMovFat.SelectedValue != null && cmbCauTpd.SelectedValue != null)
                FillDati(cmbMovFat.SelectedValue.ToString(), cmbCauTpd.SelectedValue.ToString());
        }

        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //frmUtyPeriodo f = new frmUtyPeriodo();
            //f._dayIni = DateTime.Today;
            //f._dayFin = DateTime.Today;
            //f.ShowDialog();
            //if (f._strRes == "S")
            //{
            //    AllDocs2Excel(f._dayIni, f._dayFin);
            //}
        }

        private void excelDocumentiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmUtyPeriodo f = new frmUtyPeriodo();
            f._dayIni = DateTime.Today;
            f._dayFin = DateTime.Today;
            f.ShowDialog();
            if (f._strRes == "S")
            {
                AllDocs2Excel(f._dayIni, f._dayFin);
            }
        }

        private void AllDocs2Excel(DateTime dayIni, DateTime dayFin)
        {
            string s = "";
            DataRow x;
            DataRow[] j;

            string sTpd = cmbCauTpd.SelectedValue.ToString();

            DataTable t = new clsGenTabTmp().TabTmpDocMov("DocAll");

            s = "SELECT * FROM AnaClienti";
            DataTable tCli = _clsFun.FillTabSql("TabCli", s, false, _strConSql);

            //s = "SELECT ";
            //s += "GesFatTestate.fat_tpd, ";
            //s += "GesFatTestate.fat_yfa, ";
            //s += "GesFatTestate.fat_nfa, ";
            //s += "GesFatTestate.fat_ndo, ";
            //s += "GesFatTestate.fat_ddo, ";
            //s += "GesFatTestate.fat_neg, ";
            //s += "GesFatTestate.fat_cfo, ";
            //s += "SUM(GesMovimenti.mov_imp) AS mov_imp ";
            //s += "FROM GesFatTestate ";
            //s += "LEFT OUTER JOIN GesMovimenti ON GesFatTestate.fat_yfa = GesMovimenti.mov_yfa AND GesFatTestate.fat_nfa = GesMovimenti.mov_nfa ";
            //s += "WHERE ";
            //s += "(GesFatTestate.fat_ann = 0) AND ";
            //s += "(GesFatTestate.fat_tpd = 'FV') AND ";
            //s += "GesFatTestate.fat_ddo >= " + _clsFun.DaySql(dayIni) + " AND ";
            //s += "GesFatTestate.fat_ddo <= " + _clsFun.DaySql(dayFin) + " ";
            //s += "GROUP BY GesFatTestate.fat_tpd, GesFatTestate.fat_yfa, GesFatTestate.fat_nfa, GesFatTestate.fat_ndo, GesFatTestate.fat_ddo, GesFatTestate.fat_neg, GesFatTestate.fat_cfo ";

            s = "SELECT ";
            s += "fat_yfa, ";
            s += "fat_nfa, ";
            s += "fat_tdo, ";
            s += "fat_ndo, ";
            s += "fat_ddo, ";
            s += "fat_neg, ";
            s += "fat_cfo, ";
            s += "AnaClienti.cli_des, ";
            //s += "TabStato.tab_des AS DocStd, ";
            //s += "TabPagamenti.tab_des AS DocTpg, ";
            s += "R.mov_imp + R.MovIva AS MovImp, ";
            s += "fat_no1 AS DocNo1, ";
            s += "fat_ann AS DocAnn ";
            s += "FROM GesFatTestate ";
            s += "LEFT OUTER JOIN AnaClienti ON GesFatTestate.fat_cfo = AnaClienti.cli_cod ";
            //s += "LEFT OUTER JOIN TabStato ON GesFatTestate.fat_sta = TabStato.tab_cod) ";
            //s += "LEFT OUTER JOIN TabPagamenti ON GesFatTestate.fat_tpg = TabPagamenti.tab_cod) ";
            s += "LEFT JOIN (";

            s += "SELECT ";
            s += "mov_nfa, ";
            s += "SUM(mov_imp) AS mov_imp, ";
            s += "SUM(CASE WHEN tab_ali IS NOT NULL AND tab_ali > 0 AND mov_imp > 0 THEN mov_imp * tab_ali / 100 ELSE 0 END) AS MovIva ";
            s += "FROM GesMovimenti ";
            s += "LEFT OUTER JOIN TabIva ON GesMovimenti.mov_iva = TabIva.tab_cod ";
            s += "WHERE (mov_ann = 0 AND mov_yfa='" + cmbYea.Text + "') ";
            s += "GROUP BY mov_yfa, mov_nfa ";

            s += ") R ON GesFatTestate.fat_nfa = R.mov_nfa ";

            s += "WHERE ";
            s += "fat_ann = 0 AND ";
            //s += "fat_yfa ='" + cmbYea.Text + "' AND ";
            s += "fat_tpd='" + sTpd + "' AND ";
            s += "GesFatTestate.fat_ddo >= " + _clsFun.DaySql(dayIni) + " AND ";
            s += "GesFatTestate.fat_ddo <= " + _clsFun.DaySql(dayFin) + " ";
            s += "ORDER BY fat_ddo, fat_ndo";

            DataTable tTmp = _clsFun.FillTabSql("GesFatTestate", s, false, _strConSql);

            foreach (DataRow y in tTmp.Rows)
            {
                x = t.NewRow();
                x["tmp_cfo"] = y["fat_cfo"];
                x["tmp_ndo"] = y["fat_ndo"];
                x["tmp_ddo"] = y["fat_ddo"];
                x["tmp_neg"] = y["fat_neg"];
                x["tmp_des"] = "Fatture";

                x["tmp_rag"] = "";
                x["tmp_piv"] = y["fat_ndo"];

                j = tCli.Select("cli_cod='" + y["fat_cfo"] + "'");
                if (j.Length > 0)
                {
                    x["tmp_rag"] = (string)j[0]["cli_des"];
                    x["tmp_piv"] = (string)j[0]["cli_piv"];
                    x["tmp_cfi"] = (string)j[0]["cli_cfi"];
                }
                x["tmp_imp"] = y["MovImp"];

                t.Rows.Add(x);

                //d += (decimal)y["mov_imp"];
            }

            //x = t.NewRow();

            //x["tmp_imp"] = d;

            //t.Rows.Add(x);

            //d = 0;

            s = "SELECT * FROM TabMovCausali";
            DataTable tCau = _clsFun.FillTabSql("TabCau", s, false, _strConSql);

            //s = "SELECT ";
            //s += "GesMovTestate.mot_cfo, ";
            //s += "GesMovTestate.mot_ndo, ";
            //s += "GesMovTestate.mot_ddo, ";
            //s += "GesMovTestate.mot_cau, ";
            ////s += "GesMovTestate.mot_mag, ";
            //s += "GesMovTestate.mot_neg, ";
            //s += "SUM(GesMovimenti.mov_imp) AS mov_imp ";
            //s += "FROM GesMovTestate ";
            //s += "LEFT OUTER JOIN GesMovimenti ON GesMovTestate.mot_ymo = GesMovimenti.mov_ymo AND GesMovTestate.mot_nmo = GesMovimenti.mov_nmo ";
            //s += "WHERE ";
            //s += "(GesMovTestate.mot_ddo >= " + _clsFun.DaySql(dayIni) + ") AND ";
            //s += "(GesMovTestate.mot_ddo <= " + _clsFun.DaySql(dayFin) + ") AND ";
            //s += "(GesMovTestate.mot_ann = 0) ";
            //s += "GROUP BY GesMovTestate.mot_cfo, GesMovTestate.mot_ndo, GesMovTestate.mot_ddo, GesMovTestate.mot_cau, GesMovTestate.mot_neg";


            s = "SELECT ";
            //s += "mot_ymo AS DocYea, ";
            s += "mot_nmo, ";
            s += "CASE WHEN mot_nfa IS NULL OR  mot_nfa = '' THEN '' ELSE 'S' END AS DocFat, ";
            s += "mot_ndo, ";
            s += "mot_ddo, ";
            s += "mot_cfo, ";
            s += "mot_cau, ";
            s += "mot_neg, ";
            s += "AnaClienti.cli_des AS DocCfd, ";
            s += "R.mov_imp + R.MovIva AS MovImp, ";
            s += "mot_no1 ";
            //s += "TabStato.tab_des AS DocStd, ";
            //s += "TabPagamenti.tab_des AS DocTpg, ";
            //s += "mot_ann AS DocAnn ";
            s += "FROM GesMovTestate ";
            s += "LEFT OUTER JOIN AnaClienti ON GesMovTestate.mot_cfo = AnaClienti.cli_cod ";
            //s += "LEFT OUTER JOIN TabStato ON GesMovTestate.mot_sta = TabStato.tab_cod ";
            //s += "LEFT OUTER JOIN TabPagamenti ON GesMovTestate.mot_tpg = TabPagamenti.tab_cod ";

            s += "LEFT JOIN (";

            s += "SELECT ";
            s += "mov_nmo, ";
            s += "SUM(mov_imp) AS mov_imp, ";
            s += "SUM(CASE WHEN tab_ali IS NOT NULL AND tab_ali > 0 AND mov_imp > 0 THEN mov_imp * tab_ali / 100 ELSE 0 END) AS MovIva ";
            s += "FROM GesMovimenti ";
            s += "LEFT OUTER JOIN TabIva ON GesMovimenti.mov_iva = TabIva.tab_cod ";
            s += "WHERE (mov_ann = 0 AND mov_ymo='" + cmbYea.Text + "') ";
            s += "GROUP BY mov_ymo, mov_nmo ";

            s += ") R ON GesMovTestate.mot_nmo = R.mov_nmo ";

            s += "WHERE ";
            //s += "mot_ymo >='" + cmbYea.Text + "' AND ";
            //s += "mot_cau='" + strCauTdo + "' ";
            s += "mot_ann = 0 AND ";
            s += "(GesMovTestate.mot_ddo >= " + _clsFun.DaySql(dayIni) + ") AND ";
            s += "(GesMovTestate.mot_ddo <= " + _clsFun.DaySql(dayFin) + ") AND ";
            s += "(GesMovTestate.mot_ann = 0) ";

            //s += "mot_tip='" + sTip + "' ";
            s += "ORDER BY mot_yfa, mot_nfa";

            DataTable tMot = _clsFun.FillTabSql("TabMot", s, false, _strConSql);

            if (true)
            {
                foreach (DataRow y in tMot.Rows)
                {
                    j = tCau.Select("tab_cod='" + y["mot_cau"] + "'");
                    if (j.Length > 0 && (string)j[0]["tab_cfo"] == _clsDef.TIPCLI)
                    {
                        x = t.NewRow();
                        x["tmp_cfo"] = y["mot_cfo"];
                        x["tmp_ndo"] = y["mot_ndo"];
                        x["tmp_ddo"] = y["mot_ddo"];
                        x["tmp_neg"] = y["mot_neg"];

                        x["tmp_rag"] = "";
                        j = tCli.Select("cli_cod='" + y["mot_cfo"] + "'");
                        if (j.Length > 0)
                        {
                            x["tmp_rag"] = (string)j[0]["cli_des"];
                            x["tmp_piv"] = (string)j[0]["cli_piv"];
                            x["tmp_cfi"] = (string)j[0]["cli_cfi"];
                        }

                        j = tCau.Select("tab_cod='" + y["mot_cau"] + "'");
                        if (j.Length > 0)
                        {
                            x["tmp_des"] += (string)j[0]["tab_des"];

                            //if ((string)j[0]["tab_sgm"] == "+")
                            //    x["tmp_tre"] = "MPI";
                            //if ((string)j[0]["tab_sgm"] == "-")
                            //    x["tmp_tre"] = "MME";
                        }
                        if ((string)j[0]["tab_sgm"] == "-")
                            x["tmp_imp"] = (decimal)y["MovImp"];
                        else
                            x["tmp_imp"] = (decimal)y["MovImp"] * -1;

                        //d += (decimal)x["tmp_imp"];

                        t.Rows.Add(x);
                    }
                }
            }
            DataTable tt = t.Clone();

            DataView v = new DataView(t, "", "tmp_ddo,tmp_des", DataViewRowState.CurrentRows);

            DateTime dDay = dayIni;
            decimal dImp = 0;
            decimal dTot = 0;

            foreach (DataRowView r in v)
            {
                if (dImp > 0 && DateTime.Compare((DateTime)r["tmp_ddo"], dDay) != 0)
                {
                    x = tt.NewRow();
                    x["tmp_ddo"] = dDay;
                    x["tmp_imp"] = dImp;
                    x["tmp_rag"] = "--------- " + "TOTALE AL " + dDay.ToString("dd/MM/yyyy"); ;
                    tt.Rows.Add(x);
                    dImp = 0;

                    dDay = (DateTime)r["tmp_ddo"];
                }

                dImp += (decimal)r["tmp_imp"];
                dTot += (decimal)r["tmp_imp"];
                tt.ImportRow(r.Row);
            }

            if (dImp > 0)
            {
                x = tt.NewRow();
                x["tmp_ddo"] = dDay;
                x["tmp_imp"] = dImp;
                x["tmp_rag"] = "******* " + "TOTALE AL " + dDay.ToString("dd/MM/yyyy"); ;
                tt.Rows.Add(x);
                dImp = 0;
            }

            x = tt.NewRow();
            x["tmp_imp"] = dTot;
            tt.Rows.Add(x);

            t = tt.Copy();

            /*
            string sTit = "Documenti dal " + dayIni.ToString("dd/MM/yyyy") + " dal " + dayFin.ToString("dd/MM/yyyy");
            string sFil = "";

            s = "";
            s += "tmp_des, tipo, 40, StringLiteral;";
            s += "tmp_cfo, Codice, 60, StringLiteral;";
            s += "tmp_rag, Nome, 200, StringLiteral;";
            s += "tmp_piv, Partita IVA, 70, StringLiteral;";
            s += "tmp_ndo, D.numero, 70, StringLiteral;";
            s += "tmp_ddo, Data, 70, StringLiteral;";
            s += "tmp_imp, Importo, 60, Decimal2;";

            for (int i = 0; i <= 100; i++)
            {
                sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_Documenti.xls";
                if (File.Exists(sFil))
                {
                    try
                    {
                        File.Delete(sFil);
                        break;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
                else
                    break;
            }

            DataSet ds = new DataSet();

            DataTable tTit = (new clsExcel()).FillTitStr(s, "TabTit");

            ds.Tables.Add(tTit);
            ds.Tables.Add(t);

            //(new clsExcel()).exportToExcel(ds, sFil, sTit, "Totale " + d.ToString("#,###,##0.00"), true);
            (new clsExcel()).exportToExcel(ds, sFil, sTit, "", true);

            ds.Tables.Remove(t);
            ds.Clear();
            ds.Dispose();
            */

            string sTit = "Documenti dal " + dayIni.ToString("dd/MM/yyyy") + " dal " + dayFin.ToString("dd/MM/yyyy");
            string sFoo = ""; // "11," + lblCos.Text.Replace(".", "").Replace(",", ".");
            string sFil = "";

            string sFld = "";
            sFld = "";
            //sFld += "tmp_ndo, D.numero, 70, StringLiteral;";
            //sFld += "tmp_ddo, Data, 70, DateTime;";
            //sFld += "tmp_cod, Codice, 60, StringLiteral;";
            //sFld += "tmp_des, Descrizione, 200, StringLiteral;";
            //sFld += "tmp_qta, Q.ta, 60, Decimal;";
            //sFld += "tmp_prz, Prezzo, 60, Decimal;";
            //sFld += "tmp_imp, Importo, 60, Decimal;";

            sFld += "tmp_des, tipo, 40, StringLiteral;";
            sFld += "tmp_cfo, Codice, 60, StringLiteral;";
            sFld += "tmp_rag, Nome, 200, StringLiteral;";
            sFld += "tmp_piv, Partita IVA, 70, StringLiteral;";
            sFld += "tmp_ndo, D.numero, 70, StringLiteral;";
            sFld += "tmp_ddo, Data, 70, DateTime;";
            sFld += "tmp_imp, Importo, 60, Decimal;";

            sFil = "Documenti.xls";

            (new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);
        }

        private void AllDocs2ExcelDett(DateTime dayIni, DateTime dayFin)
        {
            string s = "";
            DataRow x;
            DataRow[] j;
            int iRow = 0;

            string sTdo = cmbMovFat.SelectedValue.ToString();
            string sTpd = cmbCauTpd.SelectedValue.ToString();
            string sCfo = "";

            DataTable t = new clsGenTabTmp().TabTmpDocMovDett("DocAll");

            s = "SELECT * FROM AnaClienti";
            DataTable tCli = _clsFun.FillTabSql("TabCli", s, false, _strConSql);

            if (sTdo == "F")
            {
                sCfo = _clsDef.TIPFOR;
                if(sTpd == "FV")
                    sCfo = _clsDef.TIPCLI;

                s = "SELECT * FROM AnaArticoli LEFT OUTER JOIN TabReparti ON AnaArticoli.art_rep=TabReparti.tab_cod";
                DataTable tArt = _clsFun.FillTabSql("ART", s, false, _strConSql);

                s = "SELECT ";
                s += "fat_yfa, ";
                s += "fat_nfa, ";
                s += "fat_tdo, ";
                s += "fat_ndo, ";
                s += "fat_ddo, ";
                s += "fat_neg, ";
                s += "fat_cfo, ";
                s += "AnaClienti.cli_des, ";
                //s += "R.mov_imp + R.MovIva AS MovImp, ";
                s += "R.mov_imp AS MovImp, ";
                s += "fat_no1 AS DocNo1, ";
                s += "fat_ann AS DocAnn, ";
                s += "fat_sta ";
                s += "FROM GesFatTestate ";
                s += "LEFT OUTER JOIN AnaClienti ON GesFatTestate.fat_cfo = AnaClienti.cli_cod ";
                s += "LEFT JOIN (";

                s += "SELECT ";
                s += "mov_nfa, ";
                s += "SUM(mov_imp) AS mov_imp, ";
                s += "SUM(CASE WHEN tab_ali IS NOT NULL AND tab_ali > 0 AND mov_imp > 0 THEN mov_imp * tab_ali / 100 ELSE 0 END) AS MovIva ";
                //s += "SUM(mov_imp) AS MovIva ";
                s += "FROM GesMovimenti ";
                s += "LEFT OUTER JOIN TabIva ON GesMovimenti.mov_iva = TabIva.tab_cod ";
                s += "WHERE (mov_ann = 0 AND mov_yfa='" + cmbYea.Text + "') ";
                s += "GROUP BY mov_yfa, mov_nfa ";

                s += ") R ON GesFatTestate.fat_nfa = R.mov_nfa ";

                s += "WHERE ";
                s += "fat_ann = 0 AND ";
                s += "fat_tpd='" + sTpd + "' AND ";
                //s += "fat_ndo<>'' AND ";
                s += "GesFatTestate.fat_ddo >= " + _clsFun.DaySql(dayIni) + " AND ";
                s += "GesFatTestate.fat_ddo <= " + _clsFun.DaySql(dayFin) + " ";
                s += "ORDER BY fat_ddo, fat_ndo";

                DataTable tTmp = _clsFun.FillTabSql("GesFatTestate", s, false, _strConSql);

                foreach (DataRow y in tTmp.Rows)
                {
                    //x = t.NewRow();
                    //x["tmp_cod"] = y["fat_cfo"];
                    //x["tmp_ndo"] = y["fat_ndo"];
                    //x["tmp_ddo"] = y["fat_ddo"];
                    ////x["tmp_neg"] = y["fat_neg"];
                    //x["tmp_des"] = "Fattura";

                    //x["tmp_rag"] = "";
                    //x["tmp_piv"] = y["fat_ndo"];

                    //j = tCli.Select("cli_cod='" + y["fat_cfo"] + "'");
                    //if (j.Length > 0)
                    //{
                    //    x["tmp_des"] += " " + (string)j[0]["cli_des"];
                    //    //x["tmp_piv"] = (string)j[0]["cli_piv"];
                    //    //x["tmp_cfi"] = (string)j[0]["cli_cfi"];
                    //}
                    //x["tmp_imp"] = y["MovImp"];
                    //t.Rows.Add(x);

                    iRow++;
                    x = t.NewRow();
                    x["tmp_tri"] = "1T";
                    //x["tmp_key"] = (string)y["fat_yfa"] + (string)y["fat_nfa"];
                    x["tmp_key"] = (string)y["fat_yfa"] + (string)y["fat_ndo"];

                    //x["tmp_tmp"] = "Fat->" + (string)y["fat_yfa"] + (string)y["fat_ndo"];

                    x["tmp_row"] = iRow.ToString("000000");
                    if (DBNull.Value.Equals(y["fat_ndo"]) || (string)y["fat_ndo"] == "")
                        x["tmp_ndo"] = y["fat_nfa"];
                    else
                        x["tmp_ndo"] = y["fat_ndo"];

                    x["tmp_ddo"] = y["fat_ddo"];
                    //j = tCau.Select("tab_cod='" + y["mot_cau"] + "'");
                    //if (j.Length > 0 && (string)j[0]["tab_cfo"] == _clsDef.TIPCLI)
                    //{
                    x["tmp_cod"] = y["fat_cfo"];
                    x["tmp_des"] = y["fat_cfo"];

                    j = tCli.Select("cli_cod='" + y["fat_cfo"] + "'");
                    if (j.Length > 0)
                    {
                        x["tmp_des"] += " " + (string)j[0]["cli_des"];
                    }
                    //}
                    //j = tCau.Select("tab_cod='" + y["mot_cau"] + "'");
                    //if (j.Length > 0)
                    //{
                    x["tmp_des"] += "DOCUMENTO " + (string)x["tmp_ndo"];    // +" " + (string)j[0]["tab_des"];

                    //}
                    //if ((string)j[0]["tab_sgm"] == "-")
                    x["tmp_imp"] = (decimal)y["MovImp"];
                    //else
                    //    x["tmp_imp"] = (decimal)y["MovImp"] * -1;

                    x["tmp_fod"] = y["DocNo1"];
                    x["tmp_sta"] = y["fat_sta"];

                    t.Rows.Add(x);

                    s = "SELECT GesMovimenti.*, AnaArticoli.art_umi FROM GesMovimenti ";
                    s += "LEFT OUTER JOIN AnaArticoli ON GesMovimenti.mov_art = AnaArticoli.art_cod ";
                    s += "WHERE ";
                    s += "mov_yfa='" + y["fat_yfa"] + "' AND ";
                    s += "mov_nfa='" + y["fat_nfa"] + "' ";
                    s += "ORDER BY mov_rfa";
                    DataTable tMov = _clsFun.FillTabSql("GesMovimenti", s, false, _strConSql);

                    foreach(DataRow yy in tMov.Rows)
                    {
                        iRow++;
                        x = t.NewRow();
                        x["tmp_tri"] = "2R";
                        //x["tmp_key"] = (string)y["fat_yfa"] + (string)y["fat_nfa"];
                        x["tmp_key"] = (string)y["fat_yfa"] + (string)y["fat_ndo"];
                        x["tmp_row"] = iRow.ToString("000000");
                        x["tmp_ndo"] = "";
                        //x["tmp_ddo"] = "";
                        x["tmp_cod"] = yy["mov_art"];
                        x["tmp_des"] = yy["mov_ard"];
                        //if ((decimal)yy["mov_qkg"] > 0)
                        if (!DBNull.Value.Equals(yy["art_umi"]) && (string)yy["art_umi"] == "KG")
                            x["tmp_qta"] = yy["mov_qkg"];
                        else
                            x["tmp_qta"] = yy["mov_qta"];

                        //if(_clsDef.TIPFOR)

                        x["tmp_prz"] = (decimal)yy["mov_prv"];

                        x["tmp_imp"] = (decimal)yy["mov_imp"];

                        j = tArt.Select("art_cod='" + yy["mov_art"] + "'");
                        if(j.Length > 0)
                            x["tmp_tmp"] = (string)j[0]["tab_des"];

                        t.Rows.Add(x);
                    }

                }
            }
            else
            {
                s = "SELECT * FROM TabMovCausali";
                DataTable tCau = _clsFun.FillTabSql("TabCau", s, false, _strConSql);

                s = "SELECT ";
                s += "mot_ymo, ";
                s += "mot_nmo, ";
                s += "CASE WHEN mot_nfa IS NULL OR  mot_nfa = '' THEN '' ELSE 'S' END AS DocFat, ";
                s += "mot_ndo, ";
                s += "mot_ddo, ";
                s += "mot_cfo, ";
                s += "mot_cau, ";
                s += "mot_neg, ";
                s += "AnaClienti.cli_des AS DocCfd, ";
                s += "R.mov_imp AS MovImp, ";
                s += "R.mov_imp + R.MovIva AS MovIva, ";
                s += "mot_no1 ";
                s += "FROM GesMovTestate ";
                s += "LEFT OUTER JOIN AnaClienti ON GesMovTestate.mot_cfo = AnaClienti.cli_cod ";

                s += "LEFT JOIN (";

                s += "SELECT ";
                s += "mov_nmo, ";
                s += "SUM(mov_imp) AS mov_imp, ";
                s += "SUM(CASE WHEN tab_ali IS NOT NULL AND tab_ali > 0 AND mov_imp > 0 THEN mov_imp * tab_ali / 100 ELSE 0 END) AS MovIva ";
                s += "FROM GesMovimenti ";
                s += "LEFT OUTER JOIN TabIva ON GesMovimenti.mov_iva = TabIva.tab_cod ";
                s += "WHERE (mov_ann = 0 AND mov_ymo='" + cmbYea.Text + "') ";
                s += "GROUP BY mov_ymo, mov_nmo ";

                s += ") R ON GesMovTestate.mot_nmo = R.mov_nmo ";

                s += "WHERE ";
                s += "mot_cau = '" + sTpd + "' AND ";
                s += "(GesMovTestate.mot_ddo >= " + _clsFun.DaySql(dayIni) + ") AND ";
                s += "(GesMovTestate.mot_ddo <= " + _clsFun.DaySql(dayFin) + ") AND ";
                s += "(GesMovTestate.mot_ann = 0) ";

                s += "ORDER BY mot_yfa, mot_nfa";

                DataTable tMot = _clsFun.FillTabSql("TabMot", s, false, _strConSql);

                foreach (DataRow y in tMot.Rows)
                {
                    iRow++;
                    x = t.NewRow();
                    x["tmp_tri"] = "1T";
                    x["tmp_key"] = (string)y["mot_ymo"] + (string)y["mot_nmo"];
                    x["tmp_row"] = iRow.ToString("000000");
                    if (DBNull.Value.Equals(y["mot_ndo"]) || (string)y["mot_ndo"] == "")
                        x["tmp_ndo"] = y["mot_nmo"];
                    else
                        x["tmp_ndo"] = y["mot_ndo"];

                    x["tmp_ddo"] = y["mot_ddo"];

                    sCfo = _clsDef.TIPFOR;

                    j = tCau.Select("tab_cod='" + y["mot_cau"] + "'");
                    if (j.Length > 0 && (string)j[0]["tab_cfo"] == _clsDef.TIPCLI)
                    {
                        sCfo = _clsDef.TIPCLI;

                        x["tmp_cod"] = y["mot_cfo"];
                        x["tmp_des"] = y["mot_cfo"];

                        j = tCli.Select("cli_cod='" + y["mot_cfo"] + "'");
                        if (j.Length > 0)
                        {
                            x["tmp_des"] += " " + (string)j[0]["cli_des"];
                        }
                    }
                    j = tCau.Select("tab_cod='" + y["mot_cau"] + "'");
                    if (j.Length > 0)
                    {
                        x["tmp_des"] += "DOCUMENTO " + (string)x["tmp_ndo"] + " "  + (string)j[0]["tab_des"];

                    }
                    if ((string)j[0]["tab_sgm"] == "-")
                        x["tmp_imp"] = (decimal)y["MovImp"];
                    else
                        x["tmp_imp"] = (decimal)y["MovImp"] * -1;

                    t.Rows.Add(x);

                    s = "SELECT GesMovimenti.*, AnaArticoli.art_umi FROM GesMovimenti ";
                    s += "LEFT OUTER JOIN AnaArticoli ON GesMovimenti.mov_art = AnaArticoli.art_cod ";
                    s += "WHERE ";
                    s += "mov_ymo='" + y["mot_ymo"] + "' AND ";
                    s += "mov_nmo='" + y["mot_nmo"] + "' AND ";
                    s += "mov_rmo<>'xxxx'";

                    DataTable tMov = _clsFun.FillTabSql("GesMovimenti", s, false, _strConSql);

                    foreach(DataRow yy in tMov.Rows)
                    {
                        iRow++;
                        x = t.NewRow();
                        x["tmp_tri"] = "2R";
                        x["tmp_key"] = (string)y["mot_ymo"] + (string)y["mot_nmo"];
                        x["tmp_row"] = iRow.ToString("000000");
                        x["tmp_ndo"] = "";
                        //x["tmp_ddo"] = "";
                        x["tmp_cod"] = yy["mov_art"];
                        x["tmp_des"] = yy["mov_ard"];
                        //if ((decimal)yy["mov_qkg"] > 0)
                        if (!DBNull.Value.Equals(yy["art_umi"]) && (string)yy["art_umi"] == "KG")
                            x["tmp_qta"] = yy["mov_qkg"];
                        else
                            x["tmp_qta"] = yy["mov_qta"];

                        if(sCfo == _clsDef.TIPCLI)
                            x["tmp_prz"] = (decimal)yy["mov_prv"];
                        else
                            x["tmp_prz"] = (decimal)yy["mov_cos"];

                        x["tmp_imp"] = (decimal)yy["mov_imp"];

                        s = "SELECT lia_cos, lia_for, AnaFornitori.for_des ";
                        s += "FROM GesLisAcquisto ";
                        s += "LEFT OUTER JOIN AnaFornitori ON GesLisAcquisto.lia_for = AnaFornitori.for_cod ";
                        s += "WHERE lia_art='" + (string)yy["mov_art"] + "' "; // AND (lia_tip='F' OR lia_tip='M') ";
                        s += "ORDER BY lia_dti DESC";
                        DataTable tTmp = _clsFun.FillTabSql("", s, true, _strConSql);
                        if(tTmp.Rows.Count > 0)
                        {
                            x["tmp_for"] = (string)tTmp.Rows[0]["lia_for"];
                            x["tmp_fod"] = (string)tTmp.Rows[0]["for_des"];
                            x["tmp_cof"] = (decimal)tTmp.Rows[0]["lia_cos"];
                        }
                            
                        t.Rows.Add(x);
                    }

                }
            }

            DataTable tt = t.Clone();

            DataView v = new DataView(t, "", "tmp_key, tmp_row, tmp_ddo, tmp_des", DataViewRowState.CurrentRows);

            DateTime dDay = dayIni;
            decimal dImp = 0;
            decimal dTot = 0;
            string sKey = "";

            foreach (DataRowView r in v)
            {
                if((string)r["tmp_tri"] == "2R")
                   dTot += (decimal)r["tmp_imp"];

                //if (dImp > 0 && DateTime.Compare((DateTime)r["tmp_ddo"], dDay) != 0)
                if (sKey != (string)r["tmp_key"])
                {
                    iRow++;
                    sKey = (string)r["tmp_key"];
                    x = tt.NewRow();
                    x["tmp_tri"] = "3P";
                    x["tmp_row"] = iRow.ToString("000000");
                    x["tmp_key"] = r["tmp_key"];
                    x["tmp_ddo"] = dDay;
                    x["tmp_imp"] = (decimal)r["tmp_imp"];
                    x["tmp_des"] = "--------- " + "TOTALE DOCUMENTO " + (string)r["tmp_ndo"];

                    //x["tmp_tmp"] = r["tmp_key"];

                    tt.Rows.Add(x);
                    //dImp = 0;
                        
                    //dDay = (DateTime)r["tmp_ddo"];
                }

                //dImp += (decimal)r["tmp_imp"];
                //dTot += (decimal)r["tmp_imp"];
                tt.ImportRow(r.Row);
            }

            if (dTot > 0)
            {
                x = tt.NewRow();
                x["tmp_key"] = "YYYYYY";
                x["tmp_ddo"] = dDay;
                x["tmp_imp"] = dTot;
                x["tmp_des"] = "******* " + "TOTALE DOCUMENTI";
                tt.Rows.Add(x);
                dImp = 0;
            }

            //x = tt.NewRow();
            //x["tmp_key"] = "ZZZZZ";
            //x["tmp_imp"] = dTot;
            //tt.Rows.Add(x);

            //t = tt.Copy();
            t.Clear();

            v = new DataView(tt, "", "tmp_key, tmp_row", DataViewRowState.CurrentRows);
            foreach (DataRowView r in v)
                t.ImportRow(r.Row);

            //dataGridView1.DataSource = t;

            string sTit = "Documenti a dettaglio dal " + dayIni.ToString("dd/MM/yyyy") + " dal " + dayFin.ToString("dd/MM/yyyy");
            string sFoo = ""; // "11," + lblCos.Text.Replace(".", "").Replace(",", ".");
            string sFil = "";

            string sFld = "";
            sFld = "";
            sFld += "tmp_ndo, D.numero, 70, StringLiteral;";
            sFld += "tmp_ddo, Data, 70, DateTime;";
            sFld += "tmp_cod, Codice, 60, StringLiteral;";
            sFld += "tmp_des, Descrizione, 200, StringLiteral;";
            sFld += "tmp_qta, Q.ta, 60, Decimal;";
            sFld += "tmp_prz, Prezzo, 60, Decimal;";
            sFld += "tmp_imp, Importo, 60, Decimal;";
            //sFld += "tmp_tmp, reparto, 60, StringLiteral;";
            sFld += "tmp_cof, Costo fattura, 60, Decimal;";
            sFld += "tmp_for, Fornitore, 60, StringLiteral;";
            sFld += "tmp_fod, Descrizione, 60, StringLiteral;";
            sFld += "tmp_sta, Stato, 60, StringLiteral;";

            sFil = "DocDett.xls";

            (new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);
        }

        private void fattureElettronicheToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string sSuf = "";
            DataRow[] j = ((DataTable)cmbMovFat.DataSource).Select("tab_cod='" + cmbMovFat.SelectedValue.ToString()+"'");
            if(j.Length > 0)
                sSuf = (string)j[0]["tab_suf"];

            frmGesFattElettroniche f = new frmGesFattElettroniche();
            f._strYea = cmbYea.Text;
            f._strSuf = sSuf;
            f.ShowDialog();
        }

        private void excelElencoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ((DataTable)dgv1.DataSource == null)
                MessageBox.Show("Dati non estratti!", "ELENCO DOCUMENTI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                frmUtyPeriodo f = new frmUtyPeriodo();
                f._dayIni = new DateTime(Convert.ToInt16(cmbYea.Text), 1, 1);
                //f._dayFin = new DateTime(Convert.ToInt16(cmbYea.Text), DateTime.Today.Month, DateTime.Today.Day);
                //f._dayIni = DateTime.Today;
                //f._dayIni = new DateTime(DateTime.Today.Year,1,1);
                f._dayFin = DateTime.Today;
                f.ShowDialog();
                if (f._strRes == "S")
                {
                    FillExcel(f._dayIni, f._dayFin);
                }
            }
        }

        private void excelElencoCsvToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ElencoCsv();
        }

        private void excelElencoDettaglioToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmUtyPeriodo f = new frmUtyPeriodo();
            f._dayIni = new DateTime(DateTime.Today.Year, 1, 1); // DateTime.Today;
            f._dayFin = DateTime.Today;
            f.ShowDialog();
            if (f._strRes == "S")
            {
                Console.WriteLine("xxx");
                AllDocs2ExcelDett(f._dayIni, f._dayFin);
            }
        }
   }
}
