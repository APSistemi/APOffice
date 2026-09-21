using System;
using System.Collections;
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
    public partial class frmGesDocumento : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strCfo = "";
        private string _strIvaPar = "022";

        public string _strMovFat = "";
        public string _strCauTpd = "";
        public string _strMftYea = "";
        public string _strMftNum = "";
        public string _strMftUbi = "";                //Luogo dove vengono inseriti i dati Sede o Negozio
        public string _strReturn = "";
        public string _strDocSuf = "";

        private string _strCliLisIva = "";

        public DataSet _dasGen = new DataSet();

        private const string TABGESFAT = "GesFatTestate";
        private const string TABGESMOT = "GesMovTestate";
        private const string TABGESMOV = "GesMovimenti";

        private const string TABANACLI = "AnaClienti";
        private const string TABANAFOR = "AnaFornitori";
        private const string TABLISACQ = "GesLisAcquisto";

        private const string TABTABDOT = "TabDocTipo";
        private const string TABTABMCA = "TabMovCausali";
        private const string TABTABIVA = "TabIva";
        private const string TABTABUMI = "TabUmi";
        private const string TABTABPAG = "TabPagamenti";
        private const string TABTABSTA = "TabStato";
        private const string NRIVUOTA = "xxxx";
        private const string TABTABLIS = "TabListini";

        private Boolean _bolSeek = true;

        public frmGesDocumento()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            this.FormClosing += new FormClosingEventHandler(this.frmGesDocumento_FormClosing);
        }

        private void frmGesDocumento_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmGesDocumento_dgv1");
            clsUiIcons.SaveGridColumnWidths(dgv2, "frmGesDocumento_dgv2");
        }

        private void frmGesDocumento_Load(object sender, EventArgs e)
        {
            clsUiIcons.RestoreFormBounds(this);
            _strIvaPar = _clsFun.ParGet(clsDefine.enuParametri.Par012CodIVAxDefault, _strConSql);
            _strMftUbi = _clsFun.FileIni("R", clsDefine.enuIni.Ini13Ubicazione, "");
            if (_strMftUbi == "")
                _strMftUbi = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");

            if (_strMovFat != "F")
            {
                lblEle.Visible = false;
                txtFatEle.Visible = false;
            }

            _strConSql = _clsFun.ConSql("");
            FillTab();
            SetDgv1();
            SetDgv2();
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmGesDocumento_dgv1");
            clsUiIcons.RestoreGridColumnWidths(dgv2, "frmGesDocumento_dgv2");
            FillDati();

            btnLotto.Visible = false;
            string s = _clsFun.ParGet(clsDefine.enuParametri.Par031Lotti2Pos, _strConSql);
            if (s != "" && s.Substring(0,1) == "S")
            {

                if (_strMovFat == "F")
                    btnLotto.Visible = true;
                else if(_strMovFat == "M")
                    btnLotto.Visible = true;
            }

            ApplyModernUi();
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

                    if (letturaDaScontrinoToolStripMenuItem != null)
                        letturaDaScontrinoToolStripMenuItem.Image = clsUiIcons.GetIcon("receipt", 16);

                    if (toolStripMenuItem1 != null)
                        toolStripMenuItem1.Image = clsUiIcons.GetIcon("document", 16);

                    if (eliminaRigheToolStripMenuItem != null)
                        eliminaRigheToolStripMenuItem.Image = clsUiIcons.GetIcon("delete_rows", 16);

                    if (eliminaRigheCancellateToolStripMenuItem != null)
                        eliminaRigheCancellateToolStripMenuItem.Image = clsUiIcons.GetIcon("eraser", 16);

                    if (eliminaDocumentoToolStripMenuItem != null)
                        eliminaDocumentoToolStripMenuItem.Image = clsUiIcons.GetIcon("trash", 16);

                    if (importDaTerminalinoToolStripMenuItem != null)
                        importDaTerminalinoToolStripMenuItem.Image = clsUiIcons.GetIcon("terminal", 16);

                    if (importToolStripMenuItem != null)
                        importToolStripMenuItem.Image = clsUiIcons.GetIcon("import", 16);

                    if (documentiToolStripMenuItem != null)
                        documentiToolStripMenuItem.Image = clsUiIcons.GetIcon("document", 16);

                    if (bilanciaToolStripMenuItem != null)
                        bilanciaToolStripMenuItem.Image = clsUiIcons.GetIcon("scale", 16);

                    if (statisticheToolStripMenuItem != null)
                        statisticheToolStripMenuItem.Image = clsUiIcons.GetIcon("chart", 16);

                    if (controlloDaTerminalinoToolStripMenuItem != null)
                        controlloDaTerminalinoToolStripMenuItem.Image = clsUiIcons.GetIcon("check", 16);
                }

                if (dgv1 != null) clsUiIcons.StyleDataGridView(dgv1);
                if (dgv2 != null) clsUiIcons.StyleDataGridView(dgv2);

                if (btnPrn != null)
                {
                    clsUiIcons.StyleButton(btnPrn, clsUiIcons.GetIcon("print", 16), Color.FromArgb(239, 246, 255), Color.FromArgb(29, 78, 216));
                    btnPrn.Font = new Font("Segoe UI", 9.0f, FontStyle.Bold);
                    btnPrn.Text = " Stampa";
                }

                if (btnLotto != null)
                    clsUiIcons.StyleButton(btnLotto, clsUiIcons.GetIcon("box", 16));

                if (btnMftSco != null)
                    clsUiIcons.StyleButton(btnMftSco, clsUiIcons.GetIcon("discount", 16));
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("ApplyModernUi", ex.Message);
            }
        }

        private void frmGesDocumento_KeyDown(object sender, KeyEventArgs e)
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
            _strReturn = cmbMftSta.Text + "|" + cmbMftCfo.Text + "|" + txtMftNdo.Text + "|" + lblTotTot.Text + "|" + txtMftNo1.Text.Trim();

            Boolean b = Salva();
            if (!b)
            {
                if (MessageBox.Show("Continui l'uscita?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    this.Close();
            }
            else
            {
                if (_strMovFat == "M" && _strMftNum == _clsDef.CODNEW)
                {
                    DataRow[] j = ((DataTable)cmbMftTdc.DataSource).Select("tab_cod='" + cmbMftTdc.SelectedValue.ToString() +  "'");
                    if(j.Length > 0 && !DBNull.Value.Equals(j[0]["tab_tip"]))
                    {
                        string s = ((string)j[0]["tab_tip"]).Trim();

                        string[] a = s.Split('-');
                        if (a.Length > 1)
                        {
                            s = a[1];

                            if (s.Length == 3 && _clsFun.Numerico(s))
                            {
                                _strCauTpd = s;

                                s = "SELECT * FROM TabMovCausali WHERE tab_cod='" + _strCauTpd + "'";
                                DataTable t = _clsFun.FillTabSql(TABTABMCA, s, false, _strConSql);

                                //if (t.Rows.Count > 0)
                                //{
                                //    _strCfo = (string)t.Rows[0]["tab_cfo"];
                                //}

                                cmbMftTdc.DataSource = t;
                                cmbMftTdc.DisplayMember = "tab_des";
                                cmbMftTdc.ValueMember = "tab_cod";
                                cmbMftTdc.SelectedValue = _strCauTpd;
                                //cmbMftTdc.DropDownStyle = ComboBoxStyle.Simple;

                                lblMftNum.Text = _clsDef.CODNEW;

                                Salva();
                            }
                        }
                    }
                }

                this.Close();
            }
        }

        private void SetDgv1()
        {
            DataTable tUmi = (DataTable)cmbMovUmi.DataSource;
            DataTable tIva = (DataTable)cmbMovIva.DataSource;

            dgv1.AutoGenerateColumns = false;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            dgv1.AllowUserToResizeColumns = true;
            dgv1.AllowUserToResizeRows = false;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            if (_strMovFat == "F")
                cTbc.DataPropertyName = "mov_rfa";
            else
                cTbc.DataPropertyName = "mov_rmo";
            cTbc.Name = "Riga";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            cTbc.ToolTipText = "Doppio click per inserimento riga";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_ard";
            cTbc.Name = "Descrizione";
            cTbc.Width = 320;
            cTbc.MinimumWidth = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 90;
            cTbc.ToolTipText = "In ricerca * per articolo fornitore in fattura di acquisto";
            dgv1.Columns.Add(cTbc);

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "mov_umi";
            cCmb.Name = "UM";
            cCmb.Width = 60;
            cCmb.DataSource = tUmi;
            cCmb.ValueMember = "tab_cod";
            cCmb.DisplayMember = "tab_cod";
            cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            dgv1.Columns.Add((DataGridViewColumn)cCmb);

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "mov_iva";
            cCmb.Name = "IVA";
            cCmb.Width = 70;
            cCmb.DataSource = tIva;
            cCmb.ValueMember = "tab_cod";
            cCmb.DisplayMember = "tab_des";
            cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            dgv1.Columns.Add((DataGridViewColumn)cCmb);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_qta";
            cTbc.Name = "Q.tà";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.MaxInputLength = 10;
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_qkg";
            cTbc.Name = "Peso";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            cTbc.MaxInputLength = 10;
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            cTbc.MaxInputLength = 10;
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_prv";
            cTbc.Name = "Prezzo vendita";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            cTbc.MaxInputLength = 10;
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_sco";
            cTbc.Name = "Sconti";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.MaxInputLength = 15;
            cTbc.ReadOnly = false;
            cTbc.ToolTipText = "Se sconto valore precedere con 'V'";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_imp";
            cTbc.Name = "Importo";
            cTbc.Width = 65;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.DataPropertyName = "mov_ann";
            cCbc.Name = "Ann.";
            cCbc.Width = 35;
            cCbc.ValueType = typeof(string);
            cCbc.ReadOnly = false;
            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "MovMdy";
            cTbc.Name = "Modificato";
            cTbc.Width = 0;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.Visible = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_lot";
            cTbc.Name = "Lotto";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_ori";
            cTbc.Name = "Origine";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            if (_strMovFat == "F")
                cTbc.DataPropertyName = "mov_nmo";
            else
                cTbc.DataPropertyName = "mov_nfa";

            cTbc.Name = "DocOrigine";
            cTbc.HeaderText = "Doc. Origine";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_ubi";
            cTbc.Name = "Negozio/Sede";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
        }

        private void SetDgv2()
        {
            dgv2.AutoGenerateColumns = false;
            dgv2.AllowUserToAddRows = false;
            dgv2.ReadOnly = false;
            dgv2.AllowUserToDeleteRows = false;
            dgv2.AllowUserToResizeColumns = true;
            dgv2.RowHeadersVisible = false;
            dgv2.ScrollBars = ScrollBars.Vertical;

            DataGridViewTextBoxColumn cTbc;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 34;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_des";
            cTbc.Name = "Descrizione";
            cTbc.MinimumWidth = 70;
            cTbc.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_ali";
            cTbc.Name = "Ali";
            cTbc.Width = 34;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_imp";
            cTbc.Name = "Imponibile";
            cTbc.Width = 65;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_iva";
            cTbc.Name = "Imposta";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_tot";
            cTbc.Name = "Totale";
            cTbc.Width = 65;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_arr";
            cTbc.Name = "Arr";
            cTbc.Width = 45;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "0.000";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);
        }

        private void FillTab()
        {
            string s = "";
            //string sCfo = "";
            DataTable t = new DataTable();
            DataRow x;

            if (_strMovFat == "F")
            {
                s = "SELECT * FROM TabDocTipo WHERE tab_cod='" + _strCauTpd + "'";
                t = _clsFun.FillTabSql(TABTABDOT, s, false, _strConSql);
                if (t.Rows.Count > 0)
                {
                    this.Text += " " + (string)t.Rows[0]["tab_des"];
                    _strCfo = (string)t.Rows[0]["tab_cfo"];
                }

                s = "SELECT * FROM TabDocTpd WHERE tab_ann=0";
                t = _clsFun.FillTabSql("TabDocTipo", s, false, _strConSql);
                cmbMftTdc.DataSource = t;
                cmbMftTdc.DisplayMember = "tab_des";
                cmbMftTdc.ValueMember = "tab_cod";

                if (t.Rows.Count == 0)
                    MessageBox.Show("Tabella tipo documento fattura vuota!", "COMPILARE TABELLA", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else
                {
                    if (_strCauTpd == "" || (_strCauTpd != "PR" && _strCauTpd != "NA" && _strCauTpd != "FV"))
                        cmbMftTdc.SelectedValue = ((DataTable)cmbMftTdc.DataSource).Rows[0]["tab_cod"];
                    else
                        cmbMftTdc.SelectedValue = "FA"; // _strCauTpd;
                }
            }
            else
            {
                s = "SELECT * FROM TabMovCausali WHERE tab_cod='" + _strCauTpd + "'";
                t = _clsFun.FillTabSql(TABTABMCA, s, false, _strConSql);

                if (t.Rows.Count > 0)
                {
                    _strCfo = (string)t.Rows[0]["tab_cfo"];
                }

                //sCfo = (string)t.Rows[0]["tab_cfo"];
                cmbMftTdc.DataSource = t;
                cmbMftTdc.DisplayMember = "tab_des";
                cmbMftTdc.ValueMember = "tab_cod";
                cmbMftTdc.SelectedValue = _strCauTpd;
                cmbMftTdc.DropDownStyle = ComboBoxStyle.Simple;
                cmbMftTdc.Enabled = false;
            }

            //if (sCfo == _clsDef.TIPCLI)
            //    s = "SELECT cli_cod AS CfoCod, cli_des AS CfoDes FROM " + TABANACLI + " ORDER BY cli_des";
            //else
            //    s = "SELECT for_cod AS CfoCod, for_des AS CfoDes FROM " + TABANAFOR + " ORDER BY for_des";

            //t = _clsFun.FillTabSql(TABTABMCA, s, false, _strConSql);
            //x = t.NewRow();
            //x["CfoCod"] = "";
            //x["CfoDes"] = "  Non definito";
            //t.Rows.InsertAt(x, 0);

            //cmbMftCfo.DataSource = t;
            //cmbMftCfo.DisplayMember = "CfoDes";
            //cmbMftCfo.ValueMember = "CfoCod";
            //cmbMftCfo.SelectedValue = "";

            s = "SELECT * FROM " + TABTABIVA + " ORDER BY tab_des";
            t = _clsFun.FillTabSql(TABTABIVA, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbMovIva.DataSource = t;
            cmbMovIva.DisplayMember = "tab_des";
            cmbMovIva.ValueMember = "tab_cod";
            cmbMovIva.SelectedValue = "";

            s = "SELECT * FROM " + TABTABUMI + " ORDER BY tab_des";
            t = _clsFun.FillTabSql(TABTABUMI, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = " Non def.";
            t.Rows.InsertAt(x, 0);
            cmbMovUmi.DataSource = t;
            cmbMovUmi.DisplayMember = "tab_des";
            cmbMovUmi.ValueMember = "tab_cod";
            cmbMovUmi.SelectedValue = "";

            s = "SELECT * FROM " + TABTABPAG + " ORDER BY tab_des";
            t = _clsFun.FillTabSql(TABTABPAG, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = " Non def.";
            t.Rows.InsertAt(x, 0);
            cmbMftTpg.DataSource = t;
            cmbMftTpg.DisplayMember = "tab_des";
            cmbMftTpg.ValueMember = "tab_cod";
            cmbMftTpg.SelectedValue = "";

            s = "SELECT * FROM TabStato WHERE tab_doc=1 ORDER BY tab_des";
            t = _clsFun.FillTabSql(TABTABPAG, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = " Non def.";
            t.Rows.InsertAt(x, 0);
            cmbMftSta.DataSource = t;
            cmbMftSta.DisplayMember = "tab_des";
            cmbMftSta.ValueMember = "tab_cod";
            cmbMftSta.SelectedValue = "";

            s = "SELECT * FROM TabNegozi ORDER BY tab_cod";
            t = _clsFun.FillTabSql("TabNegozi", s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = " Non def.";
            t.Rows.InsertAt(x, 0);
            cmbMftNeg.DataSource = t;
            cmbMftNeg.DisplayMember = "tab_des";
            cmbMftNeg.ValueMember = "tab_cod";
            cmbMftNeg.SelectedValue = "";

            s = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
            if (s == "")
                s = "001";

            cmbMftNeg.SelectedValue = s;

            if (_strMovFat == "M")
            {
                DataRow[] j = ((DataTable)cmbMftTdc.DataSource).Select("tab_cod='" + cmbMftTdc.SelectedValue.ToString() +  "'");
                if (j.Length > 0 && ((string)j[0]["tab_mag"]).Trim() != "")
                {
                    string sMag = (string)j[0]["tab_mag"];
                    j = ((DataTable)cmbMftNeg.DataSource).Select("tab_mag='" + sMag + "'");
                    if(j.Length > 0)
                        cmbMftNeg.SelectedValue = (string)j[0]["tab_cod"];
                }
            }

            FillTabCfo();

            cmbMftSco.Items.Add("€");
            cmbMftSco.Items.Add("%");
            cmbMftSco.SelectedIndex = 0;
        }

        private void FillTabCfo()
        {
            if (_strCfo != "")
            {
                string s = "";
                if (_strCfo == _clsDef.TIPCLI)
                    s = "SELECT cli_cod AS CfoCod, cli_des AS CfoDes, cli_lic AS CfoLic, cli_tpa AS CfoTpa FROM AnaClienti ORDER BY cli_des";
                else
                    s = "SELECT for_cod AS CfoCod, for_des AS CfoDes, '' AS CfoLic, '' AS CfoTpa FROM AnaFornitori ORDER BY for_des";

                DataTable t = _clsFun.FillTabSql(TABTABMCA, s, false, _strConSql);
                DataRow x = t.NewRow();
                x["CfoCod"] = "";
                x["CfoDes"] = "  Non definito";
                t.Rows.InsertAt(x, 0);

                cmbMftCfo.DataSource = t;
                cmbMftCfo.DisplayMember = "CfoDes";
                cmbMftCfo.ValueMember = "CfoCod";
                cmbMftCfo.SelectedValue = "";
            }
        }

        private void FillDati()
        {
            if (((DataTable)cmbMftTdc.DataSource).Rows.Count == 0)
                MessageBox.Show("Tabella tipo documento fattura vuota!", "COMPILARE TABELLA", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                string s = "";
                string sArr = "";
                DataRow[] j;
                string sFldNri = "";
                string sMftDes = "";

                lblMftYea.Text = _strMftYea;
                lblMftNum.Text = _strMftNum;

                if (_strMovFat == "F")
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
                    s += "fat_tpg AS DocTpg, ";
                    s += "fat_no1 AS DocNo1, ";
                    s += "fat_ann AS DocAnn, ";
                    s += "fat_pvi AS DocPvi, ";
                    s += "fat_sta AS DocSta, ";
                    s += "fat_des AS DocDes, ";
                    s += "fat_ele AS DocEle ";
                    //s += "fat_arr AS DocArr ";
                    s += "FROM GesFatTestate WHERE ";
                    s += "fat_yfa='" + lblMftYea.Text + "' AND ";
                    s += "fat_nfa='" + lblMftNum.Text + "' ";
                }
                else
                {
                    s = "SELECT ";
                    s += "mot_ubi AS DocUbi, ";
                    s += "mot_ymo AS DocYea, ";
                    s += "mot_nmo AS DocNum, ";
                    s += "mot_ndo AS DocNdo, ";
                    s += "mot_ddo AS DocDdo, ";
                    s += "mot_neg AS DocNeg, ";
                    s += "mot_cfo AS DocCfo, ";
                    s += "mot_tpg AS DocTpg, ";
                    s += "mot_no1 AS DocNo1, ";
                    s += "mot_ann AS DocAnn, ";
                    s += "mot_pvi AS DocPvi, ";
                    s += "mot_sta AS DocSta, ";
                    s += "mot_des AS DocDes ";
                    s += "FROM GesMovTestate WHERE ";
                    s += "mot_ymo='" + lblMftYea.Text + "' AND ";
                    s += "mot_nmo='" + lblMftNum.Text + "' ";
                }
                DataTable t = _clsFun.FillTabSql("DOC", s, false, _strConSql);
                if (t.Rows.Count > 0)
                {
                    cmbMftCfo.SelectedValue = t.Rows[0]["DocCfo"];

                    if (_strCfo == _clsDef.TIPCLI && cmbMftCfo.SelectedValue != null)
                    {
                        j = ((DataTable)cmbMftCfo.DataSource).Select("CfoCod='" + cmbMftCfo.SelectedValue.ToString()  + "'");
                        //if (j.Length > 0 && !DBNull.Value.Equals(j[0]["CfoLic"]))
                        //    _strCliLisIva = (string)j[0]["CfoLic"];
                        if (j.Length > 0 && !DBNull.Value.Equals(j[0]["CfoLic"]))
                            _strCliLisIva = (string)j[0]["CfoLic"];
                    }

                    txtMftNdo.Text = (string)t.Rows[0]["DocNdo"];

                    if (_strMovFat == "F" && !DBNull.Value.Equals(t.Rows[0]["DocDdo"]))
                        cmbMftTdc.SelectedValue = (string)t.Rows[0]["DocTdo"];

                    if (cmbMftTdc.SelectedValue == null)
                        cmbMftTdc.SelectedValue = ((DataTable)cmbMftTdc.DataSource).Rows[0]["tab_cod"];

                    dtpMftDdt.Value = (DateTime)t.Rows[0]["DocDdo"];
                    txtMftNo1.Text = (string)t.Rows[0]["DocNo1"];
                    cmbMftTpg.SelectedValue = (string)t.Rows[0]["DocTpg"];
                    cmbMftSta.SelectedValue = (string)t.Rows[0]["DocSta"];
                    chkMftAnn.Checked = (Boolean)t.Rows[0]["DocAnn"];

                    //cmbMftNeg.SelectedValue = "";
                    if (!DBNull.Value.Equals(t.Rows[0]["DocNeg"]) && ((string)t.Rows[0]["DocNeg"]).Trim() != "")
                        cmbMftNeg.SelectedValue = (string)t.Rows[0]["DocNeg"];

                    _strMftUbi = (string)t.Rows[0]["DocUbi"];

                    sMftDes = (string)t.Rows[0]["DocDes"];
                    if (t.Columns.Contains("DocEle"))
                        txtFatEle.Text = (string)t.Rows[0]["DocEle"];
                }

                if (_strMftUbi == "")
                    _strMftUbi = cmbMftNeg.SelectedValue.ToString();

                s = "SELECT ";
                s += "mov_idx, ";
                s += "mov_ubi, ";
                s += "mov_yfa, ";
                s += "mov_nfa, ";
                s += "mov_rfa, ";
                s += "mov_ymo, ";
                s += "mov_nmo, ";
                s += "mov_rmo, ";
                s += "mov_art, ";
                s += "CASE WHEN mov_ard IS NULL OR mov_ard = '' THEN AnaArticoli.art_des ELSE mov_ard END AS mov_ard, ";
                s += "mov_iva, ";
                s += "mov_umi, ";
                s += "mov_tar, ";
                s += "mov_qta, ";
                s += "mov_qkg, ";
                s += "mov_cos, ";
                s += "mov_prv, ";
                s += "mov_sco, ";
                s += "mov_imp, ";
                s += "mov_ann, ";
                s += "mov_day, ";
                s += "mov_ori, ";
                s += "mov_lot, ";
                s += "mov_no1 ";
                s += "FROM GesMovimenti ";
                s += "LEFT OUTER JOIN AnaArticoli ON GesMovimenti.mov_art = AnaArticoli.art_cod ";
                s += "WHERE ";
                if (_strMovFat == "F")
                {
                    s += "mov_yfa='" + lblMftYea.Text + "' AND mov_nfa='" + lblMftNum.Text + "' ";
                    s += "ORDER BY mov_rfa";
                    sFldNri = "mov_rfa";
                }
                else
                {
                    s += "mov_ymo='" + lblMftYea.Text + "' AND mov_nmo='" + lblMftNum.Text + "' ";
                    s += "ORDER BY mov_rmo";
                    sFldNri = "mov_rmo";
                }

                t = _clsFun.FillTabSql(TABGESMOV, s, false, _strConSql);

                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "MovMdy",
                    Caption = "Mdy",
                    MaxLength = 1,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });

                foreach (DataRow y in t.Rows)
                {
                    if (DBNull.Value.Equals(y["mov_umi"]) || ((string)y["mov_umi"]).Trim() != "KG")
                    {
                        y["mov_umi"] = "NR";
                        y["MovMdy"] = "S";
                    }
                    if (DBNull.Value.Equals(y["mov_iva"]) || ((string)y["mov_iva"]).Trim() == "")
                    {
                        y["mov_iva"] = "022";
                        y["MovMdy"] = "S";
                    }
                }

                if (_strCfo == _clsDef.TIPCLI)
                {
                    FillCmbCliDest(sMftDes);
                }

                DataView v = new DataView(t, "", sFldNri, DataViewRowState.CurrentRows);

                dgv1.DataSource = v;

                NewRiga(false);
                dgv2.DataSource = FillIva("", sArr);
            }
        }

        private void FillCmbCliDest(String strMftDes)
        {
            if (cmbMftCfo.SelectedValue != null)
            {
                string s = "SELECT cld_cod, cld_des FROM AnaCliDestinazioni WHERE cld_ann=0 AND cld_cli='" + cmbMftCfo.SelectedValue.ToString() + "' ORDER BY cld_des";

                DataTable tt = _clsFun.FillTabSql("AnaCliDestinazioni", s, false, _strConSql);
                DataRow x = tt.NewRow();
                x["cld_cod"] = "";
                x["cld_des"] = "";
                tt.Rows.InsertAt(x, 0);

                cmbMftDes.DataSource = tt;
                cmbMftDes.DisplayMember = "cld_des";
                cmbMftDes.ValueMember = "cld_cod";
                cmbMftDes.SelectedValue = strMftDes;
            }
        }

        private DataTable FillIva(string strRange, string strArr)
        {
            string s = "";
            DataRow[] j;
            DataRow x;
            decimal dQta = 0;
            Boolean b = false;

            int iRigIni = 0;
            int iRigFin = 10000;

            if (strRange != "")
            {
                iRigIni = Convert.ToInt16(strRange.Substring(0, 3));
                iRigFin = Convert.ToInt16(strRange.Substring(4, 3));
            }
            string sFldRig = "mov_rmo";
            if (_strMovFat == "F")
                sFldRig = "mov_rfa";

            DataTable tTiv = (DataTable)cmbMovIva.DataSource;

            if (dgv2.DataSource == null)
                dgv2.DataSource = new clsGenTabTmp().TabTmpIvaRiep("TabIva");

            DataTable tIva = (DataTable)dgv2.DataSource;
            tIva.Clear();

            DataTable t = ((DataView)dgv1.DataSource).ToTable();

            decimal dTotScon = 0;           //Seck 20180302 totale valore da scontrini
            ArrayList aryScon = new ArrayList();
            Boolean bTotScon = true;

            foreach(DataRow y in t.Rows)
            {
                if (DBNull.Value.Equals(y["mov_imp"]))
                    y["mov_imp"] = 0;
                if (DBNull.Value.Equals(y["mov_ann"]))
                    y["mov_ann"] = false;

                if (!(Boolean)y["mov_ann"] && (decimal)y["mov_imp"] != 0 && ( strRange == "" ||( Convert.ToInt16(y[sFldRig]) >= iRigIni && Convert.ToInt16(y[sFldRig]) <= iRigFin)))
                { 
                    b = true;
                    if (((string)y["mov_iva"]).Trim() == "")
                        b = false;
                    if (DBNull.Value.Equals(y["mov_imp"]) || (decimal)y["mov_imp"] == 0)
                        b = false;

                    if (b)
                    {
                        j = tIva.Select("iva_cod='" + (string)y["mov_iva"] + "'");
                        if (j.Length == 0)
                        {

                            j = tTiv.Select("tab_cod='" + ((string)y["mov_iva"]).PadLeft(3,Convert.ToChar('0')) + "'");
                            if(j.Length == 0)
                                j = tTiv.Select("tab_cod='" + _strIvaPar + "'");
                            x = tIva.NewRow();
                            x["iva_cod"] = y["mov_iva"];
                            x["iva_des"] = j[0]["tab_des"];
                            x["iva_ali"] = j[0]["tab_ali"];
                            x["iva_iva"] = 0;
                            x["iva_imp"] = 0;
                            x["iva_tot"] = 0;
                            tIva.Rows.Add(x);
                            j = tIva.Select("iva_cod='" + y["mov_iva"] + "'");
                        }

                        if ((decimal)y["mov_imp"] != 0 && !DBNull.Value.Equals(j[0]["iva_ali"]))
                        {

                            if ((string)y["mov_art"] == "0156195")
                                Console.WriteLine("aaaa");

                            //j[0]["iva_iva"] = (decimal)j[0]["iva_iva"] + _clsFun.ValIva((decimal)y["mov_imp"], (decimal)j[0]["iva_ali"]);
                            //j[0]["iva_imp"] = Math.Round((decimal)j[0]["iva_imp"] + (decimal)y["mov_imp"], 2, MidpointRounding.AwayFromZero);
                            //j[0]["iva_imp"] = Math.Round((decimal)j[0]["iva_imp"] + (decimal)y["mov_imp"], 3, MidpointRounding.ToEven);
                            j[0]["iva_imp"] = (decimal)j[0]["iva_imp"] + (decimal)y["mov_imp"];

                        }
                        else
                            Console.WriteLine("aaaaaaaaa");
                    }

                    if (!DBNull.Value.Equals(y["mov_ori"]) && ((string)y["mov_ori"]).Trim() != "")
                    {
                        s = (string)y["mov_ori"];

                        string[] a = s.Split(',');

                        if (a.Length > 4)
                        {
                            string sKey = s.Substring(0, 24);

                            if (aryScon.IndexOf(sKey) < 0)
                            {
                                aryScon.Add(sKey);
                                if(a.Length > 5 )
                                    dTotScon += Convert.ToDecimal(a[5].Replace(".", ","));
                            }
                        }
                    }
                    else if ((decimal)y["mov_imp"] > 0)     //Seck 20180304 Se ci sono righe non da scontrino con righe da scontrino non faccio arrotondamenti
                        bTotScon = false;
                }

                if((string)y["mov_umi"] == "KG")
                    dQta += 1;
                else
                    dQta += (decimal)y["mov_qta"];
            }

            if (!bTotScon)
                dTotScon = 0;

            lblTotQta.Text = _clsFun.Dec2Txt(dQta, 0);

            tIva = Riepilogo(tIva, dTotScon);

            return tIva;
        }

        private DataTable Riepilogo(DataTable tIva, decimal decScoImp)
        {
            decimal dIva = 0;
            decimal dImp = 0;
            decimal dTot = 0;

            //DataView v = new DataView(tIva, "", "tab_ali", DataViewRowState.CurrentRows);
            tIva = new DataView(tIva, "", "iva_ali", DataViewRowState.CurrentRows).Table; 

            string[] aArr = { };

            foreach (DataRow y in tIva.Rows)
            {
                if ((decimal)y["iva_imp"] != 0)
                {
                    y["iva_iva"] = Math.Round(_clsFun.ValIva((decimal)y["iva_imp"], (decimal)y["iva_ali"]), 2, MidpointRounding.ToEven);
                    y["iva_iva"] = Math.Round((decimal)y["iva_iva"], 2, MidpointRounding.ToEven);
                    y["iva_tot"] = Math.Round((decimal)y["iva_imp"] + (decimal)y["iva_iva"], 2, MidpointRounding.ToEven);
                    y["iva_imp"] = (decimal)y["iva_tot"] - (decimal)y["iva_iva"];           //20180118 Seck
                }
            }

            foreach (DataRow y in tIva.Rows)
            {
                dImp += (decimal)y["iva_imp"];
                dIva += (decimal)y["iva_iva"];
                dTot += (decimal)y["iva_tot"];
            }

            if (decScoImp > 0 && dTot != decScoImp)
            {
                decimal dDif = decScoImp - dTot;
                if (dDif > 0)
                    tIva.Rows[0]["iva_iva"] = (decimal)tIva.Rows[0]["iva_iva"] + dDif;
                else
                    tIva.Rows[0]["iva_imp"] = (decimal)tIva.Rows[0]["iva_imp"] + dDif;

                tIva.Rows[0]["iva_arr"] = dDif;

                dImp = 0;
                dIva = 0;
                dTot = 0;

                foreach (DataRow y in tIva.Rows)
                {
                    dImp += (decimal)y["iva_imp"];
                    dIva += (decimal)y["iva_iva"];
                    //dTot += (decimal)y["iva_tot"];
                    dTot += (decimal)y["iva_imp"] + (decimal)y["iva_iva"];
                }
            }

            dImp = Math.Round(dImp, 2, MidpointRounding.AwayFromZero); // MidpointRounding.ToEven);

            lblTotIva.Text = _clsFun.Dec2Txt(dIva, 2);
            lblTotImp.Text = _clsFun.Dec2Txt(dImp, 2);
            lblTotTot.Text = _clsFun.Dec2Txt(dTot, 2);

            return tIva;
        }

        private void CalcScontoTot()
        {
            string s = "";
            decimal dSco = 0;
            decimal d = 0;
            decimal dTotLordo = Convert.ToDecimal(lblTotTot.Text);

            //if (!_clsFun.Numerico(txtMftSco.Text) || Convert.ToDecimal(txtMftSco.Text) == 0)
            //    return;
            //else
            dSco = 0;
            if(_clsFun.Numerico(txtMftSco.Text))
                dSco =   Convert.ToDecimal(txtMftSco.Text);
            
            //if (_strCfo == _clsDef.TIPCLI)

            decimal dTot = 0;
            DataTable t = ((DataView)dgv1.DataSource).Table;

            foreach (DataRow y in t.Rows)
            {
                if (!(Boolean)y["mov_ann"])
                    dTot += (decimal)y["mov_imp"];
            }

            if (dSco > 0)
            {
                string sTip = cmbMftSco.SelectedItem.ToString().Substring(0, 1);

                //decimal dDelta = dSco * 100 / (dTotLordo);
                decimal dDelta = dSco * 100 / (dTot);
                decimal dScoTot = 0;

                string sRig = "";

                foreach (DataRow y in t.Rows)
                {
                    if (!(Boolean)y["mov_ann"] && (decimal)y["mov_imp"] > 0)
                    {
                        d = (decimal)y["mov_imp"];
                        d = _clsFun.MenoPer(d, dDelta);

                        d = Math.Round(d, 2, MidpointRounding.AwayFromZero);
                        s = "T" + ((decimal)y["mov_imp"] - d).ToString("#0.00");
                        if (s.Length > 15)
                            s = s.Substring(0, 15);

                        if (((string)y["mov_sco"]).Trim() != "")
                            s = "+" + s;
                        
                        y["mov_sco"] = (string)y["mov_sco"] + s;
                        y["mov_imp"] = d;

                        dScoTot += d;
                        sRig = (string)y["mov_rfa"];
                        y["MovMdy"] = "S";

                    }
                }

                if(dScoTot != (dTot-dSco))
                {
                    DataRow[] j = t.Select("mov_rfa='" + sRig + "'");
                    if(j.Length > 0)
                        j[0]["mov_imp"] = (decimal)j[0]["mov_imp"] + (dTot -dSco - dScoTot);
                }
            }
            else
            {
                foreach (DataRow y in t.Rows)
                {
                    string sSco = (string)y["mov_sco"];
                    s = "";

                    if(sSco.Contains("T"))
                    {
                        string[] a = sSco.Split('+');

                        foreach(string ss in a)
                        {
                            if (ss.Length > 0)
                            {
                                if (ss.Substring(0, 1) != "T")
                                    s += ss + "*";
                            }
                        }

                        if (s != "")
                            s = s.Substring(0, s.Length - 1);

                        y["mov_sco"] = s;
                    }

                    if (!(Boolean)y["mov_ann"] && (decimal)y["mov_imp"] > 0)
                    {
                        d = 0;

                        //y["mov_sco"] = "";
                        if ((string)y["mov_umi"] == "KG")
                        {
                            if (_strCauTpd == "FA" || _strCfo == "FOR")
                                y["mov_imp"] = (decimal)y["mov_cos"] * (decimal)y["mov_qkg"];
                            else
                                y["mov_imp"] = (decimal)y["mov_prv"] * (decimal)y["mov_qkg"];
                        }
                        else
                        {
                            if (_strCauTpd == "FA" || _strCfo == "FOR")
                                y["mov_imp"] = (decimal)y["mov_cos"] * (decimal)y["mov_qta"];
                            else
                                y["mov_imp"] = (decimal)y["mov_prv"] * (decimal)y["mov_qta"];
                        }

                        if (_strCauTpd == "FA")
                            d = (decimal)y["mov_cos"];
                        else
                            d = (decimal)y["mov_prv"];

                        dSco = CalcScoRow(y, d);
                        y["mov_imp"] = (decimal)y["mov_imp"] - dSco;
                    }
                    y["MovMdy"] = "S";

                }
            }

            dgv2.DataSource = FillIva("", "");
        }

        private void btnCfo_Click(object sender, EventArgs e)
        {
            if (_strCauTpd == "FA")
            {
                frmSeekFor f = new frmSeekFor();
                f._bolAnagra = false;
                f._bolNew = true;
                f.ShowDialog();
                if (f._strCod != "")
                {
                    if (f._tabTmp.Rows.Count > 0)
                    {
                        FillTabCfo();
                        cmbMftCfo.SelectedValue = f._tabTmp.Rows[0]["tmp_for"];
                        if (cmbMftCfo.SelectedValue == null)
                        {
                            FillTabCfo();
                            cmbMftCfo.SelectedValue = f._tabTmp.Rows[0]["tmp_cli"];
                        }
                    }
                }
            }
            else
            {
                frmSeekCli f = new frmSeekCli();
                f._bolAnagra = true;
                f._bolSelect = true;
                f.ShowDialog();
                if (f._strCod != "")
                {
                    if (f._tabTmp.Rows.Count > 0)
                    {
                        cmbMftCfo.SelectedValue = f._tabTmp.Rows[0]["tmp_cli"];

                        if(cmbMftCfo.SelectedValue == null)
                        {
                            FillTabCfo();
                            cmbMftCfo.SelectedValue = f._tabTmp.Rows[0]["tmp_cli"];
                        }

                        cmbMftTpg.SelectedValue = (string)f._tabTmp.Rows[0]["tmp_tpa"];

                    }

                    FillCmbCliDest("");
                }
            }
            if ((DataView)dgv1.DataSource == null || ((DataView)dgv1.DataSource).Count == 0)
                NewRiga(true);
        }

        private void cmbMftCfo_SelectionChangeCommitted(object sender, EventArgs e)
        {
            DataRow[] j = ((DataTable)cmbMftCfo.DataSource).Select("CfoCod='" + cmbMftCfo.SelectedValue.ToString() + "'");
            if (j.Length > 0 && !DBNull.Value.Equals(j[0]["CfoLic"]))
            {
                _strCliLisIva = (string)j[0]["CfoLic"];
                cmbMftTpg.SelectedValue = (string)j[0]["CfoTpa"];
            }

            if (_strCfo == "CLI")
                FillCmbCliDest("");

            if ((DataView)dgv1.DataSource == null || ((DataView)dgv1.DataSource).Count == 0)
                NewRiga(true);
        }

        private void NewRiga(Boolean bolCtrl)
        {
            if (_strCfo != "" && (cmbMftCfo.SelectedValue == null || cmbMftCfo.SelectedValue.ToString() == ""))
            {
                if (bolCtrl)
                    MessageBox.Show("Destinatario non definito!");
            }
            else
            {
                Boolean b = true;
                string sFld = "mov_rmo";
                if (_strMovFat == "F")
                    sFld = "mov_rfa";

                DataTable t = ((DataView)dgv1.DataSource).Table;
                 
                DataView v = new DataView(t, sFld + "='" + NRIVUOTA + "'", sFld, DataViewRowState.CurrentRows);

                if(v.Count > 0)
                {
                    b = false;
                }

                if (b)
                {
                    try
                    {
                        DataRow x = t.NewRow();
                        //x["mov_idx"] = 0;
                        x["mov_yfa"] = "";
                        x["mov_nfa"] = "";
                        x["mov_ymo"] = "";
                        x["mov_nmo"] = "";
                        x["mov_rfa"] = NRIVUOTA;
                        x["mov_rmo"] = NRIVUOTA;
                        x["mov_art"] = "";
                        x["mov_ard"] = "";
                        x["mov_iva"] = "";
                        x["mov_umi"] = "";
                        x["mov_qta"] = 0;
                        x["mov_qkg"] = 0;
                        x["mov_cos"] = 0;
                        x["mov_prv"] = 0;
                        x["mov_sco"] = "";
                        x["mov_imp"] = 0;
                        x["mov_ann"] = false;
                        x["mov_day"] = DateTime.Today;
                        x["mov_no1"] = "";
                        x["MovMdy"] = "";

                        x["mov_ori"] = "";
                        x["mov_tar"] = 0;
                        x["mov_lot"] = "";
                        t.Rows.Add(x);

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                    FillDettaglio();
                }
            }
        }

        private void FillDettaglio()    //Boolean bolCls)
        {
            if(dgv1.DataSource != null && ((DataView)dgv1.DataSource).Table.Rows.Count > 1)
            {

                try
                {
                    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {
                        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                        DataRow x = r.Row;

                        if (!DBNull.Value.Equals(x["mov_art"]))
                            txtMovArt.Text = (string)x["mov_art"];

                        if (!DBNull.Value.Equals(x["mov_ard"]))
                            txtMovArd.Text = (string)x["mov_ard"];

                        if (DBNull.Value.Equals(x["mov_iva"]))
                            x["mov_iva"] = "";
                        cmbMovIva.SelectedValue = (string)x["mov_iva"];

                        if (DBNull.Value.Equals(x["mov_umi"]))
                            x["mov_umi"] = "";
                        cmbMovUmi.SelectedValue = (string)x["mov_umi"];

                        txtMovQta.Text = "0";
                        if (_clsFun.Numerico(x["mov_qta"]))
                            txtMovQta.Text = _clsFun.Dec2Txt(x["mov_qta"], 2);

                        txtMovQkg.Text = "0";
                        if (_clsFun.Numerico(x["mov_qkg"]))
                            txtMovQkg.Text = _clsFun.Dec2Txt(x["mov_qkg"], 2);

                        txtMovCos.Text = "0";
                        if (_clsFun.Numerico(x["mov_qta"]))
                            txtMovCos.Text = _clsFun.Dec2Txt(x["mov_cos"], 3);

                        txtMovImc.Text = "0";

                        if (!DBNull.Value.Equals(x["mov_cos"]) && !DBNull.Value.Equals(x["mov_qta"]))
                        {
                            if ((decimal)x["mov_cos"] != 0 && (decimal)x["mov_qta"] != 0)
                            {
                                if ((string)x["mov_umi"] == "KG")
                                    txtMovImc.Text = _clsFun.Dec2Txt((decimal)x["mov_cos"] * (decimal)x["mov_qkg"], 3);
                                else
                                    txtMovImc.Text = _clsFun.Dec2Txt((decimal)x["mov_cos"] * (decimal)x["mov_qta"], 3);
                            }
                        }
                        txtMovPrv.Text = "0";
                        if (_clsFun.Numerico(x["mov_prv"]))
                            txtMovPrv.Text = _clsFun.Dec2Txt(x["mov_prv"], 2);

                        txtMovSco.Text = "0";

                        txtMovImp.Text = "0";
                        if (_clsFun.Numerico(x["mov_imp"]))
                            txtMovImp.Text = _clsFun.Dec2Txt(x["mov_imp"], 2);

                        if(btnLotto.Visible)
                        {
                            if (_clsFun.Numerico(x["mov_lot"], "0123456789") && Convert.ToInt16(x["mov_lot"]) > 0)
                                btnLotto.Text = "Lotto " + ((string)x["mov_lot"]).Trim();
                            else
                                btnLotto.Text = "Lotto";
                        }

                        chkMovAnn.Checked = false;
                        if (!DBNull.Value.Equals(x["mov_ann"]))
                            chkMovAnn.Checked = Convert.ToBoolean(x["mov_ann"]);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

        }

        private void MovRowAgg()
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                r.Row["mov_art"] = txtMovArt.Text;
                r.Row["mov_ard"] = txtMovArd.Text;
                r.Row["mov_iva"] = cmbMovIva.SelectedValue;
                r.Row["mov_umi"] = cmbMovUmi.SelectedValue;
                r.Row["mov_qta"] = _clsFun.Txt2Dec(txtMovQta.Text);
                r.Row["mov_qkg"] = _clsFun.Txt2Dec(txtMovQkg.Text);
                r.Row["mov_cos"] = _clsFun.Txt2Dec(txtMovCos.Text);
                r.Row["mov_prv"] = _clsFun.Txt2Dec(txtMovPrv.Text);
                r.Row["mov_imp"] = _clsFun.Txt2Dec(txtMovImp.Text);
            }
        }

        private void dgv1_CurrentCellChanged(object sender, EventArgs e)
        {
            FillDettaglio();
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                if (MessageBox.Show("Inserisci riga?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {
                        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                        DataRow x = r.Row;

                        string sFldNri = "mov_rmo";
                        if (_strMovFat == "F")
                            sFldNri = "mov_rfa";

                        string sNri = (string)x[sFldNri];

                        DataView v = (DataView)dgv1.DataSource;

                        DataRowView rr = v.AddNew();

                        rr["mov_yfa"] = x["mov_yfa"];
                        rr["mov_nfa"] = x["mov_nfa"];
                        rr["mov_rfa"] = x["mov_rfa"];
                        rr["mov_ymo"] = x["mov_ymo"];
                        rr["mov_nmo"] = x["mov_nmo"];
                        rr["mov_rmo"] = x["mov_rmo"];

                        rr[sFldNri] = (Convert.ToInt16(sNri)).ToString("0000") ;

                        rr["mov_art"] = "";
                        rr["mov_ard"] = "";
                        rr["mov_iva"] = "";
                        rr["mov_umi"] = "";
                        rr["mov_qta"] = 1;
                        rr["mov_qkg"] = 0;
                        rr["mov_cos"] = 0;
                        rr["mov_prv"] = 0;
                        rr["mov_sco"] = "";
                        rr["mov_ann"] = false;
                        rr["mov_day"] = DateTime.Today;
                        rr["mov_no1"] = "";
                        rr["MovMdy"] = "";

                        Console.WriteLine("aaaa");

                        //v.Sort = sFldNri + " ASC";
                        //DataTable t = v.ToTable();

                        v = new DataView(v.ToTable(), "", sFldNri + ", mov_ard", DataViewRowState.CurrentRows);

                        int i = 0;

                        foreach(DataRowView R in v)
                        {
                            if ((string)R[sFldNri] != NRIVUOTA)
                            {
                                i++;
                                R[sFldNri] = i.ToString("0000");
                                R["MovMdy"] = "S";
                            }
                        }

                        dgv1.DataSource = v;

                        dgv1.CurrentCell = dgv1.Rows[Convert.ToInt16(sNri)-1].Cells[0];
                    }
                }

                Console.WriteLine("aaaa");
            
            }
            else if (e.ColumnIndex == 2)
            {
                string sArt = dgv1.Rows[e.RowIndex].Cells["Articolo"].Value.ToString();
                string sArd = dgv1.Rows[e.RowIndex].Cells["Descrizione"].Value.ToString();

                if (sArt.Trim() == "" || sArd.Trim() == "")
                {
                    _bolSeek = true;

                    frmSeekArt f = new frmSeekArt();
                    f.ShowDialog();
                    if (f._tabArt != null && f._tabArt.Rows.Count > 0)
                    {
                        DataTable t = f._tabArt;
                        if (t.Rows.Count > 0)
                        {
                            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                            if (cm.Position >= 0)
                            {
                                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                                DataRow x = r.Row;

                                sArt = (string)t.Rows[0]["tmp_art"];

                                t = _clsQry.ArtSeek(sArt, "SEEK");

                                if(_strCauTpd == "FA" && t.Rows.Count > 0 && !DBNull.Value.Equals(t.Rows[0]["tmp_pxc"]))
                                    x["mov_qta"] = (decimal)t.Rows[0]["tmp_pxc"];

                                FillMovRow(t, x, false);
                            }
                        }
                    }
                }
                else if(sArt.Length > 2)
                {
                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = sArt;
                    f.ShowDialog();
                    if (f._tabArt != null && f._tabArt.Rows.Count > 0 && (string)f._tabArt.Rows[0]["art_cod"] != sArt)
                    {
                        FillRow(f._tabArt);
                    }
                }
            }
        }

        private void FillRow(DataTable tabTmp)
        {
            if (tabTmp.Rows.Count > 0)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string sArt = (string)tabTmp.Rows[0]["art_cod"];
                    tabTmp = _clsQry.ArtSeek(sArt, "SEEK");
                    FillMovRow(tabTmp, x,false);
                }
            }
        }

        private void dgv1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            TextBox txb = e.Control as TextBox;
            string s = "";

            if(txb != null && _bolSeek)
            {
                e.Control.KeyPress += new KeyPressEventHandler(txtCltr_KeyPress);

                txb.PreviewKeyDown += (S, E) =>
                {
                    if (E.KeyCode == Keys.Enter)
                    {
                        if (dgv1.CurrentCell.OwningColumn.Name.Equals("Descrizione"))
                        {
                            Console.WriteLine("aaaaaaa");
                            if (txb.Text != "")
                            {
                                Boolean bBil = false;
                                DataTable t = new DataTable();
                                if (_strCliLisIva != "")
                                {
                                    string sLis = _strCliLisIva.Substring(0, 3);

                                    s = txb.Text;
                                    if (s.Length == 13 && s.Substring(0, 1) == "2")
                                    {
                                        s = s.Substring(0, 7) + "000000";
                                        bBil = true;
                                    }

                                    t = _clsQry.ArtSeek(s, "LNE" + sLis);
                                }
                                else
                                {
                                    s = txb.Text;
                                    if (s.Length == 13 && s.Substring(0, 1) == "2")
                                    {
                                        s = s.Substring(0, 7) + "000000";
                                        bBil = true;
                                    }
                                    if (_strCauTpd == "FA" && s.Length > 1 && s.Substring(0, 1) == "*")
                                    {
                                        t = _clsQry.ArtSeek(s.Substring(1), "ARF" + cmbMftCfo.SelectedValue.ToString());
                                    }
                                    else 
                                    {
                                        if (_strCauTpd == "FA" && s.Length > 1) 
                                            t = _clsQry.ArtSeek(s, "FOR" + cmbMftCfo.SelectedValue.ToString());
                                        else
                                            t = _clsQry.ArtSeek(s, "SEEK");
                                        _bolSeek = false;
                                    }
                                }

                                if (t != null && t.Rows.Count > 0)
                                {
                                    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                                    if (cm.Position >= 0)
                                    {
                                        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                                        DataRow x = r.Row;

                                        if(bBil)
                                        {
                                            s = txb.Text.Substring(8,4);
                                            if (_clsFun.Numerico(s, "0123456789"))
                                            {
                                                if ((Boolean)t.Rows[0]["tmp_ecp"])
                                                {
                                                    x["mov_qta"] = 1;
                                                    x["mov_qkg"] = Convert.ToDecimal(s)/1000;
                                                }
                                                else
                                                     x["mov_imp"] = Convert.ToDecimal(s) / 100;
                                            }
                                        }
                                        else
                                        {
                                            if (_strCauTpd == "FA" && t.Rows.Count > 0 && !DBNull.Value.Equals(t.Rows[0]["tmp_pxc"]))
                                                x["mov_qta"] = (decimal)t.Rows[0]["tmp_pxc"];
                                        }

                                        if (!DBNull.Value.Equals(t.Rows[0]["tmp_lot"]))
                                            x["mov_lot"] = ((string)t.Rows[0]["tmp_lot"]).PadLeft(5,Convert.ToChar('0'));

                                        FillMovRow(t, x, bBil);
                                    }
                                }
                            }
                        }
                    }
                };
                //SendKeys.Send("{TAB}");
                //txb = null;
                //_bolSeek = false;
            }
        }

        private void dgv1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                try
                {
                    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {
                        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                        DataRow x = r.Row;

                        //x = CalcImporto(_strCauTpd, x, "IMP");

                        if (e.ColumnIndex == 7 || e.ColumnIndex == 8)
                            x = CalcImporto(_strCauTpd, x, "PRE");
                        else if (e.ColumnIndex == 9)
                            x = CalcImporto(_strCauTpd, x, "SCO");
                        else if (e.ColumnIndex == 5 || e.ColumnIndex == 6)
                            x = CalcImporto(_strCauTpd, x, "QTA");
                        else
                            x = CalcImporto(_strCauTpd, x, "IMP");

                        CtrlNumRiga(x);
                        cm.EndCurrentEdit();

                        dgv2.DataSource = FillIva("", "");

                        if (!DBNull.Value.Equals(x["mov_ard"]) && ((string)x["mov_ard"]).Trim() != "")
                            NewRiga(true);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            else if (e.ColumnIndex >= 2)
            {
                try
                {
                    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {
                        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                        DataRow x = r.Row;

                        //x = CalcImporto(_strCauTpd, x, "IMP");

                        if (e.ColumnIndex == 7 || e.ColumnIndex == 8)
                            x = CalcImporto(_strCauTpd, x, "PRE");
                        else if (e.ColumnIndex == 9)
                            x = CalcImporto(_strCauTpd, x, "SCO");
                        else if (e.ColumnIndex == 5 || e.ColumnIndex == 6)
                            x = CalcImporto(_strCauTpd, x, "QTA");
                        else
                            x = CalcImporto(_strCauTpd, x, "IMP");

                        CtrlNumRiga(x);
                        cm.EndCurrentEdit();

                        dgv2.DataSource = FillIva("", "");

                        if (!DBNull.Value.Equals(x["mov_ard"]) && ((string)x["mov_ard"]).Trim() != "")
                            NewRiga(true);
                    }
                }                    
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv1.Columns[e.ColumnIndex].Name == "Ann." && dgv1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv1.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
            }
        }

        private void dgv1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgv1.CurrentCell is DataGridViewCheckBoxCell || dgv1.CurrentCell is DataGridViewComboBoxCell)
            {
                dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dgv1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Console.WriteLine("aaaaaaa");
            dgv1.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
        }

        private void txtCltr_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((sender as TextBox).Text != null && (sender as TextBox).Text.Trim() != "")
            {
                string s = dgv1.CurrentCell.OwningColumn.Name.ToString();

                if (s == "Q.tà" || s == "Peso" || s == "Costo" || s == "Prezzo vendita" || s == "Importo")
                {
                    if (e.KeyChar == '.')
                        e.KeyChar = ',';
                }
            }
        }

        private DataRow CalcImporto(string strCauTpd, DataRow rowMov, string strTip)
        {
            if (DBNull.Value.Equals(rowMov["mov_imp"]))
                rowMov["mov_imp"] = 0;
            if (DBNull.Value.Equals(rowMov["mov_qta"]))
                rowMov["mov_qta"] = 0;
            if (DBNull.Value.Equals(rowMov["mov_qkg"]))
                rowMov["mov_qkg"] = 0;
            if (DBNull.Value.Equals(rowMov["mov_sco"]))
                rowMov["mov_sco"] = "";
            if (DBNull.Value.Equals(rowMov["mov_imp"]))
                rowMov["mov_imp"] = 0;
            if (DBNull.Value.Equals(rowMov["mov_cos"]))
                rowMov["mov_cos"] = 0;
            if (DBNull.Value.Equals(rowMov["mov_ann"]))
                rowMov["mov_ann"] = false;

            string sFld = "mov_cos";
            decimal dSco = 0;
            decimal d = (decimal)rowMov["mov_cos"];

            if (_strCfo == _clsDef.TIPCLI)
            {
                d = (decimal)rowMov["mov_prv"];

                sFld = "mov_prv";
            }
            else if (_strMovFat == "M")
            {
                DataRow[] j = ((DataTable)cmbMftTdc.DataSource).Select("tab_cod='" + cmbMftTdc.SelectedValue.ToString() + "'");
                if (j.Length > 0)
                {
                    if ((string)j[0]["tab_cfo"] == "CLI")
                    {
                        d = (decimal)rowMov["mov_prv"];
                        sFld = "mov_prv";
                    }
                }
            }

            if (strTip == "PRE")
            {
                if ((string)rowMov["mov_umi"] == "KG" && d > 0 && (decimal)rowMov["mov_qkg"] > 0)
                    rowMov["mov_imp"] = d * (decimal)rowMov["mov_qkg"];
                else if (d > 0 && (decimal)rowMov["mov_qta"] > 0)
                    rowMov["mov_imp"] = d * (decimal)rowMov["mov_qta"];
            }
            else if (strTip == "QTA")
            {
                if ((string)rowMov["mov_umi"] == "KG" && d > 0 && (decimal)rowMov["mov_qkg"] > 0)
                    rowMov["mov_imp"] = d * (decimal)rowMov["mov_qkg"];
                else if (d > 0 && (decimal)rowMov["mov_qta"] > 0)
                    rowMov["mov_imp"] = d * (decimal)rowMov["mov_qta"];
            }
            //else if (strTip == "SCO")
            //{
            //    if ((string)rowMov["mov_umi"] == "KG" && (decimal)rowMov["mov_imp"] > 0 && !DBNull.Value.Equals(rowMov["mov_qkg"]) && (decimal)rowMov["mov_qkg"] > 0)
            //        rowMov[sFld] = (decimal)rowMov["mov_imp"] / (decimal)rowMov["mov_qkg"];
            //    else if (!DBNull.Value.Equals(rowMov["mov_imp"]) && (decimal)rowMov["mov_imp"] > 0 && (decimal)rowMov["mov_qta"] > 0)
            //        rowMov[sFld] = (decimal)rowMov["mov_imp"] / (decimal)rowMov["mov_qta"];
            //    rowMov[sFld] = Math.Round((decimal)rowMov[sFld], 3, MidpointRounding.AwayFromZero);
            //}
            else if (strTip != "SCO" && strTip != "IMP")
            {
                if ((string)rowMov["mov_umi"] == "KG" && (decimal)rowMov["mov_imp"] > 0 && (decimal)rowMov["mov_qkg"] > 0)
                    rowMov[sFld] = (decimal)rowMov["mov_imp"] / (decimal)rowMov["mov_qkg"];
                else if ((decimal)rowMov["mov_imp"] > 0 && (decimal)rowMov["mov_qta"] > 0)
                    rowMov[sFld] = (decimal)rowMov["mov_imp"] / (decimal)rowMov["mov_qta"];

                rowMov[sFld] = Math.Round((decimal)rowMov[sFld], 3, MidpointRounding.AwayFromZero);
            }

            if (((string)rowMov["mov_sco"]).Trim() != "" && strTip != "IMP")
            {
                dSco = CalcScoRow(rowMov, d);
                //rowMov["mov_imp"] = (decimal)rowMov["mov_imp"] - dSco;

                d = 0;

                if ((string)rowMov["mov_umi"] == "KG" && (decimal)rowMov[sFld] > 0 && (decimal)rowMov["mov_qkg"] > 0)
                    rowMov["mov_imp"] = (decimal)rowMov[sFld] * (decimal)rowMov["mov_qkg"];
                else if ((decimal)rowMov[sFld] > 0 && (decimal)rowMov["mov_qta"] > 0)
                    rowMov["mov_imp"] = (decimal)rowMov[sFld] * (decimal)rowMov["mov_qta"];

                rowMov["mov_imp"] = (decimal)rowMov["mov_imp"] - dSco;

            }

            rowMov["MovMdy"] = "S";

            return rowMov;
        }

        private decimal CalcScoRow(DataRow rowMov, decimal decVal)
        {
            decimal dSco = 0;

            string s = ((string)rowMov["mov_sco"]).Trim();
            if (s.Length > 0)
            {
                if (s.Substring(0, 1) == "v")
                {
                    s = "V" + s.Substring(1);
                    rowMov["mov_sco"] = s;
                }

                //Lo sconto va sull'importo

                decimal dImp = 0;
                if ((string)rowMov["mov_umi"] == "KG")
                    dImp = decVal * (decimal)rowMov["mov_qkg"];
                else
                    dImp = decVal * (decimal)rowMov["mov_qta"];

                string[] a = s.Split('+');
                foreach (string v in a)
                {
                    if (v.Trim().Length > 0)
                    {
                        decimal sc = 0;
                        string sTip = "%";
                        if (v.Substring(0, 1) == "%")
                            s = v.Substring(1);
                        else if (v.Substring(0, 1) == "V")
                        {
                            sTip = "V";
                            s = v.Substring(1);
                        }
                        if (_clsFun.Numerico(s))
                        {
                            decimal dd = Convert.ToDecimal(s.Replace(".", ","));

                            if (sTip == "%")
                            {
                                dd = dImp - _clsFun.MenoPer(dImp, dd);
                            }

                            dSco = dSco + dd;
                        }
                    }
                }
            }

            return dSco;
        }

        private void FillMovRow(DataTable tabTmp, DataRow x, Boolean bolBil)
        {
            string s = "";

            if (DBNull.Value.Equals(tabTmp.Rows[0]["tmp_iva"]) || ((string)tabTmp.Rows[0]["tmp_iva"]).Trim() == "")
                tabTmp.Rows[0]["tmp_iva"] = _strIvaPar;

            x["mov_art"] = tabTmp.Rows[0]["tmp_art"];
            x["mov_ard"] = tabTmp.Rows[0]["tmp_ard"];
            x["mov_iva"] = tabTmp.Rows[0]["tmp_iva"];
            
            x["mov_umi"] = tabTmp.Rows[0]["tmp_umi"];
            if (((string)x["mov_umi"]).Trim() == "" || ((string)x["mov_umi"]).Trim() != "KG")
                x["mov_umi"] = "NR";

            x["mov_cos"] = tabTmp.Rows[0]["tmp_cos"];
            x["mov_prv"] = tabTmp.Rows[0]["tmp_prv"];

            if ((decimal)x["mov_qta"] == 0)
                x["mov_qta"] = 1;

            if (bolBil)
            {
                if ((decimal)x["mov_imp"] != 0 &&  (decimal)x["mov_prv"] != 0)
                    x["mov_qkg"] = (decimal)x["mov_imp"] / (decimal)x["mov_prv"];
            }
            else if (_strCauTpd == "FA" || _strCfo == "FOR")
            {
                if((string)x["mov_umi"] == "KG")
                    x["mov_imp"] = (decimal)x["mov_cos"] * (decimal)x["mov_qkg"];
                else
                    x["mov_imp"] = (decimal)x["mov_cos"] * (decimal)x["mov_qta"];
            }
            else
            {
                decimal dIva = 22;
                DataRow[] j = ((DataTable)cmbMovIva.DataSource).Select("tab_cod='" + (string)x["mov_iva"] + "'");
                if (j.Length > 0)
                    dIva = Convert.ToDecimal(j[0]["tab_ali"]);
                else
                {
                    s = _clsFun.ParGet(clsDefine.enuParametri.Par012CodIVAxDefault, _strConSql);
                    if (s != "")
                    {
                        j = ((DataTable)cmbMovIva.DataSource).Select("tab_cod='" + s + "'");
                        if (j.Length > 0)
                            dIva = Convert.ToDecimal(j[0]["tab_ali"]);
                    }
                }

                decimal d = (decimal)x["mov_prv"];
                if (DBNull.Value.Equals(tabTmp.Rows[0]["tmp_liv"]) || ((string)tabTmp.Rows[0]["tmp_liv"] == "S" || ((string)tabTmp.Rows[0]["tmp_liv"]).Trim() == ""))
                    d = _clsFun.MenoIva((decimal)x["mov_prv"], dIva);
                //x["mov_imp"] = (decimal)x["mov_prv"] * (decimal)x["mov_qta"];
                x["mov_prv"] = d;
                //x["mov_imp"] = d * (decimal)x["mov_qta"];
                if ((string)x["mov_umi"] == "KG")
                    x["mov_imp"] = d * (decimal)x["mov_qkg"];
                else
                    x["mov_imp"] = d * (decimal)x["mov_qta"];

            }

            x["MovMdy"] = "S";

            CtrlNumRiga(x);

            int i = dgv1.Rows.Count - 1;
            if (i >= 0)
            {
                dgv1.Rows[i].Selected = true;
                dgv1.CurrentCell = dgv1[3, i];
            }
            if (i >= 0 && dgv1.Rows[i].Cells[0].Value.ToString() != "" && dgv1.Rows.Count - 1 == i)
            {
                NewRiga(true);
            }

            dgv2.DataSource = FillIva("", "");
        }

        private void CtrlNumRiga(DataRow rowMov)
        {
            string sFld = "";

            if (_strMovFat == "F")
            {
                sFld = "mov_rfa";
                rowMov["mov_rmo"] = _clsDef.COD04X;
            }
            else 
            {
                sFld = "mov_rmo";
                rowMov["mov_rfa"] = _clsDef.COD04X;
            }
            if ((string)rowMov[sFld] == NRIVUOTA && (((string)rowMov["mov_art"]).Trim() != "" || Convert.ToString(rowMov["mov_ard"]).Trim() != ""))
            {
                DataTable t = ((DataView)dgv1.DataSource).ToTable();
                DataView v = new DataView(t, sFld + "<>'" + NRIVUOTA + "'", sFld + " DESC", DataViewRowState.CurrentRows);
                int i = 0;

                if (v.Count > 0)
                    i = Convert.ToInt32(v[0][sFld]);
                rowMov[sFld] = (i + 1).ToString("0000");
            }
        }

        //private void txtMovArd_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        //{

        //}
        //private void dgv1_CurrentCellChanged(object sender, EventArgs e)
        //{
        //    Console.WriteLine("aaaaaaa");
        //}

        //private void dgv1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        //{
        //    if(KeyDown == Keys.Return)
        //        Console.WriteLine("aaaaaaa");
        //}

        //private void dgv1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        //{
        //    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);

        //}
        
        //private void Control_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    //if ((strings.Asc(e.KeyChar) >= Strings.Asc(Keys.A.ToString()) && Strings.Asc(e.KeyChar) >= Strings.Asc(Keys.Z.ToString())) || (Strings.Asc(e.KeyChar) >= Strings.Asc(Keys.D0.ToString()) && Strings.Asc(e.KeyChar) >= Strings.Asc(Keys.D9.ToString()) || (Strings.Asc(e.KeyChar) >= 97 && Strings.Asc(e.KeyChar) > 122)))
        //    //{
        //    //    ------
        //    //}

        //    if(e.KeyChar.Equals( Keys.Return))
        //        Console.WriteLine("aaaaaaa");
        //}

        private Boolean DocCtrl()
        {
            string sMsg = "";
            int i = 0;

            if (((DataView)dgv1.DataSource).Count > 0)
            {
                string sFld = "mov_rmo";
                if (_strMovFat == "F")
                    sFld = "mov_rfa";

                if (_strCfo != "" && (cmbMftCfo.SelectedValue == null || cmbMftCfo.SelectedValue.ToString() == ""))
                    sMsg = "Destinatario non definito";

                if (sMsg != "")
                    MessageBox.Show(sMsg, "DATI MANCANTI");
                else
                {
                    DataTable t = ((DataView)dgv1.DataSource).ToTable();
                    if (t.Rows.Count == 0)
                        sMsg = "Non ci sono righe inserite!";
                    else
                    {
                        foreach (DataRow y in t.Rows)
                        {
                            if ((string)y["MovMdy"] == "S")
                            {
                                if (!DBNull.Value.Equals(y["mov_imp"]) && Convert.ToDecimal(y["mov_imp"]) > 0 && ((string)y["mov_iva"]).Trim() == "" && (string)y["mov_art"] != "")
                                    sMsg += "Articolo " + (string)y["mov_ard"] + " valorizzato con IVA mancante! \\r\\n";
                            }
                            if ((string)y[sFld] != "")
                                i++;
                        }

                        if(i==0)
                            sMsg = "Non ci sono righe inserite!";

                    }
                }

                if (sMsg != "")
                    MessageBox.Show(sMsg, "DATI MANCANTI");
            }
            else
                sMsg = "Non ci sono righe inserite!";

            return sMsg == "";
        }

        private Boolean Salva()
        {
            Boolean b = DocCtrl();
            string s = "";
            DataRow[] j;
            Boolean bUpd = false;

            if (b)
            {
                string sPar016PathDivNegozi = _clsFun.ParGet(clsDefine.enuParametri.Par016PathDivNegozi, _strConSql);

                DataTable tLia = new DataTable();

                DataTable t = new DataTable();
                ArrayList aWhe = new ArrayList();
                ArrayList aExl = new ArrayList();

                if (_strMovFat == "F" && _strCauTpd == "FA")
                {
                    s = "SELECT * FROM GesLisAcquisto ";
                    //s += "WHERE lia_for='" + cmbMftCfo.SelectedValue.ToString() + "' AND lia_tip='F' ";
                    s += "WHERE lia_for='" + cmbMftCfo.SelectedValue.ToString() + "' ";
                    s += "ORDER BY lia_arf, lia_for, lia_dti DESC";
                    tLia = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);
                    DataColumn[] keys = new DataColumn[5];
                    keys[0] = tLia.Columns["lia_art"];
                    keys[1] = tLia.Columns["lia_tip"];
                    keys[2] = tLia.Columns["lia_arf"];
                    keys[3] = tLia.Columns["lia_dti"];
                    keys[4] = tLia.Columns["lia_dtf"];
                    tLia.PrimaryKey = keys;

                    //tLia = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);   ?????????????
                }

                if (lblMftNum.Text == _clsDef.CODNEW)
                {
                    if (_strMovFat == "F")
                        lblMftNum.Text = _clsFun.NewNum(_strMftYea, clsDefine.enuNumeratori.NumGesMovFatture, 6, _strConSql);
                    else
                        lblMftNum.Text = _clsFun.NewNum(_strMftYea, clsDefine.enuNumeratori.NumGesMovimenti, 6, _strConSql);
                }

                if (_strMovFat == "F")
                {
                    s = "SELECT * FROM GesFatTestate WHERE fat_yfa='" + lblMftYea.Text + "' AND  fat_nfa='" + lblMftNum.Text + "' AND fat_ubi='" + _strMftUbi + "'";
                    t = _clsFun.FillTabSql(TABGESFAT, s, false, _strConSql);

                    DataRow y = t.NewRow();
                    y["fat_yfa"] = lblMftYea.Text;
                    y["fat_tdo"] = cmbMftTdc.SelectedValue.ToString();
                    y["fat_nfa"] = lblMftNum.Text;
                    y["fat_tpd"] = _strCauTpd;
                    y["fat_ndo"] = txtMftNdo.Text;
                    y["fat_ddo"] = dtpMftDdt.Value;
                    y["fat_cfo"] = cmbMftCfo.SelectedValue.ToString();
                    y["fat_no1"] = txtMftNo1.Text;
                    y["fat_ubi"] = _strMftUbi;

                    y["fat_tpg"] = "";
                    if (cmbMftTpg.SelectedValue != null)
                        y["fat_tpg"] = cmbMftTpg.SelectedValue.ToString();

                    y["fat_neg"] = "";
                    if (cmbMftNeg.SelectedIndex >= 0)
                    {
                        y["fat_neg"] = cmbMftNeg.SelectedValue.ToString();

                        j = ((DataTable)cmbMftNeg.DataSource).Select("tab_cod='" + cmbMftNeg.SelectedValue.ToString() + "'");
                        if (j.Length > 0 && ((string)j[0]["tab_mag"]).Trim() != "")
                        {
                            y["fat_mag"] = (string)j[0]["tab_mag"];
                        }
                    }

                    y["fat_ann"] = chkMftAnn.Checked;
                    //y["fat_pvi"] = chkMftPvi.Checked;
                    if (cmbMftSta.SelectedValue != null)
                        y["fat_sta"] = cmbMftSta.SelectedValue.ToString();
                    y["fat_des"] = "";
                    if (cmbMftDes.SelectedValue != null)
                        y["fat_des"] = cmbMftDes.SelectedValue.ToString();
                    
                    y["fat_ele"] = txtFatEle.Text;

                    //y["fat_arr"] = FatArrotondamenti();

                    if (t.Rows.Count == 0)
                        s = _clsFun.SqlInsertRow(TABGESFAT, t, y);
                    else
                    {
                        aWhe = new ArrayList();
                        aExl = new ArrayList();
                        aWhe.Add("fat_yfa");
                        aWhe.Add("fat_nfa");
                        aWhe.Add("fat_tpd");
                        aWhe.Add("fat_ubi");
                        s = _clsFun.SqlUpdRow(TABGESFAT, t, t.Rows[0], y, aWhe, aExl);
                    }
                    if (s != "")
                    {
                        _clsFun.SqlWrite(s, _strConSql);
                        _clsFun.FileLog("GesFatTestate", lblMftYea.Text + "-" + lblMftNum.Text, s);
                        bUpd = true;
                    }

                    if (t.Rows.Count > 0)
                    {
                        //X doc creati nell'anno precedente e stampati nell'anno in corso

                        if (txtMftNdo.Text != "" && lblMftYea.Text != dtpMftDdt.Value.Year.ToString())
                        {
                            s = "UPDATE GesFatTestate SET fat_yfa='" + dtpMftDdt.Value.Year.ToString() + "' WHERE ";
                            s += "fat_idx=" + Convert.ToString(t.Rows[0]["fat_idx"]);
                            _clsFun.SqlWrite(s, _strConSql);
                            bUpd = true;
                        }
                    }
                }
                else
                {
                    s = "SELECT * FROM GesMovTestate WHERE mot_ymo='" + lblMftYea.Text + "' AND  mot_nmo='" + lblMftNum.Text + "'";
                    t = _clsFun.FillTabSql(TABGESMOT, s, false, _strConSql);

                    DataRow y = t.NewRow();
                    y["mot_ymo"] = lblMftYea.Text;
                    y["mot_nmo"] = lblMftNum.Text;
                    y["mot_day"] = DateTime.Today;
                    y["mot_ubi"] = _strMftUbi;

                    y["mot_cfo"] = "";
                    if(_strCfo != "")
                        y["mot_cfo"] = cmbMftCfo.SelectedValue.ToString();

                    y["mot_cau"] = _strCauTpd;
                    y["mot_ndo"] = txtMftNdo.Text;
                    y["mot_ddo"] = dtpMftDdt.Value;
                    y["mot_neg"] = cmbMftNeg.SelectedValue.ToString();
                    y["mot_no1"] = txtMftNo1.Text;
                    y["mot_ann"] = chkMftAnn.Checked;
                    //y["mot_pvi"] = chkMftPvi.Checked;
                    y["mot_sta"] = cmbMftSta.SelectedValue.ToString();
                    //y["mot_des"] = cmbMftDes.SelectedValue.ToString();
                    
                    y["mot_tpg"] = "";
                    if (cmbMftTpg.SelectedValue != null)
                        y["mot_tpg"] = cmbMftTpg.SelectedValue.ToString();


                     y["mot_mag"] = "01";
                    j = ((DataTable)cmbMftTdc.DataSource).Select("tab_cod='" + cmbMftTdc.SelectedValue.ToString() +  "'");
                    if (j.Length > 0 && ((string)j[0]["tab_mag"]).Trim() != "")
                    {
                        y["mot_mag"] = (string)j[0]["tab_mag"];
                        if (cmbMftNeg.SelectedValue != null)
                            y["mot_neg"] = cmbMftNeg.SelectedValue.ToString();
                    }
                    y["mot_des"] = "";
                    if (cmbMftDes.SelectedValue != null)
                        y["mot_des"] = cmbMftDes.SelectedValue.ToString();

                    if (t.Rows.Count == 0)
                    {
                        y["mot_yfa"] = _clsDef.COD04X;
                        y["mot_nfa"] = _clsDef.COD06X;

                        s = _clsFun.SqlInsertRow(TABGESMOT, t, y);
                    }
                    else
                    {
                        y["mot_yfa"] = t.Rows[0]["mot_yfa"];
                        y["mot_nfa"] = t.Rows[0]["mot_nfa"];

                        aWhe = new ArrayList();
                        aExl = new ArrayList();
                        aWhe.Add("mot_ymo");
                        aWhe.Add("mot_nmo");
                        aWhe.Add("mot_ubi");
                        s = _clsFun.SqlUpdRow(TABGESMOT, t, t.Rows[0], y, aWhe, aExl);
                    }
                    if (s != "")
                    {
                        _clsFun.SqlWrite(s, _strConSql);
                        _clsFun.FileLog(TABGESMOT, lblMftYea.Text + "-" + lblMftNum.Text, s);
                        bUpd = true;
                    }
                }

                s = "SELECT * FROM GesMovimenti WHERE ";
                if (_strMovFat == "F")
                {
                    //s += "mov_ubi = '" + _strMftUbi + "' AND ";
                    s += "mov_yfa = '" + lblMftYea.Text + "' AND ";
                    s += "mov_nfa = '" + lblMftNum.Text + "' ";
                    s += "ORDER BY mov_rfa";
                }
                else
                {
                    //s += "mov_ubi = '" + _strMftUbi + "' AND ";
                    s += "mov_ymo = '" + lblMftYea.Text + "' AND ";
                    s += "mov_nmo = '" + lblMftNum.Text + "' ";
                    s += "ORDER BY mov_rmo";
                }
                t = _clsFun.FillTabSql(TABGESMOV, s, false, _strConSql);

                aExl = new ArrayList();

                DataTable tTmp = ((DataView)dgv1.DataSource).Table;

                foreach (DataRow x in (((DataView)dgv1.DataSource).Table).Rows)
                {
                    if ((string)x["MovMdy"] == "S")
                    {
                        b = true;

                        aWhe = new ArrayList();
                        if (_strMovFat == "F")
                        {
                            aWhe.Add("mov_yfa");
                            aWhe.Add("mov_nfa");
                            aWhe.Add("mov_rfa");
                            aWhe.Add("mov_ubi");
                        }
                        else
                        {
                            aWhe.Add("mov_ymo");
                            aWhe.Add("mov_nmo");
                            aWhe.Add("mov_rmo");
                            aWhe.Add("mov_ubi");
                        }

                        if (_strMovFat == "F")
                            s = "mov_rfa='" + x["mov_rfa"] + "'";
                        else
                            s = "mov_rmo='" + x["mov_rmo"] + "'";

                        if ((string)x["mov_art"] == "0000007")
                            Console.WriteLine("xxxxxxx");

                        Boolean bIns = false;
                        j = t.Select(s);
                        if (j.Length > 0)
                            s = _clsFun.SqlUpdRow(TABGESMOV, t, j[0], x, aWhe, aExl);
                        else
                        {
                            if (_strMovFat == "F")
                            {
                                x["mov_yfa"] = lblMftYea.Text;
                                x["mov_nfa"] = lblMftNum.Text;
                                x["mov_ymo"] = _clsDef.COD04X;
                                x["mov_nmo"] = _clsDef.COD06X;
                                x["mov_ubi"] = _strMftUbi;
                            }
                            else
                            {
                                x["mov_yfa"] = _clsDef.COD04X;
                                x["mov_nfa"] = _clsDef.COD06X;
                                x["mov_ymo"] = lblMftYea.Text;
                                x["mov_nmo"] = lblMftNum.Text;
                                x["mov_ubi"] = _strMftUbi;
                            }

                            s = _clsFun.SqlInsertRow(TABGESMOV, t, x);
                            bIns = true;
                        }
                        if (s != "")
                        {
                            _clsFun.SqlWrite(s, _strConSql);
                            _clsFun.FileLog("GesMovimenti", (string)x["mov_art"], s); 
                            bUpd = true;
                        }

                        //20200129 Controllo costo in anagrafica solo in caso inserimento e se data documento > data ultimo inserimento
                        if (_strMovFat == "F" && _strCauTpd == "FA" && bIns)    
                        {
                            aWhe = new ArrayList();
                            aWhe.Add("lia_tip");
                            aWhe.Add("lia_art");
                            aWhe.Add("lia_for");
                            aExl = new ArrayList();

                            //j = tLia.Select("lia_tip='F' AND lia_art='" + x["mov_art"] + "'", "lia_dti DESC");
                            j = tLia.Select("lia_art='" + x["mov_art"] + "'", "lia_dti DESC");

                            DateTime dDti = DateTime.Today;
                            if (j.Length > 0)
                                dDti = (DateTime)j[0]["lia_dti"];

                            Console.WriteLine("zzzzzzzz");

                            //if (dtpMftDdt.Value.ToString("yyyyMMdd") > ((DateTime)j[0]["lia_dti"]).ToString("yyyyMMdd"))
                            if (_clsFun.DayDiff(dtpMftDdt.Value, dDti) > 0)
                            {
                                DataRow x1 = tLia.NewRow();
                                x1["lia_tip"] = "F";
                                x1["lia_art"] = x["mov_art"];
                                x1["lia_for"] = cmbMftCfo.SelectedValue.ToString();
                                x1["lia_dti"] = dtpMftDdt.Value;
                                x1["lia_dtf"] = _clsDef.DAYOUT;
                                x1["lia_arf"] = x["mov_art"];
                                x1["lia_cos"] = x["mov_cos"];
                                x1["lia_pxc"] = 1;
                                x1["lia_cxp"] = 1;
                                x1["lia_day"] = DateTime.Today;
                                x1["lia_prv"] = x["mov_prv"];

                                if ((decimal)x["mov_cos"] > 0)
                                {
                                    s = "";
                                    if (j.Length > 0)
                                    {
                                        if (!DBNull.Value.Equals(j[0]["lia_arf"]) && ((string)j[0]["lia_arf"]).Trim() != "")
                                            x1["lia_arf"] = (string)j[0]["lia_arf"];

                                        if ((decimal)j[0]["lia_cos"] != (decimal)x["mov_cos"])
                                        {
                                            if (DateTime.Compare((DateTime)x1["lia_dti"], (DateTime)j[0]["lia_dti"]) == 0)
                                                s = _clsFun.SqlUpdRow(TABLISACQ, tLia, j[0], x1, aWhe, aExl);
                                            else
                                                s = _clsFun.SqlInsertRow(TABLISACQ, tLia, x1);
                                        }
                                    }
                                    else
                                        s = _clsFun.SqlInsertRow(TABLISACQ, tLia, x1);

                                    if (s != "")
                                    {
                                        _clsFun.SqlWrite(s, _strConSql);
                                    }
                                }
                            }
                        }
                        if (((string)x["mov_art"]).Trim() != "")
                        {
                            s = "UPDATE AnaArticoli SET art_sta='A' WHERE art_cod='" + (string)x["mov_art"] + "'";
                            _clsFun.SqlWrite(s, _strConSql);

                            //new clsVariazioni().Variazioni((string)x["mov_art"], "Forza", "Doc acquisto fornitori", "FA");
                            //_clsFun.FileLog("AnaArticoli", (string)x["mov_art"], "art_cod");

                            if (sPar016PathDivNegozi != "")
                            {
                                ArrayList a = new ArrayList();
                                a.Add((string)x["mov_art"]);

                                new clsVariazioni().DivNegArticoli(sPar016PathDivNegozi, a, null, cmbMftNeg.SelectedValue.ToString(), "");
                            }
                        }
                    }
                }

                if (t.Rows.Count > 0)
                {
                    //X doc creati nell'anno precedente e stampati nell'anno in corso

                    if (_strMovFat == "F" && txtMftNdo.Text != "" && lblMftYea.Text != dtpMftDdt.Value.Year.ToString())
                    {
                        foreach (DataRow y in t.Rows)
                        {
                            s = "UPDATE GesMovimenti SET mov_yfa='" + dtpMftDdt.Value.Year.ToString() + "' WHERE ";
                            s += "mov_idx=" + Convert.ToString(y["mov_idx"]);
                            _clsFun.SqlWrite(s, _strConSql);
                            bUpd = true;
                        }
                    }
                }
            }

            if(bUpd)
            { 

                s = _clsFun.ParGet(clsDefine.enuParametri.Par036Div2Sede, _strConSql);

                string[] a = s.Split(',');

                if (b && a.Length > 1 && a[1].Substring(0, 1) == "S")          //Divulgazione documento
                {
                    //s = a[1];

                    //Console.WriteLine("zzzz");

                    //a = s.Split('-');

                    //Console.WriteLine("xxxx");

                    //foreach (string ss in a)
                    //{
                    //    if (ss.Contains('+'))
                    //    {
                    //        string[] aa = ss.Split('+');
                    //        string sTip = aa[0];
                    //        string sTpd = aa[1];
                    //        if (sTip == _strMovFat && sTpd == _strCauTpd)
                    //        {
                    //            s = _strMovFat + ";" + lblMftYea.Text + ";" + lblMftNum.Text;
                    //            new clsVariazioni().DivNegDocumenti(s);
                    //        }
                    //    }
                    //}

                    s = _strMovFat + ";" + lblMftYea.Text + ";" + lblMftNum.Text;
                    new clsVariazioni().DivNegDocumenti(s);
                }
            }

            return b;
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            DataRow[] j;
            string s = "";

            if ((_strCauTpd == "FV" && cmbMftCfo.SelectedValue != null && cmbMftCfo.SelectedValue.ToString() != "") && (cmbMftTpg.SelectedValue == null || cmbMftTpg.SelectedValue.ToString() == ""))
                MessageBox.Show("Tipo pagamento non definito!", "CONTROLLO STAMPA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else if (dgv1.DataSource == null || ((DataView)dgv1.DataSource).Count == 0)
                MessageBox.Show("Non ci sono dati da stampare!", "CONTROLLO STAMPA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else if (_strMftYea != dtpMftDdt.Value.Year.ToString())
                MessageBox.Show("Anno selezionato e data documento diversi!", "CONTROLLO STAMPA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                if(DateTime.Compare(dtpMftDdt.Value, DateTime.Today) != 0)
                    MessageBox.Show("Data del documento diversa dalla data odierna!", "CONTROLLO STAMPA", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (txtMftNdo.Text.Trim() == "")
                {
                    //if (_strMovFat == "F" && _strCauTpd.Substring(0, 1) == "F" && cmbMftTdc.SelectedValue.ToString() != "PR" && cmbMftTdc.SelectedValue.ToString() != "FP") //PR = Preventivo, FP = Fattura proforma
                    if (_strMovFat == "F")
                    {
                        string sTpd = cmbMftTdc.SelectedValue.ToString();

                        if (sTpd == "FA")
                            txtMftNdo.Text = _clsFun.NewNum(_strMftYea, clsDefine.enuNumeratori.NumGesDocFatture, 10, _strConSql);
                        else
                        {
                            j = ((DataTable)cmbMftTdc.DataSource).Select("tab_cod='" + sTpd + "'");
                            if (j.Length > 0 && !DBNull.Value.Equals(j[0]["tab_num"]) && ((string)j[0]["tab_num"]).Trim() != "")
                            {
                                if((string)j[0]["tab_num"] == "007")
                                    txtMftNdo.Text = _clsFun.NewNum(_strMftYea, clsDefine.enuNumeratori.NumGesDocFatture, 10, _strConSql);
                            }
                        }

                        //txtMftNdo.Text = _clsFun.NewNum(_strMftYea, clsDefine.enuNumeratori.NumGesDocFatture, 10, _strConSql);
                    }
                    if (_strMovFat == "M")
                    {
                        j = ((DataTable)cmbMftTdc.DataSource).Select("tab_cod='" + cmbMftTdc.SelectedValue.ToString() + "'");
                        if (j.Length > 0 && ((string)j[0]["tab_num"]).Trim() != "")
                        {
                            if ((string)j[0]["tab_num"] == "013")
                                txtMftNdo.Text = _clsFun.NewNum(_strMftYea, clsDefine.enuNumeratori.NumGesDocCeliaci, 10, _strConSql);
                            else if ((string)j[0]["tab_num"] == "015")
                                txtMftNdo.Text = _clsFun.NewNum(_strMftYea, clsDefine.enuNumeratori.NumGesDocRiparazioni, 10, _strConSql);
                            else if ((string)j[0]["tab_num"] == "016")
                                txtMftNdo.Text = _clsFun.NewNum(_strMftYea, clsDefine.enuNumeratori.NumGesMovInterni, 10, _strConSql);
                            else
                                MessageBox.Show("Numeratore non configurato (" + (string)j[0]["tab_num"] + ")!", "CONTROLLO NUMERATORE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                            txtMftNdo.Text = _clsFun.NewNum(_strMftYea, clsDefine.enuNumeratori.NumGesDocDdt, 10, _strConSql);
                    }
                }

                if (Salva())
                {
                    string sCfo = "";
                    string sTip = "";
                    string sMsg = "";

                    s = "";
                    if(_strMovFat == "M")
                    {
                        j = ((DataTable)cmbMftTdc.DataSource).Select("tab_cod='" + cmbMftTdc.SelectedValue + "'");
                        if (j.Length > 0)
                        {
                            sCfo = (string)j[0]["tab_cfo"];
                            s = (string)j[0]["tab_tip"];

                            string[] a = s.Split('-');
                            if(a.Length > 0)
                                sTip = a[0];
                        }
                    }

                    if(_strCfo == "" && sCfo != "")
                        MessageBox.Show("Tipo cliente/fornitore non definito!");
                    else
                    {
                        if (cmbMftDes.DataSource != null && ((DataTable)cmbMftDes.DataSource).Rows.Count > 1 && cmbMftDes.Text == "")
                            MessageBox.Show("Destinazione PV non selezionata!", "Stampa documento", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        if (_strMovFat == "F")
                            s = "FV";
                        else if (_strMovFat == "M")
                            s = "DT";
                        else
                            s = "XX";
                        s = "SELECT * FROM TabNote WHERE tab_cod='" + s + "' AND tab_ann=0 ORDER BY tab_row";
                        DataTable tNot = _clsFun.FillTabSql("TabNote", s, false, _strConSql);

                        s = "";
                        if (cmbMftCfo.SelectedValue != null)
                            s = cmbMftCfo.SelectedValue.ToString();

                        DataTable tTes = _clsQry.DocCfo(_strCfo, s);
                        tTes.Rows[0]["tmp_ndo"] = txtMftNdo.Text;
                        tTes.Rows[0]["tmp_ddo"] = dtpMftDdt.Value;
                        tTes.Rows[0]["tmp_tpg"] = cmbMftTpg.Text;

                        s = "Totali";
                        decimal dIva = 0;
                        decimal dImp = 0;
                        DataTable tIva = (DataTable)dgv2.DataSource;
                        foreach (DataRow y in tIva.Rows)
                        {
                            if ((string)y["iva_des"] != s)
                            {
                                dIva += Math.Round((decimal)y["iva_iva"], 2, MidpointRounding.AwayFromZero);
                                dImp += Math.Round((decimal)y["iva_imp"], 2, MidpointRounding.AwayFromZero);
                            }
                        }

                        j = tIva.Select("iva_des='" + s + "'");
                        if (j.Length > 0)
                            j[0].Delete();

                        DataRow x = tIva.NewRow();
                        x["iva_des"] = s;
                        x["iva_iva"] = dIva;
                        x["iva_imp"] = dImp;
                        x["iva_tot"] = dImp + dIva;
                        tIva.Rows.Add(x);

                        s = "";

                        DataTable tMov = ((DataView)dgv1.DataSource).ToTable();
                        foreach(DataRow y in tMov.Rows)
                        {
                            if (((string)y["mov_art"]).Trim() != "" && ((decimal)y["mov_prv"] == 0 || (decimal)y["mov_imp"] == 0))
                            {
                                s = "ERR";
                                break;
                            }
                        }

                        if(s != "")
                            MessageBox.Show("Ci sono righe con importo a ZERO!", "Stampa documento", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        DataSet dasGen = new DataSet();
                        dasGen.Tables.Add(tTes);
                        //dasGen.Tables.Add(((DataView)dgv1.DataSource).ToTable());
                        dasGen.Tables.Add(tMov);
                        dasGen.Tables.Add(((DataTable)dgv2.DataSource).Copy());
                        dasGen.Tables.Add(tNot);

                        frmGesDocPrint f = new frmGesDocPrint();
                        f._dasGen = dasGen;
                        f._strMovFat = _strMovFat;
                        f._strCauTpd = _strCauTpd;
                        f._strSuf = _strDocSuf;
                        f._strTipCod = cmbMftTdc.Text;
                        f._strDocTip = sTip;
                        f._strQta = lblTotQta.Text;
                        f._strDes = cmbMftDes.Text;
                        f._strCfo = _strCfo;
                        //f._bolMftPvi = chkMftPvi.Checked;
                        f.ShowDialog();

                        dasGen.Dispose();
                    }
                }
            }
        }

        private void txtMov_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                MovRowAgg();
        }

        private void cmbMov_SelectionChangeCommitted(object sender, EventArgs e)
        {
            MovRowAgg();
        }

        private void chkMovAnn_CheckedChanged(object sender, EventArgs e)
        {
            MovRowAgg();
        }

        private void letturaDaScontrinoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (cmbMftCfo.SelectedValue.ToString() == "")
                MessageBox.Show("Destinatario non definito!");
            else
            {
                frmGesDocScontrino f = new frmGesDocScontrino();
                if (cmbMftNeg.SelectedValue != null)
                    f._strNeg = cmbMftNeg.SelectedValue.ToString();
                f.ShowDialog();
                if (f._tabSco != null && f._tabSco.Rows.Count > 0)
                {
                    FillFromScontrino(f._tabSco, f._tabVep);
                }
            }
        }

        private void FillFromScontrino(DataTable tabSco, DataTable tabVep)
        {
            string s = "";

            string sMat = "";
            s = "SELECT * FROM TabPos WHERE tab_ann=0";
            DataTable t = _clsFun.FillTabSql("TabPos", s, false, _strConSql);
            if (t.Rows.Count > 0 && !DBNull.Value.Equals(t.Rows[0]["tab_mat"]))
            {
                int iPoss = Convert.ToInt16(tabSco.Rows[0]["vet_pos"]);

                s = ((string)t.Rows[0]["tab_mat"]).Trim();

                string[] a = s.Split(',');
                if (a.Length >= iPoss && a[iPoss - 1] != "")
                    sMat = a[iPoss - 1];
            }

            decimal dTotoDoc = Convert.ToDecimal(lblTotTot.Text);

            DataTable tTabIva = (DataTable)cmbMovIva.DataSource;

            t = ((DataView)dgv1.DataSource).Table;

            DataRow x;
            DataRow[] j;
            int iRig = 0;
            int iPos = 0;
            decimal d = 0;
            int iRowLast = 0;
            string sIva = "";

            //if (t.Rows.Count > 0)
            if (t.Rows.Count > 0)
            {
                if (((string)t.Rows[t.Rows.Count - 1]["mov_nmo"]).Trim() != "" || ((string)t.Rows[t.Rows.Count - 1]["mov_nfa"]).Trim() != "")
                    NewRiga(false);
            }
            else
                NewRiga(false);

            iPos = t.Rows.Count-1;
            iRig = t.Rows.Count;
            if (_strMovFat == "F")
                t.Rows[iPos]["mov_rfa"] = iRig.ToString("0000");
            else
                t.Rows[iPos]["mov_rmo"] = iRig.ToString("0000");

            s = "Scontrino";
            if ((string)tabSco.Rows[0]["vet_cau"] != "MOV")
                s = "Documento gestionale";
            else if ((string)tabSco.Rows[0]["vet_cau"] != "SCO")
                s = "Proforma";

            s += " N. " + (string)tabSco.Rows[0]["vet_sco"] + " del " + ((DateTime)tabSco.Rows[0]["vet_day"]).ToShortDateString();
            if (sMat != "")
                s += " Matr. " + sMat;

            t.Rows[iPos]["mov_ard"] = s;

            t.Rows[iPos]["MovMdy"] = "S";
            iRig = iPos+1;

            string sRigIni = "";
            string sRigFin = "";
            decimal dTot = 0;

            string sFldRig = "mov_rmo";
            if (_strMovFat == "F")
                sFldRig = "mov_rfa";

            decimal dDelta = 0;
            decimal dScoTot = 0;
            decimal dSconto = 0;

            if (tabVep != null && tabVep.Rows.Count > 0)
            {
                foreach (DataRow y in tabVep.Rows)
                {
                    if (((string)y["vep_tip"]).Substring(0, 2) == "SC")
                        dSconto += (decimal)y["vep_imp"];
                    else if (((string)y["vep_tip"]) == "PAG")
                        dScoTot += (decimal)y["vep_imp"];
                    else if (((string)y["vep_tip"]) == "RES")
                        dScoTot -= (decimal)y["vep_imp"];
                }

                if (dSconto > 0)
                    dDelta = dSconto * 100 / (dScoTot + dSconto);
            }

            foreach(DataRow y in tabSco.Rows)
            {
                iRowLast = iRig;
                iRig ++;

                sIva = _strIvaPar;

                if ((string)y["ven_art"] == "0008142")
                    Console.Write("zzzz");


                j = tTabIva.Select("tab_cod='" + y["ven_iva"] + "'");
                if (j.Length > 0)
                    sIva = (string)j[0]["tab_cod"];


                if (!DBNull.Value.Equals(y["ven_sct"]) && (string)y["ven_sct"] == "RES")
                    dTot -= (decimal)y["ven_ven"];
                else
                    dTot += (decimal)y["ven_ven"];


                x = t.NewRow();
                x["mov_yfa"] = "";
                x["mov_nfa"] = "";
                x["mov_ymo"] = "";
                x["mov_nmo"] = "";
                x["mov_rfa"] = ""; 
                x["mov_rmo"] = "";

                x[sFldRig] = iRig.ToString("0000");

                if (((string)y["ven_art"]).Length == 7)
                    x["mov_art"] = y["ven_art"];
                else
                    x["mov_art"] = "";

                x["mov_ard"] = y["ven_ard"];
                x["mov_iva"] = sIva;
                x["mov_umi"] = y["ven_umi"];
                x["mov_qta"] = y["ven_qta"];
                x["mov_qkg"] = y["ven_qkg"];
                x["mov_cos"] = 0;
                x["mov_ann"] = false;
                x["mov_day"] = DateTime.Today;

                if (((string)y["ven_art"]).Length != 7)
                    x["mov_no1"] = y["ven_art"]; 


                x["mov_prv"] = y["ven_prz"];
                if (!DBNull.Value.Equals(y["ven_sct"]) && (string)y["ven_sct"] == "RES")
                    x["mov_imp"] = (decimal)y["ven_ven"] * -1;
                else
                    x["mov_imp"] = y["ven_ven"];

                if ((decimal)x["mov_imp"] < 0)
                    Console.WriteLine("aaaa");

                if ((string)x["mov_art"] == "0156195")
                    Console.WriteLine("aaaa");

                if (dDelta > 0)
                {
                    d = (decimal)x["mov_imp"];
                    d = _clsFun.MenoPer(d, dDelta);
                    d = Math.Round(d, 3, MidpointRounding.AwayFromZero);
                    s = "V" + ((decimal)x["mov_imp"] - d).ToString();
                    if (s.Length > 15)
                        s = s.Substring(0, 15);
                    x["mov_sco"] = s;
                    x["mov_imp"] = d;
                }

                s = (string)y["ven_cau"] + ",";
                s += ((DateTime)y["ven_day"]).ToString("yyMMdd") + ",";
                s += (string)y["ven_ora"] + ",";
                s += (string)y["ven_pos"] + ",";
                s += (string)y["ven_sco"] + ",";
                s += dScoTot.ToString().Replace(",", ".");

                x["mov_ori"] = s;
                x["MovMdy"] = "S";

                decimal dIva = 22;
                j = ((DataTable)cmbMovIva.DataSource).Select("tab_cod='" + (string)x["mov_iva"] + "'");
                if (j.Length > 0)
                    dIva = Convert.ToDecimal(j[0]["tab_ali"]);

                d = _clsFun.MenoIva((decimal)x["mov_imp"], dIva);

                if ((string)x["mov_umi"] == "KG")
                {
                    if (d != 0 && (decimal)x["mov_qkg"] != 0)
                        x["mov_prv"] = d / (decimal)x["mov_qkg"];
                }
                else
                {
                    if (d != 0 && (decimal)x["mov_qta"] != 0)
                        x["mov_prv"] = d / (decimal)x["mov_qta"];
                }

                x["mov_prv"] = Math.Round((decimal)x["mov_prv"], 3, MidpointRounding.ToEven);

                x["mov_imp"] = Math.Round(d, 3);

                t.Rows.Add(x);

                if (sRigIni == "")
                    sRigIni = iRig.ToString("000");
            }

            sRigFin = iRig.ToString("000");

            DataView v = new DataView(t, "", "", DataViewRowState.CurrentRows);

            dgv1.DataSource = v;

            NewRiga(false);

            //Tentativo di arrotondamento per far corrispondere l'IVA

            if (true)
            {
                string sArr = "";

                dTotoDoc = dTotoDoc + dScoTot;

                DataTable tIva = FillIva("", "");

                decimal dTotIva = 0;
                foreach (DataRow y in tIva.Rows)
                    dTotIva += (decimal)y["iva_tot"];

                dTotIva = Math.Round(dTotIva, 2);

                Console.WriteLine("aaaa");

                if (dTotoDoc != dTotIva)
                {
                    decimal dDif = dTotoDoc - dTotIva;

                    t = (DataTable)dgv2.DataSource;

                    int iRow = 0;
                    int iRows = t.Rows.Count;
                    int i = 0;

                    while(true)
                    {
                        i++;
                        decimal dd = Convert.ToDecimal(0.01);
                        if (dDif < 0)
                            dd = dd * -1;

                        tIva.Rows[iRow]["iva_imp"] = (decimal)tIva.Rows[iRow]["iva_imp"] + dd;
                        tIva.Rows[iRow]["iva_arr"] = (decimal)tIva.Rows[iRow]["iva_arr"] + dd;

                        if(dDif > 0)
                            dDif -= Math.Abs(dd);
                        else
                            dDif += Math.Abs(dd);

                        if (dDif == 0)
                            break;

                        iRow++;

                        if (iRow > iRows-1)
                            iRow = 0;

                        if (i > 5)
                            break;
                    }

                    FillIva("", "");
                }
            }
        }

        private void dtpMftDdt_ValueChanged(object sender, EventArgs e)
        {
            //if (_clsFun.Numerico(dtpMftDdt.Value.Year.ToString()))
            //    lblMftYea.Text = dtpMftDdt.Value.Year.ToString();
        }

        private void dtpMftDdt_Validating(object sender, CancelEventArgs e)
        {
            Console.WriteLine("aaaaaaaaa");
        }

        private void btnMftSco_Click(object sender, EventArgs e)
        {
            txtMftSco.Text = txtMftSco.Text.Replace(".",",");
            CalcScontoTot();
        }

        private void documentiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmSeekDocs f = new frmSeekDocs();
            f._strImpDoc = "S";
            f.ShowDialog();
            string s = f._strImpDoc;
            if (s != "S")
            {
                string[] a = s.Split(';');

                string sMovFat = a[0];
                string sCauTpd = a[1];
                string sMftYea = a[2];
                string sMftNum = a[3];

                string sFldNri = "";

                s = "SELECT ";
                s += "mov_yfa, ";
                s += "mov_nfa, ";
                s += "mov_rfa, ";
                s += "mov_ymo, ";
                s += "mov_nmo, ";
                s += "mov_rmo, ";
                s += "mov_art, ";
                //s += "mov_ard, ";
                //s += "AnaArticoli.art_des AS MovArd, ";
                s += "CASE WHEN mov_ard IS NULL OR mov_ard = '' THEN AnaArticoli.art_des ELSE mov_ard END AS mov_ard, ";
                s += "mov_iva, ";
                s += "mov_umi, ";
                s += "mov_qta, ";
                s += "mov_qkg, ";
                s += "mov_cos, ";
                s += "mov_prv, ";
                s += "mov_sco, ";
                s += "mov_imp, ";
                s += "mov_ann, ";
                s += "mov_day, ";
                s += "mov_ori, ";
                s += "mov_no1 ";
                s += "FROM GesMovimenti ";
                s += "LEFT OUTER JOIN AnaArticoli ON GesMovimenti.mov_art = AnaArticoli.art_cod ";
                s += "WHERE ";
                if (sMovFat == "F")
                {
                    s += "mov_yfa='" + sMftYea + "' AND mov_nfa='" + sMftNum + "' ";
                    s += "ORDER BY mov_rfa";
                }
                else
                {
                    s += "mov_ymo='" + sMftYea + "' AND mov_nmo='" + sMftNum + "' ";
                    s += "ORDER BY mov_rmo";
                }

                DataTable tMov = _clsFun.FillTabSql(TABGESMOV, s, false, _strConSql);

                if (tMov.Rows.Count > 0)
                {
                    DataRow x;
                    int iRig = 0;
                    int iPos = 0;
                    decimal d = 0;

                    DataTable t = ((DataView)dgv1.DataSource).Table;

                    if (t.Rows.Count > 0)
                    {
                        if (((string)t.Rows[t.Rows.Count - 1]["mov_nmo"]).Trim() == "" && ((string)t.Rows[t.Rows.Count - 1]["mov_nfa"]).Trim() == "")
                            t.Rows[t.Rows.Count - 1].Delete();
                    }

                    iPos = t.Rows.Count - 1;
                    if (iPos < 0)
                        iPos = 0;
                    if (t.Rows.Count > 0)
                    {
                        iRig = t.Rows.Count;
                        if (sMovFat == "F")
                            t.Rows[iPos]["mov_rfa"] = iRig.ToString("0000");
                        else
                            t.Rows[iPos]["mov_rmo"] = iRig.ToString("0000");
                        t.Rows[iPos]["MovMdy"] = "S";
                        iPos += 1;
                    }
                    iRig = iPos + 1;

                    foreach (DataRow y in tMov.Rows)
                    {
                        if (!(Boolean)y["mov_ann"])
                        {
                            x = t.NewRow();
                            x["mov_yfa"] = "";
                            x["mov_nfa"] = "";
                            x["mov_ymo"] = "";
                            x["mov_nmo"] = "";
                            x["mov_rfa"] = "";
                            x["mov_rmo"] = "";

                            if (_strMovFat == "F")
                                x["mov_rfa"] = iRig.ToString("0000");
                            else
                                x["mov_rmo"] = iRig.ToString("0000");

                            x["mov_art"] = y["mov_art"];
                            x["mov_ard"] = y["mov_ard"];
                            x["mov_iva"] = y["mov_iva"];
                            x["mov_umi"] = y["mov_umi"];
                            x["mov_qta"] = y["mov_qta"];
                            x["mov_qkg"] = y["mov_qkg"];
                            x["mov_cos"] = y["mov_cos"];
                            x["mov_ann"] = false;
                            x["mov_day"] = DateTime.Today;
                            x["mov_no1"] = y["mov_no1"];

                            x["mov_prv"] = y["mov_prv"];
                            x["mov_sco"] = y["mov_sco"];
                            x["mov_imp"] = y["mov_imp"];

                            x["mov_ori"] = "";
                            x["MovMdy"] = "S";

                            t.Rows.Add(x);

                            iRig++;
                        }
                    }

                    if (_strMovFat == "F")
                        sFldNri = "mov_rfa";
                    else
                        sFldNri = "mov_rmo";

                    DataView v = new DataView(t, "", sFldNri, DataViewRowState.CurrentRows);
                    dgv1.DataSource = v;
                    NewRiga(false);

                    dgv2.DataSource = FillIva("", "");
                }
            }
        }

        private void bilanciaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRow[] j;
            string sPar035BilScontrino = _clsFun.ParGet(clsDefine.enuParametri.Par035BilScontrino, _strConSql);

            string[] a = sPar035BilScontrino.Split(',');
            if (a.Length < 2 || a[0] != "S")
                MessageBox.Show("Configurazione non attiva!", "IMPORT BILANCIA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                string sPth = a[2];

                if (!Directory.Exists(sPth))
                    MessageBox.Show("Percorso non trovato " + sPth + "!", "IMPORT BILANCIA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    frmUtyInput1 f = new frmUtyInput1();
                    f._strDes = "Inserire il codice scontrino";
                    f.ShowDialog();
                    string s = f._strTxt;

                    if (s != "")
                    {
                        DataTable tBil = new DataTable();

                        if(a[1] == "BIZ")
                            tBil = new clsBilBizWinvarp().BizScontrino(sPar035BilScontrino, s);
                        else
                             MessageBox.Show("Bilancia non configurata!", "IMPORT BILANCIA", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        if (tBil.Rows.Count == 0)
                            MessageBox.Show("Dati scontrino non trovati!", "IMPORT BILANCIA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        else
                        {
                            DataTable t = new DataTable();
                            DataTable tTiv = (DataTable)cmbMovIva.DataSource;
                            DataTable tMov = ((DataView)dgv1.DataSource).Table;

                            string sFldNri = "mov_rmo";
                            if (_strMovFat == "F")
                                sFldNri = "mov_rfa";

                            int iRig = 0;
                            if (tMov.Rows.Count > 0)
                            {
                                if (((string)tMov.Rows[tMov.Rows.Count - 1][sFldNri]).Trim() == "" || ((string)tMov.Rows[tMov.Rows.Count - 1][sFldNri]).Trim() == NRIVUOTA)
                                    tMov.Rows[tMov.Rows.Count - 1].Delete();
                                if (tMov.Rows.Count > 0)
                                    iRig = Convert.ToInt16(tMov.Rows[tMov.Rows.Count - 1][sFldNri]);
                            }

                            foreach (DataRow y in tBil.Rows)
                            {
                                iRig++;
                                Boolean b = false;

                                s = (string)y["bil_ban"] + (string)y["bil_plu"];

                                t = _clsQry.ArtSeek(s, "PLU");

                                if (t.Rows.Count > 0)
                                {
                                    b = true;

                                    if (b)
                                    {
                                        DataRow x = tMov.NewRow();
                                        x["mov_yfa"] = "";
                                        x["mov_nfa"] = "";
                                        x["mov_ymo"] = "";
                                        x["mov_nmo"] = "";
                                        x["mov_rfa"] = NRIVUOTA;
                                        x["mov_rmo"] = NRIVUOTA;

                                        x[sFldNri] = iRig.ToString("0000");

                                        if (t != null && t.Rows.Count > 0)
                                        {
                                            x["mov_art"] = (string)t.Rows[0]["tmp_art"];
                                            x["mov_ard"] = (string)t.Rows[0]["tmp_ard"];
                                            x["mov_iva"] = (string)t.Rows[0]["tmp_iva"];
                                            x["mov_umi"] = (string)t.Rows[0]["tmp_umi"];
                                            x["mov_cos"] = (decimal)t.Rows[0]["tmp_cos"];
                                        }
                                        else
                                        {
                                            x["mov_art"] = "";
                                            x["mov_ard"] = (string)y["bil_ban"] + (string)y["bil_plu"] + " non trovato";
                                            x["mov_iva"] = "";
                                            x["mov_umi"] = "";
                                            x["mov_cos"] = 0;
                                        }

                                        if (((string)x["mov_iva"]).Trim() == "")
                                            x["mov_iva"] = _strIvaPar;
                                        if (((string)x["mov_iva"]).Trim() == "")
                                            x["mov_iva"] = "010";

                                        decimal dAli = 0;
                                        j = tTiv.Select("tab_cod='" + ((string)x["mov_iva"]).PadLeft(3, Convert.ToChar('0')) + "'");
                                        if (j.Length > 0 && _clsFun.Numerico(j[0]["tab_ali"], "0123456789"))
                                            dAli = Convert.ToDecimal(j[0]["tab_ali"]);

                                        decimal dImp = _clsFun.MenoIva((decimal)y["bil_imp"], dAli);
                                        decimal dPrv = 0;
                                        if (dImp > 0 && (decimal)y["bil_qta"] > 0)
                                            dPrv = Math.Round(dImp / (decimal)y["bil_qta"], 2);

                                        x["mov_qta"] = 1;
                                        x["mov_qkg"] = (decimal)y["bil_qta"];
                                        if ((string)t.Rows[0]["tmp_umi"] == "NR")
                                        {
                                            x["mov_qta"] = (decimal)y["bil_qta"];
                                            x["mov_qkg"] = 0;
                                        }

                                        x["mov_prv"] = dPrv;                    // (decimal)y["bil_prv"];
                                        x["mov_sco"] = "";
                                        x["mov_imp"] = dImp;                    // (decimal)y["bil_imp"];
                                        x["mov_ann"] = false;
                                        x["mov_day"] = DateTime.Today;
                                        x["mov_no1"] = "";
                                        x["MovMdy"] = "S";
                                        tMov.Rows.Add(x);
                                    }
                                }
                            }
                            NewRiga(true);
                            dgv2.DataSource = FillIva("", "");

                        }
                    }
                }
            }
        }

        private void eliminaRigheCancellateToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void eliminaDocumentoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (txtMftNdo.Text != "")
                MessageBox.Show("Documento già stampato!");
            else
            {
                if (MessageBox.Show("Confermi l'eliminazione TOTALE del documento corrente?", "CANCELLAZIONE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string s = "DELETE FROM GesMovimenti WHERE ";
                    if (_strMovFat == "M")
                        s += "mov_ymo='" + lblMftYea.Text + "' AND mov_nmo='" + lblMftNum.Text + "'";
                    else
                        s += "mov_yfa='" + lblMftYea.Text + "' AND mov_nfa='" + lblMftNum.Text + "'";

                    _clsFun.SqlWrite(s, _strConSql);

                    if (_strMovFat == "M")
                    {
                        s = "DELETE FROM GesMovTestate WHERE ";
                        s += "mot_ymo='" + lblMftYea.Text + "' AND mot_nmo='" + lblMftNum.Text + "'";
                    }
                    else
                    {
                        s = "DELETE FROM GesFatTestate WHERE ";
                        s += "fat_yfa='" + lblMftYea.Text + "' AND fat_nfa='" + lblMftNum.Text + "'";
                    }

                    _clsFun.SqlWrite(s, _strConSql);

                    this.Close();
                }
            }
        }

        private void eliminaRigheToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            EliminaRighe(true);
            NewRiga(false);
        }

        private void old_eliminaRigheCancellateToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (MessageBox.Show("Eliminazione righe cancellate, confermi?", "CANCELLAZIONE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Salva();

                string s = "DELETE FROM GesMovimenti WHERE ";
                if (_strMovFat == "M")
                    s += "mov_ymo='" + lblMftYea.Text + "' AND mov_nmo='" + lblMftNum.Text + "'";
                else
                    s += "mov_yfa='" + lblMftYea.Text + "' AND mov_nfa='" + lblMftNum.Text + "'";

                _clsFun.SqlWrite(s, _strConSql);

                string sFldNri = "mov_rmo";
                if (_strMovFat == "F")
                    sFldNri = "mov_rfa";

                int i = 0;

                DataView v = (DataView)dgv1.DataSource;
                v.Sort = sFldNri + " ASC";

                ArrayList aCan = new ArrayList();

                foreach (DataRowView R in v)
                {
                    if ((Boolean)R["mov_ann"] == true)
                    {
                        //aCan.Add((string)R[sFldNri]);
                        R.Delete();
                    }
                }

                //FillDati();

                foreach (DataRowView R in v)
                {

                    if (((string)R[sFldNri]).ToLower() != NRIVUOTA)
                    {
                        i++;
                        //R[sFldNri] = i.ToString("000");
                        if (_clsFun.Numerico(R[sFldNri]))
                            R[sFldNri] = i.ToString("0000");
                        Console.WriteLine("aaa");
                        R["MovMdy"] = "S";
                    }

                }

            }

        }

        private void eliminaRigheCancellateToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            EliminaRighe(false);
        }

        private void EliminaRighe(Boolean bolAll)
        {
            string s = "";

            s = "Eliminazione righe cancellate, confermi?";
            if(bolAll)
                s = "Eliminazione di tutte le righe, confermi?";

            if (MessageBox.Show(s, "CANCELLAZIONE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Salva();

                DataView v = (DataView)dgv1.DataSource;
                string sMsg = "";

                foreach (DataRowView r in v)
                {
                    s = "";

                    //if ((Boolean)r["mov_ann"] || bolAll)
                    if (true)
                    {
                        if (_strMovFat == "F")
                        {
                            if ((string)r["mov_nmo"] != _clsDef.COD06X)
                            {
                                s = "UPDATE GesMovimenti SET mov_yfa='" + _clsDef.COD04X + "', mov_nfa='" + _clsDef.COD06X + "', mov_rfa='" + _clsDef.COD04X + "' WHERE ";
                                s += "mov_yfa='" + r["mov_yfa"] + "' AND ";
                                s += "mov_nfa='" + r["mov_nfa"] + "' AND ";
                                s += "mov_rfa='" + r["mov_rfa"] + "' AND ";
                                s += "mov_ymo='" + r["mov_ymo"] + "' AND ";
                                s += "mov_nmo='" + r["mov_nmo"] + "' AND ";
                                s += "mov_rmo='" + r["mov_rmo"] + "'";
                            }
                            else
                            {
                                s = "DELETE FROM GesMovimenti WHERE ";
                                s += "mov_yfa='" + r["mov_yfa"] + "' AND ";
                                s += "mov_nfa='" + r["mov_nfa"] + "' AND ";
                                s += "mov_rfa='" + r["mov_rfa"] + "' AND ";
                                s += "mov_ymo='" + r["mov_ymo"] + "' AND ";
                                s += "mov_nmo='" + r["mov_nmo"] + "' AND ";
                                s += "mov_rmo='" + r["mov_rmo"] + "'";
                            }
                        }
                        else
                        {
                            if ((string)r["mov_nfa"] != _clsDef.COD06X)
                            {
                                if (_strMovFat == "F")
                                {
                                    if ((string)r["mov_rfa"] != "xxxx")
                                        sMsg += "riga nel movimento " + (string)r["mov_nmo"] + " " + (string)r["mov_ard"] + _clsDef.CRLF;
                                }
                                else
                                {
                                    if ((string)r["mov_rmo"] != "xxxx")
                                        sMsg += "riga nella fattura " + (string)r["mov_nfa"] + " " + (string)r["mov_ard"] + _clsDef.CRLF;
                                }
                            }
                            else
                            {
                                s = "DELETE FROM GesMovimenti WHERE ";
                                s += "mov_yfa='" + r["mov_yfa"] + "' AND ";
                                s += "mov_nfa='" + r["mov_nfa"] + "' AND ";
                                s += "mov_rfa='" + r["mov_rfa"] + "' AND ";
                                s += "mov_ymo='" + r["mov_ymo"] + "' AND ";
                                s += "mov_nmo='" + r["mov_nmo"] + "' AND ";
                                s += "mov_rmo='" + r["mov_rmo"] + "'";
                            }
                        }

                        if (s != "")
                            _clsFun.SqlWrite(s, _strConSql);
                        if (bolAll)
                            r["mov_ann"] = true;
                    }
                }

                int i = 0;
                string sFldNri = "mov_rmo";
                if (_strMovFat == "F")
                    sFldNri = "mov_rfa";
                v = (DataView)dgv1.DataSource;
                v.Sort = sFldNri + " ASC";
                ArrayList aCan = new ArrayList();
                foreach (DataRowView R in v)
                {
                    s = (string)R["mov_rmo"];
                    Console.WriteLine(s);

                    if ((Boolean)R["mov_ann"] == true)
                        R.Delete();
                }

                //for (int ii = v.Count - 1; ii == 0; ii--)
                //{
                //    if ((Boolean)v[ii]["mov_ann"] == true)
                //        v[ii].Delete();
                //}


                foreach (DataRowView R in v)
                {
                    if (((string)R[sFldNri]).ToLower() != NRIVUOTA)
                    {
                        i++;
                        if (_clsFun.Numerico(R[sFldNri]))
                            R[sFldNri] = i.ToString("0000");
                        Console.WriteLine("aaa");
                        R["MovMdy"] = "S";
                    }
                }

                if (sMsg != "")
                    MessageBox.Show(sMsg, "RIGHE NON CANCELLABILE", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void old_eliminaDocumentoToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (txtMftNdo.Text != "")
                MessageBox.Show("Documento già stampato!");
            else
            {
                if (MessageBox.Show("Confermi l'eliminazione TOTALE del documento corrente?", "CANCELLAZIONE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string s = "DELETE FROM GesMovimenti WHERE ";
                    if (_strMovFat == "M")
                        s += "mov_ymo='" + lblMftYea.Text + "' AND mov_nmo='" + lblMftNum.Text + "'";
                    else
                        s += "mov_yfa='" + lblMftYea.Text + "' AND mov_nfa='" + lblMftNum.Text + "'";

                    _clsFun.SqlWrite(s, _strConSql);

                    if (_strMovFat == "M")
                    {
                        s = "DELETE FROM GesMovTestate WHERE ";
                        s += "mot_ymo='" + lblMftYea.Text + "' AND mot_nmo='" + lblMftNum.Text + "'";
                    }
                    else
                    {
                        s = "DELETE FROM GesFatTestate WHERE ";
                        s += "fat_yfa='" + lblMftYea.Text + "' AND fat_nfa='" + lblMftNum.Text + "'";
                    }

                    _clsFun.SqlWrite(s, _strConSql);

                    _strReturn = "CANC";

                    this.Close();
                }
            }
        }

        private void eliminaDocumentoToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            string s = "";

            if (txtMftNdo.Text != "")
                MessageBox.Show("Documento già stampato!");
            else
            {
                if (MessageBox.Show("Confermi l'eliminazione TOTALE del documento corrente?", "CANCELLAZIONE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    DataView v = (DataView)dgv1.DataSource;
                    string sMsg = "";

                    foreach(DataRowView r in v)
                    {
                        s = "";

                        if (_strMovFat == "F")
                        {
                            if ((string)r["mov_nmo"] != _clsDef.COD06X)
                            {
                                s = "UPDATE GesMovimenti SET mov_yfa='" + _clsDef.COD04X + "', mov_nfa='" + _clsDef.COD06X + "', mov_rfa='" + _clsDef.COD04X + "' WHERE ";
                                s += "mov_yfa='" + r["mov_yfa"] + "' AND ";
                                s += "mov_nfa='" + r["mov_nfa"] + "' AND ";
                                s += "mov_rfa='" + r["mov_rfa"] + "' AND ";
                                s += "mov_ymo='" + r["mov_ymo"] + "' AND ";
                                s += "mov_nmo='" + r["mov_nmo"] + "' AND ";
                                s += "mov_rmo='" + r["mov_rmo"] + "'";
                            }
                            else
                            {
                                s = "DELETE FROM GesMovimenti WHERE ";
                                s += "mov_yfa='" + r["mov_yfa"] + "' AND ";
                                s += "mov_nfa='" + r["mov_nfa"] + "' AND ";
                                s += "mov_rfa='" + r["mov_rfa"] + "' AND ";
                                s += "mov_ymo='" + r["mov_ymo"] + "' AND ";
                                s += "mov_nmo='" + r["mov_nmo"] + "' AND ";
                                s += "mov_rmo='" + r["mov_rmo"] + "'";
                            }
                        }
                        else
                        {
                            if ((string)r["mov_nfa"] != _clsDef.COD06X)
                            {
                                if (_strMovFat == "F")
                                {
                                    if ((string)r["mov_rfa"] != "xxxx")
                                        sMsg += "riga del movimento " + (string)r["mov_nmo"] + " " + (string)r["mov_ard"] + _clsDef.CRLF;
                                }
                                else
                                {
                                    if ((string)r["mov_rmo"] != "xxxx")
                                        sMsg += "riga della fattura " + (string)r["mov_nfa"] + " " + (string)r["mov_ard"] + _clsDef.CRLF;
                                }
                            }
                            else
                            {
                                s = "DELETE FROM GesMovimenti WHERE ";
                                s += "mov_yfa='" + r["mov_yfa"] + "' AND ";
                                s += "mov_nfa='" + r["mov_nfa"] + "' AND ";
                                s += "mov_rfa='" + r["mov_rfa"] + "' AND ";
                                s += "mov_ymo='" + r["mov_ymo"] + "' AND ";
                                s += "mov_nmo='" + r["mov_nmo"] + "' AND ";
                                s += "mov_rmo='" + r["mov_rmo"] + "'";
                            }
                        }

                        if(s != "")
                        {
                            _clsFun.SqlWrite(s, _strConSql);
                        }
                    }

                    if (sMsg != "")
                        MessageBox.Show(sMsg, "DOCUMENTO NON CANCELLABILE", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    else
                    {
                        if (_strMovFat == "M")
                        {
                            s = "DELETE FROM GesMovTestate WHERE ";
                            s += "mot_ymo='" + lblMftYea.Text + "' AND mot_nmo='" + lblMftNum.Text + "'";
                        }
                        else
                        {
                            s = "DELETE FROM GesFatTestate WHERE ";
                            s += "fat_yfa='" + lblMftYea.Text + "' AND fat_nfa='" + lblMftNum.Text + "'";
                        }

                        _clsFun.SqlWrite(s, _strConSql);

                        _strReturn = "CANC";

                        this.Close();
                    }
                }
            }
        }

        private void importDaTerminalinoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImpTerm();
        }

        private void ImpTerm()
        {
            string s = "";
            DataRow[] j;
            Boolean b = true;

            string sNrig = "0001";

            DataTable t = new DataTable();
            //DataTable tEcr = _clsQry.tabEcr();
            DataTable tTer = new clsGenTabTmp().TabTmpTerRil("TabTer");
            DataTable tMov = ((DataView)dgv1.DataSource).Table;
            //s = "SELECT * FROM TabReparti";
            //DataTable tRep = _clsFun.FillTabSql("TabRep", s, false, _strConSql);

            string sFldNri = "mov_rmo";
            if (_strMovFat == "F")
                sFldNri = "mov_rfa";

            if (tMov.Rows.Count > 0)
            {
                if ((string)tMov.Rows[tMov.Rows.Count - 1][sFldNri] == NRIVUOTA)
                    tMov.Rows[tMov.Rows.Count - 1].Delete();

                if(tMov.Rows.Count > 0)
                    sNrig =  (string)tMov.Rows[tMov.Rows.Count - 1][sFldNri];
            }

            int i = Convert.ToInt32(sNrig);

            frmGesImpTerm f = new frmGesImpTerm();
            f._strTip = "INV";
            f._tabTer = tTer;
            f.ShowDialog();
            if (f._tabImp != null && f._tabImp.Rows.Count > 0)
            {
                string sMem = "";

                foreach (DataRow y in f._tabImp.Rows)
                {
                    if ((string)y["ord_art"] == "001261")
                        Console.WriteLine("aaaaa");
                    i++;
                    b = false;
                    if (((string)y["ord_ean"]).Trim() != "")
                    {
                        s = Convert.ToInt64(y["ord_ean"]).ToString();
                        t = _clsQry.ArtSeek(s, "");
                        b = true;
                    }
                    else if (((string)y["ord_art"]).Trim() != "")
                    {
                        s = Convert.ToInt64(y["ord_art"]).ToString();
                        t = _clsQry.ArtSeek(s, "");
                        b = true;
                    }

                    if (b)
                    {
                        if (t != null && t.Rows.Count > 0)
                        {
                            if (sMem == "")
                            {
                                s = (string)t.Rows[0]["tmp_art"];

                                j = tMov.Select("mov_art='" + s + "'");
                                if (j.Length > 0)
                                {
                                    s = (string)t.Rows[0]["tmp_art"] + " " + (string)t.Rows[0]["tmp_ard"] + _clsDef.CRLF;
                                    s += "ARTICOLO GIA' PRESENTE, VUOI AGGIUNGERLO?";

                                    frmMsg1 f2 = new frmMsg1();
                                    f2._bolMem = true;
                                    f2._strTip = "SINO";
                                    f2._strMsg = s;
                                    f2.ShowDialog();
                                    s = f2._strRes;
                                    sMem = f2._strMem;

                                    if (s != "SI")
                                        b = false;
                                }
                            }
                            else if (sMem != "SI")
                                b = false;
                        }
                        if (b)
                        {
                            DataRow x = tMov.NewRow();
                            x["mov_yfa"] = "";
                            x["mov_nfa"] = "";
                            x["mov_ymo"] = "";
                            x["mov_nmo"] = "";
                            x["mov_rfa"] = NRIVUOTA;
                            x["mov_rmo"] = NRIVUOTA;

                            x[sFldNri] = i.ToString("0000");

                            if (t != null && t.Rows.Count > 0)
                            {
                                x["mov_art"] = (string)t.Rows[0]["tmp_art"];
                                x["mov_ard"] = (string)t.Rows[0]["tmp_ard"];
                                x["mov_iva"] = (string)t.Rows[0]["tmp_iva"];
                                x["mov_umi"] = (string)t.Rows[0]["tmp_umi"];
                                x["mov_cos"] = (decimal)t.Rows[0]["tmp_cos"];
                            }
                            else
                            {
                                x["mov_art"] = "";
                                x["mov_ard"] = ((string)y["ord_ean"]).Trim() + " non trovato";
                                x["mov_iva"] = "";
                                x["mov_umi"] = "";
                                x["mov_cos"] = 0;
                            }

                            if ((string)x["mov_umi"] == "KG")
                            {
                                x["mov_qta"] = 1;
                                x["mov_qkg"] = (decimal)y["ord_qta"];
                            }
                            else
                            {
                                x["mov_qta"] = (decimal)y["ord_qta"];
                                x["mov_qkg"] = 0;
                            }

                            x["mov_prv"] = 0;
                            x["mov_sco"] = "";

                            //if (_strMovFat == "F" && _strCauTpd == "" && _strCfo == _clsDef.TIPCLI)
                            if (_strCfo == _clsDef.TIPCLI)
                                x["mov_imp"] = (decimal)x["mov_qta"] * (decimal)x["mov_prv"];
                            else
                                x["mov_imp"] = (decimal)x["mov_qta"] * (decimal)x["mov_cos"];

                            x["mov_ann"] = false;
                            x["mov_day"] = DateTime.Today;
                            x["mov_no1"] = "";
                            x["MovMdy"] = "S";
                            tMov.Rows.Add(x);
                        }
                    }
                }
            }
        }

        private void btnLotto_Click(object sender, EventArgs e)
        {
            if (Salva())
            {
                string s = "";
                Boolean b = true;

                if(_strMovFat == "M" && _strCfo == _clsDef.TIPCLI)
                {
                    DataRow[] j = ((DataTable)cmbMftTdc.DataSource).Select("tab_cod='" + cmbMftTdc.SelectedValue.ToString() + "'");
                    if (j.Length > 0)
                    {
                        s = "";

                        CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
                        if (cm.Position >= 0)
                        {
                            DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                            DataRow x = r.Row;

                            b = false;

                            frmSeekDocLotti f = new frmSeekDocLotti();
                            f._strSelect = "S";
                            f.ShowDialog();

                            s = f._strRes;

                            //string[] a = s.Split('-');
                            //if (a.Length > 1)
                            //{
                            //    x["mov_lot"] = a[1];
                            //    x["MovMdy"] = "S";
                            //}

                            x["mov_lot"] = s;
                            x["MovMdy"] = "S";
                        }

                    }
                }
                
                if (b)
                {
                    if (cmbMftCfo.SelectedValue == null || cmbMftCfo.SelectedValue.ToString() == "")
                        MessageBox.Show("Fornitore non definito!", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    else if (txtMftNdo.Text == "")
                        MessageBox.Show("Numero documento non definito!", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    else
                    {
                        CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
                        if (cm.Position >= 0)
                        {
                            DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                            DataRow x = r.Row;

                            if (DBNull.Value.Equals(x["mov_art"]) || (string)x["mov_art"] == "")
                                MessageBox.Show("Prodotto non definito!", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            else if (DBNull.Value.Equals(x["mov_qkg"]) || (decimal)x["mov_qkg"] == 0)
                                MessageBox.Show("Peso non definito!", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            else
                            {
                                string sArt = (string)x["mov_art"];
                                string sArd = (string)x["mov_ard"];

                                //string s = "";
                                string sLot = "";
                                string sYea = "";

                                if (!DBNull.Value.Equals(x["mov_lot"]))
                                {
                                    sLot = (string)x["mov_lot"];

                                    if (_strMovFat == "F")
                                        sYea = (string)x["mov_yfa"];
                                    else
                                        sYea = (string)x["mov_ymo"];
                                }
                                s = "SELECT * FROM GesDocLotti WHERE lot_yea='" + sYea + "' AND lot_cod='" + sLot + "'";

                                DataTable t = _clsFun.FillTabSql("GesDocLotti", s, true, _strConSql);
                                DataTable tt = t.Copy();

                                t.Columns.Add(new DataColumn()
                                {
                                    DataType = Type.GetType("System.String"),
                                    ColumnName = "LotMdy",
                                    Caption = "Mdy",
                                    MaxLength = 1,
                                    ReadOnly = false,
                                    DefaultValue = (String)""
                                });
                                t.Columns.Add(new DataColumn()
                                {
                                    DataType = Type.GetType("System.String"),
                                    ColumnName = "ForDes",
                                    Caption = "ForDes",
                                    MaxLength = 50,
                                    ReadOnly = false,
                                    DefaultValue = (String)""
                                });
                                t.Columns.Add(new DataColumn()
                                {
                                    DataType = Type.GetType("System.Decimal"),
                                    ColumnName = "MovQkg",
                                    Caption = "Peso",
                                    ReadOnly = false,
                                    DefaultValue = (Decimal)0
                                });

                                DataRow xx;
                                if (t.Rows.Count == 0)
                                {
                                    xx = t.NewRow();
                                    xx["lot_cod"] = "";
                                    xx["lot_for"] = cmbMftCfo.SelectedValue.ToString();
                                    xx["lot_ndo"] = txtMftNdo.Text;
                                    xx["lot_ddo"] = dtpMftDdt.Value;
                                    xx["lot_tip"] = "";
                                    xx["lot_des"] = "";
                                    xx["lot_nat"] = "";
                                    xx["lot_all"] = "";
                                    xx["lot_mac"] = "";
                                    xx["lot_sez"] = "";
                                    xx["lot_lot"] = "";
                                    xx["lot_mna"] = "";
                                    xx["lot_sbo"] = "";
                                    xx["lot_sta"] = "";
                                    //xx["lot_mdt"] = "";
                                    xx["ForDes"] = cmbMftCfo.Text;
                                    xx["LotMdy"] = "I";

                                    //_strMovFat = (string)_rowLot["lot_dot"];
                                    //_strMovYea = (string)_rowLot["lot_doy"];
                                    //_strMftNum = (string)_rowLot["lot_don"];
                                    //_strMftMov = (string)_rowLot["lot_dom"];
                                    //_strMovRow = (string)_rowLot["lot_dor"];
                                    //_strArt =    (string)_rowLot["lot_art"];

                                    xx["lot_dot"] = _strMovFat;
                                    xx["lot_doy"] = lblMftYea.Text;
                                    xx["lot_don"] = txtMftNdo.Text;
                                    xx["lot_dom"] = lblMftNum.Text;

                                    if (_strMovFat == "F")
                                        xx["lot_dor"] = (string)x["mov_rfa"];
                                    else
                                        xx["lot_dor"] = (string)x["mov_rmo"];

                                    xx["lot_art"] = (string)x["mov_art"];

                                    xx["MovQkg"] = (decimal)x["mov_qkg"];
                                }
                                else
                                {
                                    xx = t.Rows[0];
                                    xx["lot_for"] = cmbMftCfo.SelectedValue.ToString();
                                    xx["lot_ndo"] = txtMftNdo.Text;
                                    xx["lot_ddo"] = dtpMftDdt.Value;

                                    xx["lot_dot"] = _strMovFat;
                                    xx["lot_doy"] = lblMftYea.Text;
                                    xx["lot_don"] = txtMftNdo.Text;
                                    xx["lot_dom"] = lblMftNum.Text;

                                    if (_strMovFat == "F")
                                        xx["lot_dor"] = (string)x["mov_rfa"];
                                    else
                                        xx["lot_dor"] = (string)x["mov_rmo"];

                                    xx["lot_art"] = (string)x["mov_art"];
                                    xx["MovQkg"] = (decimal)x["mov_qkg"];
                                }

                                frmGesDocLottiDettaglio f = new frmGesDocLottiDettaglio();
                                f._strArt = sArt;
                                f._strArd = sArd;
                                f._strTip = "RowDoc";
                                f._rowLot = xx;

                                f._strMovFat = _strMovFat;
                                f._strMovYea = lblMftYea.Text;
                                f._strMftNum = lblMftNum.Text;

                                if (_strMovFat == "F")
                                    f._strMovRow = (string)x["mov_rfa"];
                                else
                                    f._strMovRow = (string)x["mov_rmo"];

                                f.ShowDialog();

                                if ((string)f._rowLot["lot_cod"] != (string)x["mov_lot"])
                                {
                                    x["mov_lot"] = (string)f._rowLot["lot_cod"];
                                    btnLotto.Text = "Lotto " + ((string)x["mov_lot"]).Trim();
                                    x["MovMdy"] = "S";
                                }

                            }
                        }
                    }
                }
            }
        }

        private void statisticheToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable tMov = ((DataView)dgv1.DataSource).Table;
            DataRow x;

            string sArr = "";

            string sFldNri = "mov_rmo";
            if (_strMovFat == "F")
                sFldNri = "mov_rfa";

            string sNrig = "0001";

            if (tMov.Rows.Count > 0)
            {
                if ((string)tMov.Rows[tMov.Rows.Count - 1][sFldNri] == NRIVUOTA)
                    tMov.Rows[tMov.Rows.Count - 1].Delete();

                if (tMov.Rows.Count > 0)
                    sNrig = (string)tMov.Rows[tMov.Rows.Count - 1][sFldNri];
            }

            int i = Convert.ToInt32(sNrig);

            frmGesDocStatArticoli f = new frmGesDocStatArticoli();
            //f._dayDti = dtpIni.Value;
            //f._dayDtf = dtpFin.Value;
            //f._strStatVisProf = _strStatVisProf;
            f.ShowDialog();

            DataTable tSta = f._tabCho.Copy();

            foreach(DataRow y in tSta.Rows)
            {

                decimal dCos = _clsQry.ArtCostoUltimo("", "", (string)y["ven_art"], DateTime.Today);


                x = tMov.NewRow();
                x["mov_yfa"] = "";
                x["mov_nfa"] = "";
                x["mov_ymo"] = "";
                x["mov_nmo"] = "";
                x["mov_rfa"] = NRIVUOTA;
                x["mov_rmo"] = NRIVUOTA;

                x[sFldNri] = i.ToString("0000");

                x["mov_art"] = (string)y["ven_art"];
                x["mov_ard"] = (string)y["ven_ard"];
                x["mov_iva"] = (string)y["ven_iva"];
                x["mov_umi"] = (string)y["ven_umi"];
                x["mov_cos"] = (decimal)y["ven_cos"];

                x["mov_qta"] = (decimal)y["ven_qta"];
                x["mov_qkg"] = (decimal)y["ven_qkg"];
                //x["mov_prv"] = (decimal)y["ven_prz"];
                if (Convert.ToDecimal(y["ven_cos"]) > 0)
                {
                    x["mov_prv"] = (decimal)y["ven_cos"];
                    x["mov_imp"] = (decimal)y["ven_qta"] * (decimal)y["ven_cos"];
                }
                else if(dCos > 0)
                {
                    x["mov_prv"] = dCos;
                    x["mov_imp"] = (decimal)y["ven_qta"] * dCos;
                }
                else
                {
                    x["mov_prv"] = (decimal)y["ven_prz"];
                    x["mov_imp"] = (decimal)y["ven_ven"];
                }

                x["mov_sco"] = "";
                x["mov_ann"] = false;
                x["mov_day"] = DateTime.Today;
                x["mov_no1"] = "";
                x["MovMdy"] = "S";

                tMov.Rows.Add(x);

                i++;
            }

            NewRiga(false);
            dgv2.DataSource = FillIva("", sArr);
        }

        private void controlloDaTerminalinoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmGesTermImport f = new frmGesTermImport();
            f.ShowDialog();
            string s = f._strTrmCho;

            if(s != "")
            {
                DocCtrlDoc(s);
            }
        }

        private void DocCtrlDoc(string strTrm)
        {
            string s = "";
            DataRow[] j;
            DataRow x;

            DataTable tMsg = new clsGenTabTmp().TabTmpMsg("TabMsg");

            s = "SELECT * FROM TmpDiv WHERE tmp_day='XXXXXXXXXXXXXXXXX' AND tmp_tip='XXXXXX'";
            DataTable tTrm = _clsFun.FillTabSql("TmpDiv", s, false, _strConSql);

            string[] a = strTrm.Split(';');

            foreach (string ss in a)
            {
                if (ss.Trim() != "")
                {
                    string[] aa = ss.Split('-');

                    string sDay = aa[0];
                    string sTip = aa[1];

                    s = "SELECT * FROM TmpDiv WHERE tmp_day='" + sDay + "' AND tmp_tip='" + sTip + "'";
                    DataTable t = _clsFun.FillTabSql("TmpDiv", s, false, _strConSql);

                    foreach (DataRow y in t.Rows)
                        tTrm.ImportRow(y);
                }
            }

            DataTable tTer = new clsGenTabTmp().TabTmpOrdTermMemor("TabImp");

            DataTable tDoc = ((DataView)dgv1.DataSource).ToTable();

            int iCnt = 0;
            int iMan = 0;
            int iDif = 0;

            x = tMsg.NewRow();
            x["msg_co1"] = "";
            x["msg_co2"] = "";
            x["msg_des"] = "";
            x["msg_msg"] = "*** Controllo differenze terminalino => documento";
            tMsg.Rows.Add(x);

            foreach(DataRow y in tTrm.Rows)
            {
                iCnt++;

                s = (string)y["tmp_tmp"];

                a = s.Split(';');
                string sEan = a[3];
                string sArt = a[4];
                string sArd = a[5];
                string sQta = a[6];

                x = tTer.NewRow();
                x["ord_art"] = sArt;
                x["ord_ean"] = sEan;
                x["ord_for"] = "";

                x["ord_qta"] = 0;
                if(_clsFun.Numerico(sQta, "0123456789."))
                    x["ord_qta"] = Convert.ToDecimal(sQta.Replace(".",","));

                x["ord_msg"] = "";
                x["ord_dtt"] = "";
                tTer.Rows.Add(x);

                j = tDoc.Select("mov_art='" + sArt + "'");
                if(j.Length == 0)
                {
                    iMan++;
                    x = tMsg.NewRow();
                    x["msg_co1"] = sEan;
                    x["msg_co2"] = sArt;
                    x["msg_des"] = sArd;
                    x["msg_msg"] = "articolo non trovato nel documento";
                    tMsg.Rows.Add(x);
                }
                else
                {
                    decimal d = (decimal)j[0]["mov_qta"];

                    if((string)j[0]["mov_umi"] == "KG")
                        d = (decimal)j[0]["mov_qkg"];

                    decimal dQta = Convert.ToDecimal(sQta.Replace(".",","));

                    if(dQta != d)
                    {
                        iDif++;
                        x = tMsg.NewRow();
                        x["msg_co1"] = sEan;
                        x["msg_co2"] = sArt;
                        x["msg_des"] = sArd;
                        x["msg_msg"] = "Quantità diverse: Q.tà documento: " + d.ToString() + " - Q.tà rilevata: " + dQta.ToString();
                        tMsg.Rows.Add(x);
                    }
                }
            }

            x = tMsg.NewRow();
            x["msg_co1"] = "";
            x["msg_co2"] = "";
            x["msg_des"] = "";
            x["msg_msg"] = "Righe lette: " + iCnt.ToString() + " mancanti: " + iMan.ToString() + " differenze: " + iDif.ToString();
            tMsg.Rows.Add(x);

            /***    xyz    ***/

            iCnt = 0;
            iMan = 0;
            iDif = 0;

            x = tMsg.NewRow();
            x["msg_co1"] = "";
            x["msg_co2"] = "";
            x["msg_des"] = "";
            x["msg_msg"] = "*** Controllo differenze documento => terminalino";
            tMsg.Rows.Add(x);

            foreach (DataRow y in tDoc.Rows)
            {

                //s = (string)y["tmp_tmp"];
                //a = s.Split(';');

                string sEan = "";
                string sArt = (string)y["mov_art"];
                string sArd = (string)y["mov_ard"];
                decimal dQta = (decimal)y["mov_qta"];

                if ((string)y["mov_umi"] == "KG")
                    dQta = (decimal)y["mov_qkg"];

                if (sArt.Trim() != "")
                {
                    iCnt++;
                    j = tTer.Select("ord_art='" + sArt + "'");
                    if (j.Length == 0)
                    {
                        iMan++;
                        x = tMsg.NewRow();
                        x["msg_co1"] = sEan;
                        x["msg_co2"] = sArt;
                        x["msg_des"] = sArd;
                        x["msg_msg"] = "articolo non trovato nel documento";
                        tMsg.Rows.Add(x);
                    }
                    else
                    {
                        decimal d = (decimal)j[0]["ord_qta"];

                        //decimal dQta = Convert.ToDecimal(sQta);

                        if (dQta != d)
                        {
                            iDif++;
                            x = tMsg.NewRow();
                            x["msg_co1"] = sEan;
                            x["msg_co2"] = sArt;
                            x["msg_des"] = sArd;
                            x["msg_msg"] = "Quantità diverse: Q.tà documento: " + d.ToString() + " - Q.tà rilevata: " + dQta.ToString();
                            tMsg.Rows.Add(x);
                        }
                    }
                }
            }

            x = tMsg.NewRow();
            x["msg_co1"] = "";
            x["msg_co2"] = "";
            x["msg_des"] = "";
            x["msg_msg"] = "Righe lette: " + iCnt.ToString() + " mancanti: " + iMan.ToString() + " differenze: " + iDif.ToString();
            tMsg.Rows.Add(x);

            new clsGenPdfVenErr().PrnPdfTrmDocErr(tMsg);
        }

        private void cmbMftNeg_SelectedIndexChanged(object sender, EventArgs e)
        {
            _strMftUbi = cmbMftNeg.SelectedValue.ToString();
        }

    }
}
