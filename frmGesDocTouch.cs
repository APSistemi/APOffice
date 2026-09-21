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
using System.IO.Ports;

namespace APOffice
{
    public partial class frmGesDocTouch : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();
        clsBilCheckOutSerialPort _clsBil = new clsBilCheckOutSerialPort();

        clsResize _form_resize;

        private string _strConSql = "";
        private string _strCfo = "";
        private string _strIvaPar = "022";
        private string _strFrmPar = "";

        public string _strMovFat = "";
        public string _strCauTpd = "";
        public string _strMftYea = "";
        public string _strMftNum = "";
        public string _strMftUbi = "";
        public string _strReturn = "";
        public string _strDocSuf = "";
        public string _strBilTime = "";

        public string _strOpenPorta = "";
        public SerialPort _serialPort = new SerialPort();

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
        private const string TABTABTGR = "TabTipoGrammatura";
        
        private const string NRIVUOTA = "xxxx";

        private string _strIni11VenditaTouch = "";

        public frmGesDocTouch()
        {
            InitializeComponent();

            _form_resize = new clsResize(this);
            this.Load += _Load;
            this.Resize += _Resize;

            new clsGesGraph().SetGraph(this, 0);
        }
        private void _Load(object sender, EventArgs e)
        {
            _form_resize._get_initial_size();
        }
        private void _Resize(object sender, EventArgs e)
        {
            _form_resize._resize();
        }
        private void frmGesDocumento_Load(object sender, EventArgs e)
        {
            Video();

            _strConSql = _clsFun.ConSql("");
            _strIvaPar = _clsFun.ParGet(clsDefine.enuParametri.Par012CodIVAxDefault, _strConSql);
            _strIni11VenditaTouch = _clsFun.FileIni("R", clsDefine.enuIni.Ini11VenditaTouch, "");
            _strMftUbi = _clsFun.FileIni("R", clsDefine.enuIni.Ini13Ubicazione, "");

            FillTab();
            SetDgv1();
            SetDgv2();
            FillDati();

            if (_strOpenPorta == "S")
            {
                string sMsg = "";

                _serialPort = _clsBil.iniCom(ref sMsg);
                //_clsBil.PortaApri(_serialPort);

                if (!_serialPort.IsOpen)
                {
                    Console.WriteLine("");

                    string[] a = sMsg.Split('|');
                    string s1 = "";
                    string s2 = "";

                    if (a.Length > 0)
                        s1 = a[0];
                    else if (a.Length > 1)
                        s2 = a[2];
                    else
                        s1 = "Problemi accesso alla bilancia!";
                    MessageBox.Show(s1,s2, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }

        private void Video()
        {
            _form_resize._get_initial_size();
            _strFrmPar = _clsQry.ParForm(this, "R", "");

            this.CenterToScreen();
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
            //_strReturn = cmbMftSta.Text + "|";
            _strReturn = cmbMftSta.Text + "|" + cmbMftCfo.Text + "|" + txtMftNdo.Text + "|" + lblTotTot.Text;

            _clsQry.ParForm(this, "W", "");

            Boolean b = Salva();

            Boolean bChiudi = false;

            if (!b)
            {
                if (MessageBox.Show("Continui l'uscita?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    bChiudi = true;
            }
            else
            {
                bChiudi = true;
            }

            if (bChiudi)
            {
                if (_strOpenPorta == "S")
                {
                    if (_serialPort.IsOpen)
                        _serialPort.Close();
                }

                this.Close();
            }
        }

        private void SetDgv1()
        {
            string s = "";

            s = "SELECT * FROM TabIva ORDER BY tab_des";
            DataTable tIva = _clsFun.FillTabSql(TABTABIVA, s, false, _strConSql);

            s = "SELECT * FROM TabUmi ORDER BY tab_des";
            DataTable tUmi = _clsFun.FillTabSql(TABTABUMI, s, false, _strConSql);

            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            DataGridViewComboBoxColumn cCmb;

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "mov_idx";
            //cTbc.Name = "Idx";
            //cTbc.Width = 20;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = true;
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dgv1.Columns.Add(cTbc);

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
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_ard";
            cTbc.Name = "Descrizione";
            cTbc.Width = 280;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "mov_iva";
            //cTbc.Name = "IVA";
            //cTbc.Width = 40;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = false;
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dgv1.Columns.Add(cTbc);

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "mov_umi";
            cCmb.Name = "UM";
            cCmb.Width = 70;
            cCmb.DataSource = tUmi;
            cCmb.ValueMember = "tab_cod";
            cCmb.DisplayMember = "tab_cod";
            cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            dgv1.Columns.Add((DataGridViewColumn)cCmb);

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "mov_iva";
            cCmb.Name = "IVA";
            cCmb.Width = 80;
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
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            cTbc.MaxInputLength = 10;
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_sco";
            cTbc.Name = "Sconti";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.MaxInputLength = 15;
            cTbc.ReadOnly = false;
            cTbc.ToolTipText = "Se sconto valore precedere con 'V'";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_imp";
            cTbc.Name = "Importo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.DataPropertyName = "mov_ann";
            cCbc.Name = "Ann.";
            cCbc.Width = 30;
            cCbc.ValueType = typeof(string);
            cCbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_lot";
            cTbc.Name = "Lotto";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_tar";
            cTbc.Name = "Tara gr";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0";
            cTbc.MaxInputLength = 10;
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_ori";
            cTbc.Name = "Origine";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "MovMdy";
            cTbc.Name = "Modificato";
            cTbc.Width = 0;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_ubi";
            cTbc.Name = "Negozio/Sede";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
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

            //dgv1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //dgv1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_ali";
            cTbc.Name = "Aliquota";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_imp";
            cTbc.Name = "Imponbile";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_iva";
            cTbc.Name = "Imposta";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_tot";
            cTbc.Name = "Totale";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
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
            }
            else
            {
                //s = "SELECT * FROM TabMovCausali WHERE tab_cod='" + _strCauTpd + "'";
                ////s = "SELECT * FROM TabMovCausali ORDER BY tab_des";
                //t = _clsFun.FillTabSql(TABTABMCA, s, false, _strConSql);

                s = "SELECT * FROM TabMovCausali WHERE tab_cod='" + _strCauTpd + "'";
                //s = "SELECT * FROM TabMovCausali ORDER BY tab_des";
                t = _clsFun.FillTabSql(TABTABMCA, s, false, _strConSql);


                if (t.Rows.Count > 0)
                {
                    _strCfo = (string)t.Rows[0]["tab_cfo"];
                }

                cmbMftTdc.DataSource = t;
                cmbMftTdc.DisplayMember = "tab_des";
                cmbMftTdc.ValueMember = "tab_cod";
                cmbMftTdc.SelectedValue = _strCauTpd;
                //cmbMftTdc.DropDownStyle = ComboBoxStyle.Simple;
            }

            s = "SELECT * FROM TabIva ORDER BY tab_des";
            t = _clsFun.FillTabSql(TABTABIVA, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            _dasGen.Tables.Add(t);
            //cmbMovIva.DataSource = t;
            //cmbMovIva.DisplayMember = "tab_des";
            //cmbMovIva.ValueMember = "tab_cod";
            //cmbMovIva.SelectedValue = "";

            s = "SELECT * FROM TabUmi ORDER BY tab_des";
            t = _clsFun.FillTabSql(TABTABUMI, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = " Non def.";
            t.Rows.InsertAt(x, 0);
            _dasGen.Tables.Add(t);

            s = "SELECT * FROM TabTipoGrammatura ORDER BY tab_cod";
            t = _clsFun.FillTabSql(TABTABTGR, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = " Non def.";
            t.Rows.InsertAt(x, 0);
            _dasGen.Tables.Add(t);

            //cmbMovUmi.DataSource = t;
            //cmbMovUmi.DisplayMember = "tab_des";
            //cmbMovUmi.ValueMember = "tab_cod";
            //cmbMovUmi.SelectedValue = "";

            s = "SELECT * FROM TabPagamenti ORDER BY tab_des";
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
            t = _clsFun.FillTabSql(TABTABSTA, s, false, _strConSql);
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

            FillTabCfo();

            cmbMftSco.Items.Add("Valore");
            cmbMftSco.SelectedIndex = 0;
        }

        private void FillTabCfo()
        {
            //DataRow[] j = ((DataTable)cmbMftTdc.DataSource).Select("tab_cod='" + cmbMftTdc.SelectedValue + "'");
            //if(j.Length > 0)
            //{
            string s = "";
            //if ((string)j[0]["tab_cfo"] == _clsDef.TIPCLI)

            if (_strCfo == "")
            {
                cmbMftCfo.Enabled = false;
                btnCfo.Enabled = false;

                s = "SELECT for_cod AS CfoCod, for_des AS CfoDes FROM AnaFornitori WHERE for_cod='XXXXXXXXXXX'";

                DataTable t = _clsFun.FillTabSql(TABTABMCA, s, false, _strConSql);
                //DataRow x = t.NewRow();
                //x["CfoCod"] = "";
                //x["CfoDes"] = "  Non definito";
                //t.Rows.InsertAt(x, 0);

                cmbMftCfo.DataSource = t;
                cmbMftCfo.DisplayMember = "CfoDes";
                cmbMftCfo.ValueMember = "CfoCod";
                cmbMftCfo.SelectedValue = "";
            }
            else
            {
                if (_strCfo == _clsDef.TIPCLI)
                {
                    s = "SELECT ";
                    s += "cli_cod AS CfoCod, ";
                    s += "cli_des AS CfoDes, ";
                    s += "cli_lic AS CfoLic, ";
                    s += "cli_lir AS CfoLir, ";
                    s += "cli_etc AS CfoEtc, ";
                    s += "cli_ind AS CfoInd, ";
                    s += "cli_cap AS CfoCap, ";
                    s += "cli_loc AS CfoLoc, ";
                    s += "cli_prv AS CfoPrv ";
                    s += "FROM AnaClienti ORDER BY cli_des";
                }
                else
                    s = "SELECT for_cod AS CfoCod, for_des AS CfoDes FROM AnaFornitori ORDER BY for_des";

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
            
            //}
        }

        private void FillDati()
        {
            if (((DataTable)cmbMftTdc.DataSource).Rows.Count == 0)
                MessageBox.Show("Tabella tipo documento fattura vuota!", "COMPILARE TABELLA", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                string s = "";
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
                    s += "fat_des AS DocDes ";
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
                    txtMftNdo.Text = (string)t.Rows[0]["DocNdo"];
                    _strMftUbi = (string)t.Rows[0]["DocUbi"];

                    if (_strMovFat == "F" && !DBNull.Value.Equals(t.Rows[0]["DocDdo"]))
                        cmbMftTdc.SelectedValue = (string)t.Rows[0]["DocTdo"];

                    if (cmbMftTdc.SelectedValue == null)
                    {
                        cmbMftTdc.SelectedValue = ((DataTable)cmbMftTdc.DataSource).Rows[0]["tab_cod"];
                    }

                    if (txtMftNdo.Text.Trim() != "")
                        cmbMftTdc.Enabled = false;

                    dtpMftDdt.Value = (DateTime)t.Rows[0]["DocDdo"];
                    txtMftNo1.Text = (string)t.Rows[0]["DocNo1"];
                    cmbMftTpg.SelectedValue = (string)t.Rows[0]["DocTpg"];
                    cmbMftSta.SelectedValue = (string)t.Rows[0]["DocSta"];
                    //chkMftAnn.Checked = (Boolean)t.Rows[0]["DocAnn"];

                    cmbMftNeg.SelectedValue = "";
                    if (!DBNull.Value.Equals(t.Rows[0]["DocNeg"]))
                        cmbMftNeg.SelectedValue = (string)t.Rows[0]["DocNeg"];

                    _strMftUbi = (string)t.Rows[0]["DocUbi"];
                    if (_strMftUbi == "")
                        _strMftUbi = cmbMftNeg.SelectedValue.ToString();
                     
                    sMftDes = (string)t.Rows[0]["DocDes"];

                    //chkMftPvi.Checked = (Boolean)t.Rows[0]["DocPvi"];
                    //if (!DBNull.Value.Equals(t.Rows[0]["DocSta"]))
                    //    cmbMftSta.SelectedValue = (string)t.Rows[0]["DocSta"];
                }

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
                s += "mov_qta, ";
                s += "mov_tar, ";
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

                if (_strCfo == _clsDef.TIPCLI)
                {
                    FillCmbCliDest(sMftDes);
                }

                //MessageBox.Show(sFldNri);

                DataView v = new DataView(t, "", sFldNri, DataViewRowState.CurrentRows);

                dgv1.DataSource = v;

                //if (v.Count == 0)
                //NewRiga(false);
                dgv2.DataSource = FillIva("");

                if (v.Count > 0)
                {
                    int i = dgv1.Rows.Count - 1;
                    if (i >= 0)
                    {
                        //dgv1.Rows[i].Selected = true;
                        dgv1.CurrentCell = dgv1[3, i];
                    }
                }
            }
        }

        private void FillCmbCliDest(String strMftDes)
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

        private DataTable FillIva(string strRange)
        {
            string s = "";
            DataRow[] j;
            DataRow x;
            decimal dIva = 0;
            decimal dImp = 0;
            decimal dTot = 0;
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


            //DataTable tTiv = (DataTable)cmbMovIva.DataSource;
            DataTable tTiv = _dasGen.Tables[TABTABIVA];

            if (dgv2.DataSource == null)
                dgv2.DataSource = new clsGenTabTmp().TabTmpIvaRiep("TabIva");

            DataTable tIva = (DataTable)dgv2.DataSource;
            tIva.Clear();

            DataTable t = ((DataView)dgv1.DataSource).ToTable();

            foreach(DataRow y in t.Rows)
            {
                //if ((decimal)y["mov_imp"] < 0)
                //    Console.WriteLine("aaaa");

                if (DBNull.Value.Equals(y["mov_imp"]))
                    y["mov_imp"] = 0;


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
                            //j[0]["iva_iva"] = (decimal)j[0]["iva_iva"] + _clsFun.ValIva((decimal)y["mov_imp"], (decimal)j[0]["iva_ali"]);
                            //j[0]["iva_imp"] = Math.Round((decimal)j[0]["iva_imp"] + (decimal)y["mov_imp"], 2, MidpointRounding.AwayFromZero);
                            j[0]["iva_imp"] = Math.Round((decimal)j[0]["iva_imp"] + (decimal)y["mov_imp"], 3, MidpointRounding.ToEven);

                        }
                        else
                            Console.WriteLine("aaaaaaaaa");
                        if((string)y["mov_umi"] == "KG")
                            dQta += 1;
                        else
                            dQta += (decimal)y["mov_qta"];

                    }
                }
            }

            foreach (DataRow y in tIva.Rows)
            {
                if ((decimal)y["iva_imp"] != 0)
                {
                    //y["iva_iva"] = _clsFun.ValIva((decimal)y["iva_imp"], (decimal)y["iva_ali"]);
                    //y["iva_iva"] = Math.Round((decimal)y["iva_iva"], 2, MidpointRounding.ToEven);

                    y["iva_iva"] = _clsFun.ValIva((decimal)y["iva_imp"], (decimal)y["iva_ali"]);

                    y["iva_iva"] = Math.Round((decimal)y["iva_iva"], 2, MidpointRounding.ToEven);


                    y["iva_tot"] = (decimal)y["iva_imp"] + (decimal)y["iva_iva"];
                }
            }

            foreach (DataRow y in tIva.Rows)
            {
                dImp += (decimal)y["iva_imp"];
                dIva += (decimal)y["iva_iva"];
                dTot += (decimal)y["iva_tot"];
            }

            dImp = Math.Round(dImp, 2, MidpointRounding.AwayFromZero); // MidpointRounding.ToEven);

            lblTotQta.Text = _clsFun.Dec2Txt(dQta, 0);
            //lblTotImp.Text = _clsFun.Dec2Txt(dImp, 2);
            lblTotImp.Text = _clsFun.Dec2Txt(dImp, 2);
            lblTotIva.Text = _clsFun.Dec2Txt(dIva, 2);
            lblTotTot.Text = _clsFun.Dec2Txt(dTot, 2);

            return tIva;
        }

        private void CalcScontoTot()
        {
            string s = "";
            decimal dSco = 0;

            decimal decTotLordo = Convert.ToDecimal(lblTotTot.Text);

            if (!_clsFun.Numerico(txtMftSco.Text) || Convert.ToDecimal(txtMftSco.Text) == 0)
                return;
            else
                dSco = Convert.ToDecimal(txtMftSco.Text);

            decimal dTot = 0;
            DataTable t = ((DataView)dgv1.DataSource).Table;

            //string sFldRig = "mov_rmo";
            //if (_strMovFat == "F")
            //    sFldRig = "mov_rfa";

            //int iRigIni = 0;
            //int iRigFin = 10000;

            //if (strRange != "")
            //{
            //    iRigIni = Convert.ToInt16(strRange.Substring(0, 3));
            //    iRigFin = Convert.ToInt16(strRange.Substring(4, 3));
            //}

            foreach (DataRow y in t.Rows)
            {
                if (!(Boolean)y["mov_ann"]) // && Convert.ToInt16(y[sFldRig]) >= iRigIni && Convert.ToInt16(y[sFldRig]) <= iRigFin)
                    dTot += (decimal)y["mov_imp"];
            }

            if (txtMftSco.Text != "")
            {
                string sTip = cmbMftSco.SelectedItem.ToString().Substring(0, 1);

                //decimal dDelta = dSco * 100 / dTot;
                decimal dDelta = dSco * 100 / (decTotLordo);
                decimal dScoTot = 0;

                string sRig = "";

                foreach (DataRow y in t.Rows)
                {
                    if (!(Boolean)y["mov_ann"] && (decimal)y["mov_imp"] > 0) // && Convert.ToInt16(y[sFldRig]) >= iRigIni && Convert.ToInt16(y[sFldRig]) <= iRigFin)
                    {
                        decimal d = (decimal)y["mov_imp"];
                        d = _clsFun.MenoPer(d, dDelta);

                        d = Math.Round(d, 2, MidpointRounding.AwayFromZero);
                        s = "V" + ((decimal)y["mov_imp"] - d).ToString();
                        if (s.Length > 15)
                            s = s.Substring(0, 15);
                        y["mov_sco"] = s;
                        y["mov_imp"] = d;

                        dScoTot += d;
                        sRig = (string)y["mov_rfa"];
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
                    if (!(Boolean)y["mov_ann"] && (decimal)y["mov_imp"] > 0)
                    {
                        y["mov_sco"] = "";
                        if ((string)y["mov_umi"] == "KG")
                        {
                            if (_strCauTpd == "FA")
                                y["mov_imp"] = (decimal)y["mov_cos"] * (decimal)y["mov_qkg"];
                            else
                                y["mov_imp"] = (decimal)y["mov_prv"] * (decimal)y["mov_qkg"];
                        }
                        else
                        {
                            if (_strCauTpd == "FA")
                                y["mov_imp"] = (decimal)y["mov_cos"] * (decimal)y["mov_qta"];
                            else
                                y["mov_imp"] = (decimal)y["mov_prv"] * (decimal)y["mov_qta"];
                        }
                    }
                }
            }
            dgv2.DataSource = FillIva("");
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
                    FillTabCfo();

                    if (f._tabTmp.Rows.Count > 0)
                    {
                        //DataTable t = ((DataView)dgv2.DataSource).Table;
                        //DataRow x = t.NewRow();
                        //x["lia_for"] = f._tabTmp.Rows[0]["tmp_for"];
                        //x["CosFod"] = f._tabTmp.Rows[0]["tmp_fod"];
                        //x["lia_dti"] = DateTime.Today;
                        //x["lia_tip"] = "M";
                        //x["lia_arf"] = "";
                        //x["lia_cos"] = 0;
                        //x["LiaMdy"] = "S";
                        //t.Rows.Add(x);

                        cmbMftCfo.SelectedValue = f._tabTmp.Rows[0]["tmp_cli"];

                        Console.WriteLine("aaaaaaaaa");

                        if(cmbMftCfo.SelectedValue == null)
                        {
                            FillTabCfo();
                            cmbMftCfo.SelectedValue = f._tabTmp.Rows[0]["tmp_cli"];

                            Console.WriteLine("aaaaaaaaa");
                        }
                    }

                    FillCmbCliDest("");
                }
            }
            //if ((DataView)dgv1.DataSource == null || ((DataView)dgv1.DataSource).Count == 0)
            //    NewRiga(true);

        }

        private void cmbMftCfo_SelectionChangeCommitted(object sender, EventArgs e)
        {
            //if ((DataView)dgv1.DataSource == null || ((DataView)dgv1.DataSource).Count == 0)
            //    NewRiga(true);
        }

        private void NewRiga(Boolean bolCtrl)
        {
            if (_strCfo != "" && (cmbMftCfo.SelectedValue == null || cmbMftCfo.SelectedValue.ToString() == ""))
            {
                if (bolCtrl)
                    MessageBox.Show("Cliente destinatario non definto!", "NUOVO ARTICOLO", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                        t.Rows.Add(x);

                        PosRowLast();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                    FillDettaglio();
                }
            }
        }

        private void PosRowLast()
        {
            int nRowIndex = dgv1.Rows.Count - 1;
            int nColumnIndex = 3;

            dgv1.FirstDisplayedScrollingRowIndex = nRowIndex;

            dgv1.Rows[nRowIndex].Selected = true;
            dgv1.Rows[nRowIndex].Cells[1].Selected = true;
            dgv1.CurrentCell = dgv1[nColumnIndex, nRowIndex];
        }

        private void FillDettaglio()    //Boolean bolCls)
        {
            try
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    //txtMovArt.Text = (string)x["mov_art"];
                    //if (!DBNull.Value.Equals(x["mov_ard"]))
                    //    txtMovArd.Text = (string)x["mov_ard"];

                    //if (DBNull.Value.Equals(x["mov_iva"]))
                    //    x["mov_iva"] = "";
                    //cmbMovIva.SelectedValue = (string)x["mov_iva"];

                    //if (DBNull.Value.Equals(x["mov_umi"]))
                    //    x["mov_umi"] = "";
                    //cmbMovUmi.SelectedValue = (string)x["mov_umi"];

                    //txtMovQta.Text = "0";
                    //if (_clsFun.Numerico(x["mov_qta"]))
                    //    txtMovQta.Text = _clsFun.Dec2Txt(x["mov_qta"], 2);

                    //txtMovQkg.Text = "0";
                    //if (_clsFun.Numerico(x["mov_qkg"]))
                    //    txtMovQkg.Text = _clsFun.Dec2Txt(x["mov_qkg"], 2);

                    //txtMovCos.Text = "0";
                    //if (_clsFun.Numerico(x["mov_qta"]))
                    //    txtMovCos.Text = _clsFun.Dec2Txt(x["mov_cos"], 3);

                    //txtMovImc.Text = "0";
                    //if ((decimal)x["mov_cos"] != 0 && (decimal)x["mov_qta"] != 0)
                    //{ 
                    //    if((string)x["mov_umi"] == "KG")
                    //        txtMovImc.Text = _clsFun.Dec2Txt((decimal)x["mov_cos"] * (decimal)x["mov_qkg"], 3);
                    //    else
                    //        txtMovImc.Text = _clsFun.Dec2Txt((decimal)x["mov_cos"] * (decimal)x["mov_qta"], 3);
                    //}

                    //txtMovPrv.Text = "0";
                    //if (_clsFun.Numerico(x["mov_prv"]))
                    //    txtMovPrv.Text = _clsFun.Dec2Txt(x["mov_prv"], 2);

                    //txtMovSco.Text = "0";

                    //txtMovImp.Text = "0";
                    //if (_clsFun.Numerico(x["mov_imp"]))
                    //    txtMovImp.Text = _clsFun.Dec2Txt(x["mov_imp"], 2);

                    //chkMovAnn.Checked = false;
                    //if (!DBNull.Value.Equals(x["mov_ann"]))
                    //    chkMovAnn.Checked = Convert.ToBoolean(x["mov_ann"]);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private void MovRowAgg()
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                //DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                //r.Row["mov_art"] = txtMovArt.Text;
                //r.Row["mov_ard"] = txtMovArd.Text;
                //r.Row["mov_iva"] = cmbMovIva.SelectedValue;
                //r.Row["mov_umi"] = cmbMovUmi.SelectedValue;
                //r.Row["mov_qta"] = _clsFun.Txt2Dec(txtMovQta.Text);
                //r.Row["mov_qkg"] = _clsFun.Txt2Dec(txtMovQkg.Text);
                //r.Row["mov_cos"] = _clsFun.Txt2Dec(txtMovCos.Text);
                //r.Row["mov_prv"] = _clsFun.Txt2Dec(txtMovPrv.Text);
                //r.Row["mov_imp"] = _clsFun.Txt2Dec(txtMovImp.Text);
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

                        //sNri = "x" + sNri;
                        //DataTable t = (DataTable)dgv1.DataSource;
                        //DataTable t = ((DataView)dgv1.DataSource).ToTable();

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

                                FillMovRow(t, x);
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
                    FillMovRow(tabTmp, x);
                }
            }
        }

        private void dgv1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            TextBox txb = e.Control as TextBox;

            if(txb != null)
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
                                DataTable t = _clsQry.ArtSeek(txb.Text, "SEEK");
                                if (t != null && t.Rows.Count > 0)
                                {
                                    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                                    if (cm.Position >= 0)
                                    {
                                        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                                        DataRow x = r.Row;
                                        FillMovRow(t, x);
                                    }
                                }
                            }
                        }
                    }
                };
            }
        }

        private void dgv1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 2)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    try
                    {

                        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                        DataRow x = r.Row;

                        //x = CalcImporto(_strCauTpd, x, "IMP");

                        if (e.ColumnIndex == 7 || e.ColumnIndex == 8)
                            x = CalcImporto(_strCauTpd, x, "PRE");
                        else if (e.ColumnIndex == 5)
                            x = CalcImporto(_strCauTpd, x, "QTA");
                        else
                            x = CalcImporto(_strCauTpd, x, "IMP");

                        CtrlNumRiga(x);
                        cm.EndCurrentEdit();

                        dgv2.DataSource = FillIva("");

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }


                    //if (!DBNull.Value.Equals(x["mov_ard"]) && ((string)x["mov_ard"]).Trim() != "")
                    //    NewRiga(true);
                }
            }
        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv1.Columns[e.ColumnIndex].Name == "Ann." && dgv1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv1.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
                FillIva("");
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

                if (s == "Q.tà" || s == "Costo" || s == "Prezzo vendita" || s == "Importo")
                {
                    if (e.KeyChar == '.')
                        e.KeyChar = ',';
                }
            }
        }

        private DataRow oldCalcImporto(string strCauTpd, DataRow rowMov)
        {
            decimal d = (decimal)rowMov["mov_cos"];
            if (_strMovFat == "F" && _strCauTpd == "FV")
            {
                d = (decimal)rowMov["mov_prv"];
            }
            else if (_strMovFat == "M")
            {
                DataRow[] j = ((DataTable)cmbMftTdc.DataSource).Select("tab_cod='" + cmbMftTdc.SelectedValue.ToString() + "'");
                if(j.Length > 0)
                {
                    if((string)j[0]["tab_cfo"] == "CLI")
                        d = (decimal)rowMov["mov_prv"];
                }
            }

            if (!DBNull.Value.Equals(rowMov["mov_sco"]) && ((string)rowMov["mov_sco"]).Trim() != "")
            {
                string s = ((string)rowMov["mov_sco"]).Trim();

                if (s.Substring(0, 1) == "v")
                {
                    s = "V" + s.Substring(1);
                    rowMov["mov_sco"] = s;
                }

                string sTip = "%";
                if(s.Substring(0,1).ToUpper() == "V")
                {
                    sTip = s.Substring(0, 1).ToUpper();
                    s = s.Substring(1);
                }

                string[] a = s.Split('+');

                foreach(string v in a)
                {
                    if (_clsFun.Numerico(v))
                    {
                        decimal dSco = Convert.ToDecimal(v.Replace(".",","));

                        if (sTip == "%")
                            d = _clsFun.MenoPer(d, dSco);
                        else if (sTip == "V")
                            d = d - dSco;
                    }
                }
            }

            if (strCauTpd == "FA")
            {
                if ((string)rowMov["mov_umi"] == "KG")
                    rowMov["mov_imp"] = d * (decimal)rowMov["mov_qkg"];
                else
                    rowMov["mov_imp"] = d * (decimal)rowMov["mov_qta"];
            }
            else
            {
                if ((string)rowMov["mov_umi"] == "KG")
                    rowMov["mov_imp"] = d * (decimal)rowMov["mov_qkg"];
                else
                    rowMov["mov_imp"] = d * (decimal)rowMov["mov_qta"];
            }
            rowMov["MovMdy"] = "S";

            return rowMov;
        }

        private DataRow CalcImporto(string strCauTpd, DataRow rowMov, string strTip)
        {
            string sFld = "mov_cos";
            decimal dSco = 0;
            decimal d = (decimal)rowMov["mov_cos"];

            if (DBNull.Value.Equals(rowMov["mov_imp"]))
                rowMov["mov_imp"] = 0;

            if (_strMovFat == "F" && (_strCauTpd == "FV" ||_strCauTpd == "PR"))
            {
                //if (strTip == "PRE")
                //    d = (decimal)rowMov["mov_prv"];
                //else
                //    d = (decimal)rowMov["mov_imp"];
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

            if (!DBNull.Value.Equals(rowMov["mov_sco"]) && ((string)rowMov["mov_sco"]).Trim() != "")
            {
                string s = ((string)rowMov["mov_sco"]).Trim();

                if (s.Substring(0, 1) == "v")
                {
                    s = "V" + s.Substring(1);
                    rowMov["mov_sco"] = s;
                }

                string sTip = "%";
                if (s.Substring(0, 1).ToUpper() == "V")
                {
                    sTip = s.Substring(0, 1).ToUpper();
                    s = s.Substring(1);
                }


                //Lo sconto va sull'importo
                string[] a = s.Split('+');
                foreach (string v in a)
                {
                    if (_clsFun.Numerico(v))
                    {
                        dSco = Convert.ToDecimal(v.Replace(".", ","));

                        decimal dImp = d;
                        if(!DBNull.Value.Equals(rowMov["mov_qta"]))
                            dImp = d *(decimal)rowMov["mov_qta"];


                        if (sTip == "%")
                        {
                            dSco = _clsFun.MenoPer(dImp, dSco);
                            dSco = dImp - dSco;
                        }
                        //else if (sTip == "V")
                        //    d = d - dSco;
                    }
                }
            }

            //if (strCauTpd == "FA")
            //{
            //    if ((string)rowMov["mov_umi"] == "KG")
            //        rowMov["mov_imp"] = d * (decimal)rowMov["mov_qkg"];
            //    else
            //        rowMov["mov_imp"] = d * (decimal)rowMov["mov_qta"];
            //}
            //else
            //{
            //    if ((string)rowMov["mov_umi"] == "KG")
            //        rowMov["mov_imp"] = d * (decimal)rowMov["mov_qkg"];
            //    else
            //        rowMov["mov_imp"] = d * (decimal)rowMov["mov_qta"];
            //}

            if (strTip == "PRE")
            {
                if ((string)rowMov["mov_umi"] == "KG" && d > 0 && !DBNull.Value.Equals(rowMov["mov_qkg"]) && (decimal)rowMov["mov_qkg"] > 0) 
                    rowMov["mov_imp"] = d * (decimal)rowMov["mov_qkg"];
                else if(d > 0 && !DBNull.Value.Equals(rowMov["mov_qta"]) && (decimal)rowMov["mov_qta"] > 0)
                    rowMov["mov_imp"] = d * (decimal)rowMov["mov_qta"];

                rowMov["mov_imp"] = (decimal)rowMov["mov_imp"] - dSco;

            }

            else if (strTip == "QTA" && (string)rowMov["mov_umi"] == "NR" && (decimal)rowMov["mov_qta"] > 0)   
            {
                rowMov["mov_imp"] = d * (decimal)rowMov["mov_qta"]; //Seck 20180226 per il calcolo dell'importo dal peso

                rowMov["mov_imp"] = (decimal)rowMov["mov_imp"] - dSco;
                rowMov[sFld] = Math.Round((decimal)rowMov[sFld], 3, MidpointRounding.AwayFromZero);
            }
            else
            {
                if ((string)rowMov["mov_umi"] == "KG" && !DBNull.Value.Equals(rowMov["mov_imp"]) && (decimal)rowMov["mov_imp"] > 0 && !DBNull.Value.Equals(rowMov["mov_qkg"]) && (decimal)rowMov["mov_qkg"] > 0)
                    rowMov[sFld] = (decimal)rowMov["mov_imp"] / (decimal)rowMov["mov_qkg"];
                else if (!DBNull.Value.Equals(rowMov["mov_imp"]) && (decimal)rowMov["mov_imp"] > 0 && !DBNull.Value.Equals(rowMov["mov_qta"]) && (decimal)rowMov["mov_qta"] > 0)
                    rowMov[sFld] = (decimal)rowMov["mov_imp"] / (decimal)rowMov["mov_qta"];

                rowMov["mov_imp"] = (decimal)rowMov["mov_imp"] - dSco;
                rowMov[sFld] = Math.Round((decimal)rowMov[sFld], 3, MidpointRounding.AwayFromZero);
            }

            rowMov["MovMdy"] = "S";

            return rowMov;
        }


        private void FillMovRow(DataTable tabTmp, DataRow x)
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

            if (_strCauTpd == "FA")
            {
                if((string)x["mov_umi"] == "KG")
                    x["mov_imp"] = (decimal)x["mov_cos"] * (decimal)x["mov_qkg"];
                else
                    x["mov_imp"] = (decimal)x["mov_cos"] * (decimal)x["mov_qta"];
            }
            else
            {
                decimal dIva = 22;
                //DataRow[] j = ((DataTable)cmbMovIva.DataSource).Select("tab_cod='" + (string)x["mov_iva"] + "'");
                DataRow[] j = (_dasGen.Tables[TABTABIVA]).Select("tab_cod='" + (string)x["mov_iva"] + "'");
                if (j.Length > 0)
                    dIva = Convert.ToDecimal(j[0]["tab_ali"]);
                else
                {
                    s = _clsFun.ParGet(clsDefine.enuParametri.Par012CodIVAxDefault, _strConSql);
                    if(s != "")
                    {
                        j = (_dasGen.Tables[TABTABIVA]).Select("tab_cod='" + s + "'");
                        if (j.Length > 0)
                            dIva = Convert.ToDecimal(j[0]["tab_ali"]);
                    }
                }

                //decimal d = _clsFun.MenoIva((decimal)x["mov_prv"], dIva);
                decimal d = (decimal)x["mov_prv"]; 
                x["mov_prv"] = d;
                if ((string)x["mov_umi"] == "KG")
                    x["mov_imp"] = d * (decimal)x["mov_qkg"];
                else
                {
                    x["mov_imp"] = d * (decimal)x["mov_qta"];
                }
            }

            x["MovMdy"] = "S";

            CtrlNumRiga(x);

            //int i = dgv1.Rows.Count - 1;
            //if (i >= 0)
            //{
            //    dgv1.Rows[i].Selected = true;
            //    dgv1.CurrentCell = dgv1[3, i];
            //}

            //if (i >= 0 && dgv1.Rows[i].Cells[0].Value.ToString() != "" && dgv1.Rows.Count - 1 == i)
            //{
            //    NewRiga(true);
            //}

            dgv2.DataSource = FillIva("");
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

            if (b)
            {
                DataTable tLia = new DataTable();

                DataTable t = new DataTable();
                ArrayList aWhe = new ArrayList();
                ArrayList aExl = new ArrayList();

                if (_strMovFat == "F" && _strCauTpd == "FA")
                {
                    s = "SELECT * FROM GesLisAcquisto ";
                    s += "WHERE lia_for='" + cmbMftCfo.SelectedValue.ToString() + "' AND lia_tip='F' ";
                    s += "ORDER BY lia_arf, lia_for, lia_dti DESC";
                    tLia = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);
                    DataColumn[] keys = new DataColumn[5];
                    keys[0] = tLia.Columns["lia_tip"];
                    keys[2] = tLia.Columns["lia_for"];
                    keys[1] = tLia.Columns["lia_arf"];
                    keys[3] = tLia.Columns["lia_dti"];
                    keys[4] = tLia.Columns["lia_dtf"];
                    tLia.PrimaryKey = keys;

                    tLia = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);
                }

                if (lblMftNum.Text == _clsDef.CODNEW)
                {
                    if (_strMovFat == "F")
                        lblMftNum.Text = _clsFun.NewNum(DateTime.Today.Year.ToString("0000"), clsDefine.enuNumeratori.NumGesMovFatture, 6, _strConSql);
                    else
                        lblMftNum.Text = _clsFun.NewNum(DateTime.Today.Year.ToString("0000"), clsDefine.enuNumeratori.NumGesMovimenti, 6, _strConSql);
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
                    y["fat_tpg"] = cmbMftTpg.SelectedValue.ToString();
                    y["fat_ubi"] = _strMftUbi;

                    y["fat_neg"] = "";
                    if (cmbMftNeg.SelectedIndex >= 0)
                        y["fat_neg"] = cmbMftNeg.SelectedValue.ToString();

                    if (cmbMftSta.SelectedValue != null)
                        y["fat_sta"] = cmbMftSta.SelectedValue.ToString();
                    y["fat_des"] = "";
                    if (cmbMftDes.SelectedValue != null)
                        y["fat_des"] = cmbMftDes.SelectedValue.ToString();

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
                    }

                    if (t.Rows.Count > 0)
                    {
                        //X doc creati nell'anno precedente e stampati nell'anno in corso

                        if (txtMftNdo.Text != "" && lblMftYea.Text != dtpMftDdt.Value.Year.ToString())
                        {
                            s = "UPDATE GesFatTestate SET fat_yfa='" + dtpMftDdt.Value.Year.ToString() + "' WHERE ";
                            s += "fat_idx=" + Convert.ToString(t.Rows[0]["fat_idx"]);
                            _clsFun.SqlWrite(s, _strConSql);
                        }
                    }
                }
                else
                {
                    s = "SELECT * FROM " + TABGESMOT + " WHERE mot_ymo='" + lblMftYea.Text + "' AND  mot_nmo='" + lblMftNum.Text + "' AND mot_ubi='" + _strMftUbi + "'";
                    t = _clsFun.FillTabSql(TABGESMOT, s, false, _strConSql);

                    DataRow y = t.NewRow();
                    y["mot_ymo"] = lblMftYea.Text;
                    y["mot_nmo"] = lblMftNum.Text;
                    y["mot_day"] = DateTime.Today;
                    y["mot_cfo"] = ""; // cmbMftCfo.SelectedValue.ToString();
                    if(_strCfo != "")
                        y["mot_cfo"] = cmbMftCfo.SelectedValue.ToString();
                    //y["mot_cau"] = _strCauTpd;        //20180125 tipo documento modificabile 
                    y["mot_cau"] = cmbMftTdc.SelectedValue.ToString();
                    y["mot_ndo"] = txtMftNdo.Text;
                    y["mot_ddo"] = dtpMftDdt.Value;
                    y["mot_neg"] = cmbMftNeg.SelectedValue.ToString();
                    y["mot_no1"] = txtMftNo1.Text;
                    y["mot_sta"] = cmbMftSta.SelectedValue.ToString();
                    y["mot_ubi"] = _strMftUbi;

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
                    }
                }

                s = "SELECT * FROM GesMovimenti WHERE ";
                if (_strMovFat == "F")
                {
                    s += "mov_ubi = '" + _strMftUbi + "' AND ";
                    s += "mov_yfa = '" + lblMftYea.Text + "' AND ";
                    s += "mov_nfa = '" + lblMftNum.Text + "' "; 
                    s += "ORDER BY mov_rfa";
                }
                else
                {
                    s += "mov_ubi = '" + _strMftUbi + "' AND ";
                    s += "mov_ymo = '" + lblMftYea.Text + "' AND ";
                    s += "mov_nmo = '" + lblMftNum.Text + "' ";
                    s += "ORDER BY mov_rmo";
                }
                t = _clsFun.FillTabSql(TABGESMOV, s, false, _strConSql);
                DataRow[] j;

                aExl = new ArrayList();

                DataTable tTmp = ((DataView)dgv1.DataSource).ToTable();

                foreach (DataRow x in (((DataView)dgv1.DataSource).ToTable()).Rows)
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
                        }
                        if (s != "")
                        {
                            _clsFun.SqlWrite(s, _strConSql);
                            _clsFun.FileLog("GesMovimenti", (string)x["mov_art"], s);
                        }

                        if (_strMovFat == "F" && _strCauTpd == "FA")
                        {
                            aWhe = new ArrayList();
                            aWhe.Add("lia_tip");
                            aWhe.Add("lia_art");
                            aWhe.Add("lia_for");
                            aExl = new ArrayList();

                            j = tLia.Select("lia_tip='F' AND lia_art='" + x["mov_art"] + "'");

                            DataRow x1 = tLia.NewRow();
                            x1["lia_tip"] = "F";
                            x1["lia_art"] = x["mov_art"];
                            x1["lia_for"] = cmbMftCfo.SelectedValue.ToString();
                            x1["lia_dti"] = DateTime.Today;
                            x1["lia_dtf"] = _clsDef.DAYOUT;
                            x1["lia_arf"] = "";//x["mov_nfa"];
                            x1["lia_cos"] = x["mov_cos"];
                            x1["lia_pxc"] = 1;
                            x1["lia_cxp"] = 1;
                            x1["lia_day"] = DateTime.Today;
                            x1["lia_prv"] = x["mov_prv"]; ;

                            if((decimal)x["mov_cos"] > 0)
                            {
                                s = "";
                                if (j.Length > 0)
                                {
                                    if (!DBNull.Value.Equals(j[0]["lia_arf"]))
                                        x1["lia_arf"] = (string)j[0]["lia_arf"];

                                    if ((decimal)j[0]["lia_cos"] != (decimal)x["mov_cos"])
                                    {
                                        if (DateTime.Compare((DateTime)x1["lia_dti"],(DateTime)j[0]["lia_dti"]) == 0)
                                            s = _clsFun.SqlUpdRow(TABLISACQ, tLia, j[0], x1, aWhe, aExl);
                                        else
                                            s = _clsFun.SqlInsertRow(TABLISACQ, tLia, x1);
                                    }
                                }
                                else
                                    s = _clsFun.SqlInsertRow(TABLISACQ, tLia, x1);

                                if(s != "")
                                    _clsFun.SqlWrite(s, _strConSql);
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
                        }
                    }
                }

            }

            return b;
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource == null || ((DataView)dgv1.DataSource).Count == 0)
                MessageBox.Show("Non ci sono dati da stampare!");
            else
            {
                DataRow[] j;
                if (txtMftNdo.Text.Trim() == "")
                {
                    Console.WriteLine("aaaa");

                    if (_strMovFat == "F" && _strCauTpd == "FA")
                        Console.WriteLine("Non stampare");
                    else if (_strMovFat == "F" && _strCauTpd.Substring(0,1) == "F" && cmbMftTdc.SelectedValue.ToString() != "PR")
                        txtMftNdo.Text = _clsFun.NewNum(DateTime.Today.Year.ToString("0000"), clsDefine.enuNumeratori.NumGesDocFatture, 10, _strConSql);
                    if (_strMovFat == "M")
                    {
                        j = ((DataTable)cmbMftTdc.DataSource).Select("tab_cod='" + _strCauTpd + "'");
                        if(j.Length == 0)
                            MessageBox.Show("Tipo documento non definito correttamente!", "Stampa documento", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        else
                        {
                            if(!((string)j[0]["tab_tip"]).Contains("P"))
                                txtMftNdo.Text = _clsFun.NewNum(DateTime.Today.Year.ToString("0000"), clsDefine.enuNumeratori.NumGesDocDdt, 10, _strConSql);
                        }
                    }
                }

                if (Salva())
                {
                    string s = "";
                    //DataRow[] j = ((DataTable)cmbMftTdc.DataSource).Select("tab_cod='" + cmbMftTdc.SelectedValue + "'");
                    //if (j.Length == 0)
                    if(_strCfo == "")
                        MessageBox.Show("Tipo cliente/fornitore non definito!");
                    else
                    {
                        if (((DataTable)cmbMftDes.DataSource).Rows.Count > 1 && cmbMftDes.Text == "")
                            MessageBox.Show("Destinazione PV non selezionata!", "Stampa documento", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        if (_strMovFat == "F")
                            s = "FV";
                        else if (_strMovFat == "M")
                            s = "DT";
                        else
                            s = "XX";
                        s = "SELECT * FROM TabNote WHERE tab_cod='" + s + "' AND tab_ann=0 ORDER BY tab_row";

                        s = "SELECT * FROM TabNote WHERE tab_cod='FV' AND tab_ann=0 ORDER BY tab_row";
                        DataTable tNot = _clsFun.FillTabSql("TabNote", s, false, _strConSql);

                        DataTable tTes = _clsQry.DocCfo(_strCfo, cmbMftCfo.SelectedValue.ToString());
                        tTes.Rows[0]["tmp_ndo"] = txtMftNdo.Text;
                        if (txtMftNdo.Text.Trim() == "")
                            tTes.Rows[0]["tmp_ndo"] = lblMftNum.Text;

                        tTes.Rows[0]["tmp_ddo"] = dtpMftDdt.Value;
                        tTes.Rows[0]["tmp_tpg"] = cmbMftTpg.Text;

                        DataTable tMov = new clsGenTabTmp().TabTmpDocMov("DocMov");

                        s = "Totali";
                        decimal dIva = 0;
                        decimal dImp = 0;
                        DataTable tIva = (DataTable)dgv2.DataSource;
                        foreach (DataRow y in tIva.Rows)
                        {
                            if ((string)y["iva_des"] != s)
                            {
                                dIva += (decimal)y["iva_iva"];
                                dImp += (decimal)y["iva_imp"];
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

                        DataSet dasGen = new DataSet();
                        dasGen.Tables.Add(tTes);
                        dasGen.Tables.Add(((DataView)dgv1.DataSource).ToTable());
                        dasGen.Tables.Add(((DataTable)dgv2.DataSource).Copy());
                        dasGen.Tables.Add(tNot);

                        string sDocTip = "";
                        j = ((DataTable)cmbMftTdc.DataSource).Select("tab_cod='" + _strCauTpd + "'");
                        if(j.Length > 0)
                            sDocTip = (string)j[0]["tab_tip"];

                        frmGesDocPrint f = new frmGesDocPrint();
                        f._dasGen = dasGen;
                        f._strMovFat = _strMovFat;
                        f._strCauTpd = _strCauTpd;
                        f._strSuf = _strDocSuf;
                        f._strTipCod = cmbMftTdc.Text;
                        f._strQta = lblTotQta.Text;
                        f._strDes = cmbMftDes.Text;
                        f._strPrnPeso = "S";
                        f._strDocTip = sDocTip;
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

            //DataTable tTabIva = (DataTable)cmbMovIva.DataSource;
            DataTable tTabIva = _dasGen.Tables[TABTABIVA];

            t = ((DataView)dgv1.DataSource).Table;

            DataRow x;
            DataRow[] j;
            int iRig = 0;
            int iPos = 0;
            decimal d = 0;
            int iRowLast = 0;
            string sIva = "";

            //if(t.Rows.Count > 0)
            //{
            //    if (((string)t.Rows[t.Rows.Count - 1]["mov_nmo"]).Trim() != "" || ((string)t.Rows[t.Rows.Count - 1]["mov_nfa"]).Trim() != "")
            //    {
            //        NewRiga(false);
            //    }
            //}
            iPos = t.Rows.Count-1;
            iRig = t.Rows.Count;
            if (_strMovFat == "F")
                t.Rows[iPos]["mov_rfa"] = iRig.ToString("0000");
            else
                t.Rows[iPos]["mov_rmo"] = iRig.ToString("0000");

            s = "Scontrino N. " + (string)tabSco.Rows[0]["vet_sco"] + " del " + ((DateTime)tabSco.Rows[0]["vet_day"]).ToShortDateString();
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

                j = tTabIva.Select("tab_cod='" + y["ven_iva"] + "'");
                if (j.Length > 0)
                    sIva = (string)j[0]["tab_cod"];

                dTot += (decimal)y["ven_ven"];

                x = t.NewRow();
                x["mov_yfa"] = "";
                x["mov_nfa"] = "";
                x["mov_ymo"] = "";
                x["mov_nmo"] = "";
                x["mov_rfa"] = ""; 
                x["mov_rmo"] = "";

                //if (_strMovFat == "F")
                //    x["mov_rfa"] = iRig.ToString("0000");
                //else
                //    x["mov_rmo"] = iRig.ToString("0000");

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
                x["mov_imp"] = y["ven_ven"];

                if ((decimal)x["mov_imp"] < 0)
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
                s += (string)y["ven_sco"];

                x["mov_ori"] = s;
                x["MovMdy"] = "S";

                decimal dIva = 22;
                j = (_dasGen.Tables[TABTABIVA]).Select("tab_cod='" + (string)x["mov_iva"] + "'");
                if (j.Length > 0)
                    dIva = Convert.ToDecimal(j[0]["tab_ali"]);
                //else
                //{
                //    //s = _clsFun.ParGet(clsDefine.enuParametri.Par012CodIVAxDefault, _strConSql);
                //    if (s != "")
                //    {
                //        j = ((DataTable)cmbMovIva.DataSource).Select("tab_cod='" + _strIvaPar + "'");
                //        if (j.Length > 0)
                //            dIva = Convert.ToDecimal(j[0]["tab_ali"]);
                //    }
                //}

                d = _clsFun.MenoIva((decimal)x["mov_imp"], dIva);

                if ((string)x["mov_umi"] == "KG")
                {
                    //x["mov_prv"] = Math.Round(d, 2) / (decimal)x["mov_qkg"];
                    if (d != 0 && (decimal)x["mov_qkg"] != 0)
                        x["mov_prv"] = d / (decimal)x["mov_qkg"];
                }
                else
                {
                    //x["mov_prv"] = Math.Round(d, 2) / (decimal)x["mov_qta"];
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

            //NewRiga(false);

            //Tentativo di arrotondamento per far corrispondere l'IVA

            if (true)
            {

                dTotoDoc = dTotoDoc + dScoTot;

                decimal dPrv = (decimal)t.Rows[iRowLast]["mov_prv"];   //  3.735
                decimal dImp = (decimal)t.Rows[iRowLast]["mov_imp"];   //  26.144
                decimal dQta = (decimal)t.Rows[iRowLast]["mov_qta"];   //  7

                int i = 0;

                s = (string)t.Rows[iRowLast]["mov_art"];


                while (true)
                {
                    i++;

                    //DataTable tIva = FillIva(sRigIni + "," + sRigFin);
                    DataTable tIva = FillIva("");

                    decimal dTotIva = 0;
                    foreach (DataRow y in tIva.Rows)
                        dTotIva += (decimal)y["iva_tot"];

                    dTotIva = Math.Round(dTotIva, 2);

                    Console.WriteLine("aaaa");

                    //if (dScoTot != dTotIva)
                    //{
                    //    d = dScoTot - dTotIva;

                    if (dTotoDoc != dTotIva)
                    {
                        d = dTotoDoc - dTotIva;

                        decimal dd = Convert.ToDecimal(0.001);
                        if (d < 0)
                            dd = dd * -1;

                        s = (string)t.Rows[iRowLast]["mov_art"];

                        //t.Rows[1]["mov_prv"] = (decimal)t.Rows[1]["mov_prv"] + dd;
                        //t.Rows[1]["mov_imp"] = (decimal)t.Rows[1]["mov_prv"] * (decimal)t.Rows[1]["mov_qta"];

                        t.Rows[iRowLast]["mov_imp"] = (decimal)t.Rows[iRowLast]["mov_imp"] + dd;
                        t.Rows[iRowLast]["mov_prv"] = (decimal)t.Rows[iRowLast]["mov_imp"] / (decimal)t.Rows[iRowLast]["mov_qta"];
                        t.Rows[iRowLast]["mov_prv"] = Math.Round((decimal)t.Rows[iRowLast]["mov_prv"], 3, MidpointRounding.ToEven);

                        Console.WriteLine("aaaa");

                    }
                    else
                        break;

                    if (i > 40)
                        break;
                }

            }

            //Console.WriteLine("aaaa");

            ////string sFldRig = "mov_rmo";
            ////if (_strMovFat == "F")
            ////    sFldRig = "mov_rfa";
            ////foreach(DataRow y in t.Rows)
            ////{
            ////    if ((string)y["mov_ard"] != "" && !(Boolean)y["mov_ann"] && Convert.ToInt16(y[sFldRig]) >= Convert.ToInt16(sRigIni) && Convert.ToInt16(y[sFldRig]) <= Convert.ToInt16(sRigFin))
            ////    {
            ////        decimal dd = Convert.ToDecimal(0.01);
            ////        if (d < 0)
            ////            dd = dd * -1;
            ////        y["mov_imp"] = (decimal)y["mov_imp"] + dd;
            ////        y["mov_prv"] = (decimal)y["mov_imp"] / (decimal)y["mov_qta"];
            ////        d -= dd;
            ////    }
            ////    if (d == 0)
            ////        break;
            ////}

            ////decimal dd = Math.Round( d / (decimal)t.Rows[1]["mov_qta"], 3,MidpointRounding.AwayFromZero);
            //decimal dPrv = (decimal)t.Rows[1]["mov_prv"];
            //decimal dImp = (decimal)t.Rows[1]["mov_imp"];
            //decimal dQta = (decimal)t.Rows[1]["mov_qta"];

            //Console.WriteLine("aaaa");

            //decimal dIva = 1 + (Convert.ToDecimal(t.Rows[1]["mov_iva"])/100);

            ////decimal dd1 = Math.Round((dImp * dIva), 3, MidpointRounding.AwayFromZero);

            //decimal dd1 = dImp * dIva;

            //decimal dImpNew = (dImp * dIva) + d;

            ////decimal ddd = dd2 / dIva / dQta; 

            //Console.WriteLine("aaaa");

            //Boolean b = true;

            //while (b)
            //{
            //    decimal dd = Convert.ToDecimal(0.001);
            //    if (d < 0)
            //        dd = dd * -1;

            //    dPrv += dd;

            //    decimal v1 = Math.Round(dPrv * dQta * dIva, 3, MidpointRounding.AwayFromZero);


            //    decimal ddd = v1 - dImpNew;

            //    Console.WriteLine("aaaa");

            //    if (Math.Abs(ddd) < Convert.ToDecimal(0.01))
            //        break;

            //}

            //Console.WriteLine("aaaa");

            //t.Rows[1]["mov_prv"] = dPrv;
            //t.Rows[1]["mov_imp"] = dPrv * (decimal)t.Rows[1]["mov_qta"];

            ////decimal v1 = Math.Round(dPrv * dQta * dIva, 3, MidpointRounding.AwayFromZero);
            ////decimal v2 = Math.Round(ddd * dQta * dIva, 3, MidpointRounding.AwayFromZero);

            ////t.Rows[1]["mov_prv"] = (decimal)t.Rows[1]["mov_prv"] + dd;
            ////t.Rows[1]["mov_imp"] = (decimal)t.Rows[1]["mov_prv"] * (decimal)t.Rows[1]["mov_qta"];

            ////t.Rows[1]["mov_imp"] = (decimal)t.Rows[1]["mov_imp"] + d;
            ////t.Rows[1]["mov_prv"] = (decimal)t.Rows[1]["mov_imp"] / (decimal)t.Rows[1]["mov_qta"];

            ////t.Rows[1]["mov_prv"] = Math.Round((decimal)t.Rows[1]["mov_prv"], 3, MidpointRounding.AwayFromZero);

            ////t.Rows[1]["mov_prv"] = Math.Round((decimal)t.Rows[1]["mov_prv"], 3, MidpointRounding.AwayFromZero);

            ////Boolean b = true;

            ////decimal dIva = 1 + (Convert.ToDecimal(t.Rows[1]["mov_iva"])/100);

            ////while (b)
            ////{
            ////    decimal dd = Convert.ToDecimal(0.001);
            ////    if(d < 0)
            ////        dd = dd *-1;

            ////    dPrv += dd;

            ////    decimal v1 = Math.Round( dPrv * dQta * dIva, 3, MidpointRounding.AwayFromZero);

            ////    decimal v2 = Math.Round((dImp + d) * dIva, 3, MidpointRounding.AwayFromZero);

            ////    if (v1 >= v2)
            ////        break;
            ////}

            ////Console.WriteLine("aaaa");

            ////t.Rows[1]["mov_prv"] = dPrv;
            ////t.Rows[1]["mov_imp"] = dPrv / (decimal)t.Rows[1]["mov_qta"]; 

            //}

            dgv2.DataSource = FillIva("");
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

        private void eliminaRigheToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (txtMftNdo.Text != "")
                MessageBox.Show("Documento già stampato!");
            else
            {
                if (MessageBox.Show("Confermi l'eliminazione delle righe del documento?", "CANCELLAZIONE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string s = "DELETE FROM GesMovimenti WHERE ";
                    if (_strMovFat == "M")
                        s += "mov_ymo='" + lblMftYea.Text + "' AND mov_nmo='" + lblMftNum.Text + "'";
                    else
                        s += "mov_yfa='" + lblMftYea.Text + "' AND mov_nfa='" + lblMftNum.Text + "'";

                    _clsFun.SqlWrite(s, _strConSql);

                    FillDati();
                }
            }
        }

        private void importDaDocumentiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmSeekDocs f = new frmSeekDocs();
            f._strImpDoc = "S";
            f.ShowDialog();
            string s = f._strImpDoc;
            if(s != "S")
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

                    //if (t.Rows.Count > 0)
                    //{
                    //    if (((string)t.Rows[t.Rows.Count - 1]["mov_nmo"]).Trim() != "" || ((string)t.Rows[t.Rows.Count - 1]["mov_nfa"]).Trim() != "")
                    //    {
                    //        NewRiga(false);
                    //    }
                    //}
                    
                    //if (((string)t.Rows[t.Rows.Count - 1]["mov_nmo"]).Trim() == "" && ((string)t.Rows[t.Rows.Count - 1]["mov_nfa"]).Trim() == "")
                    //    t.Rows[0].Delete();

                    if (t.Rows.Count > 0)
                    {
                        if (((string)t.Rows[t.Rows.Count - 1]["mov_nmo"]).Trim() == "" && ((string)t.Rows[t.Rows.Count - 1]["mov_nfa"]).Trim() == "")
                            t.Rows[0].Delete();
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
                    }
                    iRig = iPos +1;

                    foreach (DataRow y in tMov.Rows)
                    {
                        if (!(Boolean)y["mov_ann"])
                        {
                            iRig++;

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
                        }
                    }

                    if(_strMovFat == "F") 
                        sFldNri = "mov_rfa";
                    else
                        sFldNri = "mov_rmo";

                    DataView v = new DataView(t, "", sFldNri, DataViewRowState.CurrentRows);
                    dgv1.DataSource = v;
                    //NewRiga(false);

                    dgv2.DataSource = FillIva("");
                }
            }
        }

        private void eliminaRigheCancellateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Eliminazione righe cancellate, confermi?", "CANCELLAZIONE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Salva();

                string s = "DELETE FROM GesMovimenti WHERE ";
                if (_strMovFat == "M")
                    s += "mov_ymo='" + lblMftYea.Text + "' AND mov_nmo='" + lblMftNum.Text + "'";
                else
                    s += "mov_yfa='" + lblMftYea.Text + "' AND mov_nfa='" + lblMftNum.Text + "'";
                //s += " AND mov_ann=1";

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

        private void btnArtNew_Click(object sender, EventArgs e)
        {
            //if (_strCfo == _clsDef.TIPCLI)
            //{
            if (_strCfo != "" && (cmbMftCfo.SelectedValue == null || cmbMftCfo.SelectedValue.ToString() == ""))
                MessageBox.Show("Cliente destinatario non definto!", "NUOVO ARTICOLO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                string sCfo = "";
                string sCfd = "";
                string sInd = "";
                string sLic = "";
                string sLir = "";
                string sCpa = "";

                if (_strCfo != "")
                {
                    DataRow[] j = ((DataTable)cmbMftCfo.DataSource).Select("CfoCod='" + cmbMftCfo.SelectedValue.ToString() + "'");
                    if (j.Length > 0)
                    {
                        sCfo = cmbMftCfo.SelectedValue.ToString();
                        sCfd = cmbMftCfo.Text;
                        sInd = (string)j[0]["CfoInd"];
                        sLir = (string)j[0]["CfoLir"];
                        sLic = (string)j[0]["CfoLic"];
                        sCpa = (string)j[0]["CfoEtc"];
                    }
                }

                frmGesDocVenPesato f = new frmGesDocVenPesato();
                f._strNegLic = sLic;
                f._strNegLir = sLir;
                f._strConSql = _strConSql;
                f._strCli = sCfo; // cmbMftCfo.SelectedValue.ToString();
                f._strCld = sCfd; // cmbMftCfo.Text;
                f._strInd = sInd;
                f._strCpa = sCpa;
                f._strOpenPorta = _strOpenPorta;
                f._serialPort = _serialPort;
                f._strBilTime = _strBilTime;
                f._tabTgr = _dasGen.Tables[TABTABTGR];
                f.ShowDialog();

                if (f._tabRow.Rows.Count > 0)
                {
                    string sTipEti = f._strTipEti;

                    foreach(DataRow y in f._tabRow.Rows)
                    {
                        string sArt = (string)y["tmp_art"];
                        string sUmi = (string)y["tmp_umi"];
                        string sPxc = (string)y["tmp_pxc"];
                        string sLot = (string)y["tmp_lot"];
                        decimal dTar = (decimal)y["tmp_tar"];
                        decimal dQkg = (decimal)y["tmp_qkg"];
                        decimal dPrv = (decimal)y["tmp_pve"];
                        decimal dPrp = (decimal)y["tmp_prp"];
                        int iGsc = Convert.ToInt16(y["tmp_gsc"]);

                        NewRiga(true);

                        CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                        if (cm.Position >= 0)
                        {
                            DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                            DataRow x = r.Row;

                            if(sUmi == "NR")
                                x["mov_qta"] = dQkg * 1000;
                            else
                                x["mov_qkg"] = dQkg;

                            x["mov_lot"] = sLot;
                            x["mov_tar"] = dTar;

                            DataTable tArt = _clsQry.SeekArtEan(f._strArt, "LNE" + sLic);

                            DataTable t = _clsQry.ArtSeek(sArt, "SEEK");
                            /* da prendere dalla pesata */
                            //t.Rows[0]["tmp_prv"] = (decimal)tArt.Rows[0]["tmp_prv"];
                            t.Rows[0]["tmp_prv"] = dPrv;
                            //dPce = (decimal)t.Rows[0]["tmp_prv"];

                            FillMovRow(t, x);
                        }
                        if (f._tabRow.Rows.Count == 1 && f._strPrnLabelAuto == "S")
                        {
                            EtcSingolo(f._strCpa, dPrp, iGsc, sTipEti);
                        }
                        if(_strCfo != "" && dPrv == 0)
                            MessageBox.Show("Prezzo di cessione non definito!", "VALORI A ZERO", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }
                }
            }
            //}
            //else
            //    MessageBox.Show("Funzione prevista solo per clienti!", "NUOVO ARTICOLO", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void btnArtMod_Click(object sender, EventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                if (DBNull.Value.Equals(x["mov_lot"]))
                    x["mov_lot"] = "";

                string sCfo = "";
                string sCfd = "";
                string sLir = "";
                string sCpa = "";
                string sInd = "";

                if (_strCfo != "")
                {
                    DataRow[] j = ((DataTable)cmbMftCfo.DataSource).Select("CfoCod='" + cmbMftCfo.SelectedValue.ToString() + "'");
                    if (j.Length > 0)
                    {
                        sCfo = cmbMftCfo.SelectedValue.ToString();
                        sCfd = cmbMftCfo.Text;

                        sLir = (string)j[0]["CfoLir"];
                        sInd = "";
                        if (!DBNull.Value.Equals(j[0]["CfoInd"]))
                            sInd += ((string)j[0]["CfoInd"]).Trim();
                        if (!DBNull.Value.Equals(j[0]["CfoCap"]))
                            sInd += " - " + ((string)j[0]["CfoCap"]).Trim();
                        if (!DBNull.Value.Equals(j[0]["CfoLoc"]))
                            sInd += " " + ((string)j[0]["CfoLoc"]).Trim();
                        if (!DBNull.Value.Equals(j[0]["CfoPrv"]))
                            sInd += " (" + ((string)j[0]["CfoPrv"]).Trim() + ")";

                        sCpa = (string)j[0]["CfoEtc"];
                    }
                }

                DataTable t = _clsQry.SeekArtEan((string)x["mov_art"], "LNE" + sLir);

                //t.Rows[0]["tmp_prv"] = (decimal)x["mov_prv"];

                decimal dTar = 0;
                if(!DBNull.Value.Equals(x["mov_tar"]))
                    dTar = ((decimal)x["mov_tar"]) / 1000;

                int iGio = 0;
                if (!DBNull.Value.Equals(t.Rows[0]["tmp_gsc"]) && _clsFun.Numerico((string)t.Rows[0]["tmp_gsc"]) && Convert.ToInt16(t.Rows[0]["tmp_gsc"]) > 0)
                    iGio = Convert.ToInt16(t.Rows[0]["tmp_gsc"]);

                frmGesDocVenPesato f = new frmGesDocVenPesato();
                f._strConSql = _strConSql;
                f._strCli = sCfo; // cmbMftCfo.SelectedValue.ToString();
                f._strCld = sCfd; // cmbMftCfo.Text;
                f._strInd = sInd;
                f._strCpa = sCpa;
                f._strNegLir = sLir;

                f._decTar = (decimal)x["mov_tar"];
                f._strUmi = (string)x["mov_umi"];

                f._decPes = (decimal)x["mov_qKg"] + dTar;
                if ((string)x["mov_umi"] == "NR")
                    f._decPes = (decimal)x["mov_qta"];

                f._strLot = (string)x["mov_lot"];
                f._tabArt = t.Copy();
                f._intScaGio = iGio;
                f._strOpenPorta = _strOpenPorta;
                f._serialPort = _serialPort;
                f._strBilTime = _strBilTime;
                f._tabTgr = _dasGen.Tables[TABTABTGR];

                f.ShowDialog();

                if (f._strArt != "" && f._decPes > 0 && (_strCfo == "" || f._decPve > 0))
                {
                    x["mov_lot"] = f._strLot;
                    x["mov_tar"] = f._decTar;

                    if(f._strUmi == "NR")
                        x["mov_qta"] = f._decPes * 1000;
                    else
                        x["mov_qkg"] = f._decPes;

                    t = _clsQry.ArtSeek(f._strArt, "SEEK");
                    t.Rows[0]["tmp_prv"] = (decimal)x["mov_prv"];

                    FillMovRow(t, x);
                }
            }
        }

        private void btnRowDbl_Click(object sender, EventArgs e)
        {
            string sFldRow = "mov_rmo";
            if (_strMovFat == "F")
                sFldRow = "mov_rfa";

            DataView v = (DataView)dgv1.DataSource;

            if (v.Count == 0)
                MessageBox.Show("NON CI SONO RIGHE DA DUPLICARE!", "DUPLICA RIGA", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                DataView vv = new DataView(v.ToTable(), "", sFldRow + " DESC", DataViewRowState.CurrentRows);

                string sRig = (Convert.ToInt16(vv[0][sFldRow]) + 1).ToString("0000");

                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;

                    string s = "";

                    DataTable t = v.Table;

                    DataRow x = t.NewRow();
                    foreach (DataColumn c in t.Columns)
                    {
                        x[c.ColumnName] = r[c.ColumnName];
                    }
                    x[sFldRow] = sRig;
                    x["MovMdy"] = "S";
                    t.Rows.Add(x);

                    PosRowLast();
                }

                dgv2.DataSource = FillIva("");
            }
        }

        private void btnEtiIng_Click(object sender, EventArgs e)
        {
            EtcSingolo("", 0, 0, "");
        }

        private void EtcSingolo(string strCpa, decimal decPrv, int intGsc, string strTipEti)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {                
                string s = "";
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;
                DataRow[] j;

                string sCli = "";
                string sCld = "";
                if (_strCfo != "")
                {
                    sCli = cmbMftCfo.SelectedValue.ToString(); ;
                    sCld = cmbMftCfo.Text;
                }

                string sInd = "";

                string sArt = (string)x["mov_art"];
                string sArd = (string)x["mov_ard"];
                string sUmi = (string)x["mov_umi"];
                //string sPxc = (string)x["mov_pxc"];
                decimal dQkg = (decimal)x["mov_qkg"];
                decimal dQta = (decimal)x["mov_qta"];
                //decimal dTar = (decimal)x["mov_tar"];

                //decimal dPrv = 0; // (decimal)x["mov_prv"];

                decimal dImp = (decimal)x["mov_imp"];
                string sLot = "";
                if (DBNull.Value.Equals(x["mov_lot"]))
                    x["mov_lot"] = "";

                sLot = (string)x["mov_lot"];

                string sLir = "";
                string sCpa = "";       //Parametri del cliente (flag)

                if (_strCfo != "")
                {
                    j = ((DataTable)cmbMftCfo.DataSource).Select("CfoCod='" + cmbMftCfo.SelectedValue.ToString() + "'");
                    if (j.Length > 0)
                    {
                        sLir = (string)j[0]["CfoLir"];  //Listino di rivendita

                        sInd = "";
                        if (!DBNull.Value.Equals(j[0]["CfoInd"]))
                            sInd += ((string)j[0]["CfoInd"]).Trim();
                        if (!DBNull.Value.Equals(j[0]["CfoCap"]))
                            sInd += " - " + ((string)j[0]["CfoCap"]).Trim();
                        if (!DBNull.Value.Equals(j[0]["CfoLoc"]))
                            sInd += " " + ((string)j[0]["CfoLoc"]).Trim();
                        if (!DBNull.Value.Equals(j[0]["CfoPrv"]))
                            sInd += " (" + ((string)j[0]["CfoPrv"]).Trim() + ")";

                        sCpa = (string)j[0]["CfoEtc"];

                    }
                }

                if(sLir != "")
                    sLir = "LNE" + sLir;

                if (strCpa != "")
                    sCpa = strCpa;

                DataTable tArt = _clsQry.SeekArtEan((string)x["mov_art"], sLir);

                if (tArt.Rows.Count > 0)
                {
                    string sPxc = "0";
                    if (!DBNull.Value.Equals(tArt.Rows[0]["tmp_pxc"]))
                        sPxc = Convert.ToString(tArt.Rows[0]["tmp_pxc"]);

                    if(sUmi != "KG")
                    {
                        dImp = decPrv;
                        dQkg = 0;
                        decimal dPne = 0;
                        decimal dTgv = 0;
                        if(!DBNull.Value.Equals(tArt.Rows[0]["tmp_pne"]))
                            dPne = (decimal)tArt.Rows[0]["tmp_pne"];

                        j = (_dasGen.Tables[TABTABTGR]).Select("tab_cod='" + (string)tArt.Rows[0]["tmp_tgr"] + "'");
                        if (j.Length > 0 && !DBNull.Value.Equals(j[0]["tab_val"]))
                            dTgv = (decimal)j[0]["tab_val"];

                        if (dTgv == 0)
                            MessageBox.Show("Tipo grammatura o contenuto non corretti sull'anagrafica articolo!", "CONTROLLO ARTICOLO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                        if (dPne > 0 && dTgv > 0)
                            decPrv = dImp / dPne / dTgv;
                        dQkg = dPne;

                    }
                    else
                        dImp = decPrv * dQkg;


                    string sGsc = (string)tArt.Rows[0]["tmp_gsc"];

                    if (intGsc > 0)
                        sGsc = intGsc.ToString();

                    if (!_clsFun.Numerico(sGsc, "0123456789") || Convert.ToInt16(sGsc) == 0)
                        MessageBox.Show("Giorni conservazione in anagrafica articoli mancanti!", "CONTROLLO ARTICOLO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    else
                    {
                        int iQta = 1;
                        if (sUmi == "NR")
                            iQta = Convert.ToInt16(dQta);

                        string sPrn = "";
                        string sTip = "";
                        string[] a = _strIni11VenditaTouch.Split(';');
                        if (a.Length > 4)
                            sPrn = a[4];
                        if (a.Length > 5)
                            sTip = a[5];                        
                        
                        for (int i = 0; i < iQta; i++)
                        {
                            if (_strCfo == "")
                            {
                                Boolean b = true;
                                string sEan = ""; //_clsQry.GenLotto2Ean(sLot, sArt, Convert.ToInt64(dQkg * 1000).ToString("00000"));

                                //if (sEan == "" && MessageBox.Show("Barcode non trovato, continui con la stampa dell'etichetta?", "STAMPA ETICHETTA", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) 
                                //    b = false;

                                if (b)
                                {

                                    frmGesMovIngLabel05 f = new frmGesMovIngLabel05();
                                    f._strEtiTipo = sTip;
                                    f._strConSql = _strConSql;
                                    f._strCliCod = sCli;
                                    f._strCliDes = sCld;
                                    f._strCliInd = sInd;
                                    f._strCliPar = sCpa;
                                    f._strArtCod = sArt;
                                    f._strEanCod = sEan;
                                    f._strArtDes = sArd;
                                    f._strArtUmi = sUmi;
                                    f._decMovQkg = dQkg;
                                    f._decMovPrv = decPrv;
                                    f._decMovImp = dImp;
                                    f._dayArtSca = DateTime.Today.AddDays(Convert.ToInt16(sGsc));
                                    f._strLotCod = sLot;
                                    f.ShowDialog();
                                }
                            }
                            else if (strTipEti == "EAN128")
                            {
                                frmGesMovIngLabel03 f = new frmGesMovIngLabel03();
                                f._strEtiTipo = sTip;
                                f._strConSql = _strConSql;
                                f._strCliCod = sCli;
                                f._strCliDes = sCld;
                                f._strCliInd = sInd;
                                f._strCliPar = sCpa;
                                f._strArtCod = sArt;
                                f._strArtDes = sArd;
                                f._strArtUmi = sUmi;
                                //f._strArtUmi = _strUmi;
                                //f._decMovQkg = dPes;
                                //f._decMovPrv = dPrv;
                                //f._decMovImp = dImp;
                                f._dayDaySca = DateTime.Today.AddDays(Convert.ToInt16(sGsc));
                                f._strLotCod = sLot;            // txtLot.Text;
                                f._strArtPxc = sPxc;            //Pezzi per confezione
                                f.ShowDialog();
                            }
                            else if (sPrn != "")
                            {
                                clsGesMovIngLabel03 cls = new clsGesMovIngLabel03();
                                cls._strPrn = sPrn;
                                cls._strConSql = _strConSql;
                                cls._strCliDes = sCld;
                                cls._strCliInd = sInd;
                                cls._strCliPar = sCpa;
                                cls._strArtCod = sArt;
                                cls._strArtDes = sArd;
                                cls._strArtUmi = sUmi;
                                cls._decMovQkg = dQkg;
                                cls._decMovPrv = decPrv;
                                cls._decMovImp = dImp;
                                cls._dayArtSca = DateTime.Today.AddDays(Convert.ToInt16(sGsc));
                                cls._strLotCod = sLot;
                                cls.Stampa();
                            }
                            else
                            {
                                frmGesMovIngLabel01 f = new frmGesMovIngLabel01();
                                f._strEtiTipo = sTip;
                                f._strConSql = _strConSql;
                                f._strCliCod = sCli;
                                f._strCliDes = sCld;
                                f._strCliInd = sInd;
                                f._strCliPar = sCpa;
                                f._strArtCod = sArt;
                                f._strArtDes = sArd;
                                f._strArtUmi = sUmi;
                                f._decMovQkg = dQkg;
                                f._decMovPrv = decPrv;
                                f._decMovImp = dImp;
                                f._dayArtSca = DateTime.Today.AddDays(Convert.ToInt16(sGsc));
                                f._strLotCod = sLot;
                                f.ShowDialog();
                            }
                        }
                    }
                }

                if ((string)x["mov_lot"] != sLot)
                {
                    x["mov_lot"] = sLot;
                    x["MovMdy"] = "S";
                }
            }
        }

        private void btnEtiArt_Click(object sender, EventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                string sCli = cmbMftCfo.Text;
                string sArt = (string)x["mov_art"];
                string sArd = (string)x["mov_ard"];
                decimal dPrv = (decimal)x["mov_prv"];

                decimal dQta = 0;
                decimal dQkg = 0;
                decimal dImp = 0;

                DataTable t = ((DataView)dgv1.DataSource).ToTable();

                foreach (DataRow y in t.Rows)
                {
                    if ((string)y["mov_art"] == sArt && !(Boolean)y["mov_ann"])
                    {
                        dQta += 1;
                        dQkg += (decimal)y["mov_qkg"];
                        dImp += (decimal)y["mov_imp"];
                    }
                }

                frmGesMovIngLabel02 f = new frmGesMovIngLabel02();
                f._strConSql = _strConSql;
                f._strCliDes = sCli;
                f._strArtCod = sArt;
                f._strArtDes = sArd;
                f._decMovQta = dQta;
                f._decMovQkg = dQkg;
                f._decMovPrv = dPrv;
                f._decMovImp = dImp;
                f.ShowDialog();
            }
        }

        private void lottoSuDocumentoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new frmSeekDocLotti().ShowDialog();
        }

        private void lottiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmSeekDocLotti f = new frmSeekDocLotti();
            f._strSelect = "S";
            f.ShowDialog();
            if (f._strRes != "")
            {
                string sLot = f._strRes;

                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    x["mov_lot"] = sLot;
                    x["MovMdy"] = "S";
                }
            }
        }

        private void cmbMftTdc_SelectionChangeCommitted(object sender, EventArgs e)
        {
            _strCauTpd = cmbMftTdc.SelectedValue.ToString();
        }

        private void iminaDocumentoToolStripMenuItem_Click(object sender, EventArgs e)
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

        private void btnArtIns_Click(object sender, EventArgs e)
        {
            string s = "";

            NewRiga(true);

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

                        string sArt = (string)t.Rows[0]["tmp_art"];

                        t = _clsQry.ArtSeek(sArt, "SEEK");

                        if (_strCfo == _clsDef.TIPCLI)
                        {
                            DataRow[] j = ((DataTable)cmbMftCfo.DataSource).Select("CfoCod='" + cmbMftCfo.SelectedValue.ToString() + "'");
                            if (j.Length > 0)
                            {
                                s = (string)j[0]["CfoLic"];
                                t = _clsQry.ArtPrezzo(t, s);
                            }
                        }

                        FillMovRow(t, x);
                    }
                }
            }
        }

    }
}
