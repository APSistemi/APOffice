using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmAnaTessera : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private const string TABTESANA = "AnaTessere";
        private const string TABTESLEG = "AnaTessLegami";
        private const string TABSTAVET = "GesNegVet";
        private const string TABTABCAU = "TabCauCassa";

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public string _strNeg = "";
        public string _strCod = "";
        public string _strDes = "";
        public string _strGru = "";
        public DataTable _tabTmp = new DataTable();

        public Boolean _bolCodNew = false;
        public Boolean _bolCodSostituito = false;

        public frmAnaTessera()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmAnaTessere_Load(object sender, EventArgs e)
        {
            ApplyModernUi();
            _strConSql = _clsFun.ConSql(""); 
            _strConSqlSta = _clsFun.ConSql("3");

            string s = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
            if (s == "")
            {
                MessageBox.Show("Configurazione negozio mancante sui parametri ini!", "CONTROLLO ACCESSO DATI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }

            SetDgv1();
            SetDgv2();
            FillTabs();
            FillData();
            FillMov();
            FillGru();

            pnlAnag.Enabled = false;
            if(txtTesCod.Text != "")
                pnlAnag.Enabled = true;

            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmAnaTessera_dgv1");
            clsUiIcons.RestoreGridColumnWidths(dgv2, "frmAnaTessera_dgv2");
        }

        private void frmAnaTessera_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmAnaTessera_dgv1");
            clsUiIcons.SaveGridColumnWidths(dgv2, "frmAnaTessera_dgv2");
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
                    if (excelToolStripMenuItem != null)
                        excelToolStripMenuItem.Image = clsUiIcons.GetIcon("excel", 16);
                }

                this.BackColor = Color.FromArgb(243, 244, 246);

                if (btnSos != null)
                {
                    clsUiIcons.StyleStatButton(btnSos, "Sostituzione", "sync", 14,
                        Color.FromArgb(254, 252, 232), Color.FromArgb(254, 240, 138),
                        Color.FromArgb(250, 204, 21), Color.FromArgb(113, 63, 18), Color.FromArgb(202, 138, 4));
                }

                if (btnCollega != null)
                {
                    clsUiIcons.StyleStatButton(btnCollega, "Collega", "link", 14,
                        Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254),
                        Color.FromArgb(96, 165, 250), Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));
                }

                if (btnRett != null)
                {
                    clsUiIcons.StyleStatButton(btnRett, "Rettifiche", "edit", 14,
                        Color.FromArgb(245, 243, 255), Color.FromArgb(221, 214, 254),
                        Color.FromArgb(167, 139, 250), Color.FromArgb(76, 29, 149), Color.FromArgb(124, 58, 237));
                }

                if (dgv1 != null)
                    clsUiIcons.StyleDataGridView(dgv1);
                if (dgv2 != null)
                    clsUiIcons.StyleDataGridView(dgv2);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaTessera.ApplyModernUi", ex.Message);
            }
        }

        private void frmAnaTessera_KeyDown(object sender, KeyEventArgs e)
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
            Boolean b = Salva();
            _strDes = txtTesDes.Text;
            _strGru = cmbTesGru.SelectedValue.ToString();
            
            if (!b)
            {
                if (MessageBox.Show("Continui l'uscita?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                      == DialogResult.Yes)
                    this.Close();
            }
            else this.Close();
        }

        private void txtTesCod_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Return)
                CtrlTes(txtTesCod.Text);

        }
    
        private void txtTesCod_Validated(object sender, EventArgs e)
        {
            CtrlTes(txtTesCod.Text);
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

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_day";
            cTbc.Name = "Data";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Format = "dd/MM/yy";
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_cau";
            cTbc.Name = "Causale";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_ora";
            cTbc.Name = "Ora";
            cTbc.Width = 35;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.DefaultCellStyle.Format = "##.##";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_pos";
            cTbc.Name = "Cassa";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);
            //dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_sco";
            cTbc.Name = "Scontrino";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "TipDes";
            cTbc.Name = "Descrizione";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VepImp";
            cTbc.Name = "Importo";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0.00";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VepPun";
            cTbc.Name = "Punti";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VepPua";
            cTbc.Name = "Punti articoli";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VepArt";
            cTbc.Name = "Articolo";
            cTbc.Width = 170;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
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
            cTbc.DataPropertyName = "let_cod";
            cTbc.Name = "C.legame";
            cTbc.Width = 45;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "let_tes";
            cTbc.Name = "Tessera";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tes_des";
            cTbc.Name = "Nome";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv2.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.DataPropertyName = "tes_ann";
            cCbc.Name = "Ann.";
            cCbc.Width = 40;
            cCbc.ValueType = typeof(string);
            cCbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv2.Columns.Add(cCbc);

        }

        private void FillTabs()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabProfessioni";
            s = "SELECT * FROM " + p + " WHERE tab_ann=0";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbTesPro.DataSource = t;
            cmbTesPro.DisplayMember = "tab_des";
            cmbTesPro.ValueMember = "tab_cod";
            cmbTesPro.SelectedValue = "";

            p = "TabFidGruppi";
            s = "SELECT * FROM " + p + " WHERE tab_ann=0";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbTesGru.DataSource = t;
            cmbTesGru.DisplayMember = "tab_des";
            cmbTesGru.ValueMember = "tab_cod";
            cmbTesGru.SelectedValue = "";

            p = "TabFidCampagne";
            s = "SELECT * FROM " + p + " WHERE tab_ann=0 ORDER BY tab_dti DESC";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.Add(x);
            cmbFidCam.DataSource = t;
            cmbFidCam.DisplayMember = "tab_des";
            cmbFidCam.ValueMember = "tab_cod";
            //cmbFidCam.SelectedValue = "";
        }

        private void FillData()
        {
            if (_strCod == _clsDef.CODNEW)
            {
                _strCod = "";
                txtTesCod.Text = _strCod;
                txtTesCod.Select();
            }
            else if (_tabTmp.Rows.Count > 0)
            {
                txtTesCod.ReadOnly = true;
                _strCod = (string)_tabTmp.Rows[0]["tmp_tes"];

                string s = "SELECT * FROM AnaTessere WHERE tes_cod = '" + _strCod + "'";
                DataTable t = _clsFun.FillTabSql(TABTESANA, s, true, _strConSql);

                if (t.Rows.Count > 0)
                {
                    txtTesCod.Text = (string)t.Rows[0]["tes_cod"];
                    if (!_bolCodSostituito)
                    {
                        txtTesDes.Text = (string)t.Rows[0]["tes_des"];
                        txtTesInd.Text = (string)t.Rows[0]["tes_ind"];
                        txtTesLoc.Text = (string)t.Rows[0]["tes_loc"];
                        txtTesCap.Text = (string)t.Rows[0]["tes_cap"];
                        txtTesPrv.Text = (string)t.Rows[0]["tes_prv"];
                        txtTesTel.Text = (string)t.Rows[0]["tes_tel"];
                        txtTesCel.Text = (string)t.Rows[0]["tes_cel"];
                        txtTesMai.Text = (string)t.Rows[0]["tes_mai"];

                        if ((string)t.Rows[0]["tes_sex"] == "M")
                            rdbTesSxM.Checked = true;

                        dtpTesDna.Value = (DateTime)t.Rows[0]["tes_dna"];

                        cmbTesPro.SelectedValue = (string)t.Rows[0]["tes_pro"];
                        cmbTesGru.SelectedValue = (string)t.Rows[0]["tes_gru"];

                        chkTesAnn.Checked = (Boolean)t.Rows[0]["tes_ann"];
                    }
                    txtTesDes.Select();
                }
            }
        }

        private void FillGru()
        {
            string s = "SELECT * FROM " + TABTESLEG + " WHERE ";
            if(_strCod == "")
                s += "let_tes = '99999999999999999999'";
            else
                s += "let_tes = '" + _strCod + "'";

            DataTable t = _clsFun.FillTabSql(TABTESLEG, s, false, _strConSql);

            if(t.Rows.Count > 0)
            {
                s = "SELECT ";
                s += "AnaTessLegami.*, ";
                s += "AnaTessere.tes_ann, ";
                s += "AnaTessere.tes_des ";
                s += "FROM AnaTessLegami ";
                s += "INNER JOIN AnaTessere ON AnaTessLegami.let_tes = AnaTessere.tes_cod ";
                s += "WHERE let_cod='" + (string)t.Rows[0]["let_cod"] + "'";
                t = _clsFun.FillTabSql(TABTESLEG, s, false, _strConSql);
            }

            dgv2.DataSource = t;
        }

        private Boolean Salva()
        {
            Boolean b = true;
            string s = "";
            DataRow y;

            if (txtTesDes.Text == "")
            {
                MessageBox.Show("Nome mancante!");
                return false;
            }

            //if (txtTesCod.Text == _clsDef.CODNEW)
            //{
            //    txtTesCod.Text = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaClienti, 5, _strConSql);
            //    _strCod = txtTesCod.Text;
            //}

            _strCod = txtTesCod.Text;

            s = "SELECT * FROM " + TABTESANA + " WHERE tes_cod = '" + txtTesCod.Text + "'";

            DataTable t = _clsFun.FillTabSql(TABTESANA, s, true, _strConSql);

            y = t.NewRow();

            y["tes_cod"] = txtTesCod.Text;
            y["tes_des"] = txtTesDes.Text;
            y["tes_ind"] = txtTesInd.Text;
            y["tes_loc"] = txtTesLoc.Text;
            y["tes_cap"] = txtTesCap.Text;
            y["tes_prv"] = txtTesPrv.Text;
            y["tes_tel"] = txtTesTel.Text;
            y["tes_cel"] = txtTesCel.Text;
            y["tes_mai"] = txtTesMai.Text;

            if (rdbTesSxM.Checked)
                y["tes_sex"] = "M";
            else
                y["tes_sex"] = "F";

            y["tes_dna"]   =  dtpTesDna.Value;

            y["tes_pro"] = cmbTesPro.SelectedValue;
            y["tes_gru"] = cmbTesGru.SelectedValue;

            y["tes_ann"] =  chkTesAnn.Checked;

            if (t.Rows.Count == 0)
                s = _clsFun.SqlInsertRow(TABTESANA, t, y);
            else
            {
                ArrayList aWhe = new ArrayList();
                aWhe.Add("tes_cod");
                ArrayList aExl = new ArrayList();
                s = _clsFun.SqlUpdRow(TABTESANA, t, t.Rows[0], y, aWhe, aExl);
            }

            if (s != "")
            {
                _clsFun.SqlWrite(s, _strConSql);
                _clsFun.FileLog(TABTESANA, txtTesCod.Text, s);
            }

            return (b);
        }
 
        private void CtrlTes(string strTes)
        {
            DataRow[] j;

            string s = strTes.Substring(0, strTes.Length - 1) + new clsCtrlCodici().FindMod10Digit(strTes.Substring(0, strTes.Length - 1));
            if (s != strTes)
            {
                MessageBox.Show("Codice non corretto! (" + s + ")");
                txtTesCod.Select();
            }
            else
            {
                s = "SELECT * FROM AnaTessere WHERE tes_cod='" + strTes + "'";
                DataTable t = _clsFun.FillTabSql(TABTESANA, s, false, _strConSql);
                if (t.Rows.Count > 0)
                {
                    MessageBox.Show("Codice tessera già presente!");
                    txtTesCod.Select();
                }
                else
                {
                    if (txtTesCod.Text != "")
                        pnlAnag.Enabled = true;
                }
            }
        }

        private void FillMov()
        {
            if (txtTesCod.Text != "")
            {
                DataRow x;
                DataRow[] j = ((DataTable)cmbFidCam.DataSource).Select("tab_cod='" + cmbFidCam.SelectedValue.ToString() + "'");
                if (j.Length > 0)
                {
                    DateTime dDti = (DateTime)j[0]["tab_dti"];
                    DateTime dDtf = (DateTime)j[0]["tab_dtf"];

                    //dDti = new DateTime(2019, 3, 4);

                    decimal dPun = 0;
                    decimal dImp = 0;

                    string s = "";

                    s = "SELECT * FROM " + TABTABCAU;
                    DataTable tCau = _clsFun.FillTabSql(TABTABCAU, s, false, _strConSql);

                    s = "SELECT ";
                    s += "GesNegVep.vep_cau, ";
                    s += "GesNegVep.vep_neg, ";
                    s += "GesNegVep.vep_day, ";
                    s += "GesNegVep.vep_ora, ";
                    s += "GesNegVep.vep_pos, ";
                    s += "GesNegVep.vep_sco, ";
                    s += "GesNegVep.vep_cod, ";
                    s += "GesNegVep.vep_tri, ";
                    s += "GesNegVep.vep_tip, ";
                    s += "GesNegVep.vep_off, ";
                    s += "GesNegVep.vep_imp, ";
                    s += "GesNegVep.vep_val, ";
                    //s += "GesNegVep.vep_ppo, ";
                    //s += "GesNegVep.vep_ord, ";
                    s += "GesNegVep.vep_ean, ";
                    s += "GesNegVet.vet_fid ";
                    s += "FROM GesNegVep ";
                    s += "LEFT OUTER JOIN GesNegVet ON ";
                    s += "GesNegVep.vep_sco = GesNegVet.vet_sco AND ";
                    s += "GesNegVep.vep_pos = GesNegVet.vet_pos AND ";
                    s += "GesNegVep.vep_ora = GesNegVet.vet_ora AND ";
                    s += "GesNegVep.vep_day = GesNegVet.vet_day AND ";
                    s += "GesNegVep.vep_neg = GesNegVet.vet_neg AND ";
                    s += "GesNegVep.vep_cau = GesNegVet.vet_cau ";
                    s += "WHERE ";
                    s += "GesNegVep.vep_day >= " + _clsFun.DaySql(dDti) + " AND ";
                    s += "GesNegVep.vep_day <= " + _clsFun.DaySql(dDtf) + " AND ";
                    s += "GesNegVet.vet_fid = '" + txtTesCod.Text + "' ";
                    s += "ORDER BY GesNegVet.vet_day DESC, GesNegVet.vet_ora DESC";
                    DataTable t = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlSta);

                    DataTable tVep = t.Clone();
                    tVep.Columns.Add(new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "TipDes",
                        Caption = "Descrizione",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    });
                    tVep.Columns.Add(new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "VepPun",
                        Caption = "Punti",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    });
                    tVep.Columns.Add(new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "VepPua",
                        Caption = "Punti",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    });
                    tVep.Columns.Add(new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "VepImp",
                        Caption = "Importo",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    });
                    tVep.Columns.Add(new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "VepArt",
                        Caption = "Articolo",
                        MaxLength = 100,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    });

                    DataColumn[] keys = new DataColumn[7];
                    keys[0] = tVep.Columns["vep_cau"];
                    keys[1] = tVep.Columns["vep_day"];
                    keys[2] = tVep.Columns["vep_ora"];
                    keys[3] = tVep.Columns["vep_pos"];
                    keys[4] = tVep.Columns["vep_sco"];
                    keys[5] = tVep.Columns["vep_tip"];
                    keys[6] = tVep.Columns["vep_ean"];
                    tVep.PrimaryKey = keys;

                    foreach (DataRow y in t.Rows)
                    {
                        string sSql = "";
                        sSql = "vep_cau='" + y["vep_cau"] + "' AND ";
                        sSql += "vep_day=" + _clsFun.DayMdb((DateTime)y["vep_day"]) + " AND ";
                        sSql += "vep_ora='" + y["vep_ora"] + "' AND ";
                        sSql += "vep_pos='" + y["vep_pos"] + "' AND ";
                        sSql += "vep_sco='" + y["vep_sco"] + "' AND ";
                        sSql += "vep_tip='" + y["vep_tip"] + "' AND ";
                        sSql += "vep_ean='" + y["vep_ean"] + "'";
                        j = tVep.Select(sSql);
                        if (j.Length == 0)
                        {
                            x = tVep.NewRow();
                            x["vep_day"] = y["vep_day"];
                            x["vep_cau"] = y["vep_cau"];
                            x["vep_ora"] = y["vep_ora"];
                            x["vep_pos"] = y["vep_pos"];
                            x["vep_sco"] = y["vep_sco"];
                            x["vep_tip"] = y["vep_tip"];
                            x["vep_ean"] = y["vep_ean"];
                            x["VepPun"] = 0;
                            x["VepImp"] = 0;
                            
                            j = tCau.Select("tab_cod='" + (string)y["vep_cod"] + "'");
                            if (j.Length > 0)
                                x["TipDes"] = (string)j[0]["tab_des"];

                            if ((string)y["vep_cau"] != "PUN" && (string)y["vep_ean"] != "" && (string)y["vep_ean"] != "PUNTITESSERAX")
                            {
                                DataTable tArt = _clsQry.SeekEanArt((string)y["vep_ean"]);
                                if(tArt.Rows.Count > 0)
                                    x["VepArt"] = ((string)tArt.Rows[0]["tmp_ard"]).Trim() + "-" + (string)tArt.Rows[0]["tmp_art"];
                            }

                            tVep.Rows.Add(x);
                        }

                        if ((string)y["vep_tip"] == "RES")
                            Console.WriteLine("aaaa");

                        j = tVep.Select(sSql);
                        if ((string)y["vep_tip"] == "PAG" || (string)y["vep_tip"] == "SCT" || (string)y["vep_tip"] == "RES")
                        {
                            j[0]["VepImp"] = (decimal)j[0]["VepImp"] + (decimal)y["vep_imp"];
                            if ((string)y["vep_tip"] == "PAG")
                                dImp += (decimal)y["vep_imp"];
                            else
                                dImp -= (decimal)y["vep_imp"];
                        }
                        else if (((string)y["vep_cau"]).Trim() != "PUN" && ((string)y["vep_ean"]).Trim() != "" && (string)y["vep_ean"] != "PUNTITESSERAX")
                        {
                            j[0]["VepPua"] = (decimal)j[0]["VepPua"] + (decimal)y["vep_imp"];
                            //dPun += (decimal)y["vep_imp"]; NON CONTEGGIATI PERCHE' GIA' COMPRESI NEL TOTALE SCONTRINO
                        }
                        else if ((string)y["vep_tip"] == "PUM")
                        {
                            j[0]["VepPun"] = (decimal)j[0]["VepPun"] + (decimal)y["vep_imp"];
                            if ((decimal)y["vep_imp"] > 0)
                                dPun -= (decimal)y["vep_imp"];
                            else
                                dPun += (decimal)y["vep_imp"];
                        }
                        else if ((string)y["vep_tip"] == "PUN")
                        {
                            j[0]["VepPun"] = (decimal)j[0]["VepPun"] + (decimal)y["vep_imp"];
                            dPun += (decimal)y["vep_imp"];
                        }
                    }

                    dgv1.DataSource = tVep;

                    lblTotPun.Text = dPun.ToString("##,###,##0");
                    lblTotImp.Text = dImp.ToString("##,###,##0.00");
                }
            }
        }

        private void btnSos_Click(object sender, EventArgs e)
        {
            frmUtyTessSostituzione f = new frmUtyTessSostituzione();
            f._strOldTes = txtTesCod.Text;
            f._strOldDes = txtTesDes.Text;
            f._strOldPun = lblTotPun.Text;
            f.ShowDialog();
            if (f._bolFatto)
            {
                _strCod = f._strNewTes;
                _bolCodSostituito = true;

                //chkTesAnn.Checked = true;

                _tabTmp.Clear();

                DataRow k = _tabTmp.NewRow();
                k["tmp_tes"] = f._strNewTes;
                k["tmp_ted"] = f._strNewDes;
                _tabTmp.Rows.Add(k);

                FillData();
                FillMov();
                FillGru();
            }
        }

        private void btnRett_Click(object sender, EventArgs e)
        {
            RettifichePunti();
            FillMov();
        }

        private void RettifichePunti()
        {
            frmUtyTessRettifiche f = new frmUtyTessRettifiche();
            f._strTesNeg = _strNeg;
            f._strTesCod = txtTesCod.Text;
            f.ShowDialog();
        }

        private void btnCollega_Click(object sender, EventArgs e)
        {
            string s = "";
            Boolean b = true;
            DataRow[] j;
            DataRow x;
            string sGru = "";

            if (txtCollega.Text == "")
                MessageBox.Show("Codice tessera non inserito!");
            else if (txtCollega.Text.Length != 13)
                MessageBox.Show("Codice tessera non corretto!");
            else if (!_clsFun.Numerico(txtCollega.Text))
                MessageBox.Show("Codice tessera non corretto!");
            else if (txtCollega.Text == txtTesCod.Text)
                MessageBox.Show("Codice tessera uguale al codice da collegare!");
            else
            {
                s = txtCollega.Text.Substring(0, txtCollega.Text.Length - 1) + new clsCtrlCodici().FindMod10Digit(txtCollega.Text.Substring(0, txtCollega.Text.Length - 1));
                if (s != txtCollega.Text)
                {
                    b = false;
                    MessageBox.Show("Barcode non corretto! (" + s + ")");
                }

                if(b)
                {

                    s = "SELECT * FROM " + TABTESANA + " WHERE ";
                    s += "tes_cod='" + txtCollega.Text + "'";
                    DataTable t = _clsFun.FillTabSql(TABTESLEG, s, false, _strConSql);
                    if (t.Rows.Count == 0)
                        MessageBox.Show("Codice tessera inesistente!");
                    else
                    {
                        t = (DataTable)dgv2.DataSource;
                        if (t.Rows.Count > 0)
                        {
                            sGru = (string)t.Rows[0]["let_cod"];
                            j = t.Select("let_tes='" + txtCollega.Text + "'");
                            if (j.Length > 0)
                            {
                                b = false;
                                MessageBox.Show("Tessera da collegare già presente!");
                            }
                        }
                        if (b)
                        {
                            s = "SELECT * FROM " + TABTESLEG + " WHERE ";
                            s += "let_tes='" + txtCollega.Text + "' AND ";
                            s += "let_ann=0";
                            t = _clsFun.FillTabSql(TABTESLEG, s, false, _strConSql);
                            if (t.Rows.Count > 0)
                            {
                                s = "SELECT * FROM " + TABTESLEG + " WHERE ";
                                s += "let_cod='" + t.Rows[0]["let_cod"] + "' AND ";
                                s += "let_ann=0";
                                t = _clsFun.FillTabSql(TABTESLEG, s, false, _strConSql);

                                j = t.Select("let_tes<>'" + txtCollega.Text + "'");
                                if (j.Length > 0)
                                {
                                    MessageBox.Show("Tessera già collegata con tessera " + (string)j[0]["let_tes"] + "!");
                                    b = false;
                                }
                            }
                        }
                    }
                }
                if(b)
                {
                    if(sGru == "")
                        sGru = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumGesTesGru, 5, _strConSql); //Raggruppamento vecchia e nuova tessera

                    s = "SELECT * FROM " + TABTESLEG + " WHERE ";
                    s += "let_tes='" + txtTesCod.Text + "'";
                    DataTable t = _clsFun.FillTabSql(TABTESLEG, s, false, _strConSql);
                    if (t.Rows.Count == 0)                    
                    {
                        x = t.NewRow();
                        x["let_cod"] = sGru;
                        x["let_tes"] = txtTesCod.Text;
                        x["let_ann"] = false;

                        s = _clsFun.SqlInsertRow(TABTESLEG, t, x);
                        _clsFun.SqlWrite(s, _strConSql);
                    }

                    x = t.NewRow();
                    x["let_cod"] = sGru;
                    x["let_tes"] = txtCollega.Text;
                    x["let_ann"] = false;

                    s = _clsFun.SqlInsertRow(TABTESLEG, t, x);
                    _clsFun.SqlWrite(s, _strConSql);

                    FillGru();
                }

            }

        }

        private void cmbFidCam_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillMov();
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FillXls();
        }

        private void FillXls()
        {
            DataTable t = (DataTable)dgv1.DataSource;

            string sTit = DateTime.Today.ToString("dd/MM/yyyy") + " Articoli nuovi " + txtTesCod.Text;
            string sFil = "Tessera " + txtTesCod.Text;

            string sFld = "";
            sFld += "vep_cau, Mov, 70, StringLiteral;";
            sFld += "vep_day, Data, 70, StringLiteral;";
            sFld += "vep_ora, Ora, 30, StringLiteral;";
            sFld += "vep_pos, Pos, 90, StringLiteral;";
            sFld += "vep_sco, Sc, 70, StringLiteral;";
            sFld += "vep_tip, Tipo, 70, StringLiteral;";
            sFld += "TipDes, Descriz., 70, StringLiteral;";
            sFld += "VepPun, Punti, 70, StringLiteral;";
            sFld += "VepImp, Importo, 70, StringLiteral;";

            (new clsExcel()).exportToCsv1(sFld, t, sFil, sTit, "");

        }

    }
}
