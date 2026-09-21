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

namespace APOffice
{
    public partial class frmGesDivCtrlTab : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private const string TABTABFOR = "TabTabFornitori";
        private const string TABIVA = "TabIva";
        private const string TABUMI = "TabUmi";
        private const string TABTGR = "TabTipoGrammatura";
        private const string TABEC3 = "TabEcrLv3";
        private const string TABREP = "TabReparti";

        public DataTable _tabTab = new DataTable("tabTab");
        public DataTable _tabDiv = new DataTable("tabTab");
        public DataTable _tabDie = new DataTable("tabTab");
        public string _strTip = "";
        public string _strForCod = "";
        public Boolean _bolMdy = false;

        private DataSet _dasGen = new DataSet();

        private string _strConSql = "";

        public frmGesDivCtrlTab()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
        }

        private void frmGesDivCtrlTab_Load(object sender, EventArgs e)
        {
            string s = "";
            _strConSql = _clsFun.ConSql("");
            FillTabs();
            SetDgv1();
            SetDgv2();

            s = "tab_cod=''";
            if(_strTip == "ALL")
                s = "";

            DataView v = new DataView(_tabTab, s, "tab_tip", DataViewRowState.CurrentRows);
            DataTable t = v.ToTable();
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "TabMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            dgv1.DataSource = t;
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Boolean b = Salva();
            if (!b)
            {
                if (MessageBox.Show("Abbandono delle modifiche, confermi?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                      == DialogResult.Yes)
                    this.Close();
            }
            else this.Close();
        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_tip";
            cTbc.Name = "Tabella";
            cTbc.Width = 35;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_cof";
            cTbc.Name = "Codice fornitore";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_cod";
            cTbc.Name = "Codice da associare";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "TabArt";
            cTbc.Name = "Ricerca Articolo per il reparto";
            cTbc.Width = 180;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "TabMsg";
            cTbc.Name = "Messaggio";
            cTbc.Width = 180;
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

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);
        }

        private void FillTabs()
        {
            DataRow x;
            //string p = "";
            string s = "";
            DataTable t = new DataTable();

            s = "SELECT * FROM " + TABIVA;
            t = _clsFun.FillTabSql(TABIVA, s, false, _strConSql);
            if (_dasGen.Tables.IndexOf(t.TableName) >= 0)
                _dasGen.Tables.Remove(t.TableName);
            _dasGen.Tables.Add(t);

            s = "SELECT * FROM " + TABUMI;
            t = _clsFun.FillTabSql(TABUMI, s, false, _strConSql);
            if (_dasGen.Tables.IndexOf(t.TableName) >= 0)
                _dasGen.Tables.Remove(t.TableName);
            _dasGen.Tables.Add(t);

            s = "SELECT * FROM " + TABTGR;
            t = _clsFun.FillTabSql(TABTGR, s, false, _strConSql);
            x = t.NewRow();
            if (_dasGen.Tables.IndexOf(t.TableName) >= 0)
                _dasGen.Tables.Remove(t.TableName);
            _dasGen.Tables.Add(t);

            s = "SELECT tab_lv1 + tab_lv2 + tab_cod AS tab_cod, tab_des FROM " + TABEC3;
            t = _clsFun.FillTabSql(TABEC3, s, false, _strConSql);
            if (_dasGen.Tables.IndexOf(t.TableName) >= 0)
                _dasGen.Tables.Remove(t.TableName);
            _dasGen.Tables.Add(t);

            s = "SELECT * FROM " + TABREP + " WHERE tab_ann=0";
            t = _clsFun.FillTabSql(TABREP, s, false, _strConSql);
            if (_dasGen.Tables.IndexOf(t.TableName) >= 0)
                _dasGen.Tables.Remove(t.TableName);
            _dasGen.Tables.Add(t);
        }

        //private void dgv1_DoubleClick(object sender, EventArgs e)
        //{
        //    txtSeek.Text = "";
        //    FillTab();
        //}

        private void FillTab(Boolean bolSeek)
        {
            if (dgv1.DataSource != null)
            {
                try
                {
                    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {

                        DataTable tTab = new DataTable();

                        if (dgv2.DataSource != null)
                            tTab = ((DataView)dgv2.DataSource).Table;
                        
                        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                        string sTip = (string)r.Row["tab_tip"];

                        string p = "Tab" + sTip;
                        if (sTip == "REP")
                            p = TABREP;
                        else if (sTip == "TGR")
                            p = TABTGR;
                        else if (sTip == "EC3")
                            p = TABEC3;

                        string s = "";
                        if (txtSeek.Text != "")
                            s = "tab_des LIKE '%" + txtSeek.Text + "%'";

                        if (p.ToUpper() != tTab.TableName.ToUpper() || bolSeek)
                        {
                            tTab = _dasGen.Tables[p];

                            DataView v = new DataView(tTab, s, "tab_des", DataViewRowState.CurrentRows);

                            dgv2.DataSource = v;
                            txtSeek.Text = "";
                        }
                        txtSeek.Select();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }
        }

        private Boolean Salva()
        {
            Boolean b = true;
            DataRow[] j;

            string s = "SELECT * FROM " + TABTABFOR + " WHERE tab_for='" + _strForCod + "'";
            DataTable tTab = _clsFun.FillTabSql(TABTABFOR, s, false, _strConSql);

            DataTable t = (DataTable)dgv1.DataSource;

            ArrayList aWhe = new ArrayList();
            aWhe.Add("tab_for");
            aWhe.Add("tab_tip");
            aWhe.Add("tab_cof");
            ArrayList aExl = new ArrayList();

            foreach(DataRow y in t.Rows)
            {
                if((string)y["TabMdy"] == "S")
                {

                    j = tTab.Select("tab_for='" + y["tab_for"] + "' AND tab_tip='" + y["tab_tip"] + "' AND tab_cof='" + y["tab_cof"] + "'");
                    if(j.Length > 0)
                        s = _clsFun.SqlUpdRow(tTab.TableName, tTab, j[0], y, aWhe, aExl);
                    else
                        s = _clsFun.SqlInsertRow(tTab.TableName,tTab, y);
                    if (s != "")
                        _clsFun.SqlWrite(s, _strConSql);
                    _bolMdy = true;
                }
            }

            return b;
        }

        private void txtSeek_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                FillTab(true);
        }

        private void dgv1_Click(object sender, EventArgs e)
        {
            //if (dgv2.DataSource != null)
            //{
            //    CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
            //    if (cm.Position >= 0)
            //    {
            //        DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
            //        string sCod = (string)r.Row["tab_cod"];

            //        cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            //        if (cm.Position >= 0)
            //        {
            //            r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
            //            r.Row["tab_cod"] = sCod;
            //            r.Row["TabMdy"] = "S";
            //        }
            //    }
            //}
        }

        private void dgv1_CurrentCellChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            //    if (cm.Position >= 0)
            //    {
            //        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
            //        DataRow x = r.Row;

            //        string sTab = (string)x["mov_art"];



            //    }
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //}

            FillTab(false);
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv1.DataSource != null)
            {
                if (e.ColumnIndex == 3)
                {
                    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {
                        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                        string sCof = (string)r.Row["tab_cof"];
                        string sTip = (string)r.Row["tab_tip"];

                        if (sTip == "REP")
                        {
                            frmUtyTab f = new frmUtyTab();
                            f._tabTab = _tabDiv;
                            f._strSql = "for_rep='" + sCof + "'";
                            f.ShowDialog();
                        }
                        else if (sTip == "EC3")
                        {
                            frmUtyTab f = new frmUtyTab();
                            f._tabTab = _tabDiv;
                            f._strSql = "for_ecr='" + sCof + "'";
                            f.ShowDialog();
                        }
                    }
                }
                else if (e.ColumnIndex == 2 && dgv2.DataSource != null)
                {
                    CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {
                        DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
                        string sCod = (string)r.Row["tab_cod"];

                        cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                        if (cm.Position >= 0)
                        {
                            r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                            r.Row["tab_cod"] = sCod;
                            r.Row["TabMdy"] = "S";
                        }
                    }

                }
            }

        }

        private void dgv2_DoubleClick(object sender, EventArgs e)
        {
            CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
                string sCod = (string)r.Row["tab_cod"];

                cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    r.Row["tab_cod"] = sCod;
                    r.Row["TabMdy"] = "S";
                }
            }
        }

        private void articoliToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new frmAnaArticolo().ShowDialog();
        }

        private void abbinamentoAutomaticoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_tabDiv == null)
                MessageBox.Show("Tabella EAN fornitore in divulgazione non agganciata");
            else
                Abbina();
        }

        private void Abbina()
        {
            string s = "";
            DataRow[] j;
            DataRow[] j2;
            DataRow[] j3;

            s = "SELECT ";
            s += "AnaArticoli.art_cod, ";
            s += "AnaArticoli.art_des, ";
            s += "AnaArticoli.art_ec1, ";
            s += "AnaArticoli.art_ec2, ";
            s += "AnaArticoli.art_ec3, ";
            s += "AnaBarcode.ean_ean ";
            s += "FROM AnaArticoli LEFT OUTER JOIN AnaBarcode ON AnaArticoli.art_cod = AnaBarcode.ean_art ";
            s += "WHERE AnaBarcode.ean_ean <> '' ";
            s += "ORDER BY AnaBarcode.ean_dti DESC";
            DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tArt.Columns["ean_ean"];
            tArt.PrimaryKey = keys;

            DataTable t = (DataTable)dgv1.DataSource;

            foreach(DataRow y in t.Rows)
            {
                if((string)y["tab_tip"] == "EC3")
                {
                    j = _tabDiv.Select("for_ecr='" + (string)y["tab_cof"] + "'");

                    for(int i = 0; i < j.Length; i++)
                    {
                        j2 = _tabDie.Select("die_for=" + y["tab_for"] + " AND die_arf='" + (string)j[i]["div_arf"] + "'");

                        for (int ii = 0; ii < j2.Length; ii++)
                        {
                            j3 = tArt.Select("ean_ean='" + (string)j2[ii]["die_ean"] + "'");
                            if (j3.Length > 0)
                            {
                                y["tab_cod"] = (string)j3[0]["art_ec1"] + (string)j3[0]["art_ec2"] + (string)j3[0]["art_ec3"];
                                y["TabMdy"] = "S";
                                break;
                            }
                        }
                    }
                }
            }

        }

        private void btnEcr_Click(object sender, EventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                //(string)r.Row["tab_cod"] == "" && 

                if ((string)r.Row["tab_tip"] == "EC3")
                {
                    frmGesTabEcr f = new frmGesTabEcr();
                    f._bolCho = true;
                    f.ShowDialog();
                    string sEcr = f._strRes;

                    if (sEcr != "")
                    {
                        r.Row["tab_cod"] = sEcr;
                        r.Row["TabMdy"] = "S";
                    }
                    FillTabs();

                }
            }
        }

    }
}
