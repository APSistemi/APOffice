using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmGesStatPunti : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private const string TABCAM = "TabFidCampagne";

        public DateTime _dayDti = DateTime.Today;
        public DateTime _dayDtf = DateTime.Today;

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public DataSet _dasGen = new DataSet();

        public frmGesStatPunti()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
        }

        private void frmGesStatPunti_Load(object sender, EventArgs e)
        {
            dtpDti.Value = _dayDti;
            dtpDtf.Value = _dayDtf;
            SetDgv1();
            FillTab();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmGesStaArticoli_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void btnEstrai_Click(object sender, EventArgs e)
        {
            FillDati();
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
            DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_fid";
            cTbc.Name = "Tessera";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            cTbc.ToolTipText = "Doppio click per dettaglio tessera";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_neg";
            cTbc.Name = "Negozio";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_day";
            cTbc.Name = "Data";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "vep_off";
            cTbc.DataPropertyName = "VepCam";
            cTbc.Name = "Campagna";
            cTbc.Width = 180;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            //cCmb = new DataGridViewComboBoxColumn();
            //cCmb.DataPropertyName = "vep_off";
            //cCmb.Name = "Campagna";
            //cCmb.Width = 200;
            //cCmb.DataSource = tCam;
            //cCmb.ValueMember = "tab_cod";
            //cCmb.DisplayMember = "tab_des";
            //cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            //dgv1.Columns.Add((DataGridViewColumn)cCmb);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_tip";
            cTbc.Name = "Tipo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_ora";
            cTbc.Name = "Ora";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_sco";
            cTbc.Name = "Scontrino/Data";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            cTbc.ToolTipText = "Formato data AMMGG (Doppio click per dettaglio movimento)";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_imp";
            cTbc.Name = "Punti";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VepMdy";
            cTbc.Name = "M";
            cTbc.Width = 25;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillTab()
        {
            string p = "TabNegozi";
            string s = "SELECT * FROM TabNegozi ORDER BY tab_cod";
            DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
            DataRow x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Tutti";
            t.Rows.InsertAt(x, 0);
            cmbNeg.DataSource = t;
            cmbNeg.DisplayMember = "tab_des";
            cmbNeg.ValueMember = "tab_cod";
            cmbNeg.SelectedValue = "";

            p = TABCAM;
            s = "SELECT * FROM TabFidCampagne WHERE tab_ann=0";
            DataTable tCam = _clsFun.FillTabSql(TABCAM, s, false, _strConSql);

            if (_dasGen.Tables.IndexOf(p) >= 0)
                _dasGen.Tables.Remove(p);
            _dasGen.Tables.Add(tCam);
        }

        private void FillDati()
        {
            string s = "";
            DataRow[] j;
            //string sScoKey = "";
            decimal d = 0;
            //decimal dTotVen = 0;

            string sNeg = cmbNeg.SelectedValue.ToString();
            //string sRep = cmbRep.SelectedValue.ToString();

            s = "SELECT AnaArticoli.art_cod, AnaArticoli.art_des, AnaArticoli.art_iva, AnaArticoli.art_sfr, AnaArticoli.art_rep, TabReparti.tab_des AS RepDes, TabEcrLv3.tab_des AS EcrDes ";
            s += "FROM AnaArticoli ";
            s += "LEFT OUTER JOIN TabReparti ON AnaArticoli.art_rep = TabReparti.tab_cod ";
            s += "LEFT OUTER JOIN TabEcrLv3 ON AnaArticoli.art_ec1 = TabEcrLv3.tab_lv1 AND AnaArticoli.art_ec2 = TabEcrLv3.tab_lv2 AND AnaArticoli.art_ec3 = TabEcrLv3.tab_cod ";
            DataTable tArt = _clsFun.FillTabSql("Art", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            s = "SELECT vep_cau, vep_neg, vep_day, vep_ora, vep_pos, vep_sco, vep_cod, vep_tri, vep_tip, vep_off, vep_imp, vep_val, vep_ppo, vep_ord, vep_ean, vep_fid ";
            s += "FROM GesNegVep ";
            s += "WHERE ";
            s += "(GesNegVep.vep_cau = 'PUN') AND ";
            s += "(vep_day >= " + _clsFun.DaySql(dtpDti.Value) + ") AND (vep_day <= " + _clsFun.DaySql(dtpDtf.Value) + ") ";
            if (sNeg != "")
                s += " AND vep_neg='" + sNeg + "' ";
            DataTable t = _clsFun.FillTabSql("Sta", s, false, _strConSqlSta);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "VepCam",
                Caption = "Campagna",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "VepMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "VepDay",
                Caption = "Data",
                ReadOnly = false
            });

            keys = new DataColumn[9];
            keys[2] = t.Columns["vep_neg"];
            keys[1] = t.Columns["vep_cau"];
            keys[0] = t.Columns["vep_day"];
            keys[4] = t.Columns["vep_ora"];
            keys[3] = t.Columns["vep_pos"];
            keys[5] = t.Columns["vep_sco"];
            keys[6] = t.Columns["vep_cod"];
            keys[7] = t.Columns["vep_ean"];
            keys[8] = t.Columns["vep_fid"];
            t.PrimaryKey = keys;

            DataTable tCam = _dasGen.Tables[TABCAM];

            foreach (DataRow y in t.Rows)
            {
                j = tCam.Select("tab_cod='" + y["vep_off"] + "'");
                if(j.Length > 0)
                {
                    y["VepCam"] = (string)j[0]["tab_cod"] + " - " + (string)(string)j[0]["tab_des"];
                }
                y["VepDay"] = (DateTime)y["vep_day"];
            }

            dgv1.DataSource = t;
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                string s = "";

                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string sCod = Convert.ToString(x["vep_fid"]);

                    s = "SELECT AnaTessere.*,  AnaTessLegami.let_cod ";
                    s += "FROM AnaTessere INNER JOIN AnaTessLegami ON AnaTessere.tes_cod = AnaTessLegami.let_cod ";
                    s += "WHERE ";
                    s += "let_tes='" + sCod + "'";

                    DataTable t = _clsFun.FillTabSql("", s, true, _strConSql);

                    if (t.Rows.Count > 0)
                    {
                        sCod = (string)t.Rows[0]["let_cod"];

                        t = new clsGenTabTmp().TabTmpTes("AnaTes");
                        DataRow k = t.NewRow();
                        k["tmp_tes"] = sCod;
                        k["tmp_ted"] = "";
                        t.Rows.Add(k);

                        frmAnaTessera2 f = new frmAnaTessera2();
                        f._tabTmp = t.Copy();
                        f.ShowDialog();
                    }

                }
            }
            if (e.ColumnIndex == 6)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string sNeg = (string)r.Row["vep_neg"];
                    string sCau = (string)r.Row["vep_cau"];
                    DateTime dDay = (DateTime)r.Row["vep_day"];
                    string sOra = (string)r.Row["vep_ora"];
                    string sPos = (string)r.Row["vep_pos"];
                    string sSco = (string)r.Row["vep_sco"];

                    frmGesStatScoDettaglio f = new frmGesStatScoDettaglio();
                    f._strNeg = sNeg;
                    f._strCau = sCau;
                    f._dayDay = dDay;
                    f._strOra = sOra;
                    f._strPos = sPos;
                    f._strSco = sSco;
                    f.ShowDialog();
                }
            }
        }

        private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = (string)x["VepCam"];

                    DataTable tCam = _dasGen.Tables[TABCAM];

                    DataRow[] j;
                    string sCam = "";

                    if(s.Length > 5)
                        sCam = s.Substring(0, 3).Trim();

                    if(sCam.Length == 0)
                        s = (string)tCam.Rows[0]["tab_cod"] + " - " + (string)tCam.Rows[0]["tab_des"];
                    else
                    {
                        sCam = (Convert.ToInt16(sCam)+1).ToString("000");

                        j = tCam.Select("tab_cod='" + sCam + "'");
                        if (j.Length > 0)
                            s = (string)j[0]["tab_cod"] + " - " + (string)j[0]["tab_des"];
                        else
                            s = "";
                    }

                    x["VepCam"] = s;
                    x["VepMdy"] = "S";

                    if(s != "")
                    {
                        sCam = s.Substring(0,3);
                        j = tCam.Select("tab_cod='" + sCam + "'");
                        if(j.Length > 0)
                        {
                            DateTime dDtf = (DateTime)j[0]["tab_dtf"];

                            if (DateTime.Compare(dDtf, DateTime.Today) < 0)
                                x["vep_day"] = dDtf;
                            else
                                x["vep_day"] = x["VepDay"];
                        }
                    }

                }
            }
        }

        private void btnAgg_Click(object sender, EventArgs e)
        {
            string s = "";
            DataTable t = (DataTable)dgv1.DataSource;

            foreach(DataRow y in t.Rows)
            {

                if((string)y["VepMdy"] == "S")
                {
                    s = "UPDATE GesNegVep SET ";
                    s += "vep_day=" + _clsFun.DaySql((DateTime)y["vep_day"]) + ",";
                    s += "vep_off='" + (string)y["vep_off"] + "' ";
                    s += "WHERE ";
                    s += "vep_neg='" + y["vep_neg"] + "' AND ";
                    s += "vep_cau='" + y["vep_cau"] + "' AND ";
                    s += "vep_day=" + _clsFun.DaySql((DateTime)y["VepDay"]) + " AND ";
                    s += "vep_ora='" + y["vep_ora"] + "' AND ";
                    s += "vep_pos='" + y["vep_pos"] + "' AND ";
                    s += "vep_sco='" + y["vep_sco"] + "' AND ";
                    s += "vep_cod='" + y["vep_cod"] + "' AND ";
                    s += "vep_ean='" + y["vep_ean"] + "' AND ";
                    s += "vep_fid='" + y["vep_fid"] + "'";
                    _clsFun.SqlWrite(s, _strConSqlSta);

                    s = "UPDATE GesNegVet SET ";
                    s += "vet_day=" + _clsFun.DaySql((DateTime)y["vep_day"]) + " ";
                    s += "WHERE ";
                    s += "vet_neg='" + y["vep_neg"] + "' AND ";
                    s += "vet_cau='" + y["vep_cau"] + "' AND ";
                    s += "vet_day=" + _clsFun.DaySql((DateTime)y["VepDay"]) + " AND ";
                    s += "vet_ora='" + y["vep_ora"] + "' AND ";
                    s += "vet_pos='" + y["vep_pos"] + "' AND ";
                    s += "vet_sco='" + y["vep_sco"] + "'";
                    _clsFun.SqlWrite(s, _strConSqlSta);
                }
            }
        }

    }
}
