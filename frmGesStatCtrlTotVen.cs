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
    public partial class frmGesStatCtrlTotVen : Form
    {
        private const string TABSTAVET = "GesNegVet";
        private const string TABSTAVEN = "GesNegVen";
        private const string TABMOVFAT = "GesFatTestate";

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public DataTable _tabVet = new DataTable();
        public DataTable _tabVep = new DataTable();

        private string _strConSql = "";
        private string _strConSqlStat = "";
        private string _strStatVisProf = "S";

        public frmGesStatCtrlTotVen()
        {
            InitializeComponent(); _strConSql = _clsFun.ConSql("");
            _strConSqlStat = _clsFun.ConSql("3");

        }

        private void frmGesStaCtrlTotVen_Load(object sender, EventArgs e)
        {

        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void procediToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FillDati();
        }
        private void FillDati()
        {
            DataRow[] j;
            DataRow x;

            string s = "";
            decimal d = 0;
            decimal dRes = 0;
            decimal dTot = 0;
            decimal dPag = 0;
            decimal dVen = 0;

            DataTable tTmp = new clsGenTabTmp().TabTmpVenCtrl("TabTmp");
            DataColumn[] keys = new DataColumn[2];
            keys[0] = tTmp.Columns["tmp_pos"];
            keys[1] = tTmp.Columns["tmp_sco"];
            tTmp.PrimaryKey = keys;

            s = "SELECT * ";
            s += "FROM GesNegVet INNER JOIN GesNegVen ON ";
            s += "GesNegVet.vet_neg = GesNegVen.ven_neg AND ";
            s += "GesNegVet.vet_cau = GesNegVen.ven_cau AND ";
            s += "GesNegVet.vet_day = GesNegVen.ven_day AND ";
            s += "GesNegVet.vet_ora = GesNegVen.ven_ora AND ";
            s += "GesNegVet.vet_pos = GesNegVen.ven_pos AND ";
            s += "GesNegVet.vet_sco = GesNegVen.ven_sco ";
            s += "WHERE ";
            s += "GesNegVet.vet_day >= " + _clsFun.DaySql(dtpDay.Value) + " AND ";
            s += "GesNegVet.vet_day <= " + _clsFun.DaySql(dtpDay.Value) + " ";
            if (_strStatVisProf == "N")
                s += " AND GesNegVet.vet_cau<>'PRO' ";
            s += "ORDER BY ";
            s += "GesNegVen.ven_cau,";
            s += "GesNegVen.ven_pos,";
            s += "GesNegVen.ven_ora,";
            s += "GesNegVen.ven_sco";
            DataTable t = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlStat);

            string sKey = "";


            foreach(DataRow y in t.Rows)
            {
                if (sKey != (string)y["ven_cau"] + (string)y["ven_pos"] + (string)y["ven_ora"] + (string)y["ven_sco"])
                {
                    sKey = (string)y["ven_cau"] + (string)y["ven_pos"] + (string)y["ven_ora"] + (string)y["ven_sco"];
                    dTot += (decimal)y["vet_imp"];

                }

                j = tTmp.Select("tmp_pos='" + y["ven_pos"] + "' AND tmp_sco='" + y["ven_sco"] + "'");
                if(j.Length == 0)
                {
                    x = tTmp.NewRow();
                    x["tmp_cau"] = y["vet_cau"];
                    x["tmp_cau"] = y["vet_cau"];
                    x["tmp_cau"] = y["vet_cau"];
                    x["tmp_cau"] = y["vet_cau"];
                    x["tmp_pos"] = y["ven_pos"];
                    x["tmp_sco"] = y["ven_sco"];
                    x["tmp_vet"] = y["vet_imp"];
                    x["tmp_vep"] = 0;
                    x["tmp_ven"] = 0;
                    tTmp.Rows.Add(x);
                }

                j = tTmp.Select("tmp_pos='" + y["ven_pos"] + "' AND tmp_sco='" + y["ven_sco"] + "'");
                if ((string)y["ven_sct"] == "RES")
                {
                    j[0]["tmp_ven"] = (decimal)j[0]["tmp_ven"] - (decimal)y["ven_ven"];
                    j[0]["tmp_vnd"] = (decimal)j[0]["tmp_vet"] + (decimal)j[0]["tmp_ven"];
                    dRes += (decimal)y["ven_ven"];
                }
                else
                {
                    j[0]["tmp_ven"] = (decimal)j[0]["tmp_ven"] + (decimal)y["ven_ven"];
                    j[0]["tmp_vnd"] = (decimal)j[0]["tmp_vet"] - (decimal)j[0]["tmp_ven"];
                }

            }

            dgv1.DataSource = tTmp;

            lblRes.Text = dRes.ToString();
            lblTot.Text = dTot.ToString();
            lblPag.Text = dPag.ToString();
            lblVen.Text = dVen.ToString();
        }

    }
}
