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
    public partial class frmGesStatCtrlPagamenti : Form
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
        private string _strStatVisProf = "";

        public frmGesStatCtrlPagamenti()
        {
            InitializeComponent();
            _strConSql = _clsFun.ConSql("");
            _strConSqlStat = _clsFun.ConSql("3");
        }

        private void frmGesStatCtrlPagamenti_Load(object sender, EventArgs e)
        {

        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void procediToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dgv1.DataSource = VenOpe();
        }

        public DataTable VenOpe()
        {
            string s = "";
            DataRow x;
            DataRow[] j;
            DataRow[] j2;
            DataRow[] j3;
            decimal d = 0;

            //string sPos = strPos.PadLeft(2, Convert.ToChar('0'));

            s = "SELECT * FROM TabCauCassa";
            DataTable tCau = _clsFun.FillTabSql("Cau", s, false, _strConSql);

            s = "SELECT ";
            s += "GesNegVet.vet_usr, ";
            s += "GesNegVet.vet_pos, ";
            s += "GesNegVet.vet_ora, ";
            s += "GesNegVet.vet_sco, ";
            s += "GesNegVet.vet_imp, ";
            s += "GesNegVep.vep_cod, ";
            s += "GesNegVep.vep_tip, ";
            s += "GesNegVep.vep_imp ";
            //s += "GesNegVet.vep_imt ";
            s += "FROM GesNegVep ";
            s += "INNER JOIN GesNegVet ON ";
            s += "GesNegVep.vep_neg = GesNegVet.vet_neg AND ";
            s += "GesNegVep.vep_cau = GesNegVet.vet_cau AND ";
            s += "GesNegVep.vep_day = GesNegVet.vet_day AND ";
            s += "GesNegVep.vep_ora = GesNegVet.vet_ora AND ";
            s += "GesNegVep.vep_pos = GesNegVet.vet_pos AND ";
            s += "GesNegVep.vep_sco = GesNegVet.vet_sco ";
            s += "WHERE ";
            s += "(GesNegVep.vep_day = " + _clsFun.DaySql(dtpDay.Value) + ") ";
            //s += "(GesNegVep.vep_tip = 'PAG' OR GesNegVep.vep_tip = 'RES')";
            //s += "GROUP BY GesNegVep.vep_cod, GesNegVep.vep_tip ";
            s += "ORDER BY GesNegVet.vet_pos, GesNegVet.vet_sco";
            DataTable tVep = _clsFun.FillTabSql("Vep", s, false, _strConSqlStat);

            //_clsFun.FileLog("LogMio.log", "aaa", s);
            //_clsFun.FileLog("LogMio.log", "aaa", tVep.Rows.Count.ToString());

            //s = "SELECT * FROM GesNegMca ";
            //s += "WHERE ";
            //s += "mca_day = " + _clsFun.DaySql(dayDay) + " AND ";
            //s += "mca_usr = '" + strUsr + "' AND ";
            //s += "mca_pos = '" + sPos + "'";
            //DataTable tMca = _clsFun.FillTabSql("Mca", s, false, _strConSqlStat);

            tVep.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "PagCod",
                Caption = "Codice",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (string)""
            });

            tVep.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "PagDes",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (string)""
            });

            tVep.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "PagTip",
                Caption = "Descrizione",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (string)""
            });

            tVep.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "PagKey",
                Caption = "Descrizione",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (string)""
            });

            tVep.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "VepImp",
                Caption = "",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });

            tVep.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "ScoDif",
                Caption = "",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });

            foreach (DataRow y in tVep.Rows)
            {
                j = tCau.Select("tab_cod='" + y["vep_cod"] + "'");
                if (j.Length > 0)
                {
                    y["PagCod"] = (string)j[0]["tab_ppo"];
                    y["PagDes"] = (string)j[0]["tab_des"];
                    y["PagTip"] = (string)j[0]["tab_tip"];
                    y["PagKey"] = (string)j[0]["tab_key"];
                }
            }

            //foreach (DataRow y in tMca.Rows)
            //{
            //    x = tVep.NewRow();
            //    x["PagDes"] = y["mca_cau"];
            //    if ((string)x["PagDes"] == "VER")
            //        x["PagDes"] = "Versamento";
            //    else if ((string)x["PagDes"] == "PRE")
            //        x["PagDes"] = "Prelevamento";

            //    if ((string)y["mca_ora"] != "")
            //        x["PagDes"] += "(" + (string)y["mca_ora"] + ")";

            //    x["VepImp"] = y["mca_imp"];
            //    tVep.Rows.Add(x);
            //}

            //d = 0;
            //j = tVep.Select("vep_tip='RES'");
            //if (j.Length > 0)
            //    d = (decimal)j[0]["VepImp"];
            //if (d > 0)
            //{
            //    j = tVep.Select("vep_cod='001'");
            //    if (j.Length > 0)
            //        j[0]["VepImp"] = (decimal)j[0]["VepImp"] - d;
            //}

            //s = "SELECT ";
            //s += "A.vep_day, A.vep_pos, A.vep_sco, GesNegVep_1.vep_tip, GesNegVep_1.vep_imp ";
            //s += "FROM ";
            //s += "(SELECT vep_day, vep_pos, vep_sco ";
            //s += "FROM GesNegVep ";
            //s += "WHERE (vep_tip = 'SCT')) AS A ";
            //s += "INNER JOIN GesNegVep AS GesNegVep_1 ON GesNegVep_1.vep_day = A.vep_day AND GesNegVep_1.vep_pos = A.vep_pos AND GesNegVep_1.vep_sco = A.vep_sco ";
            //s += "WHERE (GesNegVep_1.vep_day = " + _clsFun.DaySql(dayDay) + ")";
            //DataTable t = _clsFun.FillTabSql("vep", s, false, _strConSqlStat);
            //if (t.Rows.Count > 0)
            //{
            //    j = t.Select("vep_tip='SCT'");

            //    if (j.Length > 0)
            //    {
            //        for (int i = 0; i < j.Length; i++)
            //        {
            //            s = "vep_pos='" + j[i]["vep_pos"] + "' AND ";
            //            s += "vep_sco='" + j[i]["vep_sco"] + "'";
            //            j2 = t.Select(s);
            //            if (j2.Length > 0)
            //            {
            //                j3 = tVep.Select("vep_tip='" + (string)j2[0]["vep_tip"] + "'");
            //                if (j3.Length > 0)
            //                    j3[0]["VepImp"] = (decimal)j3[0]["VepImp"] - (decimal)j[i]["vep_imp"];
            //            }
            //        }
            //    }

            //    //foreach (DataRow y in t.Rows)
            //    //{
            //    //    j = tVep.Select("vep_tip='SCT'");
            //    //    if (j.Length > 0)
            //    //        d = (decimal)j[0]["VepImp"];
            //    //    if (d > 0)
            //    //    {
            //    //        j = tVep.Select("vep_cod='001'");
            //    //        if (j.Length > 0)
            //    //            j[0]["VepImp"] = (decimal)j[0]["VepImp"] - d;
            //    //    }
            //    //}
            //}

            //t = tVep.Clone();
            //foreach (DataRow y in tVep.Rows)
            //{
            //    if ((string)y["vep_tip"] == "PAG")
            //        t.ImportRow(y);
            //}

            if (tVep.Rows.Count > 0)
            {
                string sPos = (string)tVep.Rows[0]["vet_pos"];
                string sSco = (string)tVep.Rows[0]["vet_sco"];

                d = 0;

                foreach (DataRow y in tVep.Rows)
                {


                    if((string)y["vet_sco"] != sSco)
                    {
                        j = tVep.Select("vet_pos='"+ sPos+"' AND vet_sco='" + sSco + "'");
                        if (j.Length > 0)
                        {
                            j[0]["VepImp"] = d;
                            j[0]["ScoDif"] = (decimal)j[0]["vet_imp"]-d;
                        }
                        sPos = (string)y["vet_pos"];
                        sSco = (string)y["vet_sco"];
                        d = 0;

                    }

                    if ((string)y["PagTip"] != "PUN" && (string)y["vep_tip"] != "SCT")
                    {

                        if ((string)y["PagTip"] == "RES")
                            d -= (decimal)y["vep_imp"];
                        else
                            d += (decimal)y["vep_imp"];
                    }

                }
            }

            return tVep;
        }

    }
}
