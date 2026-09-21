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
    public partial class frmGesStatCtrlTotali : Form
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

        public frmGesStatCtrlTotali()
        {
            InitializeComponent(); 
            _strConSql = _clsFun.ConSql("");
            _strConSqlStat = _clsFun.ConSql("3");
        }

        private void frmGesStatCtrlTotali_Load(object sender, EventArgs e)
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

            string s = "";
            decimal d = 0;

            s = "SELECT * ";
            s += "FROM GesNegVen ";
            s += "WHERE ";
            s += "(GesNegVen.ven_day = " + _clsFun.DaySql(dtpDay.Value) + ") ";
            DataTable tVen = _clsFun.FillTabSql("Ven", s, false, _strConSqlStat);

            s = "SELECT * ";
            s += "FROM GesNegVep ";
            s += "WHERE ";
            s += "(GesNegVep.vep_day = " + _clsFun.DaySql(dtpDay.Value) + ") AND vep_tip <> 'PUN'";
            DataTable tVep = _clsFun.FillTabSql("Vep", s, false, _strConSqlStat);

            s = "SELECT ";
            s += "GesNegVet.vet_usr, ";
            s += "GesNegVet.vet_pos, ";
            s += "GesNegVet.vet_ora, ";
            s += "GesNegVet.vet_sco, ";
            s += "GesNegVet.vet_imp ";
            s += "FROM GesNegVet ";
            s += "WHERE ";
            s += "vet_cau='SCO' AND ";
            s += "(GesNegVet.vet_day = " + _clsFun.DaySql(dtpDay.Value) + ") ";
            s += "ORDER BY GesNegVet.vet_pos, GesNegVet.vet_sco";
            DataTable tVet = _clsFun.FillTabSql("Vet", s, false, _strConSqlStat);

            tVet.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_ven",
                Caption = "Vendite",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            tVet.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_ved",
                Caption = "Vendite",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            tVet.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_pag",
                Caption = "Vendite",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            tVet.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_pad",
                Caption = "Vendite",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            tVet.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_res",
                Caption = "Resi",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            tVet.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_sco",
                Caption = "Sconti",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            progressBar1.Value = 0;
            progressBar1.Maximum = tVet.Rows.Count;
            progressBar1.Minimum = 0;

            decimal dTot = 0;
            decimal dPag = 0;
            decimal dVen = 0;
            decimal dValScon = 0;

            foreach(DataRow y in tVet.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                decimal dRes = 0;

                dValScon = 0;
                s = (string)y["vet_sco"];

                if (s == "00437")
                    Console.WriteLine("aaaa");

                j = tVep.Select("vep_pos='" + y["vet_pos"] + "' AND vep_sco='" + y["vet_sco"] + "'");
                if (j.Length > 0)
                {
                    dTot += (decimal)y["vet_imp"];

                    for (int i = 0; i < j.Length; i++)
                    {
                        if ((string)j[i]["vep_tip"] == "RES") // Resto || (string)j[i]["vep_tip"] == "SCT" || (string)j[i]["vep_tip"] == "SCV")
                        {
                            Console.WriteLine("aaa");
                            y["tmp_pag"] = (decimal)y["tmp_pag"] - (decimal)j[i]["vep_imp"];
                            dPag -= (decimal)j[i]["vep_imp"];
                        }
                        else if ((string)j[i]["vep_tip"] == "PAG")
                        {
                            y["tmp_pag"] = (decimal)y["tmp_pag"] + (decimal)j[i]["vep_imp"];
                            dPag += (decimal)j[i]["vep_imp"];
                        }

                        if ((string)j[i]["vep_tip"] == "SCT" || (string)j[i]["vep_tip"] == "SCV")
                        {
                            dValScon += (decimal)j[i]["vep_imp"];
                            y["tmp_sco"] = (decimal)j[i]["vep_imp"];
                        }
                    }
                }

                j = tVen.Select("ven_pos='" + y["vet_pos"] + "' AND ven_sco='" + y["vet_sco"] + "' AND ven_sct='RES'");
                if (j.Length > 0)
                {
                    for (int i = 0; i < j.Length; i++)
                    {
                        dRes += (decimal)j[i]["ven_ven"];
                    }
                }

                j = tVen.Select("ven_pos='" + y["vet_pos"] + "' AND ven_sco='" + y["vet_sco"] + "'");
                if (j.Length > 0)
                {
                    d = (decimal)y["vet_imp"] + dValScon;

                    decimal dDelta = dValScon * 100 / d;
                    decimal dd = 0;

                    //dDelta = dValScon * 100 / d;

                    Console.WriteLine("aaaa");

                    for (int i = 0; i < j.Length; i++)
                    {
                        d = (decimal)j[i]["ven_ven"];

                        if(dDelta > 0)
                            d = _clsFun.MenoPer(d, dDelta);

                        if ((string)j[i]["ven_sct"] == "RES")
                        {
                            y["tmp_ven"] = (decimal)y["tmp_ven"] - d;
                            dVen -= d;
                            dd -= d;
                            y["tmp_res"] = (decimal)y["tmp_res"] + d;
                        }
                        else
                        {
                            y["tmp_ven"] = (decimal)y["tmp_ven"] + d;
                            dVen += d;
                            dd += d;
                        }

                        if (i == j.Length - 1)
                        {
                            if ((decimal)y["vet_imp"] != dd)
                            {
                                Console.WriteLine("aaaaaaa");

                                d = (decimal)y["vet_imp"] - dd;

                                y["tmp_ven"] = (decimal)y["tmp_ven"] + d;
                                dVen += d;

                            }
                        }
                    }
                }

                if ((decimal)y["tmp_ven"] != (decimal)y["vet_imp"])
                    Console.WriteLine("aaaaaaaaa");


                y["tmp_pad"] = (decimal)y["vet_imp"] - (decimal)y["tmp_pag"];
                y["tmp_ved"] = (decimal)y["vet_imp"] - (decimal)y["tmp_ven"];
            }

            dgv1.DataSource = tVet;

            lblTot.Text = dTot.ToString();
            lblPag.Text = dPag.ToString();
            lblVen.Text = dVen.ToString();
        }
    }
}
