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
    public partial class frmUtyArtScontoVenduto : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public frmUtyArtScontoVenduto()
        {
            InitializeComponent();
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
        }

        private void frmUtyArtScontoVenduto_Load(object sender, EventArgs e)
        {

        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FillDati();
        }

        private void FillDati()
        {
            string s = "SELECT ";
            s += "vet_idx, ";
            s += "vet_day, ";
            s += "vet_neg, ";
            s += "vet_cau, ";
            s += "vet_ora, ";
            s += "vet_pos, ";
            s += "vet_sco, ";
            s += "vet_imp, ";
            s += "vet_sct ";
            s += "FROM GesNegVet WHERE vet_day>=" + _clsFun.DaySql(dtpDti.Value) + " AND vet_day<=" + _clsFun.DaySql(dtpDtf.Value);
            DataTable t = _clsFun.FillTabSql("GesNegVet", s, false, _strConSqlSta);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "VetMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "VetSco",
                Caption = "Sco",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            DataTable t1 = t.Clone();

            //dgv1.DataSource = t;

            decimal d = 0;
            decimal dImp = 0;
            decimal dSco = 0;

            s = "SELECT ven_idx, ven_day, ven_neg, ven_cau, ven_ora, ven_pos, ven_sco, ven_art, ven_ven, ven_sct, ven_scn ";
            s += "FROM GesNegVen WHERE ";
            s += "ven_neg='9999999999999999'";
            DataTable tVen = _clsFun.FillTabSql("GesNegVen", s, false, _strConSqlSta);
            DataTable t2 = tVen.Clone();

            progressBar1.Value = 0;
            progressBar1.Maximum = (int)t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                s = "SELECT ven_idx, ven_day, ven_neg, ven_cau, ven_ora, ven_pos, ven_sco, ven_art, ven_ven, ven_sct, ven_scn ";
                s += "FROM GesNegVen WHERE ";
                s += "ven_neg='" + y["vet_neg"] + "' AND ";
                s += "ven_cau='" + y["vet_cau"] + "' AND ";
                s += "ven_day=" + _clsFun.DaySql((DateTime)y["vet_day"]) + " AND ";
                s += "ven_ora='" + y["vet_ora"] + "' AND ";
                s += "ven_pos='" + y["vet_pos"] + "' AND ";
                s += "ven_sco='" + y["vet_sco"] + "'";

                tVen = _clsFun.FillTabSql("GesNegVen", s, false, _strConSqlSta);
                
                dImp = 0;

                if(tVen.Rows.Count > 0)
                {
                    foreach(DataRow yy in tVen.Rows)
                    {
                        if((string)yy["ven_sct"] == "RES")
                            dImp -= (decimal)yy["ven_ven"];
                        else
                            dImp += (decimal)yy["ven_ven"];
                    }
                }

                if((decimal)y["vet_imp"] < dImp)
                {
                    dSco = (decimal)y["vet_imp"] - dImp;
                    y["vet_sct"] = dSco;
                    y["VetMdy"] = "S";

                    int iRow = -1;
                    if (dSco < 0)
                    {
                        decimal dTot = 0;

                        d = Math.Abs((Decimal)dSco) / ((Decimal)y["vet_imp"] + Math.Abs(dSco)) * 100;
                        foreach (DataRow k in tVen.Rows)
                        {
                            k["ven_scn"] = Math.Round((Decimal)k["ven_ven"] * d / 100,2,MidpointRounding.ToEven);
                            k["ven_ven"] = (decimal)k["ven_ven"] - (decimal)k["ven_scn"];

                            dTot += (decimal)k["ven_scn"];
                            t2.ImportRow(k);

                            iRow++;
                        }

                        d = dTot +  dSco;

                        if (d != 0)
                        {
                            Console.WriteLine("aaaaaaaaa");

                            t2.Rows[t2.Rows.Count - 1].Delete();

                            tVen.Rows[iRow]["ven_scn"] = (decimal)tVen.Rows[iRow]["ven_scn"] + d; // (d * -1);
                            tVen.Rows[iRow]["ven_ven"] = (decimal)tVen.Rows[iRow]["ven_ven"] + d; // (d * -1);
                            t2.ImportRow(tVen.Rows[iRow]);

                        }
                        y["VetSco"] = dTot-d;

                        t1.ImportRow(y);

                    }
                }
            }

            dgv1.DataSource = t1;
            dgv2.DataSource = t2;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Salva();
        }

        private void Salva()
        {
            string s = "";

            DataTable tVet = (DataTable)dgv1.DataSource;
            DataTable tVen = (DataTable)dgv2.DataSource;

            foreach (DataRow y in tVet.Rows)
            {
                s = "UPDATE GesNegVet SET vet_sct=" + ((decimal)y["vet_sct"]).ToString().Replace(",", ".") + " WHERE ";
                s += "vet_idx=" + Convert.ToString(y["vet_idx"]) + "";
                _clsFun.SqlWrite(s, _strConSqlSta);
            }

            Console.WriteLine("aaaaaa");

            foreach (DataRow y in tVen.Rows)
            {
                s = "UPDATE GesNegVen SET ";
                s += "ven_scn=" + ((decimal)y["ven_scn"]).ToString().Replace(",", ".") + ", ";
                s += "ven_ven=" + ((decimal)y["ven_ven"]).ToString().Replace(",", ".") + " ";
                s += "WHERE ";
                s += "ven_idx=" + Convert.ToString(y["ven_idx"]) + "";
                _clsFun.SqlWrite(s, _strConSqlSta);
            }
        }

    }
}
