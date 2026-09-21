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
    public partial class frmUtyIva : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        private string _strConSqlSta = "";

        public frmUtyIva()
        {
            InitializeComponent();
        }

        private void frmUtiIva_Load(object sender, EventArgs e)
        {
            _strConSqlSta = _clsFun.ConSql("3");
            dtpAl.Value = new DateTime(2016, 6, 24);
            dtpDal.Value = new DateTime(2016, 6, 24);
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void estraiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FillDati();
        }

        private void FillDati()
        {
            string s = "";

            s = "SELECT ";
            s += "GesNegVen.ven_day, GesNegVen.ven_art, GesNegVen.ven_iva, B.ven_iva AS IVA ";
            s += "FROM GesNegVen ";
            s += "LEFT OUTER JOIN ";
            s += "(SELECT ven_art, ven_iva FROM GesNegVen ";
            s += "WHERE (ven_day > " + _clsFun.DaySql(dtpAl.Value) + ") ";
            s += "GROUP BY ven_art, ven_iva) AS B ON GesNegVen.ven_art = B.ven_art ";
            s += "WHERE (GesNegVen.ven_day < " + _clsFun.DaySql(dtpAl.Value) + ") AND (B.ven_iva IS NOT NULL)";
            s += "ORDER BY ven_art";

            DataTable t = _clsFun.FillTabSql("MOV", s, false, _strConSqlSta);

            dgv1.DataSource = t;


        }

        private void aggiornaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Salva();
        }

        private void Salva()
        {
            string s = "";
            string sArt = "";
            DataTable t = (DataTable)dgv1.DataSource;

            foreach(DataRow y in t.Rows)
            {
                if(sArt != (string)y["ven_art"])
                {
                    sArt = (string)y["ven_art"];

                    s = "UPDATE GesNegVen ";
                    s += "SET ven_iva='" + y["IVA"] + "' WHERE ";
                    s += "ven_day <= " + _clsFun.DaySql(dtpAl.Value) + " AND ven_art='" + y["ven_art"] + "'";
                    _clsFun.SqlWrite(s, _strConSqlSta);
                }

            }

            MessageBox.Show("Fine");
        }

    }
}
