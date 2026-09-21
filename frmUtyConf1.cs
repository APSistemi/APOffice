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
    public partial class frmUtyConf1 : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        public frmUtyConf1()
        {
            InitializeComponent();
        }

        private void frmConfUty1_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql(""); 

        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string s = "SELECT ean_art FROM AnaBarcode WHERE (LEN(AnaBarcode.ean_ean) < 4)";

            DataTable t = _clsFun.FillTabSql("EAN", s, false, _strConSql);

            dgv1.DataSource = t;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string s = "";
            DataTable t = (DataTable)dgv1.DataSource;

            foreach(DataRow y in t.Rows)
            {
                string sArt = ((string)y["ean_art"]).Trim();

                if(sArt != "")
                {
                    s = "DELETE GesLisAcquisto WHERE lia_art='" + sArt + "'";
                    _clsFun.SqlWrite(s, _strConSql);

                    s = "DELETE AnaBarcode WHERE ean_art='" + sArt + "'";
                    _clsFun.SqlWrite(s, _strConSql);

                }
            }

            MessageBox.Show("Finito");


        }
    }
}
