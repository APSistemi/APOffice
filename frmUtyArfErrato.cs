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
    public partial class frmUtyArfErrato : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        public frmUtyArfErrato()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0); ;
            _strConSql = _clsFun.ConSql("");
        }

        private void frmUtyArfErrato_Load(object sender, EventArgs e)
        {

        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void frmUtyArfErrato_KeyDown(object sender, KeyEventArgs e)
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

        private void FillDati()
        {
            string s = "";

            s = "SELECT lia_art, cnt ";
            s += "FROM (SELECT lia_art, COUNT(*) AS cnt ";
            s += "FROM GesLisAcquisto ";
            s += "WHERE (lia_for = '" + txtFor.Text + "') ";
            s += "GROUP BY lia_art) AS A ";
            s += "WHERE (cnt > 1) ";
            s += "ORDER BY cnt DESC ";

            DataTable t = _clsFun.FillTabSql("Tab", s, false, _strConSql);

            dgv1.DataSource = t;

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Aggiorna();
        }

        private void Aggiorna()
        {
            string s = "";
            DataRow[] j;

            s = "SELECT * FROM GesLisAcquisto WHERE lia_for='" + txtFor.Text + "'";
            DataTable tArf = _clsFun.FillTabSql("Arf", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[2];
            keys[1] = tArf.Columns["lia_art"];
            keys[0] = tArf.Columns["lia_arf"];
            tArf.PrimaryKey = keys;

            DataTable t = (DataTable)dgv1.DataSource;

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                if(Convert.ToInt16(y["cnt"])> 1)
                {
                    s = "DELETE AnaBarcode WHERE ean_art='" + y["lia_art"] +"'";
                    _clsFun.SqlWrite(s, _strConSql);

                    j = tArf.Select("lia_art='" + y["lia_art"] + "'");
                    if(j.Length > 0)
                    {
                        for(int i = 1; i < j.Length; i++)
                        {
                            s = "DELETE GesLisAcquisto WHERE lia_art='" + j[i]["lia_art"] + "' AND lia_arf='" + j[i]["lia_arf"] + "'";
                            _clsFun.SqlWrite(s, _strConSql);

                        }

                    }

                }

            }

        }
    }
}
