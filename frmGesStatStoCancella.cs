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
    public partial class frmGesStatStoCancella : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strConSqlSta = "";
        private string _strConSqlStaSto = "";

        public frmGesStatStoCancella()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
            //_strConSqlStaSto = _clsFun.ConSql("5");
        }

        private void frmGesStaStoCancella_Load(object sender, EventArgs e)
        {
            FillTab();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmGesStaStoCancella_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }        
        private void Esci()
        {
            this.Close();
        }

        private void FillTab()
        {
            Console.WriteLine("zzzz");

            for (int i = DateTime.Today.Year - 6; i < DateTime.Today.Year; i++)
                cmbYea.Items.Add(i.ToString());
            //cmbYea.SelectedIndex = 4;
        }

        private void btnCan_Click(object sender, EventArgs e)
        {
            if (chkSta.Checked)
                CancellaSta();
            if (chkDoc.Checked)
                CancellaDoc("FAT");
            if (chkDve.Checked)
                CancellaDoc("VEN");
            if (chkMov.Checked)
                CancellaMov("MOV");
        }

        private void CancellaDoc(string strTip)
        {
            lblMsg1.Text = "Documenti acquisto";
            if(strTip == "VEN")
                lblMsg1.Text = "Documenti vendita";

            string s = "SELECT * FROM GesFatTestate WHERE ";
            s += "fat_yfa<='" + cmbYea.Text + "' AND ";
            if (strTip == "FAT")
                s += "fat_tpd='FA' AND (fat_tdo='FA' OR fat_tdo='NA')";
            else if (strTip == "VEN")
                s += "fat_tpd='FV' AND (fat_tdo='FA' OR fat_tdo='NA')";

            DataTable t = _clsFun.FillTabSql("GesFatTestate", s, false, _strConSql);

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            int i = 0;

            foreach (DataRow y in t.Rows)
            {
                i++;

                lblMsg2.Text = i.ToString();

                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                s = "DELETE FROM GesMovimenti WHERE ";
                s += "mov_yfa='" + y["fat_yfa"] + "' AND ";
                s += "mov_nfa='" + y["fat_nfa"] + "'";
                _clsFun.SqlWrite(s, _strConSql);

                s = "DELETE FROM GesFatTestate WHERE ";
                s += "fat_idx=" + Convert.ToString(y["fat_idx"]);
                _clsFun.SqlWrite(s, _strConSql);
            }
        }

        private void CancellaMov(string strTip)
        {
            lblMsg1.Text = "Movimenti";

            string s = "SELECT * FROM GesMovTestate WHERE mot_ymo<='" + cmbYea.Text + "'";

            DataTable t = _clsFun.FillTabSql("GesFatTestate", s, false, _strConSql);

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            int i = 0;

            foreach (DataRow y in t.Rows)
            {
                i++;

                lblMsg2.Text = i.ToString();

                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                s = "DELETE FROM GesMovimenti WHERE ";
                s += "mov_ymo='" + y["mot_ymo"] + "' AND ";
                s += "mov_nmo='" + y["mot_nmo"] + "'";
                _clsFun.SqlWrite(s, _strConSql);

                s = "DELETE FROM GesMovTestate WHERE ";
                s += "mot_idx=" + Convert.ToString(y["mot_idx"]);
                _clsFun.SqlWrite(s, _strConSql);
            }
        }

        private void CancellaSta()
        {
            string s = "";
            lblMsg1.Text = "Statistiche di vendita";
            lblMsg2.Text = "";

            progressBar1.Value = 0;
            progressBar1.Maximum = 4;
            progressBar1.Minimum = 0;

            for (int n = 1; n <= 4; n++)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                string sTab = "GesNegMca";
                string sFld = "mca_day";
                if (n == 2)
                {
                    sTab = "GesNegVet";
                    sFld = "vet_day";
                }
                else if (n == 3)
                {
                    sTab = "GesNegVep";
                    sFld = "vep_day";
                }
                else if (n == 4)
                {
                    sTab = "GesNegVen";
                    sFld = "ven_day";
                }

                s = "SELECT COUNT(*) AS cnt FROM " + sTab + " WHERE YEAR(" + sFld + ") <= " + cmbYea.Text;
                DataTable t = _clsFun.FillTabSql(sTab, s, false, _strConSqlSta);

                if (t.Rows.Count > 0)
                {
                    Int64 i = Convert.ToInt64(t.Rows[0]["cnt"]);

                    while (i > 0)
                    {
                        s = "DELETE TOP (10000) FROM " + sTab + " WHERE YEAR(" + sFld + ") <= " + cmbYea.Text;
                        _clsFun.SqlWrite(s, _strConSqlSta);

                        i -= 10000;

                        lblMsg1.Text = "Statistiche " + sTab + " " + n.ToString() + "(" + i.ToString() + ")";
                        Application.DoEvents();
                    }
                }
            }


        }

    }
}
