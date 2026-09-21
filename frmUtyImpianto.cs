using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    public partial class frmUtyImpianto : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        private string sBtnExp = "Esporta tabella da ";
        private string sBtnImp = "Importa tabella in ";

        public frmUtyImpianto()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
        }

        private void frmUtyImpianto_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");

            cmbMdf.Items.Add("APOffice");
            cmbMdf.Items.Add("Prova");
            cmbMdf.SelectedIndex = 0;
            btnExp.Text = sBtnExp + cmbMdf.SelectedItem.ToString();
            btnImp.Text = sBtnImp + cmbMdf.SelectedItem.ToString();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnExp_Click(object sender, EventArgs e)
        {
            TabEsporta();
        }

        private void btnImp_Click(object sender, EventArgs e)
        {
            TabImporta();
        }

        private void cmbMdf_SelectionChangeCommitted(object sender, EventArgs e)
        {
            //int i = _strConSql.IndexOf("Catalog=");
            //i += 8;
            //string s = _strConSql.Substring(i);
            //i = s.IndexOf(";");
            //s = s.Substring(0, i);

            //if (s != cmbMdf.SelectedItem.ToString())
            //{
            //    _strConSql = _strConSql.Replace(s, cmbMdf.SelectedItem.ToString());
            //    btnExp.Text = sBtnExp + cmbMdf.SelectedItem.ToString();
            //    btnImp.Text = sBtnImp + cmbMdf.SelectedItem.ToString();
            //}
            ConSql();
        }

        private void cmbMdf_Leave(object sender, EventArgs e)
        {
            ConSql();
        }

        private void ConSql()
        {
            int i = _strConSql.IndexOf("Catalog=");
            i += 8;
            string s = _strConSql.Substring(i);
            i = s.IndexOf(";");
            s = s.Substring(0, i);

            string sMdf = cmbMdf.Text;        // SelectedItem.ToString();

            if (s != sMdf)
            {
                _strConSql = _strConSql.Replace(s, sMdf);
                btnExp.Text = sBtnExp + " " + sMdf;
                btnImp.Text = sBtnImp + " " + sMdf;

                if (!_clsFun.TestMdf(_strConSql))
                    MessageBox.Show("DataBase MDF non apribile!");
            }
        }

        private void TabEsporta()
        {
            string sChar = ";";
            string s = "SELECT DISTINCT TABLE_NAME FROM INFORMATION_SCHEMA.COLUMNS"; // WHERE TABLE_NAME LIKE 'TEST%'";

            DataTable t = _clsFun.FillTabSql("TmpTab", s, false, _strConSql);

            dgv1.DataSource = t;

            StreamWriter sw = new StreamWriter(lblPath.Text, false);

            string sRig = "";

            foreach (DataRow y in t.Rows)
            {
                if (!chkTab.Checked || ((string)y["TABLE_NAME"]).Substring(0, 3).ToLower() == "tab")
                {
                    sRig = (string)y["TABLE_NAME"] + sChar + "T" + sChar;

                    s = "SELECT * FROM " + (string)y["TABLE_NAME"];
                    DataTable tTab = _clsFun.FillTabSql((string)y["TABLE_NAME"], s, false, _strConSql);

                    foreach (DataColumn c in tTab.Columns)
                    {
                        sRig += c.ColumnName + sChar;
                    }
                    sw.Write(sRig + _clsDef.CRLF);

                    foreach (DataRow x in tTab.Rows)
                    {
                        sRig = (string)y["TABLE_NAME"] + sChar + "R" + sChar;

                        foreach (DataColumn c in tTab.Columns)
                        {
                            sRig += Convert.ToString(x[c.ColumnName]).Trim() + sChar;
                        }
                        sw.Write(sRig + _clsDef.CRLF);
                    }

                }
            }

            ((TextWriter)sw).Flush();
            sw.Close();
            sw.Dispose();

            MessageBox.Show("Generato file " + lblPath.Text);
        }

        private void TabImporta()
        {
            string s = "";
            //string sChar = ";";

            if (!File.Exists(lblPath.Text))
                MessageBox.Show(lblPath.Text, "NON TROVATO");
            else
            {
                using (StreamReader sr = new StreamReader(lblPath.Text))
                {
                    FileInfo fInfo = new FileInfo(lblPath.Text);
                    progressBar1.Value = 0;
                    progressBar1.Maximum = (int)fInfo.Length;
                    progressBar1.Minimum = 0;

                    string sRig = "";
                    //string sKeyTab = "";
                    DataTable t = new DataTable();
                    //DataRow[] j;
                    string[] aCol = sRig.Split(';');

                    while ((sRig = sr.ReadLine()) != null)
                    {
                        progressBar1.Increment(sRig.Length);
                        System.Windows.Forms.Application.DoEvents();

                        string[] a = sRig.Split(';');

                        string sTab = a[0];
                        string sTri = a[1];

                        if (sTri == "T")
                        {
                            aCol = sRig.Split(';');
                            s = "Select * FROM " + sTab;
                            t = _clsFun.FillTabSql(sTab, s, false, _strConSql);
                        }
                        else
                        {
                            string[] aRow = sRig.Split(';');

                            DataRow x = t.NewRow();

                            for(int i = 2; i < aCol.Length-1; i++)
                            {
                                Console.WriteLine("aaaa");

                                x[aCol[i]] = aRow[i];
                            }

                            s = _clsFun.SqlInsertRow(sTab, t, x);

                            _clsFun.SqlWrite(s, _strConSql);
                        }
                    }
                }
                MessageBox.Show("Fine importazione da file " + lblPath.Text);
            }

        }

    }
}
