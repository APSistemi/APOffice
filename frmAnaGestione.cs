using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmAnaGestione : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        DataSet _dasGen = new DataSet();

        private const string TABCNF = "AnaConfigurazione";

        private string _strConSql = "";

        public frmAnaGestione()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
        }

        private void frmAnaGestione_Load(object sender, EventArgs e)
        {
            FillTabs();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmAnaGestione_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void Esci()
        {
            Boolean b = Salva();
            if (!b)
            {
                if (MessageBox.Show("Continui l'uscita?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                      == DialogResult.Yes)
                    this.Close();
            }
            else this.Close();
        }

        private void FillTabs()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabNegozi";
            s = "SELECT * FROM TabNegozi WHERE tab_ann=0 ORDER BY tab_cod";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            //x = t.NewRow();
            //x["tab_cod"] = "";
            //x["tab_des"] = "  Non definito";
            //t.Rows.InsertAt(x, 0);
            cmbTabNeg.DataSource = t;
            cmbTabNeg.DisplayMember = "tab_des";
            cmbTabNeg.ValueMember = "tab_cod";
            //cmbTabNeg.SelectedValue = "";

            s = "SELECT * FROM AnaConfigurazione ORDER BY cnf_cod";
            t = _clsFun.FillTabSql(TABCNF, s, false, _strConSql);
            _dasGen.Tables.Add(t);
        }

        private void FillDati()
        {
            txtCnfRag.Text = "";
            txtCnfInd.Text = "";
            txtCnfNcv.Text = "";
            txtCnfLoc.Text = "";
            txtCnfCap.Text = "";
            txtCnfPrv.Text = "";
            txtCnfPiv.Text = "";
            txtCnfCfi.Text = "";
            txtCnfBan.Text = "";
            txtCnfIba.Text = "";
            txtCnfTel.Text = "";
            txtCnfWeb.Text = "";
            txtCnfMai.Text = "";
            txtCnfV01.Text = "";
            txtCnfPec.Text = "";
            txtCnfAgp.Text = "";

            if (cmbTabNeg.SelectedValue != null)
            {
                lblCod.Text = cmbTabNeg.SelectedValue.ToString();

                string s = cmbTabNeg.SelectedValue.ToString();
                DataTable t = _dasGen.Tables[TABCNF];
                DataRow[] j = t.Select("cnf_cod='" + s + "'");

                if (j.Length > 0)
                {
                    txtCnfRag.Text = (string)j[0]["cnf_rag"];
                    txtCnfInd.Text = (string)j[0]["cnf_ind"];
                    txtCnfNcv.Text = (string)j[0]["cnf_ncv"];
                    txtCnfLoc.Text = (string)j[0]["cnf_loc"];
                    txtCnfCap.Text = (string)j[0]["cnf_cap"];
                    txtCnfPrv.Text = (string)j[0]["cnf_prv"];
                    txtCnfPiv.Text = (string)j[0]["cnf_piv"];
                    txtCnfCfi.Text = (string)j[0]["cnf_cfi"];
                    txtCnfBan.Text = (string)j[0]["cnf_ban"];
                    txtCnfIba.Text = (string)j[0]["cnf_iba"];
                    txtCnfTel.Text = (string)j[0]["cnf_tel"];
                    txtCnfWeb.Text = (string)j[0]["cnf_web"];
                    txtCnfMai.Text = (string)j[0]["cnf_mai"];
                    txtCnfV01.Text = (string)j[0]["cnf_v01"];
                    txtCnfPec.Text = (string)j[0]["cnf_pec"];
                    txtCnfAgp.Text = (string)j[0]["cnf_agp"];
                }
            }
        }

        private void cmbTabNeg_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillDati();
        }

        private void txtCnf_Validated(object sender, EventArgs e)
        {
            TextBox txt = (TextBox)sender;
            Txt2Tab(txt);
        }

        private void Txt2Tab(TextBox txt)
        {
            DataTable t = _dasGen.Tables[TABCNF];
            string s = cmbTabNeg.SelectedValue.ToString();
            DataRow[] j = t.Select("cnf_cod='" + s + "'");

            if (j.Length > 0 && s == lblCod.Text)
            {
                if (txt.Name == "txtCnfRag")
                    j[0]["cnf_rag"] = txt.Text;
                else if (txt.Name == "txtCnfInd")
                    j[0]["cnf_ind"] = txt.Text;
                else if (txt.Name == "txtCnfNcv")
                    j[0]["cnf_ncv"] = txt.Text;
                else if (txt.Name == "txtCnfLoc")
                    j[0]["cnf_loc"] = txt.Text;
                else if (txt.Name == "txtCnfCap")
                    j[0]["cnf_cap"] = txt.Text;
                else if (txt.Name == "txtCnfPrv")
                    j[0]["cnf_prv"] = txt.Text;
                else if (txt.Name == "txtCnfPiv")
                    j[0]["cnf_piv"] = txt.Text;
                else if (txt.Name == "txtCnfCfi")
                    j[0]["cnf_cfi"] = txt.Text;
                else if (txt.Name == "txtCnfBan")
                    j[0]["cnf_ban"] = txt.Text;
                else if (txt.Name == "txtCnfTel")
                    j[0]["cnf_tel"] = txt.Text;
                else if (txt.Name == "txtCnfIba")
                    j[0]["cnf_iba"] = txt.Text;
                else if (txt.Name == "txtCnfWeb")
                    j[0]["cnf_web"] = txt.Text;
                else if (txt.Name == "txtCnfMai")
                    j[0]["cnf_mai"] = txt.Text;
                else if (txt.Name == "txtCnfV01")
                    j[0]["cnf_v01"] = txt.Text;
                else if (txt.Name == "txtCnfPec")
                    j[0]["cnf_pec"] = txt.Text;
                else if (txt.Name == "txtCnfAgp")
                    j[0]["cnf_agp"] = txt.Text;
            }
        }

        private Boolean Salva()
        {
            Boolean b = true;
            string s = "";
            DataRow[] j;

            s = "SELECT * FROM AnaConfigurazione";
            DataTable t = _clsFun.FillTabSql("AnaConfigurazione", s, false, _strConSql);

            DataTable tCnf = _dasGen.Tables[TABCNF];

            ArrayList aWhe = new ArrayList();
            aWhe.Add("cnf_cod");
            ArrayList aExl = new ArrayList();

            foreach(DataRow y in tCnf.Rows)
            {
                j = t.Select("cnf_cod='" + (string)y["cnf_cod"] + "'");

                if (j.Length == 0)
                    s = _clsFun.SqlInsertRow(TABCNF, t, y);
                else
                    s = _clsFun.SqlUpdRow(TABCNF, t, j[0], y, aWhe, aExl);

                if (s != "")
                    _clsFun.SqlWrite(s, _strConSql);
            }

            return (b);
        }
    }
}
