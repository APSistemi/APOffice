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
    public partial class frmUtyTessSostituzione : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private const string TABANATES = "AnaTessere";
        private const string TABTESLEG = "AnaTessLegami";
        private const string TABSTAVET = "GesNegVet";
        private const string TABSTAVEP = "GesNegVep";

        public string _strOldTes = "";
        public string _strOldDes = "";
        public string _strOldPun = "";
        public string _strNewTes = "";
        public string _strNewDes = "";

        public Boolean _bolFatto = false;

        public frmUtyTessSostituzione()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmUtyTessSostituzione_Load(object sender, EventArgs e)
        {
            lblOldTes.Text = _strOldTes;
            lblOldDes.Text = _strOldDes;
            lblOldPun.Text = _strOldPun;
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void Esci()
        { 
            this.Close();
        }

        private void frmUtyTessSostituzione_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void btnSeek_Click(object sender, EventArgs e)
        {
            frmSeekTessere2 f = new frmSeekTessere2();
            f._bolSelect = true;
            f._bolAnagra = false;
            f.ShowDialog();
            if(f._tabTmp.Rows.Count > 0)
            {
                if ((string)f._tabTmp.Rows[0]["tmp_tes"] == lblOldTes.Text)
                    MessageBox.Show("Tessera uguale a quella preselezionata (" + (string)f._tabTmp.Rows[0]["tmp_tes"] + ")!");
                else
                {
                    txtNewCod.Text = (string)f._tabTmp.Rows[0]["tmp_tes"];
                    lblNewDes.Text = (string)f._tabTmp.Rows[0]["tmp_ted"];
                }
            }
        }

        private void txtNewCod_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Return && txtNewCod.Text != "")
            {
                DataTable t = new clsQuery().TessSeek(txtNewCod.Text);
                if(t.Rows.Count > 0)
                {
                    txtNewCod.Text = (string)t.Rows[0]["tes_cod"];
                    lblNewDes.Text = (string)t.Rows[0]["tes_des"];
                }
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Procedi con la sostituzione?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Sostituisci();
            }
        }

        private void Sostituisci()
        {
            string _strConSql = _clsFun.ConSql(""); 
            string _strConSqlSta = _clsFun.ConSql("3"); 

            string s = "";
            string sOra = DateTime.Now.ToString("HHmm");

            s = "SELECT * FROM GesNegVet WHERE vet_fid='XXXXXXXXXXXXXXX'";
            DataTable tVet = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlSta);

            DataRow x = tVet.NewRow();
            x["vet_usr"] = "";
            x["vet_cau"] = "PUN";
            x["vet_neg"] = "001";
            x["vet_day"] = DateTime.Today;
            x["vet_ora"] = sOra;
            x["vet_pos"] = "00";
            x["vet_sco"] = "00000";
            x["vet_fid"] = txtNewCod.Text;
            x["vet_imp"] = 0;
            x["vet_pun"] = Convert.ToDecimal(lblOldPun.Text);
            x["vet_sct"] = 0;
            x["vet_scp"] = 0;
            x["vet_art"] = 0;
            s = _clsFun.SqlInsertRow(TABSTAVET, tVet, x);
            _clsFun.SqlWrite(s, _strConSqlSta);

            s = "SELECT * FROM GesNegVep WHERE vep_sco='XXXXXXXXXXXXXXX'";
            DataTable tVep = _clsFun.FillTabSql(TABSTAVEP, s, false, _strConSqlSta);

            x = tVep.NewRow();
            x["vep_cau"] = "PUN";
            x["vep_neg"] = "001";
            x["vep_day"] = DateTime.Today;
            x["vep_ora"] = sOra;
            x["vep_pos"] = "00";
            x["vep_sco"] = "00000";
            x["vep_cod"] = "006";
            x["vep_tri"] = "3";
            x["vep_tip"] = "PUN";
            x["vep_off"] = "001";
            x["vep_imp"] = Convert.ToDecimal(lblOldPun.Text);
            x["vep_val"] = 0;
            x["vep_ppo"] = "";
            x["vep_ord"] = "";
            s = _clsFun.SqlInsertRow(TABSTAVEP, tVep, x);
            _clsFun.SqlWrite(s, _strConSqlSta);

            s = "UPDATE " + TABANATES + " SET tes_ann=1 WHERE tes_cod='" + lblOldTes.Text + "'";
            _clsFun.SqlWrite(s, _strConSql);

            s = "SELECT * FROM " + TABTESLEG + " WHERE let_tes='" + txtNewCod.Text + "'";
            DataTable t = _clsFun.FillTabSql(TABTESLEG, s, false, _strConSql);
            x = t.NewRow();
            x["let_cod"] = "";
            x["let_tes"] = txtNewCod.Text;
            x["let_ann"] = false;

            string sCod = "";

            if(t.Rows.Count > 0)
            {
                sCod = (string)t.Rows[0]["let_cod"];
            }
            else
            {
                sCod = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumGesTesGru, 5, _strConSql);
                x["let_cod"] = sCod;
                x["let_tes"] = txtNewCod.Text;
                x["let_ann"] = false;

                s = _clsFun.SqlInsertRow(TABTESLEG, t, x);
                _clsFun.SqlWrite(s, _strConSql);

                s = "SELECT * FROM " + TABTESLEG + " WHERE let_tes='" + lblOldTes.Text + "'";
                t = _clsFun.FillTabSql(TABTESLEG, s, false, _strConSql);
                if(t.Rows.Count > 0)
                {
                    s = "UPDATE " + TABTESLEG + " SET let_cod='" + sCod + "' WHERE let_tes='" + lblOldTes.Text + "'";
                    _clsFun.SqlWrite(s, _strConSqlSta);
                }
                else
                {
                    x["let_cod"] = sCod;
                    x["let_tes"] = lblOldTes.Text;
                    x["let_ann"] = false;
                    s = _clsFun.SqlInsertRow(TABTESLEG, t, x);
                    _clsFun.SqlWrite(s, _strConSql);
                }
            }
            
            _bolFatto = true;
            _strNewTes = txtNewCod.Text;
            _strNewDes = lblNewDes.Text;

            Esci();
        }
    }
}
