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
    public partial class frmUtyTessRettifiche : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public string _strTesCod = "";
        public string _strTesNeg = "";
        public string _strTesCam = "";

        private const string TABANATES = "AnaTessere";
        private const string TABTESLEG = "AnaTessLegami";
        private const string TABSTAVET = "GesNegVet";
        private const string TABSTAVEP = "GesNegVep";

        public frmUtyTessRettifiche()
        {
            InitializeComponent();
        }

        private void frmUtyTessRettifiche_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");

            _strTesNeg = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");

            lblTesCod.Text = _strTesCod;
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void Esci()
        {
            this.Close();
        }
        private void frmUtyTessRettifiche_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void chkSgn_Click(object sender, EventArgs e)
        {
            if (chkSgn.Checked)
            {
                chkSgn.Text = "+";
                lblDes.Text = "Punti da aggiungere";
            }
            else
            {
                chkSgn.Text = "-";
                lblDes.Text = "Punti da sottrarre";
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (!_clsFun.Numerico(txtPun.Text) || Convert.ToDecimal(txtPun.Text) <= 0)
                MessageBox.Show("Valore punti non definito correttamente!");
            else
            {
                Salva();
                Esci();
            }
        }

        private void Salva()
        {
            string s = "";
            DataRow[] j;
            DataRow x;

            s = "SELECT tab_cod, tab_key FROM TabCauCassa WHERE tab_key='PUN' OR tab_key='PUM'";
            DataTable tCauPag = _clsFun.FillTabSql("tPag", s, false, _strConSql);

            //string sOra = DateTime.Now.ToString("HHmm");
            //string sOra = "0000";
            string sCau = "PUN";

            string sTip = "PUN";
            string sCod = "";

            DateTime dDay = DateTime.Today;
            string sOra = DateTime.Now.ToString("HHmm");

            if (_strTesCam != "")
            {
                s = "SELECT * FROM TabFidCampagne WHERE tab_cod='" + _strTesCam + "'";
                DataTable tCam = _clsFun.FillTabSql("", s, true, _strConSql);
                if (tCam.Rows.Count > 0)
                {
                    int iDif = DateTime.Compare((DateTime)tCam.Rows[0]["tab_dtf"], DateTime.Today);

                    if (iDif < 0)
                    {
                        dDay = (DateTime)tCam.Rows[0]["tab_dtf"];
                    }
                }
            }

            j = tCauPag.Select("tab_key='" + sTip + "'");
            if(j.Length > 0)
            {
                sCod = (string)j[0]["tab_cod"];

                if (!chkSgn.Checked)
                {
                    //sTip = "PUM";
                    sTip = "PUM";
                    j = tCauPag.Select("tab_key='" + sTip + "'");
                    if(j.Length > 0)
                        sCod = (string)j[0]["tab_cod"];
                }
            }

            if(sCod == "")
                 MessageBox.Show("Tipo movimento non definito in tabella pagamenti POS!", "CONTROLLO CAUSALE", MessageBoxButtons.OK, MessageBoxIcon.Question);
            else
            {
                s = "SELECT * FROM GesNegVet WHERE ";
                s += "vet_neg='" + _strTesNeg + "' AND ";
                s += "vet_cau='" + sCau + "' AND ";
                s += "vet_day=" + _clsFun.DaySql(dDay) + " AND ";
                s += "vet_ora='" + sOra + "' AND ";
                s += "vet_pos='00' AND ";
                s += "vet_sco='00000' AND ";
                s += "vet_fid='" + _strTesCod + "'";
                DataTable tVet = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlSta);

                if(tVet.Rows.Count > 0)
                {

                }
                else
                { 
                    x = tVet.NewRow();
                    x["vet_usr"] = "";
                    x["vet_cau"] = sCau;
                    x["vet_neg"] = _strTesNeg;
                    x["vet_day"] = dDay;
                    x["vet_ora"] = sOra;
                    x["vet_pos"] = "00";
                    x["vet_sco"] = "00000";
                    x["vet_fid"] = _strTesCod;      // lblTesCod.Text;
                    x["vet_imp"] = 0;
                    x["vet_pun"] = Convert.ToDecimal(txtPun.Text);
                    x["vet_sct"] = 0;
                    x["vet_scp"] = 0;
                    x["vet_art"] = 0;
                    s = _clsFun.SqlInsertRow(TABSTAVET, tVet, x);
                    _clsFun.SqlWrite(s, _strConSqlSta);
                }

                s = "SELECT * FROM GesNegVep WHERE ";
                s += "vep_neg='" + _strTesNeg + "' AND ";
                s += "vep_cau='" + sCau + "' AND ";
                s += "vep_day=" + _clsFun.DaySql(dDay) + " AND ";
                s += "vep_ora='" + sOra + "' AND ";
                s += "vep_pos='00' AND ";
                s += "vep_sco='00000' AND ";
                //s += "vep_cod='" + sCod +  "' AND ";
                s += "vep_ean='" + _clsDef.PUNTITESSERAX + "' AND ";
                s += "vep_fid='" + _strTesCod + "'";
                DataTable tVep = _clsFun.FillTabSql(TABSTAVEP, s, false, _strConSqlSta);

                if (tVep.Rows.Count > 0)
                {
                    s = "UPDATE GesNegVep SET ";
                    s += "vep_cod='" + sCod + "', ";
                    s += "vep_tip='" + sTip + "', ";
                    s += "vep_imp=" + txtPun.Text + " ";
                    s += "WHERE ";
                    s += "vep_neg='" + _strTesNeg + "' AND ";
                    s += "vep_cau='" + sCau + "' AND ";
                    s += "vep_day=" + _clsFun.DaySql(dDay) + " AND ";
                    s += "vep_ora='" + sOra + "' AND ";
                    s += "vep_pos='00' AND ";
                    s += "vep_sco='00000' AND ";
                    //s += "vep_cod='" + sTip + "' AND ";
                    s += "vep_ean='" + _clsDef.PUNTITESSERAX + "' AND ";
                    s += "vep_fid='" + _strTesCod + "'";
                    _clsFun.SqlWrite(s, _strConSqlSta);
                }
                else
                {
                    x = tVep.NewRow();
                    x["vep_cau"] = sCau;
                    x["vep_neg"] = _strTesNeg;
                    x["vep_day"] = dDay;
                    x["vep_ora"] = sOra;
                    x["vep_pos"] = "00";
                    x["vep_sco"] = "00000";
                    x["vep_cod"] = sCod;
                    x["vep_tri"] = "3";
                    x["vep_tip"] = sTip;
                    x["vep_off"] = "";                                  //20181004 Non ricordo a cosa serve
                    x["vep_imp"] = Convert.ToDecimal(txtPun.Text);
                    x["vep_val"] = 0;
                    x["vep_ppo"] = "";
                    x["vep_ord"] = "";
                    x["vep_ean"] = _clsDef.PUNTITESSERAX;
                    x["vep_fid"] = _strTesCod;
                    s = _clsFun.SqlInsertRow(TABSTAVEP, tVep, x);
                    _clsFun.SqlWrite(s, _strConSqlSta);
                }
            }
        }
    }
}
