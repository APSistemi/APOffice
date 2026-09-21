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
    public partial class frmUser : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string TABUSR = "TabUtenti";
        private string USRCNF = "CNF";
        private string USRCLI = "CLI";

        public bool _bolSta = false;
        public string _strTip = "";
        public string _strConSql = "";
        public string _strUsrCod = "";
        public string _strUsrDes = "";
        public string _strPwd = "";

        public frmUser()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
        }

        private void frmUser_Load(object sender, EventArgs e)
        {
            String s = FillTab();
            if (_strTip == USRCNF || _strTip == USRCLI)
            {
                cmbUsr.Enabled = false;
                cmbUsr.SelectedValue = s;
            }
            ((Control)txtPwd).Select();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void Esci()
        {
            if (CtrlPwd())
            {
                if (_strTip != USRCNF && _strTip != USRCLI)
                    _clsFun.FileIni("W", clsDefine.enuIni.IniUserLast, cmbUsr.SelectedValue.ToString());
                this.Close();
            }
            else if (MessageBox.Show("Dati non corretti, abbandoni?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                this.Close();
        }

        private string FillTab()
        {
            string s = "SELECT * FROM TabUtenti WHERE tab_ann=0 ORDER BY tab_des";
            if (_strTip == USRCNF)
                s = "SELECT * FROM TabUtenti WHERE tab_liv='01'";

            DataTable t = _clsFun.FillTabSql(TABUSR, s, false, _strConSql);
            cmbUsr.DataSource = t;
            cmbUsr.DisplayMember = "tab_des";
            cmbUsr.ValueMember = "tab_cod";
            cmbUsr.SelectedValue = _clsFun.FileIni("R", clsDefine.enuIni.IniUserLast, "");

            if (_strTip == USRCNF && t.Rows.Count > 0)
                s = (string)t.Rows[0]["tab_cod"];

            return s;
        }

        private void txtPwd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Return || !CtrlPwd())
                return;
            Esci();
        }

        private void cmbUsr_SelectionChangeCommitted(object sender, EventArgs e)
        {
            ((Control)txtPwd).Select();
        }
        private bool CtrlPwd()
        {
            if (_strTip == USRCLI)
            {
                if(_strPwd == txtPwd.Text)
                    _bolSta = true;
            }
            else
            {
                DataRow[] j = ((DataTable)cmbUsr.DataSource).Select("tab_cod='" + cmbUsr.SelectedValue.ToString() + "'");
                if (j.Length > 0 && ((string)j[0]["tab_pwd"]).Trim() == txtPwd.Text)
                {
                    _strUsrCod = (string)j[0]["tab_cod"];
                    _strUsrDes = (string)j[0]["tab_des"];
                    _bolSta = true;
                }
            }
            return _bolSta;
        }

    }
}
