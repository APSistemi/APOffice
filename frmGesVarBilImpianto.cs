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
    public partial class frmGesVarBilImpianto : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public Boolean _bolOk = false;
        public Boolean _bolAll = false;
        public string _strReb = "";

        private string _strConSql = "";

        public frmGesVarBilImpianto()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
        }

        private void frmGesVarBilImpianto_Load(object sender, EventArgs e)
        {
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmGesVarBilImpianto_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
                Esci();
        }    
        private void Esci()
        {
            this.Close();
        }

        private void FillDati()
        {
            string s = "SELECT * FROM TabRepBilance ORDER BY tab_des";
            DataTable t = _clsFun.FillTabSql("TabRepBilance", s, false, _strConSql);

            DataRow x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbReb.DataSource = t;
            cmbReb.DisplayMember = "tab_des";
            cmbReb.ValueMember = "tab_cod";
            cmbReb.SelectedValue = "";
        }

        private void cmbReb_SelectionChangeCommitted(object sender, EventArgs e)
        {
            _strReb = cmbReb.SelectedValue.ToString();
        }

        private void btnFill_Click(object sender, EventArgs e)
        {
            _bolOk = true;
            if (rdbAll.Checked)
                _bolAll = true;
            Esci();
        }
    }
}
