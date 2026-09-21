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
    public partial class frmUtyInvNegozio : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        public frmUtyInvNegozio()
        {
            InitializeComponent();
            _strConSql = _clsFun.ConSql(""); 
        }

        private void frmUtyInvNegozio_Load(object sender, EventArgs e)
        {
            FillTabs();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FillTabs()
        {
            string p = "TabNegozi";
            string s = "SELECT * FROM TabNegozi ORDER BY tab_cod";
            DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
            //DataRow x = t.NewRow();
            //x["tab_cod"] = "";
            //x["tab_des"] = "  Tutti";
            //t.Rows.InsertAt(x, 0);
            cmbIntNeg.DataSource = t;
            cmbIntNeg.DisplayMember = "tab_des";
            cmbIntNeg.ValueMember = "tab_cod";
            cmbIntNeg.SelectedValue = "001";
        }

        private void FillDati()
        {
            string s = "";
            s = "SELECT * FROM GesInvTestate LEFT OUTER JOIN TabReparti ON GesInvTestate.int_rep = TabReparti.tab_cod ";
            s += "ORDER BY int_num,int_day DESC";
            DataTable t = _clsFun.FillTabSql("GesInvTestate", s, false, _strConSql);

            dgv1.DataSource = t;
        }

        private void dgv1_CurrentCellChanged(object sender, EventArgs e)
        {
            try
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string sNeg = (string)x["int_neg"];
                    string sNum = (string)x["int_num"];

                    DataRow[] j;
                    DataRow y;

                    //string s = "SELECT * FROM GesInventario WHERE (inv_neg='" + sNeg + "' OR inv_neg=null) AND inv_num='" + sNum + "'";
                    string s = "SELECT * FROM GesInventario WHERE inv_num='" + sNum + "'";
                    DataTable tInv = _clsFun.FillTabSql("GesInventario", s, false, _strConSql);

                    dgv2.DataSource = tInv;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Confermi l'aggiornamento con codice negozio '" + cmbIntNeg.SelectedValue.ToString() + "'?", "Cambio codice negozio", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                AggNegozio();
            }
        }

        private void AggNegozio()
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                string sNum = (string)x["int_num"];

                string sNeg = cmbIntNeg.SelectedValue.ToString();
                string s = "UPDATE GesInventario SET inv_neg='" + sNeg + "' WHERE inv_num='" + sNum + "'";
                _clsFun.SqlWrite(s, _strConSql);

                s = "UPDATE GesInvTestate SET int_neg='" + sNeg + "' WHERE int_num='" + sNum + "'";
                _clsFun.SqlWrite(s, _strConSql);

                MessageBox.Show("Fine");

            }
        }
    }
}
