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
    public partial class frmUtyTab : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public DataTable _tabTab = new DataTable("tabTab");
        public string _strSql = "";
        public string _strTip = "";
        public string _strCod = "";
        public string _strRes = "";

        public frmUtyTab()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmUtyTab_Load(object sender, EventArgs e)
        {
            SetDgv1();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmUtyTab_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
                Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            if (_strTip == "LotNaz")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "tab_des";
                cTbc.Name = "Descrizione";
                cTbc.Width = 260;
                cTbc.ValueType = typeof(string);
                //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                cTbc.ReadOnly = true;
                dgv1.Columns.Add(cTbc);
            }
            else
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "div_arf";
                cTbc.Name = "Codice fornitore";
                cTbc.Width = 70;
                cTbc.ValueType = typeof(string);
                //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                cTbc.ReadOnly = true;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "div_ard";
                cTbc.Name = "Descrizione";
                cTbc.Width = 260;
                cTbc.ValueType = typeof(string);
                //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                cTbc.ReadOnly = true;
                dgv1.Columns.Add(cTbc);
            }
        }

        private void FillDati()
        {
            DataView v = new DataView((DataTable)_tabTab, _strSql, "", DataViewRowState.CurrentRows);
            dgv1.DataSource = v;
            lblCnt.Text = v.Count.ToString();
        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Scelto();
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            Scelto();
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Scelto();
        }

        private void dgv1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                Scelto();
        }

        private void dgv1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
                Scelto();
        }

        private void Scelto()
        {
            if (_strTip != "")
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    if (_strTip == "LotNaz")
                    {
                        _strCod = Convert.ToString(x["tab_cod"]);
                        _strRes = Convert.ToString(x["tab_des"]);
                        Esci();
                    }
                }
            }
        }

    }
}
