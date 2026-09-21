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
    public partial class frmUtyTessere : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private const string TABANA = "AnaTessere";
        private string _strConSql = "";

        public frmUtyTessere()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
        }

        private void frmUtyTessere_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            FillTabs();
            SetDgv1();
        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tes_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 150;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tes_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 180;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tes_gru";
            cTbc.Name = "Gruppo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);
        }

        private void FillTabs()
        {
            string p = "TabFidGruppi";
            string s = "SELECT * FROM " + p + " WHERE tab_ann=0";
            DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
            //x = t.NewRow();
            //x["tab_cod"] = "";
            //x["tab_des"] = "  Non definito";
            //t.Rows.InsertAt(x, 0);
            cmbTesGru.DataSource = t;
            cmbTesGru.DisplayMember = "tab_des";
            cmbTesGru.ValueMember = "tab_cod";
            //cmbTesGru..SelectedValue = "";
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnFill_Click(object sender, EventArgs e)
        {
            if (txtIni.Text == "") 
                MessageBox.Show("Prefisso tessera non definito!");
            else if (!_clsFun.Numerico(txtQta.Text))
                MessageBox.Show("Q.tà tessere da generare non definito!");
            else if (Convert.ToInt32(txtQta.Text) == 0)
                MessageBox.Show("Q.tà tessere da generare a ZERO!");
            else
                FillTes();
        }

        private void FillTes()
        {
            string s = "SELECT * FROM " + TABANA + " WHERE tes_cod >= '" + txtIni.Text + "'" ;
            DataTable t = _clsFun.FillTabSql(TABANA, s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = t.Columns["tes_cod"];
            t.PrimaryKey = keys;

            DataTable tTmp = t.Clone();

            DataRow x;
            DataRow[] j;

            Int32 iQta = Convert.ToInt32(txtQta.Text);
            Int32 iIni = Convert.ToInt32(txtIni.Text.Substring(9,3));
            string sIni = txtIni.Text.Substring(0, 9);

            for(int i = iIni; i <= iIni + iQta; i++)
            {
                s = sIni + i.ToString("000");
                s = s + new clsCtrlCodici().FindMod10Digit(s);

                j = t.Select("tes_cod='" + s + "'");
                if (j.Length > 0)
                {
                    MessageBox.Show("Tessera " + s + " già presente!");
                    break;
                }
                x = tTmp.NewRow();
                x["tes_cod"] = s;
                x["tes_des"] = "Nuova";
                x["tes_gru"] = cmbTesGru.SelectedValue.ToString();
                tTmp.Rows.Add(x);

            }

            dgv1.DataSource = tTmp;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource == null || ((DataTable)dgv1.DataSource).Rows.Count == 0)
                MessageBox.Show("Non ci sono tessere da salvare!");
            else
                Salva();
        }

        private void Salva()
        {
            string s = "";
            DataTable t = (DataTable)dgv1.DataSource;

            foreach(DataRow y in t.Rows)
            {
                s = _clsFun.SqlInsertRow(TABANA, t, y);
                _clsFun.SqlWrite(s, _strConSql);
            }

            MessageBox.Show("Fine lavoro!");
            t.Clear();
        }
    }
}
