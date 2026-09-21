using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace APOffice
{
    public partial class frmGesVarTerm : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public string _strDiv = "";
        public string _strTer = "";
        public Boolean _bolOk = false;

        private string _strConSql = "";

        public frmGesVarTerm()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
        }

        private void frmGesVarTerm_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            //SetDgv1();
            FillTabs();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private void btnFill_Click(object sender, EventArgs e)
        {
            if (cmbDiv.DataSource == null)
                MessageBox.Show("Tipo di divulgazione non definita!");
            else if (cmbTer.DataSource == null)
                MessageBox.Show("Tipo di terminalino non definito!");
            else
            {
                _strDiv = cmbDiv.SelectedValue.ToString();
                _strTer = cmbTer.SelectedValue.ToString();
                _bolOk = true;
                Esci();
            }
        }

        //private void SetDgv1()
        //{
        //    dgv1.AutoGenerateColumns = false;
        //    //dgv1.VirtualMode = true;
        //    //dgv1.Dock = DockStyle.Fill;
        //    dgv1.AllowUserToAddRows = false;
        //    dgv1.ReadOnly = false;
        //    dgv1.AllowUserToDeleteRows = false;
        //    //dgv1.DisplayedRowCount() = true;

        //    DataGridViewTextBoxColumn cTbc;
        //    //DataGridViewCheckBoxColumn cCbc;
        //    //DataGridViewComboBoxColumn cCmb;

        //    cTbc = new DataGridViewTextBoxColumn();
        //    cTbc.DataPropertyName = "tab_msg";
        //    cTbc.Name = "Messaggio";
        //    cTbc.Width = 140;
        //    cTbc.ValueType = typeof(string);
        //    cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        //    cTbc.ReadOnly = true;
        //    dgv1.Columns.Add(cTbc);
        //}

        private void FillTabs()
        {
            //DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabForImport";
            s = "SELECT * FROM " + p + " WHERE tab_ann=0";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            if (t.Rows.Count == 0)
                MessageBox.Show("Tabella terminalino non configurata!");
            else
            {
                //x = t.NewRow();
                //x["tab_cod"] = "";
                //x["tab_des"] = "  Non definito";
                //t.Rows.InsertAt(x, 0);
                cmbDiv.DataSource = t;
                cmbDiv.DisplayMember = "tab_des";
                cmbDiv.ValueMember = "tab_cod";
                cmbDiv.SelectedIndex = 0;
            }

            p = "TabTerm";
            s = "SELECT * FROM " + p + " WHERE tab_ann=0";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            if (t.Rows.Count == 0)
                MessageBox.Show("Tabella terminalino non configurata!");
            else
            {
                //x = t.NewRow();
                //x["tab_cod"] = "";
                //x["tab_des"] = "  Non definito";
                //t.Rows.InsertAt(x, 0);
                cmbTer.DataSource = t;
                cmbTer.DisplayMember = "tab_des";
                cmbTer.ValueMember = "tab_cod";
                cmbTer.SelectedIndex = 0;
            }
        }

        //private void FillDati()
        //{
        //    string sFor = cmbDiv.SelectedValue.ToString();
        //    string s = "SELECT * FROM " + TABLISACQ + " WHERE lia_for='" + sFor + "'";
        //    DataTable t = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);
        //    dgv1.DataSource = t;
        //}
    }
}
