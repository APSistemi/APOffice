using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    public partial class frmSeekDivForTestate : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";
        private string _strConSqlMdb = "";
        public string _strRes = "";

        public frmSeekDivForTestate()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmAnaDivForTest_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            _strConSqlMdb = _clsFun.ConMdb(Path.GetDirectoryName(_clsDef.FILEINI) + "\\ApForTracciati.mdb");
            SetDgv1();

            FillData();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmAnaDivForTest_KeyDown(object sender, KeyEventArgs e)
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

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trt_tip";
            cTbc.Name = "Tipo";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trk_des";
            cTbc.Name = "Fornitore";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trt_cod";
            cTbc.Name = "Tracciato";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "trt_pth";
            //cTbc.Name = "Path";
            //cTbc.Width = 200;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = true;
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trt_fil";
            cTbc.Name = "File";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "trt_leg";
            //cTbc.Name = "Legami";
            //cTbc.Width = 160;
            //cTbc.ValueType = typeof(string);
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "trt_reg";
            //cTbc.Name = "Regola";
            //cTbc.Width = 160;
            //cTbc.ValueType = typeof(string);
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc);
        }

        private void FillData()
        {
            string s = "";

            s = "SELECT ";
            s += "AnaForDivTestate.trt_for, ";
            s += "AnaForDivTestate.trt_cod, ";
            s += "AnaForDivTipi.trk_des, ";
            s += "AnaForDivTestate.trt_tip, ";
            s += "AnaForDivTestate.trt_fil ";
            s += "FROM AnaForDivTestate ";
            s += "INNER JOIN AnaForDivTipi ON AnaForDivTestate.trt_for = AnaForDivTipi.trk_cod;";

            DataTable t = _clsFun.FillTabMdb("AnaForDivTestate", s, false, _strConSqlMdb);

            dgv1.DataSource = t;
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                string sFor = (string)x["trt_for"];
                string sCod = (string)x["trt_cod"];

                _strRes = sFor + "," + sCod;

                Esci();

            }
        }

    }
}
