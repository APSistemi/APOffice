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
    public partial class frmGesStatScontrini : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public DateTime _dayDay = new DateTime(2050, 1, 1, 0, 0, 0);
        public String _strNeg = "";

        public frmGesStatScontrini()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmStatScontrini_Load(object sender, EventArgs e)
        {
            this.Text += " al " + _dayDay.ToString("dd/MM/yyyy");

            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
            SetDgv1();
            FillDati();
        }

        private void frmStatScontrini_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void FillDati()
        {
            decimal dVen = 0;
            decimal dPro = 0;

            string s = "SELECT * FROM GesNegVet WHERE ";
            s += "GesNegVet.vet_neg='" + _strNeg + "' AND ";
            s += "GesNegVet.vet_day >= " + _clsFun.DaySql(_dayDay) + " AND ";
            s += "GesNegVet.vet_day <= " + _clsFun.DaySql(_dayDay) + " ";
            DataTable t = _clsFun.FillTabSql("Vet", s, false, _strConSqlSta);

            foreach(DataRow y in t.Rows)
            {
                if ((string)y["vet_cau"] == "SCO")
                    dVen += (decimal)y["vet_imp"];
                else if ((string)y["vet_cau"] == "PRO")
                    dPro += (decimal)y["vet_imp"];
            }

            dgv1.DataSource = t;

            lblTotVen.Text = dVen.ToString("#####,##0.00");
            lblTotPro.Text = dPro.ToString("#####,##0.00");
        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = true;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_cau";
            cTbc.Name = "Causale";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "ven_day";
            //cTbc.Name = "Data";
            //cTbc.Width = 50;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_neg";
            cTbc.Name = "Negozio";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_pos";
            cTbc.Name = "Cassa";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_ora";
            cTbc.Name = "Ora";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_sco";
            cTbc.Name = "Scontrino";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_fid";
            cTbc.Name = "Tessera";
            cTbc.Width = 120;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_imp";
            cTbc.Name = "Importo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_sct";
            cTbc.Name = "Sconto su totale";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.MaxInputLength = 80;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_pun";
            cTbc.Name = "Punti";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_scp";
            cTbc.Name = "Punti articoli";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
        }

        private void dgv1_Click(object sender, EventArgs e)
        {
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                string sNeg = (string)r.Row["vet_neg"];
                string sCau = (string)r.Row["vet_cau"];
                DateTime dDay = (DateTime)r.Row["vet_day"];
                string sOra = (string)r.Row["vet_ora"];
                string sPos = (string)r.Row["vet_pos"];
                string sSco = (string)r.Row["vet_sco"];

                frmGesStatScoDettaglio f = new frmGesStatScoDettaglio();
                f._strNeg = sNeg;
                f._strCau = sCau;
                f._dayDay = dDay;
                f._strOra = sOra;
                f._strPos = sPos;
                f._strSco = sSco;
                f.ShowDialog();

                if (f._tabVet != null && ((DataTable)f._tabVet).Rows.Count > 0)
                {
                    r.Row["vet_pun"] = f._tabVet.Rows[0]["vet_pun"];
                    r.Row["vet_imp"] = f._tabVet.Rows[0]["vet_imp"];
                }
            }
        }
        private void btnTip_Click(object sender, EventArgs e)
        {
            frmGesStaTipoPagamenti f = new frmGesStaTipoPagamenti();
            f._dayDay = _dayDay;
            f.ShowDialog();
        }

        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable t = (DataTable)dgv1.DataSource;

            string sPar = this.Text;

            new clsStampe().PrintStaScontrini(t, sPar);
        }

    }
}
