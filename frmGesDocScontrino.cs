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
    public partial class frmGesDocScontrino : Form
    {
        private const string TABSTAVET = "GesNegVet";
        private const string TABSTAVEN = "GesNegVen";
        private const string TABSTAVEP = "GesNegVep";

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public string _strNeg = "";

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public DataTable _tabSco = new DataTable(TABSTAVET);
        public DataTable _tabVep = new DataTable(TABSTAVEP);

        public frmGesDocScontrino()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);; 
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
        }

        private void frmGesDocScontrino_Load(object sender, EventArgs e)
        {
            SetDgv1();
            dtpDay.Select();

            //dtpDay.Value = new DateTime(2018, 2, 1);
            //nudPos.Value = 3;
            //nudNum.Value = 490;

            //dtpDay.Value = new DateTime(2017,3,23);
            //nudPos.Value = 1;
            //nudNum.Value = 1;
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmGesDocScontrino_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            string s = "";

            TimeSpan d = dtpDay.Value.Date.Subtract(DateTime.Today);

            if (d.Days == 0)
            {
                DataTable tCnf = _clsQry.ConfSeek("", "VEN");
                if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16((object)clsDefine.enuPos.posNcr745x).ToString("00"))
                {
                    string sDay = dtpDay.Value.ToString("yyyyMMdd");
                    string sPos = nudPos.Value.ToString("00");
                    string sNum = nudNum.Value.ToString("00000");

                    s = sDay + "|" + sPos + "|" + sNum;
                    new frmGesChiusure().ChiuNcr745x(tCnf, "C:\\Server\\Data", s);
                }
                else if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16((object)clsDefine.enuPos.posDitron).ToString("00"))
                {
                    string sDay = dtpDay.Value.ToString("yyyyMMdd");
                    string sPos = nudPos.Value.ToString("00");
                    string sNum = nudNum.Value.ToString("00000");

                    s = sDay + "|" + sPos + "|" + sNum;
                    new frmGesChiusure().DitronLeggiScontrino(tCnf, s);
                    //    new frmGesChiusure().ChiuDitron(tCnf, "C:\\NetPos\\SH", s);
                }

            }

            string sCau = "SCO";
            if (rdbCauPro.Checked)
                sCau = "PRO";

            FillDati(sCau);
            dgv1.Select();
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
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_day";
            cTbc.Name = "Data";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_ora";
            cTbc.Name = "Ora";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##.##";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_pos";
            cTbc.Name = "Cassa";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            //dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_sco";
            cTbc.Name = "Scontrino";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_imp";
            cTbc.Name = "Importo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#,##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_pun";
            cTbc.Name = "Punti";
            cTbc.Width = 35;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#,##0";
        }

        private void FillDati(string strCau)
        {
            string s = "";

            string sNeg = "";
            if (_strNeg == "")
                sNeg = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
            else
                sNeg = _strNeg;

            DateTime dDay = dtpDay.Value;
            string sPos = nudPos.Value.ToString("00");
            string sNum = nudNum.Value.ToString("00000");

            s = "SELECT * FROM GesNegVet WHERE ";
            s += "GesNegVet.vet_cau = '" + strCau + "' AND ";
            s += "GesNegVet.vet_day = " + _clsFun.DaySql(dDay) + " AND ";
            s += "GesNegVet.vet_pos = '" + sPos + "' AND ";
            s += "GesNegVet.vet_sco = '" + sNum + "'";
            if (sNeg != "")
                s += " AND GesNegVet.vet_neg = '" + sNeg + "'";
            DataTable t = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlSta);

            dgv1.DataSource = t;
        }

        private void dgv1_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Return)
                Scelto();
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            Scelto();
        }

        private void Scelto()
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                DateTime dDay = (DateTime)x["vet_day"];
                string sNeg = (string)x["vet_neg"];
                string sOra = (string)x["vet_ora"];
                string sPos = (string)x["vet_pos"];
                string sNum = (string)x["vet_sco"];
                decimal dImp = (decimal)x["vet_imp"];

                string s = "";
                s = "SELECT * ";
                s += "FROM GesNegVet INNER JOIN GesNegVen ON ";
                s += "GesNegVet.vet_neg = GesNegVen.ven_neg AND ";
                s += "GesNegVet.vet_cau = GesNegVen.ven_cau AND ";
                s += "GesNegVet.vet_day = GesNegVen.ven_day AND ";
                s += "GesNegVet.vet_ora = GesNegVen.ven_ora AND ";
                s += "GesNegVet.vet_pos = GesNegVen.ven_pos AND ";
                s += "GesNegVet.vet_sco = GesNegVen.ven_sco ";
                s += "WHERE ";
                s += "(GesNegVet.vet_neg = " + sNeg + ") AND ";
                s += "(GesNegVet.vet_day = " + _clsFun.DaySql(dDay) + ") AND ";
                s += "(GesNegVet.vet_ora = '" + sOra + "') AND ";
                s += "(GesNegVet.vet_pos = '" + sPos + "') AND ";
                s += "(GesNegVet.vet_sco = '" + sNum + "') ";
                s += "ORDER BY ven_idx";

                _tabSco = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlSta);

                s = "SELECT * ";
                s += "FROM GesNegVet INNER JOIN GesNegVep ON ";
                s += "GesNegVet.vet_neg = GesNegVep.vep_neg AND ";
                s += "GesNegVet.vet_cau = GesNegVep.vep_cau AND ";
                s += "GesNegVet.vet_day = GesNegVep.vep_day AND ";
                s += "GesNegVet.vet_ora = GesNegVep.vep_ora AND ";
                s += "GesNegVet.vet_pos = GesNegVep.vep_pos AND ";
                s += "GesNegVet.vet_sco = GesNegVep.vep_sco ";
                s += "WHERE ";
                s += "(GesNegVet.vet_neg = " + sNeg + ") AND ";
                s += "(GesNegVet.vet_day = " + _clsFun.DaySql(dDay) + ") AND ";
                s += "(GesNegVet.vet_ora = '" + sOra + "') AND ";
                s += "(GesNegVet.vet_pos = '" + sPos + "') AND ";
                s += "(GesNegVet.vet_sco = '" + sNum + "')";

                _tabVep = _clsFun.FillTabSql(TABSTAVEP, s, false, _strConSqlSta);

                if(_tabVep.Rows.Count == 0)
                {
                    x = _tabVep.NewRow();
                    x["vep_imp"] = dImp;
                    x["vep_tip"] = "PAG";
                    _tabVep.Rows.Add(x);
                }

                Esci();
            }
        }

    }
}
