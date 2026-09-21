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
    public partial class frmAnaCfoDocumenti : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public string _strCod = "";
        public string _strDes = "";
        public string _strCfo = "";

        //private const string TABANA = "AnaClienti";
        private string _strConSql = "";

        public frmAnaCfoDocumenti()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0); ;
        }

        private void frmAnaCfoDocumenti_Load(object sender, EventArgs e)
        {
            dtpDti.Value = new DateTime(DateTime.Now.Year, 1, 1);
            _strConSql = _clsFun.ConSql("");
            SetDgv1();
            FillTabs();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmAnaCfoDocumenti_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void btnEstrai_Click(object sender, EventArgs e)
        {
            FillDati();
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

            //dgv1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //dgv1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            //s += "fat_yfa AS DocYea, ";
            //s += "fat_nfa AS DocNum, ";
            //s += "fat_ndo AS DocNdo, ";
            //s += "fat_ddo AS DocDdo, ";
            //s += "fat_cfo AS DocCfo, ";
            //s += "AnaClienti.cli_des AS DocCfd, ";
            //s += "fat_ann AS DocAnn ";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocYea";
            cTbc.Name = "Anno";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocTdo";
            cTbc.Name = "T";
            cTbc.Width = 23;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocFat";
            cTbc.Name = "Fatturato";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.Visible = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocNeg";
            cTbc.Name = "Negozio";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocCfd";
            cTbc.Name = "Cliente/Fornitore";
            cTbc.Width = 170;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocNdo";
            cTbc.Name = "Documento";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocDdo";
            cTbc.Name = "Data";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "MovImp";
            cTbc.Name = "Importo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(decimal);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0.00";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocTpg";
            cTbc.Name = "Pagamento";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocStd";
            cTbc.Name = "Stato";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocNo1";
            cTbc.Name = "Note";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.DataPropertyName = "DocAnn";
            cCbc.Name = "Annullato";
            cCbc.Width = 40;
            cCbc.ValueType = typeof(string);
            cCbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "FatEle";
            cTbc.Name = "Fattura Elettronica";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocNum";
            cTbc.Name = "Movimento";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocUbi";
            cTbc.Name = "Negozio/Sede";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);
        }
        
        private void FillTabs()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabDocTpd";
            s = "SELECT * FROM TabDocTpd WHERE tab_ann=0 ORDER BY tab_cod";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbTpd.DataSource = t;
            cmbTpd.DisplayMember = "tab_des";
            cmbTpd.ValueMember = "tab_cod";
            cmbTpd.SelectedValue = "";
        }

        private void FillDati()
        {
            string s = "";
            string sCfo = "";
            //string sTip = "";
            DataRow[] j;
            decimal d = 0;

            if (_strCfo == _clsDef.TIPCLI)
            {
                s = "SELECT ";
                s += "fat_ubi AS DocUbi, ";
                s += "fat_yfa AS DocYea, ";
                s += "fat_nfa AS DocNum, ";
                s += "fat_tdo AS DocTdo, ";
                s += "fat_ndo AS DocNdo, ";
                s += "fat_ddo AS DocDdo, ";
                s += "fat_neg AS DocNeg, ";
                s += "fat_cfo AS DocCfo, ";
                s += "fat_ele AS FatEle, ";
                s += "AnaClienti.cli_des AS DocCfd, ";
                s += "TabStato.tab_des AS DocStd, ";
                s += "TabPagamenti.tab_des AS DocTpg, ";
                s += "R.mov_imp + R.MovIva AS MovImp, ";
                s += "fat_no1 AS DocNo1, ";
                s += "fat_ann AS DocAnn ";
                s += "FROM (((GesFatTestate ";
                s += "LEFT OUTER JOIN AnaClienti ON GesFatTestate.fat_cfo = AnaClienti.cli_cod) ";
                s += "LEFT OUTER JOIN TabStato ON GesFatTestate.fat_sta = TabStato.tab_cod) ";
                s += "LEFT OUTER JOIN TabPagamenti ON GesFatTestate.fat_tpg = TabPagamenti.tab_cod) ";
                s += "LEFT JOIN (";

                s += "SELECT ";
                s += "mov_yfa, ";
                s += "mov_nfa, ";
                s += "SUM(mov_imp) AS mov_imp, ";
                s += "SUM(CASE WHEN tab_ali IS NOT NULL AND tab_ali > 0 AND mov_imp > 0 THEN mov_imp * tab_ali / 100 ELSE 0 END) AS MovIva ";
                s += "FROM GesMovimenti ";
                s += "LEFT OUTER JOIN TabIva ON GesMovimenti.mov_iva = TabIva.tab_cod ";
                //s += "WHERE (mov_ann = 0 AND mov_yfa='" + cmbYea.Text + "') ";
                s += "WHERE (mov_ann = 0) ";
                s += "GROUP BY mov_yfa, mov_nfa ";

                s += ") R ON GesFatTestate.fat_yfa = R.mov_yfa AND GesFatTestate.fat_nfa = R.mov_nfa ";

                s += "WHERE ";
                s += "fat_cfo ='" + _strCod + "' AND ";
                s += "GesFatTestate.fat_ddo >= " + _clsFun.DaySql(dtpDti.Value) + " AND ";
                s += "GesFatTestate.fat_ddo <= " + _clsFun.DaySql(dtpDtf.Value) + " ";

                if (!DBNull.Value.Equals(cmbTpd.SelectedValue) && cmbTpd.SelectedValue.ToString() != "")
                    s += "AND fat_tdo='" + cmbTpd.SelectedValue.ToString() + "' ";
                //s += "fat_tip='" + sTip + "' ";
                s += "ORDER BY fat_ddo, fat_ndo";
            }

            DataTable t = _clsFun.FillTabSql("DOC", s, false, _strConSql);

            dgv1.DataSource = t;

            d = 0;
            if (t.Columns.Contains("MovImp"))
            {
                foreach (DataRow y in t.Rows)
                {
                    if (!(Boolean)y["DocAnn"])
                    {
                        if ((string)y["DocTdo"] == "NA")
                            d -= (decimal)y["MovImp"];
                        else
                            d += (decimal)y["MovImp"];
                    }
                }
            }

            if (t.Rows.Count > 0)
            {
                int i = dgv1.Rows.Count - 1;
                if (i >= 0)
                    dgv1.CurrentCell = dgv1[3, i];
            }

            lblVal.Text = d.ToString("###,##0.00");
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable t = (DataTable)dgv1.DataSource;

            string sTit = "DOCUMENTI " + _strDes + " dal " + dtpDti.Value.ToShortDateString() + " al " + dtpDti.Value.ToShortDateString();
            string sFil = "";
            string sFoo = ",,,,'Totali'," + lblVal.Text.ToString().Replace(".", "").Replace(",", ".")+",,";

            string sFld = "";
            sFld += "DocYea, Anno, 30, StringLiteral;";
            sFld += "DocTdo, T.Documento, 70, StringLiteral;";
            sFld += "DocNdo, N.Documento, 30, StringLiteral;";
            sFld += "DocDdo, Data, 90, DateTime;";
            sFld += "DocTpg, T.pagamento, 30, StringLiteral;";
            sFld += "MovImp, Importo, 90, Decimal;";
            sFld += "DocStd, Stato, 90, StringLiteral;";
            sFld += "DocNo1, Note, 50, StringLiteral;";

            sFil = "DocCliente.xls";

            (new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);
        }


    }
}
