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
    public partial class frmGesTermImport : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        public string _strTrmCho = "";

        public frmGesTermImport()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
        }

        private void frmGesTerminalino_Load(object sender, EventArgs e)
        {
            SetDgv1();
            SetDgv2();
            dgv2.DataSource = new clsGenTabTmp().TabTmpTermDett("DivRow");
            FillTabs();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmGesTerminalino_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void Esci()
        {

            //_strTrmCho = (string)x["tmp_day"] + ";" + (string)x["tmp_tip"];

            //string sDay = (string)x["tmp_day"];
            //string sTip = (string)x["tmp_tip"];

            //string s = "UPDATE TmpDiv SET tmp_sta='N' WHERE tmp_tip='" + sTip + "' AND tmp_day='" + sDay + "'";
            //_clsFun.SqlWrite(s, _strConSql);

            //Esci();

            if (dgv1.DataSource != null)
            {
                DataTable t = (DataTable)dgv1.DataSource;
                foreach(DataRow y in t.Rows)
                {
                    if ((string)y["tab_cho"] == "S")
                    {
                        string sDay = (string)y["tmp_day"];
                        string sTip = (string)y["tmp_tip"];

                        _strTrmCho += sDay + "-" + sTip + ";";

                        string s = "UPDATE TmpDiv SET tmp_sta='N' WHERE tmp_tip='" + sTip + "' AND tmp_day='" + sDay + "'";
                        _clsFun.SqlWrite(s, _strConSql);
                    }
                }
            }

            this.Close();
        }

        private void btnFill_Click(object sender, EventArgs e)
        {
            FillDati();
        }
        
        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            Scelto();
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
            cTbc.DataPropertyName = "tab_cho";
            cTbc.Name = "Scelto";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_day";
            cTbc.Name = "Data";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_ora";
            cTbc.Name = "ora";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_tip";
            cTbc.Name = "Tipo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_sta";
            cTbc.Name = "Stato";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void SetDgv2()
        {
            dgv2.AutoGenerateColumns = false;
            //dgv2.VirtualMode = true;
            //dgv2.Dock = DockStyle.Fill;
            dgv2.AllowUserToAddRows = false;
            dgv2.ReadOnly = false;
            dgv2.AllowUserToDeleteRows = false;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trm_ean";
            cTbc.Name = "Barcode";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trm_art";
            cTbc.Name = "Codice";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trm_ard";
            cTbc.Name = "Descrizione";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trm_qta";
            cTbc.Name = "Q.tà";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);
        }

        private void FillTabs()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            //string sEcr = _clsFun.ParGet(clsDefine.enuParametri.ParEcrDefault, _strConSql);

            t = new clsGenTabTmp().TabTmpSta("TabStato");
            //t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "N";
            x["tab_des"] = "Elaborati";
            t.Rows.Add(x);
            x = t.NewRow();
            x["tab_cod"] = "S";
            x["tab_des"] = "Nuovi";
            t.Rows.Add(x);
            cmbSta.DataSource = t;
            cmbSta.DisplayMember = "tab_des";
            cmbSta.ValueMember = "tab_cod";
            cmbSta.SelectedValue = "S";

            s = "SELECT * FROM TabNegozi ORDER BY tab_cod";
            t = _clsFun.FillTabSql("TabNegozi", s, false, _strConSql);
            //t.Rows.InsertAt(x, 0);
            cmbNeg.DataSource = t;
            cmbNeg.DisplayMember = "tab_des";
            cmbNeg.ValueMember = "tab_cod";
            cmbNeg.SelectedValue = "001";

            t = new clsGenTabTmp().TabTmpSta("TabStato");
            //t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "ALL";
            x["tab_des"] = "Tutti";
            t.Rows.Add(x);
            x = t.NewRow();
            x["tab_cod"] = "ETI";
            x["tab_des"] = "Etichette";
            t.Rows.Add(x);
            x = t.NewRow();
            x["tab_cod"] = "INV";
            x["tab_des"] = "Inventario";
            t.Rows.Add(x);
            x = t.NewRow();
            x["tab_cod"] = "ORD";
            x["tab_des"] = "Ordini";
            t.Rows.Add(x);
            cmbTip.DataSource = t;
            cmbTip.DisplayMember = "tab_des";
            cmbTip.ValueMember = "tab_cod";
            cmbTip.SelectedValue = "ALL";
        }

        private void FillDati()
        {
            string s = "";

            string sSta = cmbSta.SelectedValue.ToString();
            string sNeg = cmbNeg.SelectedValue.ToString();
            string sTip = cmbTip.SelectedValue.ToString();

            s = "SELECT ";
            s += "tmp_day, ";
            s += "tmp_tip, ";
            s += "tmp_sta, ";
            s += "tmp_neg ";
            s += "FROM TmpDiv ";
            s += "WHERE tmp_sta='" + sSta + "' AND tmp_neg='" + sNeg + "' ";
            if (sTip != "ALL")
                s += " AND tmp_tip='" + sTip + "'";
            s += "GROUP BY tmp_day, tmp_tip, tmp_sta, tmp_neg ";
            s += "ORDER BY tmp_day DESC";
            DataTable t = _clsFun.FillTabSql("TmpDiv", s, false, _strConSql);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "tab_day",
                Caption = "Data",
                ReadOnly = false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tab_ora",
                Caption = "Ora",
                MaxLength = 5,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tab_cho",
                Caption = "Scelto",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            foreach(DataRow y in t.Rows)
            {
                s = (string)y["tmp_day"];
                y["tab_day"] = _clsFun.Str2Day(s.Substring(0, 8));
                y["tab_ora"] = s.Substring(8, 2) + ":" + s.Substring(10, 2);
            }

            dgv1.DataSource = t;
        }

        private void dgv1_CurrentCellChanged(object sender, EventArgs e)
        {
            FillRighe();
        }

        private void FillRighe()
        {
            try
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;
                    string sDay = (string)x["tmp_day"];

                    DataTable tDiv = (DataTable)dgv2.DataSource;
                    tDiv.Clear();

                    string s = "SELECT * FROM TmpDiv WHERE tmp_day='" + sDay + "'";
                    DataTable t = _clsFun.FillTabSql("TmpDiv", s, false, _strConSql);

                    foreach(DataRow y in t.Rows)
                    {
                        s = (string)y["tmp_tmp"];

                        string[] a = s.Split(';');

                        if(a.Length > 7)
                        {
                            DataRow xx = tDiv.NewRow();
                            xx["trm_ean"] = a[3];
                            xx["trm_art"] = a[4];
                            xx["trm_ard"] = a[5];

                            s = a[6].Trim().Replace(".",",");

                            Console.WriteLine("zzzzzzzzzzz");

                            xx["trm_qta"] = Convert.ToDecimal(s);
                            tDiv.Rows.Add(xx);
                        }
                    }

                    dgv2.DataSource = tDiv;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void Scelto()
        {
            if (dgv1.DataSource != null)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    if ((string)x["tab_cho"] == "S")
                        x["tab_cho"] = "";
                    else
                        x["tab_cho"] = "S";
                }
            }
        }

    }
}
