using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    public partial class frmGesTabEcr : Form
    {
        private const string TABLIV = "TabEcrLv";
        private const string TABLV1 = "TabEcrLv1";
        private const string TABLV2 = "TabEcrLv2";
        private const string TABLV3 = "TabEcrLv3";

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private DataSet _dasGen = new DataSet();
        private string _strConSql = "";
        public string _strRes = "";
        public Boolean _bolCho = false;

        public frmGesTabEcr()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
        }

        private void frmGesTabEcr_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            SetDgv(dgv1);
            SetDgv(dgv2);
            SetDgv(dgv3);
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmGesTabEcr_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            Salva();
            this.Close();
        }

        private void SetDgv(DataGridView dgv)
        {
            dgv.AutoGenerateColumns = false;
            //dgv.VirtualMode = true;
            //dgv.Dock = DockStyle.Fill;
            dgv.AllowUserToAddRows = false;
            dgv.ReadOnly = false;
            dgv.AllowUserToDeleteRows = false;
            //dgv.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 30;
            if (dgv.Name == "dgv3")
                cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            dgv.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            dgv.Columns.Add(cTbc);

            if (dgv.Name == "dgv3")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "tab_tip";
                cTbc.Name = "Tipo";
                cTbc.Width = 30;
                cTbc.MaxInputLength = 1;
                cTbc.ValueType = typeof(string);
                dgv.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "tab_liv";
                cTbc.Name = "Livello";
                cTbc.Width = 30;
                cTbc.MaxInputLength = 1;
                cTbc.ValueType = typeof(string);
                dgv.Columns.Add(cTbc);
            }

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "tab_ann";
            cCbc.Name = "Ann.";
            cCbc.Width = 40;
            dgv.Columns.Add(cCbc);
        }

        private void FillDati()
        {
            DataColumn c;
            string p = "";
            string s = "";
            DataTable t;

            for (int i = 1; i <= 3; i++)
            {
                p = "TabEcrLv" + i.ToString();
                s = "SELECT * FROM " + p + " ORDER BY tab_cod";
                t = _clsFun.FillTabSql(p, s, false, _strConSql);

                foreach (DataColumn dc in t.Columns)
                {
                    if (dc.DataType == typeof(string))
                    {
                        dc.MaxLength = -1;
                    }
                }

                c = new DataColumn();
                c.DataType = System.Type.GetType("System.String");
                c.ColumnName = "TabMdy";
                c.Caption = "Modify";
                c.MaxLength = 1;
                c.ReadOnly = false;
                c.DefaultValue = "";
                t.Columns.Add(c);

                _dasGen.Tables.Add(t);
            }

            dgv1.DataSource = new DataView(_dasGen.Tables[TABLV1], "", "", DataViewRowState.CurrentRows);
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            string sTab = TABLIV + btn.Name.Substring(6, 1);
            if (sTab == TABLV1)
                NewRec(sTab, dgv1);
            if (sTab == TABLV2)
                NewRec(sTab, dgv2);
            if (sTab == TABLV3)
                NewRec(sTab, dgv3);
            //if (sTab == TABLV4)
            //    GesCfo("NEW");
        }

        private void NewRec(string strTab, DataGridView dgv)
        {
            string sDef = _clsFun.ParGet(clsDefine.enuParametri.ParEcrDefault, _strConSql);
            string[] aDef = !string.IsNullOrEmpty(sDef) ? sDef.Split(',') : new string[0];

            int i = 0;
            string sWhe = "";
            string sLv1 = "";
            string sLv2 = "";
            if (dgv.Name == "dgv1")
            {
                sWhe = "tab_des=''";
            }
            if (dgv.Name == "dgv2")
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm == null || cm.Position < 0 || cm.Position >= cm.List.Count)
                {
                    MessageBox.Show("Selezionare prima un livello 1 (Settore)!", "ATTENZIONE", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                DataRowView r = cm.List[cm.Position] as DataRowView;
                sLv1 = Convert.ToString(r.Row["tab_cod"]);
                sWhe = "tab_des='' AND tab_lv1='" + sLv1 + "'";
            }
            if (dgv.Name == "dgv3")
            {
                CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
                if (cm == null || cm.Position < 0 || cm.Position >= cm.List.Count)
                {
                    MessageBox.Show("Selezionare prima un livello 2 (Famiglia)!", "ATTENZIONE", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                DataRowView r = cm.List[cm.Position] as DataRowView;
                sLv1 = Convert.ToString(r.Row["tab_lv1"]);
                sLv2 = Convert.ToString(r.Row["tab_cod"]);
                sWhe = "tab_des='' AND tab_lv1='" + sLv1 + "' AND tab_lv2='" + sLv2 + "'";
            }

            DataView v = new DataView(_dasGen.Tables[strTab], sWhe, "", DataViewRowState.CurrentRows);
            if (v.Count > 0)
            {
                MessageBox.Show("Ci sono righe senza descrizione!");
                i = 0;

                foreach (DataGridViewRow r in dgv.Rows)
                {
                    if (r.Cells[0].Value != null && r.Cells[0].Value.ToString() == ((v[0]["tab_cod"]).ToString()))
                    {
                        dgv.CurrentCell = r.Cells[0];
                        dgv.Rows[i].Selected = true;
                        i++;
                        break;
                    }
                }

                if (dgv.RowCount > 0)
                {
                    dgv.CurrentCell = dgv.Rows[dgv.RowCount - 1].Cells[1];
                    dgv.Rows[dgv.RowCount - 1].Selected = true;
                    dgv.BeginEdit(true);
                }
            }
            else
            {
                string sCod = "";

                sWhe = "";
                if (aDef.Length > 0 && !string.IsNullOrEmpty(aDef[0]))
                    sWhe = "tab_cod <> '" + aDef[0] + "'";

                if (dgv.Name == "dgv2")
                {
                    sWhe += (sWhe.Length > 0 ? " AND " : "") + "tab_lv1='" + sLv1 + "'";
                }
                if (dgv.Name == "dgv3")
                {
                    sWhe += (sWhe.Length > 0 ? " AND " : "") + "tab_lv1='" + sLv1 + "' AND tab_lv2='" + sLv2 + "'";
                }

                DataView vExisting = new DataView(_dasGen.Tables[strTab], sWhe, "tab_cod DESC", DataViewRowState.CurrentRows);

                int targetLen = 3;
                int maxNum = 0;
                foreach (DataRowView drv in vExisting)
                {
                    string codeStr = Convert.ToString(drv["tab_cod"]).Trim();
                    if (!string.IsNullOrEmpty(codeStr))
                    {
                        if (codeStr.Length > targetLen) targetLen = codeStr.Length;
                        if (int.TryParse(codeStr, out int numVal))
                        {
                            if (numVal > maxNum) maxNum = numVal;
                        }
                    }
                }

                sCod = (maxNum + 1).ToString().PadLeft(targetLen, '0');

                DataRow y = _dasGen.Tables[strTab].NewRow();
                y["tab_cod"] = sCod;
                y["tab_des"] = "";
                y["tab_ann"] = false;
                y["TabMdy"] = "S";
                if (dgv.Name == "dgv2")
                {
                    y["tab_lv1"] = sLv1;
                }
                if (dgv.Name == "dgv3")
                {
                    y["tab_lv1"] = sLv1;
                    y["tab_lv2"] = sLv2;
                }

                _dasGen.Tables[strTab].Rows.Add(y);

                if (dgv.Name == "dgv1")
                    dgv.DataSource = _dasGen.Tables[TABLV1];
                else
                {
                    if (dgv.Name == "dgv2")
                        v = new DataView(_dasGen.Tables[TABLV2], "tab_lv1='" + sLv1 + "'", "tab_cod", DataViewRowState.CurrentRows);
                    else if (dgv.Name == "dgv3")
                        v = new DataView(_dasGen.Tables[TABLV3], "tab_lv1='" + sLv1 + "' AND tab_lv2='" + sLv2 + "'", "tab_cod", DataViewRowState.CurrentRows);
                    dgv.DataSource = v;
                }

                dgv.ClearSelection();

                foreach (DataGridViewRow r in dgv.Rows)
                {
                    if (r.Cells[0].Value != null && r.Cells[0].Value.ToString() == sCod)
                    {
                        dgv.CurrentCell = r.Cells[1];
                        r.Selected = true;
                        break;
                    }
                }

                if (dgv.CurrentCell != null)
                {
                    dgv.BeginEdit(true);
                }
            }
        }

        private void dgv1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            CurrencyManager cm = this.dgv1.BindingContext[this.dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;

            if (cm.Position >= 0)
            {

                DataRowView r = cm.List[e.RowIndex] as DataRowView;
                string sCod = (string)r.Row["tab_cod"];

                DataView v = new DataView(_dasGen.Tables[TABLV2], "tab_lv1='" + sCod + "'", "tab_cod", DataViewRowState.CurrentRows);
                dgv2.DataSource = v;

                dgv3.DataSource = null;
            }
        }

        private void dgv2_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            CurrencyManager cm = this.dgv2.BindingContext[this.dgv2.DataSource, this.dgv2.DataMember] as CurrencyManager;

            if (cm.Position >= 0)
            {

                DataRowView r = cm.List[e.RowIndex] as DataRowView;
                string sCod = Convert.ToString(r.Row["tab_cod"]);
                string sLv1 = Convert.ToString(r.Row["tab_lv1"]);

                DataView v = new DataView(_dasGen.Tables[TABLV3], "tab_lv1='" + sLv1 + "' AND tab_lv2='" + sCod + "'", "tab_cod", DataViewRowState.CurrentRows);

                dgv3.DataSource = v;
            }
        }

        private void dgv3_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            CurrencyManager cm = dgv3.BindingContext[dgv3.DataSource, dgv3.DataMember] as CurrencyManager;

            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[e.RowIndex] as DataRowView;
                string sCod = Convert.ToString(r.Row["tab_cod"]);
                string sLv2 = Convert.ToString(r.Row["tab_lv2"]);
                string sLv1 = Convert.ToString(r.Row["tab_lv1"]);

                //DataView v = new DataView(_dasGen.Tables[TABLV4], "int_con='" + sLv1 + sLv2 + sCod + "'", "int_con", DataViewRowState.CurrentRows);
                //dgv4.DataSource = v;
            }
        }

        private void dgv_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            CurrencyManager cm = dgv.BindingContext[dgv.DataSource, dgv.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;
                x["TabMdy"] = "S";
            }
        }

        private void Salva()
        {
            try
            {
                if (dgv1 != null) dgv1.EndEdit();
                if (dgv2 != null) dgv2.EndEdit();
                if (dgv3 != null) dgv3.EndEdit();
            }
            catch { }

            DataRow[] j;
            string sSql = "";

            for (int i = 1; i <= 3; i++)
            {
                string p = TABLIV + i.ToString();
                string s = "SELECT * FROM " + p;
                DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);

                DataTable tTmp = t.Clone();

                foreach (DataRow y in _dasGen.Tables[p].Rows)
                {
                    if ((string)y["TabMdy"] != "")
                    {
                        s = "";
                        if (!DBNull.Value.Equals(y["tab_des"]))
                            s = (string)y["tab_des"];
                        y["tab_des"] = s;

                        sSql = "tab_cod='" + y["tab_cod"] + "' AND ";
                        for (int i2 = 1; i2 < i; i2++)
                            sSql += "tab_lv" + i2.ToString() + "='" + y["tab_lv" + i2.ToString()] + "' AND ";

                        sSql = sSql.Substring(0, sSql.Length - 5);
                        j = t.Select(sSql);
                        s = "";
                        if (j.Length == 0)
                            s = _clsFun.SqlInsertRow(p, t, y);
                        else
                        {
                            ArrayList aWhe = new ArrayList();
                            aWhe.Add("tab_cod");
                            for (int i2 = 1; i2 < i; i2++)
                                aWhe.Add("tab_lv" + i2.ToString());
                            ArrayList aExl = new ArrayList();

                            s = _clsFun.SqlUpdRow(p, t, j[0], y, aWhe, aExl);
                        }
                        if (s != "")
                        {
                            _clsFun.SqlWrite(s, _strConSql);
                            tTmp.ImportRow(y);
                        }
                    }
                }

                s = _clsFun.ParGet(clsDefine.enuParametri.Par016PathDivNegozi, _strConSql);
                if (tTmp != null && tTmp.Rows.Count > 0 && s != "")
                    new clsVariazioni().DivNegTabelle(s, tTmp);
            }
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel1();
        }

        private void Excel1()
        {
            string s = "SELECT ";
            s += "TabEcrLv3.tab_lv1 AS Lv1Cod, ";
            s += "TabEcrLv1.tab_des AS Lv1Des, ";
            s += "TabEcrLv3.tab_lv2 AS Lv2Cod, ";
            s += "TabEcrLv2.tab_des AS Lv2Des, ";
            s += "TabEcrLv3.tab_cod AS Lv3Cod, ";
            s += "TabEcrLv3.tab_des AS Lv3Des ";
            s += "FROM TabEcrLv3 ";
            s += "INNER JOIN TabEcrLv2 ON TabEcrLv3.tab_lv1 = TabEcrLv2.tab_lv1 AND TabEcrLv3.tab_lv2 = TabEcrLv2.tab_cod ";
            s += "INNER JOIN TabEcrLv1 ON TabEcrLv2.tab_lv1 = TabEcrLv1.tab_cod ";
            s += "ORDER BY Lv1Des, Lv2Des, TabEcrLv3.tab_des";

            DataTable t = _clsFun.FillTabSql("TabEcr", s, false, _strConSql);


            string sTit = "CLASSIFICAZIONE MERCEOLOGICA AL " + DateTime.Today.ToString("dd/MM7yyyy");
            string sFil = "";

            s = "";
            s += "Lv1Cod, Settore, 30, StringLiteral;";
            s += "Lv1Des, Descrizione, 130, StringLiteral;";
            s += "Lv2Cod, Famiglia, 30, StringLiteral;";
            s += "Lv2Des, Descrizione, 130, StringLiteral;";
            s += "Lv3Cod, SottoFam, 30, StringLiteral;";
            s += "Lv3Des, Descrizione, 130, StringLiteral;";

            for (int i = 0; i <= 100; i++)
            {
                sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_ecr.xls";
                if (File.Exists(sFil))
                {
                    try
                    {
                        File.Delete(sFil);
                        break;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
                else
                    break;
            }

            DataSet ds = new DataSet();

            DataTable tTit = (new clsExcel()).FillTitStr(s, "TabTit");

            ds.Tables.Add(tTit);
            ds.Tables.Add(t);

            (new clsExcel()).exportToExcel(ds, sFil, sTit, "", true);

            ds.Tables.Remove(t);
            ds.Clear();
            ds.Dispose();
        }

        private void dgv3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if(_bolCho)
            {

                if (e.ColumnIndex == 1)
                {
                    CurrencyManager cm = dgv3.BindingContext[dgv3.DataSource, dgv3.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {
                        DataRowView r = cm.List[dgv3.CurrentRow.Index] as DataRowView;
                        DataRow x = r.Row;

                        _strRes = (string)x["tab_lv2"] + (string)x["tab_lv1"] + (string)x["tab_cod"];
                        Esci();
                    }
                }
            }
        }

    }
}
