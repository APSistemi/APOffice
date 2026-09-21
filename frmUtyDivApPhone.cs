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
    public partial class frmUtyDivApPhone : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        public frmUtyDivApPhone()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
        }

        private void frmUtyDivApPhone_Load(object sender, EventArgs e)
        {
            FillTabs();

            SetDgv0();
            SetDgv1();
            dgv1.DataSource = new clsGenTabTmp().TabTmpTermDett("DivRow");
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmUtyDivApPhone_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void btnFill_Click(object sender, EventArgs e)
        {
            string s = cmbTip.Text;

            if(s == "")
                MessageBox.Show("Tipo non definito!", "CONTROLLO DATI", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
                FillDati();
        }

        private void SetDgv0()
        {
            dgv0.AutoGenerateColumns = false;
            //dgv0.VirtualMode = true;
            //dgv0.Dock = DockStyle.Fill;
            dgv0.AllowUserToAddRows = false;
            dgv0.ReadOnly = false;
            dgv0.AllowUserToDeleteRows = false;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_cho";
            cTbc.Name = "Selezione";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv0.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_day";
            cTbc.Name = "Data";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv0.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_ora";
            cTbc.Name = "ora";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv0.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_tip";
            cTbc.Name = "Tipo";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv0.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_sta";
            cTbc.Name = "Attivo";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv0.Columns.Add(cTbc);
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
            cTbc.DataPropertyName = "trm_ean";
            cTbc.Name = "Barcode";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trm_art";
            cTbc.Name = "Codice";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trm_ard";
            cTbc.Name = "Descrizione";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trm_qta";
            cTbc.Name = "Q.tà";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trm_tmp";
            cTbc.Name = "Riga";
            cTbc.Width = 270;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillTabs()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            string sEcr = _clsFun.ParGet(clsDefine.enuParametri.ParEcrDefault, _strConSql);

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
            x = t.NewRow();
            x["tab_cod"] = "ADV";
            x["tab_des"] = "Var. prezzi";
            t.Rows.Add(x);
            cmbTip.DataSource = t;
            cmbTip.DisplayMember = "tab_des";
            cmbTip.ValueMember = "tab_cod";
            cmbTip.SelectedValue = "ALL";
        }

        private void oldFillDati()
        {
            string sTip = cmbTip.Text;

            string s = "SELECT * FROM TmpDiv WHERE tmp_tip='" + sTip + "'";
            DataTable t = _clsFun.FillTabSql("TmpDiv", s, false, _strConSql);

            DataTable tTmp = t.Clone();

            foreach (DataRow y in t.Rows)
            {
                string sDti = dtpDti.Value.ToString("yyyyMMdd");
                string sDtf = dtpDtf.Value.ToString("yyyyMMdd");

                string sDay = ((string)y["tmp_day"]).Substring(0, 8);

                if (Convert.ToInt64(((string)y["tmp_day"]).Substring(0, 8)) >= Convert.ToInt64(sDti) && Convert.ToInt64(((string)y["tmp_day"]).Substring(0, 8)) <= Convert.ToInt64(sDtf))
                    tTmp.ImportRow(y);

            }

            dgv1.DataSource = tTmp;

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
            s += "WHERE ";
            s += "tmp_sta='" + sSta + "' AND ";
            s += "tmp_neg='" + sNeg + "' AND ";
            s += "LEFT(tmp_day, 8) >='" + dtpDti.Value.ToString("yyyyMMdd") + "' AND ";
            s += "LEFT(tmp_day, 8) <='" + dtpDtf.Value.ToString("yyyyMMdd") + "' ";
            if (sTip != "ALL")
                s += " AND tmp_tip='" + sTip + "'";
            s += "GROUP BY tmp_day, tmp_tip, tmp_sta, tmp_neg ";
            s += "ORDER BY tmp_day";
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
                Caption = "Cho",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            foreach (DataRow y in t.Rows)
            {
                s = (string)y["tmp_day"];
                y["tab_day"] = _clsFun.Str2Day(s.Substring(0, 8));
                y["tab_ora"] = s.Substring(8, 2) + ":" + s.Substring(10, 2);
            }

            dgv0.DataSource = t;
        }

        private void oldSetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;
            dgv1.SelectionMode = DataGridViewSelectionMode.CellSelect;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_sta";
            cTbc.Name = "Stato";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.ToolTipText = "S da elaborare, N elaborato - Doppio click per modificare";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_tip";
            cTbc.Name = "Tipo";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_day";
            cTbc.Name = "Data";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_tmp";
            cTbc.Name = "Dati";
            cTbc.Width = 500;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void dgv0_DoubleClick(object sender, EventArgs e)
        {
            CurrencyManager cm = dgv0.BindingContext[dgv0.DataSource, dgv0.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv0.CurrentRow.Index] as DataRowView;
                if ((string)r.Row["tab_cho"] == "S")
                    r.Row["tab_cho"] = "";
                else
                    r.Row["tab_cho"] = "S";
            }
        }

        private void chkAll_Click(object sender, EventArgs e)
        {
            if (dgv0.DataSource != null)
            {
                string sSta = "S";
                if (chkAll.Checked)
                    sSta = "N";
                DataTable t = (DataTable)dgv0.DataSource;
                foreach (DataRow y in t.Rows)
                    y["tmp_sta"] = sSta;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (cmbTip.Text == "Tutti")
                MessageBox.Show("Tipo divulgazione da impostare non selezionata!", "Divulgazione ApPhone", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else if (dgv0.DataSource == null || ((DataTable)dgv0.DataSource).Rows.Count == 0)
                MessageBox.Show("Non sono presenti righe!", "Divulgazione ApPhone", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                string sSta = cmbSta.SelectedValue.ToString();
                string sTip = cmbTip.SelectedValue.ToString();

                DataTable t = (DataTable)dgv0.DataSource;

                string sKey = (string)t.Rows[0]["tmp_tip"];

                Boolean b = true;

                foreach (DataRow y in t.Rows)
                {
                    if ((string)y["tmp_tip"] != sKey)
                    {
                        b = false;
                        break;
                    }
                }

                if (b)
                {
                    if (MessageBox.Show("Confermi l'aggiornamento delle divulgazioni selezionate?", "AGGIORNAMENTO DIVULGAZIONI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        Salva(sSta, sTip);
                }
                else
                    MessageBox.Show("Deve essere selezionato solo un tipo di divulgazione!", "Divulgazione ApPhone", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Salva(string strSta, string strTip)
        {
            string s = "";

            //string sTip = cmbTip.Text;

            DataTable t = (DataTable)dgv0.DataSource;

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if ((string)y["tab_cho"] == "S")
                {
                    s = "UPDATE TmpDiv SET ";
                    s += "tmp_sta='" + strSta + "', ";
                    s += "tmp_tip='" + strTip + "' ";
                    s += "WHERE tmp_day='" + Convert.ToString(y["tmp_day"] + "'");
                    _clsFun.SqlWrite(s, _strConSql);
                }
            }

            FillDati();
        }


        private void btnCanc_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Confermi la cancellazione delle divulgazioni selezionate?", "CANCELLAZIONE DIVULGAZIONI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Elimina();
            }
        }

        private void Elimina()
        {
            string s = "";

            DataTable t = (DataTable)dgv0.DataSource;

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if ((string)y["tab_cho"] == "S")
                {
                    s = "DELETE TmpDiv WHERE tmp_day='" + Convert.ToString(y["tmp_day"] + "'");
                    _clsFun.SqlWrite(s, _strConSql);
                }
            }

            FillDati();
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.ColumnIndex == 0)
            {
                Console.WriteLine("xxx");

                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    if((string)x["tmp_sta"] == "S")
                        x["tmp_sta"] = "N";
                    else
                        x["tmp_sta"] = "S";
                }
            }
        }

        private void invioASedeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InvioASede();
        }

        private void InvioASede()
        {
            string s = "";

            clsVariazioni clsVar = new clsVariazioni();

            string sPar = _clsFun.ParGet(clsDefine.enuParametri.Par036Div2Sede, _strConSql);

            string[] a = sPar.Split(',');

            Boolean b = false;

            if (a.Length > 3 && a[3].Trim() == "S")
            {
                string sPath = a[0].Trim();
                string sNeg = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");

                sPar = sPath + "|" + sNeg;

                DataTable t = (DataTable)dgv0.DataSource;

                foreach (DataRow y in t.Rows)
                {
                    if ((string)y["tab_cho"] == "S")
                    {
                        b = true;

                        s = clsVar.DivNeg2SedeApPhoneTmp(sPar, y);
                        if (s != "")
                        {
                            b = false;
                            MessageBox.Show("Percorso non attivo!", "DIVULAGAZIONE A SEDE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }

            if(b)
                MessageBox.Show("Invio effettuato!", "DIVULGAZIONE A SEDE", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void dgv0_CurrentCellChanged(object sender, EventArgs e)
        {
            FillRighe();
        }

        private void FillRighe()
        {
            try
            {
                DataTable tDiv = (DataTable)dgv1.DataSource;
                tDiv.Clear();

                CurrencyManager cm = dgv0.BindingContext[dgv0.DataSource, dgv0.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv0.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;
                    string sDay = (string)x["tmp_day"];

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
                            xx["trm_qta"] = Convert.ToDecimal(a[6].Replace(".",","));
                            xx["trm_tmp"] = s;
                            tDiv.Rows.Add(xx);
                        }
                    }

                    dgv1.DataSource = tDiv;
                    lblCnt.Text = tDiv.Rows.Count.ToString("#0");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

    }
}
