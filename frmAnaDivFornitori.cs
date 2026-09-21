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
    public partial class frmAnaDivFornitori : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";
        private string _strConSqlMdb = "";

        public frmAnaDivFornitori()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmGesDivFornitori1_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            _strConSqlMdb = _clsFun.ConMdb(Path.GetDirectoryName(_clsDef.FILEINI) + "\\ApForTracciati.mdb");
            SetDgv1();
            SetDgv2();
            FillTab("");
            //FillDati();
        }
        private void frmAnaDivFornitori_KeyDown(object sender, KeyEventArgs e)
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
            Salva();
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
            cTbc.DataPropertyName = "trt_for";
            cTbc.Name = "Ente";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trt_cod";
            cTbc.Name = "Tracciato";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trt_pth";
            cTbc.Name = "Path";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trt_fil";
            cTbc.Name = "File";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trt_leg";
            cTbc.Name = "Legami";
            cTbc.Width = 160;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "trt_reg";
            cTbc.Name = "Regola";
            cTbc.Width = 160;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
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
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tra_fld";
            cTbc.Name = "Field";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tra_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tra_pos";
            cTbc.Name = "Posizione";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tra_len";
            cTbc.Name = "Lunghezza";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tra_reg";
            cTbc.Name = "Regola";
            cTbc.Width = 300;
            cTbc.MaxInputLength = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "tra_ann";
            cCbc.Name = "Ann.";
            cCbc.Width = 40;
            dgv2.Columns.Add(cCbc);
        }

        private void FillTab(string strTab)
        {
            if (strTab == "" || strTab == "AnaForDivTipi")
            {
                string p = "AnaForDivTipi";
                string s = "SELECT * FROM AnaForDivTipi WHERE trk_sta='" + _clsDef.STAATT + "' ORDER BY trk_des";
                //DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
                DataTable t = _clsFun.FillTabMdb(p, s, false, _strConSqlMdb);

                DataRow x = t.NewRow();
                x["trk_cod"] = "";
                x["trk_des"] = "  Non definito";
                t.Rows.InsertAt(x, 0);
                cmbTip.DataSource = t;
                cmbTip.DisplayMember = "trk_des";
                cmbTip.ValueMember = "trk_cod";
                cmbTip.SelectedValue = "";
            }
        }

        private void FillDati(string strTip)
        {
            string s = "SELECT ";
            s += "AnaForDivTestate.* ";
            //s += "AnaFornitori.for_des AS TrtFod ";
            s += "FROM AnaForDivTestate ";
            //s += "LEFT JOIN AnaFornitori ON AnaForDivTestate.trt_for= AnaFornitori.for_cod ";
            s += "WHERE trt_for='" + strTip + "' ";
            s += "ORDER BY trt_for, trt_cod";
            //DataTable t = _clsFun.FillTabSql("AnaForDivTestate", s, false, _strConSql);
            DataTable t = _clsFun.FillTabMdb("AnaForDivTestate", s, false, _strConSqlMdb);
            dgv1.DataSource = t;
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            string sTipCod = _clsDef.CODNEW;
            string sTipDes = "";
            string sTipSta = _clsDef.STAATT;

            string s = cmbTip.SelectedValue.ToString();

            DataTable tTip = (DataTable)cmbTip.DataSource;

            DataRow[] j = tTip.Select("trk_cod='" + s + "'");

            frmAnaDivForTesta f = new frmAnaDivForTesta();

            if (j.Length > 0)
                f._rowTrk = j[0];

            f.ShowDialog();

            if ((string)f._rowTrk["trk_des"] != cmbTip.Text)
            {
                s = (string)f._rowTrk["trk_cod"];
                FillTab("AnaForDivTipi");
                cmbTip.SelectedValue = s;
            }
            FillDati(cmbTip.SelectedValue.ToString());
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                DataRow[] jTrk = ((DataTable)cmbTip.DataSource).Select("trk_cod='" + cmbTip.SelectedValue.ToString() + "'");

                if (jTrk.Length > 0)
                {
                    //string sFor = (string)jTrt[0]["trt_for"];
                    //string sFod = (string)jTrt[0]["trt_des"];
                    //string sSta = (string)jTrt[0]["trt_sta"];

                    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {
                        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                        DataRow x = r.Row;

                        //string sFor = (string)x["trt_for"];
                        //string sFod = cmbTip.Text;
                        //string sCod = (string)x["trt_cod"];
                        //string sSta = (string)x["trt_sta"];

                        frmAnaDivForTesta f = new frmAnaDivForTesta();
                        f._rowTrk = jTrk[0];
                        //f._strTip = sFor;
                        //f._strTid = sFod;
                        //f._strSta = sSta;
                        //f._strCod = sCod;
                        f._rowTrt = x;
                        f.ShowDialog();

                        if ((string)f._rowTrk["trk_des"] != cmbTip.Text)
                        {
                            string s = (string)f._rowTrk["trk_cod"];
                            FillTab("AnaForDivTipi");
                            cmbTip.SelectedValue = s;
                        }

                        FillDati(cmbTip.SelectedValue.ToString());

                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void dgv1_RowLeave(object sender, DataGridViewCellEventArgs e)
        {
            Salva();
        }

        private void dgv1_CurrentCellChanged(object sender, EventArgs e)
        {
            try
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string sFor = (string)x["trt_for"];
                    string sCod = (string)x["trt_cod"];

                    DataRow[] j;
                    DataRow y;

                    string s = "SELECT * FROM AnaForDivDettaglio WHERE tra_for='" + sFor + "' AND tra_cod='" + sCod + "'  ORDER BY tra_fld";
                    //DataTable tTra = _clsFun.FillTabSql("AnaForDivDettaglio", s, false, _strConSql);
                    DataTable tTra = _clsFun.FillTabMdb("AnaForDivDettaglio", s, false, _strConSqlMdb);

                    DataTable tTmp = tTra.Clone();

                    s = "SELECT * FROM GesDivForArticoli WHERE div_for='" + _clsDef.COD06X + "'";
                    DataTable t = _clsFun.FillTabSql("GesDivForArticoli", s, false, _strConSql);
                    //DataTable t = _clsFun.FillTabMdb("GesDivForArticoli", s, false, _strConSqlMdb);

                    foreach (DataColumn c in t.Columns)
                    {
                        y = tTmp.NewRow();
                        y["tra_fld"] = c.ColumnName;

                        j = tTra.Select("tra_fld='" + c.ColumnName + "'");
                        if (j.Length > 0)
                        {
                            y["tra_des"] = j[0]["tra_des"];
                            y["tra_pos"] = j[0]["tra_pos"];
                            y["tra_len"] = j[0]["tra_len"];
                            y["tra_reg"] = j[0]["tra_reg"];
                            y["tra_ann"] = j[0]["tra_ann"];
                        }

                        y["tra_for"] = sFor;
                        y["tra_cod"] = sCod;
                        tTmp.Rows.Add(y);
                    }

                    dgv2.DataSource = tTmp;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void Salva()
        {
            if (dgv1.DataSource != null)
            {
                //string sFor = cmbTip.SelectedValue.ToString();

                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    DataRow[] j;

                    string sFor = (string)x["trt_for"];
                    string sCod = (string)x["trt_cod"];

                    string s = "SELECT * FROM AnaForDivDettaglio WHERE tra_for='" + sFor + "' AND tra_cod='" + sCod + "'";

                    DataTable tTra = _clsFun.FillTabMdb("AnaForDivDettaglio", s, false, _strConSqlMdb);
                    DataColumn[] keys = new DataColumn[3];
                    keys[0] = tTra.Columns["tra_cod"];
                    keys[1] = tTra.Columns["tra_for"];
                    keys[2] = tTra.Columns["tra_fld"];
                    tTra.PrimaryKey = keys;

                    DataTable t = (DataTable)dgv2.DataSource;

                    foreach (DataRow y in t.Rows)
                    {
                        //y["tra_for"] = sFor;
                        //y["tra_cod"] = sCod;

                        if (!DBNull.Value.Equals(y["tra_cod"]) && (string)y["tra_cod"] == sCod) // && !DBNull.Value.Equals(y["tra_len"]) && (decimal)y["tra_len"] != 0)
                        {
                            s = (string)y["tra_fld"];

                            j = tTra.Select("tra_fld='" + y["tra_fld"] + "'");
                            if (j.Length == 0)
                                s = _clsFun.SqlInsertRow("AnaForDivDettaglio", tTra, y);
                            else
                                s = _clsFun.SqlUpdRowIdx("AnaForDivDettaglio", tTra, j[0], y, null);
                            if (s != "")
                                _clsFun.MdbWrite(s, _strConSqlMdb);
                        }
                    }
                }
            }
        }

        private void cmbTip_SelectionChangeCommitted(object sender, EventArgs e)
        {
            Salva();
            FillDati(cmbTip.SelectedValue.ToString());
        }

        private void cmbTip_Enter(object sender, EventArgs e)
        {
            //Salva();
        }

        private void cmbTip_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            //Salva();
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            Salva();

            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                frmAnaDivForTest f = new frmAnaDivForTest();
                f._rowTes = x;
                f._tabTab = (DataTable)dgv2.DataSource;
                f.ShowDialog();
            }
        }

        private void importTracciatoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource == null || dgv1.Rows.Count == 0 || dgv1.CurrentRow == null)
            {
                MessageBox.Show("Selezionare prima un tracciato di destinazione dall'elenco superiore.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm != null && cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                string sNewFor = (string)x["trt_for"];
                string sNewCod = (string)x["trt_cod"];

                frmSeekDivForTestate f = new frmSeekDivForTestate();
                f.ShowDialog();

                if (f._strRes != "")
                {
                    if (MessageBox.Show("Procedi?", "IMPORTAZIONE TRACCIATO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        string s = "";
                        DataRow[] j;
                        string[] a = f._strRes.Split(',');

                        string sFor = a[0];
                        string sCod = a[1];

                        s = "SELECT * FROM AnaForDivDettaglio WHERE tra_for='" + sFor + "' AND tra_cod='" + sCod + "'";
                        DataTable t = _clsFun.FillTabMdb("AnaForDivDettaglio", s, false, _strConSqlMdb);

                        DataTable tTrc = (DataTable)dgv2.DataSource;

                        foreach (DataRow y in tTrc.Rows)
                        {
                            j = t.Select("tra_fld='" + y["tra_fld"] + "'");
                            if (j.Length > 0)
                            {
                                foreach (DataColumn c in tTrc.Columns)
                                {
                                    y[c.ColumnName] = j[0][c.ColumnName];
                                }
                            }

                            y["tra_for"] = sNewFor;
                            y["tra_cod"] = sNewCod;
                        }
                    }
                }

            }
        }
    }
}
