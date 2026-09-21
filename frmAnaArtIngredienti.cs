using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmAnaArtIngredienti : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private const string TABARTING = "AnaArtIngredienti";

        private string _strConSql = "";

        public string _strArtCod = "";
        //public string _strDes = "";
        //public string _strEan = "";
        //public decimal _decPes = 0;
        //public decimal _decTar = 0;
        //public decimal _decPrv = 0;
        //public decimal _decImp = 0;

        public frmAnaArtIngredienti()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmAnaArtIngredienti_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            SetDgv1();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Salva();
            Esci();
        }
        private void frmAnaArtIngredienti_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void nuovaRigaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable t = (DataTable)dgv1.DataSource;

            string sRow = "001";
            string sTip = "";

            if (t.Rows.Count > 0)
            {
                DataView v = new DataView(t, "", "ing_row DESC", DataViewRowState.CurrentRows);
                if (v.Count > 0)
                {
                    sRow = (Convert.ToDecimal(v[0]["ing_row"]) + 1).ToString("000");
                    sTip = (string)v[0]["ing_tip"];
                }
            }

            DataRow x = t.NewRow();
            x["ing_art"] = _strArtCod;
            x["ing_tip"] = sTip;
            x["ing_crt"] = "";
            x["ing_row"] = sRow;
            x["ing_txt"] = "";
            x["ing_ann"] = false;
            t.Rows.Add(x);
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
            cTbc.DataPropertyName = "ing_tip";
            cTbc.Name = "Tipo";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.ToolTipText = "BIL=bilancia, ETI=etichetta";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ing_crt";
            cTbc.Name = "Carattere";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.ToolTipText = "G=grassetto";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ing_row";
            cTbc.Name = "Riga";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ing_txt";
            cTbc.Name = "Testo";
            cTbc.Width = 500;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 200;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "ing_ann";
            cCbc.Name = "Ann.";
            cCbc.Width = 25;
            dgv1.Columns.Add(cCbc);
        }

        private void FillDati()
        {

            string s = "";

            DataTable t = _clsQry.SeekArtEan(_strArtCod, "");
            if (t.Rows.Count > 0)
            {
                DataRow r0 = t.Rows[0];
                string sDes = r0["tmp_ard"] != DBNull.Value ? Convert.ToString(r0["tmp_ard"]) : "";

                decimal dPes = 0m;
                if (r0["tmp_pne"] != DBNull.Value)
                    decimal.TryParse(Convert.ToString(r0["tmp_pne"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dPes);

                decimal dTar = 0m;
                if (r0["tmp_tar"] != DBNull.Value)
                    decimal.TryParse(Convert.ToString(r0["tmp_tar"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dTar);

                decimal dPrv = 0m;
                if (r0["tmp_prv"] != DBNull.Value)
                    decimal.TryParse(Convert.ToString(r0["tmp_prv"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dPrv);

                decimal dImp = dPrv * dPes;

                double dGsc = 0;
                if (r0["tmp_gsc"] != DBNull.Value)
                {
                    string sGscRaw = Convert.ToString(r0["tmp_gsc"]).Trim();
                    if (!string.IsNullOrEmpty(sGscRaw))
                    {
                        double.TryParse(sGscRaw.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dGsc);
                    }
                }

                string sEan = r0["tmp_ean"] != DBNull.Value ? Convert.ToString(r0["tmp_ean"]).Trim() : "";

                if (dTar > 0)
                    dTar = dTar / 1000;

                decimal dVal = 0;
                if (dPes > 0 && dPrv > 0)
                    dVal = (dPes - dTar) * dPrv;

                string sPve = "0000";
                if (dVal > 0)
                    sPve = (dVal * 100).ToString("0000");

                if (sEan != "" && sEan.Length >= 8)
                {
                    sEan = sEan.Substring(0, 8) + sPve;
                    sEan = sEan + new clsCtrlCodici().FindMod10Digit(sEan);
                }

                lblDes.Text = sDes;
                lblPes.Text = dPes.ToString("#0.00");
                lblTar.Text = dTar.ToString("#0.000");
                lblPrv.Text = dPrv.ToString("#0.00");
                lblImp.Text = dImp.ToString("#0.00");
                lblEan.Text = sEan;

                try
                {
                    dtpDsc.Value = DateTime.Now.AddDays(dGsc);
                }
                catch
                {
                    dtpDsc.Value = DateTime.Today;
                }
            }

            s = "";
            s += "SELECT * FROM AnaArtIngredienti WHERE ";
            s += "ing_art='" + _strArtCod + "' ";
            s += "ORDER BY ing_row";
            t = _clsFun.FillTabSql(TABARTING, s, false, _strConSql);

            dgv1.DataSource = t;
        }

        private void Salva()
        {
            if (dgv1.CurrentCell is DataGridViewCell)
            {
                dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }

            string s = "";
            DataRow x;
            DataRow[] j;

            s = "SELECT * FROM AnaArtIngredienti WHERE ";
            s += "ing_art='" + _strArtCod + "' ";
            s += "ORDER BY ing_row";
            DataTable tIng = _clsFun.FillTabSql(TABARTING, s, false, _strConSql);

            DataTable t = (DataTable)dgv1.DataSource;

            ArrayList aWhe = new ArrayList();
            aWhe.Add("ing_art");
            aWhe.Add("ing_row");
            ArrayList aExl = new ArrayList();

            foreach(DataRow y in t.Rows)
            {
                x = tIng.NewRow();
                x["ing_art"] = y["ing_art"];
                x["ing_tip"] = y["ing_tip"];
                x["ing_crt"] = y["ing_crt"];
                x["ing_row"] = y["ing_row"];
                x["ing_txt"] = y["ing_txt"];
                x["ing_ann"] = y["ing_ann"];

                j = tIng.Select("ing_row='" + y["ing_row"] + "'");
                if (j.Length > 0)
                    s = _clsFun.SqlUpdRow(TABARTING, tIng, j[0], y, aWhe, aExl);
                else
                    s = _clsFun.SqlInsertRow(TABARTING, tIng, y);

                if(s != "")
                    _clsFun.SqlWrite(s, _strConSql);
            }
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            DataTable t = (DataTable)dgv1.DataSource;

            string sRow = "001"; 
            string sTip = "";

            if(t.Rows.Count > 0)
            {
                DataView v = new DataView(t, "", "ing_row DESC", DataViewRowState.CurrentRows);
                if(v.Count > 0)
                {
                    sRow = (Convert.ToDecimal(v[0]["ing_row"]) + 1).ToString("000");
                    sTip = (string)v[0]["ing_tip"];
                }
            }

            DataRow x = t.NewRow();
            x["ing_art"] = _strArtCod;
            x["ing_tip"] = sTip;
            x["ing_crt"] = "";
            x["ing_row"] = sRow;
            x["ing_txt"] = "";
            x["ing_ann"] = false;
            t.Rows.Add(x);
        }

        private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = (string)x["ing_tip"];
                    if (s == "")
                        x["ing_tip"] = "BIL";
                    else if (s == "BIL")
                        x["ing_tip"] = "ETI";
                    else
                        x["ing_tip"] = "";
                }
            }
            if (e.ColumnIndex == 1)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;
                    string s = (string)x["ing_crt"];
                    if (s == "")
                        x["ing_crt"] = "G";
                    else
                        x["ing_crt"] = "";
                }
            }

        }

        private void stampaEtichetteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);

            DataTable t = (DataTable)dgv1.DataSource;
            DataView v = new DataView(t, "ing_tip='ETI' AND ing_ann=0", "ing_row", DataViewRowState.CurrentRows);

            if( v.Count == 0)
                MessageBox.Show("Non sono presenti valori da stampare!", "CONTROLLO STAMPA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else if(!_clsFun.Numerico(txtQta.Text))
                MessageBox.Show("Quantità non definita correttamente!", "CONTROLLO STAMPA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                DataTable tTmp = v.ToTable();

                decimal dQta = Convert.ToDecimal(txtQta.Text);

                //frmAnaArtIngLabel f = new frmAnaArtIngLabel();
                //f._tabTab = tTmp.Copy();

                //f._strDes = lblDes.Text;
                //f._strPes = lblPes.Text.PadLeft(10,Convert.ToChar(' '));
                //f._strTar = lblTar.Text.PadLeft(10, Convert.ToChar(' '));
                //f._strPrv = lblPrv.Text.PadLeft(10, Convert.ToChar(' '));
                //f._strImp = lblImp.Text.PadLeft(10, Convert.ToChar(' '));
                //f._strEan = lblEan.Text.PadLeft(10, Convert.ToChar(' '));
                //f._strDsc = dtpDsc.Value.ToShortDateString();

                //f._decEtiNum = dQta;
                //f.ShowDialog();
            }
        }


    }
}
