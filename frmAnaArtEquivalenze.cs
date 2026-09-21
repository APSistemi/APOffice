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
    public partial class frmAnaArtEquivalenze : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsVariazioni _clsVar = new clsVariazioni();
        clsQuery _clsQry = new clsQuery();

        private const string TABTABEQV = "TabEquPrezzi";
        private const string TABLISVEN = "GesLisVendita";

        private string _strConSql = "";

        public string _strTip = "";
        public string _strEqp = "";
        public string _strArt = "";
        public decimal _decPrv = 0;
        public DataTable _tabTab = new DataTable("tabTab");

        public frmAnaArtEquivalenze()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
        }

        private void frmAnaArtEquivalenze_Load(object sender, EventArgs e)
        {
            txtPrv.Text = _decPrv.ToString("#0.00");

            lblDes.Visible = false;
            txtPrv.Visible = false;
            btnAgg.Visible = false;
            lblAll.Visible = false;
            chkAll.Visible = false;            

            if (_strTip == "Prezzi")
            {
                lblDes.Visible = true;
                txtPrv.Visible = true;
                btnAgg.Visible = true;
            }
            else if (_strTip == "Selezione")
            {
                lblDes.Visible = true;
                lblDes.Text = "Prezzo di offerta";
                txtPrv.Visible = true;
                btnAgg.Visible = true;

                lblAll.Visible = true;
                chkAll.Visible = true; 
            }

            SetDgv1();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmAnaArtEquivalenze_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
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
            //dgv1.DisplayedRowCount() = true;
            dgv1.SelectionMode = DataGridViewSelectionMode.CellSelect;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "ArtCho";
            cCbc.Name = "Scelto";
            cCbc.Width = 25;

            if (_strTip != "Selezione")
                cCbc.Visible = false;

            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_cod";
            cTbc.Name = "Articolo";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 400;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_prv";
            cTbc.Name = "Prezzo vendita";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";
        }

        private void FillDati()
        {
            string s = "SELECT art_cod, art_des, art_cos, art_prv FROM AnaArticoli WHERE art_eqp='" + _strEqp + "'";

            DataTable tTmp = new clsGenTabTmp().TabTmpArt("Art");

            DataTable t = _clsFun.FillTabSql("", s, false, _strConSql);

            DataRow x;

            foreach (DataRow y in t.Rows)
            {
                tTmp.Clear();
                x = tTmp.NewRow();
                x["tmp_art"] = y["art_cod"];
                tTmp.Rows.Add(x);
                tTmp = _clsQry.ArtCosto("", "", tTmp, DateTime.Today);
                if (tTmp.Rows.Count > 0)
                    y["art_cos"] = (decimal)tTmp.Rows[0]["tmp_cos"];
                tTmp = _clsQry.ArtPrezzo(tTmp, "");
                if (tTmp.Rows.Count > 0)
                    y["art_prv"] = (decimal)tTmp.Rows[0]["tmp_prv"];
            }

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "ArtCho",
                Caption = "Scelto",
                ReadOnly = false,
                DefaultValue = (Boolean)true
            });

            dgv1.DataSource = t;
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = (string)x["art_cod"];
                    f.ShowDialog();
                }
            }
        }

        private void btnAgg_Click(object sender, EventArgs e)
        {
            if (_strTip == "Selezione")
            {
                if (dgv1.DataSource != null)
                {
                    _tabTab = ((DataTable)dgv1.DataSource).Clone();

                    foreach (DataRow y in ((DataTable)dgv1.DataSource).Rows)
                    {
                        if ((Boolean)y["ArtCho"])
                        {
                            y["art_prv"] = Convert.ToDecimal(txtPrv.Text.Replace(".", ","));
                            if ((string)y["art_cod"] == _strArt)
                                _decPrv = (decimal)y["art_prv"];
                        }
                        _tabTab.ImportRow(y);
                    }
                    Esci();
                }
            }
            else
                Salva();
        }

        private void Salva()
        {
            string s = "";
            DataTable t = (DataTable)dgv1.DataSource;

            decimal d = 0;
            if (_clsFun.Numerico(txtPrv.Text, "0123456789,."))
            {
                d = Convert.ToDecimal(txtPrv.Text.Replace(".",","));

                foreach (DataRow y in t.Rows)
                {
                    SalvaVen((string)y["art_cod"], d);

                    y["art_prv"] = d;

                    if ((string)y["art_cod"] == _strArt)
                        _decPrv = (decimal)y["art_prv"];

                }
                MessageBox.Show("Aggiornamento effettuato", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else
                MessageBox.Show("Prezzo vendita non valido!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void SalvaVen(string strArt, decimal decPrv)
        {
            Boolean b = false;
            string s = "SELECT * FROM GesLisVendita WHERE liv_art = '" + strArt + "' ORDER BY liv_dti DESC";
            DataTable t = _clsFun.FillTabSql(TABLISVEN, s, false, _strConSql);
            DataRow[] j;
            Boolean bVar = false;

            ArrayList aWhe = new ArrayList();
            aWhe.Add("liv_lis");
            aWhe.Add("liv_art");
            aWhe.Add("liv_dti");
            ArrayList aExl = new ArrayList();

            DataRow x = t.NewRow();
            x["liv_lis"] = _clsDef.LISPOS;
            x["liv_art"] = strArt;
            x["liv_prv"] = decPrv;
            x["liv_dti"] = DateTime.Today;
            x["liv_dtf"] = _clsDef.DAYOUT;
            x["liv_sta"] = "";
            x["liv_ann"] = 0;
            x["liv_day"] = DateTime.Today;

            s = "liv_lis='" + x["liv_lis"] + "' AND liv_art='" + x["liv_art"] + "' AND liv_dti='" + x["liv_dti"] + "'";
            j = t.Select(s);
            if (j.Length > 0){
                x["liv_idx"] = j[0]["liv_idx"];
                s = _clsFun.SqlUpdRowIdx(TABLISVEN, t, j[0], x, aExl);
            }
            else
            {
                s = "";
                j = t.Select("liv_lis='" + x["liv_lis"] + "' AND liv_art='" + x["liv_art"] + "'","liv_dti DESC");

                if(j.Length == 0 || (decimal)j[0]["liv_prv"] != decPrv)
                    s = _clsFun.SqlInsertRow(TABLISVEN, t, x);
            }
            if (s != "")
            {
                _clsFun.SqlWrite(s, _strConSql);
                _clsVar.Variazioni(strArt, s, "Equivalenze articolo", _clsDef.VARALL);

                bVar = true;
            }
        }

        private void chkAll_Click(object sender, EventArgs e)
        {
            bool bSta = false;
            if (chkAll.Checked)
                bSta = true;
            DataTable t = (DataTable)dgv1.DataSource;
            foreach (DataRow y in t.Rows)
                y["ArtCho"] = bSta;
        }

    }
}
