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
    public partial class frmGesVarPromo : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();
        clsVariazioni _clsVar = new clsVariazioni();

        private string _strConSql = "";

        public frmGesVarPromo()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0); ;
        }

        private void frmGesPromo_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            SetDgv1();

            FillDati("ALL");
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmGesPromo_KeyDown(object sender, KeyEventArgs e)
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
            //DataTable tSta = _clsFun.FillTabSql(TABTABSTD, "SELECT * FROM " + TABTABSTD + " WHERE tab_var=1", false, _strConSql);

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
            cTbc.DataPropertyName = "liv_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 300;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.ToolTipText = "Doppio click per anagrafica articolo";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "liv_dti";
            cTbc.Name = "Data inizio";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "liv_dtf";
            cTbc.Name = "Data fine";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "liv_prv";
            cTbc.Name = "Prezzo vendita";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "liv_sta";
            cTbc.Name = "Stato";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "liv_ann";
            cCbc.Name = "Annullato";
            cCbc.Width = 30;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cCbc);
        }

        private void FillDati(string strTip)
        {
            string s = "SELECT * FROM GesLisVendita WHERE liv_lis='" + _clsDef.LISPRO + "'";

            s = "SELECT ";
            s += "GesLisVendita.liv_art, ";
            s += "GesLisVendita.liv_prv, ";
            s += "GesLisVendita.liv_dti, ";
            s += "GesLisVendita.liv_dtf, ";
            s += "GesLisVendita.liv_ann, ";
            s += "GesLisVendita.liv_day, ";
            s += "GesLisVendita.liv_sta, ";
            s += "AnaArticoli.art_des ";
            s += "FROM GesLisVendita INNER JOIN AnaArticoli ON GesLisVendita.liv_art = AnaArticoli.art_cod ";
            s += "WHERE (GesLisVendita.liv_lis = '" + _clsDef.LISPRO + "') ";

            if(strTip != "ALL")
            {
                s += " AND ";
                s += "GesLisVendita.liv_dti <= " + _clsFun.DaySql(dtpDti.Value) + " AND ";
                s += "GesLisVendita.liv_dtf >= " + _clsFun.DaySql(dtpDti.Value) + " ";
            }

            s += "ORDER BY liv_dtf DESC ";

            DataTable t = _clsFun.FillTabSql("GesLisVendita", s, false, _strConSql);
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
                    f._strArtCod = (string)x["liv_art"];
                    f.ShowDialog();
                }
            }
        }

        private void btnEstrai_Click(object sender, EventArgs e)
        {
            FillDati("");
        }

        private void btnEti_Click(object sender, EventArgs e)
        {
            DataTable t = (DataTable)dgv1.DataSource;
            if (t == null || t.Rows.Count == 0) return;

            List<Tuple<string, string, string, string>> items = new List<Tuple<string, string, string, string>>();
            foreach (DataRow y in t.Rows)
            {
                if (y["liv_art"] != DBNull.Value && y["liv_dti"] != DBNull.Value)
                {
                    items.Add(new Tuple<string, string, string, string>((string)y["liv_art"], "PRO" + ((DateTime)y["liv_dti"]).ToString("yyyyMMdd"), "Promo", _clsDef.VARETI));
                }
            }

            _clsVar.VariazioniBatch(items);
        }

    }
}
