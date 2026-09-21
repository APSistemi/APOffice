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
    public partial class frmPrnArtMerceologie : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        public frmPrnArtMerceologie()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            SetDgv1();
        }

        private void frmPrnArtMerceologie_Load(object sender, EventArgs e)
        {
            FillTabs();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmPrnArtMerceologie_KeyDown(object sender, KeyEventArgs e)
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
            dgv1.ReadOnly = true;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_cod";
            cTbc.Name = "Articolo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "StaDes";
            cTbc.Name = "Stato";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_iva";
            cTbc.Name = "IVA";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_umi";
            cTbc.Name = "UM";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_tgr";
            cTbc.Name = "Grammatura";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_pxc";
            cTbc.Name = "Pxc";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 9;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.000";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_pne";
            cTbc.Name = "P.netto";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 9;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.000";

            //cCbc = new DataGridViewCheckBoxColumn();
            //cCbc.ValueType = typeof(Boolean);
            //cCbc.DataPropertyName = "art_ann";
            //cCbc.Name = "Annullato";
            //cCbc.Width = 25;
            //dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_dtm";
            cTbc.Name = "Data Modifica";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "RepDes";
            cTbc.Name = "Reparto";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "EcrDe2";
            cTbc.Name = "Merceologia 2";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "EcrDe3";
            cTbc.Name = "Merceologia 3";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_plu";
            cTbc.Name = "PLU";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "RebDes";
            cTbc.Name = "Rep. bilance";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "art_bil";
            cCbc.Name = "Bilancia";
            cCbc.Width = 25;
            dgv1.Columns.Add(cCbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "art_bpz";
            cCbc.Name = "A corpo";
            cCbc.Width = 25;
            dgv1.Columns.Add(cCbc);
        }

        private void FillTabs()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabReparti";
            s = "SELECT * FROM TabReparti WHERE tab_ann=0 ORDER BY tab_des";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbRep.DataSource = t;
            cmbRep.DisplayMember = "tab_des";
            cmbRep.ValueMember = "tab_cod";
            cmbRep.SelectedValue = "";
        }

        private void FillDati()
        {
            //String sFor = cmbArtFor.SelectedValue.ToString();

            string s = "";

            s = "SELECT ";
            s += "AnaArticoli.art_cod,  ";
            s += "AnaArticoli.art_des,  ";
            s += "AnaArticoli.art_sta,  ";
            s += "AnaArticoli.art_iva,  ";
            s += "AnaArticoli.art_umi,  ";
            s += "AnaArticoli.art_tgr,  ";
            s += "AnaArticoli.art_pxc,  ";
            s += "AnaArticoli.art_pne,  ";
            s += "AnaArticoli.art_rep,  ";
            s += "AnaArticoli.art_ec1,  ";
            s += "AnaArticoli.art_ec2,  ";
            s += "AnaArticoli.art_ec3,  ";
            s += "AnaArticoli.art_reb,  ";
            s += "AnaArticoli.art_plu,  ";
            s += "AnaArticoli.art_bpz,  ";
            s += "AnaArticoli.art_dti,  ";
            s += "AnaArticoli.art_dtm,  ";
            s += "AnaArticoli.art_bil,  ";
            s += "AnaArticoli.art_cos,  ";
            s += "AnaArticoli.art_prv,  ";
            s += "TabEcrLv1.tab_des AS EcrDe1,    ";
            s += "TabEcrLv2.tab_des AS EcrDe2,    ";
            s += "TabEcrLv3.tab_des AS EcrDe3,    ";
            s += "TabEcrLv3.tab_tip AS Ec3Tip,    ";
            s += "TabReparti.tab_des AS RepDes,   ";
            s += "TabStato.tab_des AS StaDes,     ";
            s += "TabRepBilance.tab_des AS RebDes ";
            s += "FROM AnaArticoli               ";

            s += "LEFT OUTER JOIN TabEcrLv1 ON AnaArticoli.art_ec1 = TabEcrLv1.tab_cod ";
            s += "LEFT OUTER JOIN TabEcrLv2 ON AnaArticoli.art_ec1 = TabEcrLv2.tab_lv1 AND AnaArticoli.art_ec2 = TabEcrLv2.tab_cod ";
            s += "LEFT OUTER JOIN TabRepBilance ON AnaArticoli.art_reb = TabRepBilance.tab_cod ";
            s += "LEFT OUTER JOIN TabStato ON AnaArticoli.art_sta = TabStato.tab_cod  ";
            //s += "LEFT OUTER JOIN TabEcrLv2 ON AnaArticoli.art_ec1 = TabEcrLv3.tab_lv1 AND AnaArticoli.art_ec2 = TabEcrLv3.tab_lv2 ";
            s += "LEFT OUTER JOIN TabEcrLv3 ON AnaArticoli.art_ec1 = TabEcrLv3.tab_lv1 AND AnaArticoli.art_ec2 = TabEcrLv3.tab_lv2 AND AnaArticoli.art_ec3 = TabEcrLv3.tab_cod  ";
            s += "LEFT OUTER JOIN TabReparti ON AnaArticoli.art_rep = TabReparti.tab_cod ";
            if (cmbRep.SelectedValue != null && cmbRep.SelectedValue.ToString() != "")
                s += "WHERE art_rep='" + cmbRep.SelectedValue.ToString() + "'";

            //s += "WHERE art_umi='KG' AND art_plu='' AND art_rep <>'003' AND art_rep <>'007' AND art_rep <>'005' AND art_rep <>'006'";

            DataTable t = _clsFun.FillTabSql("TabAcq", s, false, _strConSql);

            dgv1.DataSource = t;

            lblCnt.Text = t.Rows.Count.ToString();
        }

        private void btnFill_Click(object sender, EventArgs e)
        {
            FillDati();
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv1.Columns[e.ColumnIndex].Name == "Descrizione")
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = ((string)x["art_cod"]).Trim();

                    if (s != "")
                    {
                        frmAnaArticolo f = new frmAnaArticolo();
                        f._strArtCod = s;
                        f.ShowDialog();
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string s = "";

            clsVariazioni cls = new clsVariazioni();

            DataTable t = (DataTable)dgv1.DataSource;

            progressBar1.Value = 0;
            progressBar1.Maximum = (int)t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                s = "UPDATE AnaArticoli SET art_umi='NR' WHERE art_cod='" + y["art_cod"] + "'";
                _clsFun.SqlWrite(s, _strConSql);

                if((string)y["art_sta"] == _clsDef.STAATT)
                    cls.Variazioni((string)y["art_cod"], s, "Utility", _clsDef.VARPOS);
            }

        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s = "";
            DataRow[] j;
            string sTit = DateTime.Today.ToString("dd/MM/yyyy") + " Articoli";
            string sFil = "ArtMerceologie.xls";
            string sFoo = "";

            s = "SELECT ean_art, ean_ean FROM AnaBarcode ORDER BY ean_ean, ean_art, ean_dti DESC";
            DataTable tEan = _clsFun.FillTabSql("AnaBarcode", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[2];
            keys[0] = tEan.Columns["ean_art"];
            keys[1] = tEan.Columns["ean_ean"];
            tEan.PrimaryKey = keys;

            //DataTable t = ((DataTable)dgv1.DataSource).Copy();
            DataTable t = new DataView((DataTable)dgv1.DataSource, "art_sta='A'", "art_des", DataViewRowState.CurrentRows).ToTable();
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ean",
                Caption = "Barcode",
                MaxLength = 13,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            foreach(DataRow y in t.Rows)
            {
                j = tEan.Select("ean_art='" + y["art_cod"] + "'");
                if (j.Length > 0)
                    y["tmp_ean"] = j[0]["ean_ean"];
            }

            string sFld = "";
            sFld += "art_cod, Articolo, 70, StringLiteral;";
            sFld += "art_des, Descrizione, 70, StringLiteral;";
            sFld += "StaDes, Stato, 30, StringLiteral;";
            sFld += "art_iva, IVA, 90, StringLiteral;";
            sFld += "art_umi, UM, 70, StringLiteral;";
            sFld += "art_tgr, Gram, 70, StringLiteral;";
            sFld += "art_pxc, Pezzi, 70, Decimal;";
            sFld += "art_pne, Peso netto, 70, Decimal;";
            sFld += "RepDes, Reparto, 70, StringLiteral;";
            sFld += "EcrDe2, ECR, 70, StringLiteral;";
            sFld += "EcrDe3, ECR, 70, StringLiteral;";
            sFld += "Ec3Tip, Tipo, 70, StringLiteral;";
            sFld += "art_plu, PLU, 70, StringLiteral;";
            sFld += "RebDes, R.Bil, 70, StringLiteral;";
            sFld += "art_cos, Costo, 70, Decimal;";
            sFld += "art_prv, Prezzo, 70, Decimal;";
            sFld += "tmp_ean, Barcode, 70, StringLiteral;";

            //(new clsExcel()).exportToCsv1(sFld, t, sFil, sTit, "");
            (new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);

        }

        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            clsGenPdfArtEcr cls = new clsGenPdfArtEcr();
            cls._tabArt = (DataTable)dgv1.DataSource;
            cls.PrnPdfArtEcr();
        }

        private void lISTINOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            clsGenPdfArtEcr cls = new clsGenPdfArtEcr();
            cls._strConSql = _strConSql;
            cls._tabArt = (DataTable)dgv1.DataSource;
            cls.PrnPdfArtListino();
        }

    }
}
