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
    public partial class frmUtyCtrlArtPos : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";
        private string _strConMdb = "";

        public frmUtyCtrlArtPos()
        {
            InitializeComponent();
        }

        private void frmUtyCtrlArtPos_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");

            SetDgv1();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmUtyCtrlArtPos_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void btnMdb_Click(object sender, EventArgs e)
        {
            OpenFileDialog o = new OpenFileDialog();
            DialogResult dr = o.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    lblPath.Text = o.FileName;
                    _strConMdb = _clsFun.ConMdb(lblPath.Text);

                }
                catch (Exception ex)
                {
                    MessageBox.Show("File non disponibile: " + o.FileName + " - " + ex.Message);
                }
            }
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

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_ean";
            cTbc.Name = "Barcode";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_ar1";
            cTbc.Name = "Codice";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_de1";
            cTbc.Name = "Descrizione";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_ar2";
            cTbc.Name = "Codice POS";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_de2";
            cTbc.Name = "Descrizione POS";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);


            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "tmp_ann";
            cCbc.Name = "Ann";
            cCbc.Width = 25;
            dgv1.Columns.Add(cCbc);
        }

        private void FillDati()
        {
            string s = "";
            string p = "";
            DataRow x;
            DataRow[] j;


            p = "AnaArticoli";
            s = "SELECT ";
            s += "AnaArticoli.art_cod, ";
            s += "AnaArticoli.art_des, ";
            s += "AnaArticoli.art_sta, ";
            s += "AnaArticoli.art_iva, ";
            s += "AnaArticoli.art_umi, ";
            s += "AnaArticoli.art_rep, ";
            s += "AnaArticoli.art_pve, ";
            s += "AnaBarcode.ean_ean, ";
            s += "AnaBarcode.ean_ann ";
            s += "FROM AnaArticoli INNER JOIN AnaBarcode ON AnaArticoli.art_cod = AnaBarcode.ean_art;";
            DataTable tMdb = _clsFun.FillTabMdb(p, s, false, _strConMdb);

            s = "SELECT ";
            s += "AnaArticoli.art_cod, ";
            s += "AnaArticoli.art_des, ";
            s += "AnaArticoli.art_sta, ";
            s += "AnaArticoli.art_umi, ";
            s += "AnaBarcode.ean_ean, ";
            s += "AnaArticoli.art_dtm ";
            s += "FROM AnaArticoli INNER JOIN AnaBarcode ON AnaArticoli.art_cod = AnaBarcode.ean_art";
            DataTable tSql = _clsFun.FillTabSql(p, s, false, _strConSql);
            DataColumn[] keys = new DataColumn[2];
            keys[0] = tSql.Columns["ean_ean"];
            keys[1] = tSql.Columns["art_cod"];

            DataTable tTmp = new clsGenTabTmp().TabTmpArtCtrl("TmpArt");

            progressBar1.Value = 0;
            progressBar1.Maximum = tMdb.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in tMdb.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                if ((string)y["ean_ean"] == "5000112557688")
                    Console.WriteLine("aaaaaaaaa");

                j = tSql.Select("ean_ean='" + (string)y["ean_ean"] + "'");
                if(j.Length > 0)
                {
                    if((string)y["art_cod"] != (string)j[0]["art_cod"])
                    {
                        x = tTmp.NewRow();
                        x["tmp_ean"] = y["ean_ean"];
                        x["tmp_ar1"] = j[0]["art_cod"];
                        x["tmp_de1"] = j[0]["art_des"];
                        x["tmp_ar2"] = y["art_cod"];
                        x["tmp_de2"] = y["art_des"];
                        x["tmp_ann"] = y["ean_ann"];
                        tTmp.Rows.Add(x);
                    }
                }
                else
                {
                    x = tTmp.NewRow();
                    x["tmp_ean"] = y["ean_ean"];
                    //x["tmp_ar1"] = j[0]["art_cod"];
                    //x["tmp_de1"] = j[0]["art_des"];
                    x["tmp_ar2"] = y["art_cod"];
                    x["tmp_de2"] = y["art_des"];
                    //x["tmp_ann"] = y["ean_ann"];

                    x["tmp_ann"] = true;

                    tTmp.Rows.Add(x);
                }

            }

            dgv1.DataSource = tTmp;
        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = (string)x["tmp_ar1"];
                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = s;
                    f.ShowDialog();
                }
            }

        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable t = (DataTable)dgv1.DataSource;

            DataView v = new DataView(t, "", "", DataViewRowState.CurrentRows);

            DataRow[] j;
            t = v.Table.Clone();

            string s = "";
            decimal d = 0;

            foreach (DataRowView r in v)
            {
                //d += (decimal)r["MovImp"];
                t.ImportRow(r.Row);
            }

            string sTit = "DIFFERENZE CASSA/APOFFICE AL " + DateTime.Now.ToString("dd/MM/yyyy");
            string sFil = "";

            string sFld = "";
            //s += "DocNum, Num, 30, StringLiteral;";
            sFld += "tmp_ean, barcode, 50, StringLiteral;";
            sFld += "tmp_ar1, APO codice, 50, StringLiteral;";
            sFld += "tmp_de1, APO descrizione, 180, StringLiteral;";
            sFld += "tmp_ar2, POS codice, 50, Decimal2;";
            sFld += "tmp_de2, POS descrizione, 120, StringLiteral;";

            //sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd_HHmm") + "_Documenti.xls";

            DataSet ds = new DataSet();

            DataTable tTit = (new clsExcel()).FillTitStr(s, "TabTit");

            ds.Tables.Add(tTit);
            ds.Tables.Add(t);

            //(new clsExcel()).exportToExcel(ds, sFil, sTit, "3|Totale  " + d.ToString("#,###,##0.00"), true);
            (new clsExcel()).exportToCsv1(sFld, t, sFil, sTit, "");


            ds.Tables.Remove(t);
            ds.Clear();
            ds.Dispose();

        }

        private void btnAnn_Click(object sender, EventArgs e)
        {
            string s = "";

            DataTable t = (DataTable)dgv1.DataSource;

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                //if (!(Boolean)y["tmp_ann"])
                if ((Boolean)y["tmp_ann"])
                {
                    s = "UPDATE AnaBarcode SET ean_ann=1 WHERE ean_ean='" + y["tmp_ean"] + "'";
                    _clsFun.MdbWrite(s, _strConMdb);
                }
            }

        }


    }
}
