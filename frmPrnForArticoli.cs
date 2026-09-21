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
    public partial class frmPrnForArticoli : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        public frmPrnForArticoli()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            SetDgv1();
        }

        private void frmPrnForArticoli_Load(object sender, EventArgs e)
        {
            FillTabs();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmPrnForArticoli_KeyDown(object sender, KeyEventArgs e)
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
            cTbc.DataPropertyName = "art_sta";
            cTbc.Name = "Stato articolo";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_arf";
            cTbc.Name = "Articolo fornitore";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 170;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 20;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_dti";
            cTbc.Name = "Data Modifica";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_pxc";
            cTbc.Name = "P x c";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 9;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_cxp";
            cTbc.Name = "C x p";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 9;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_stf";
            cTbc.Name = "Stato fornitore";
            cTbc.Width = 15;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_tip";
            cTbc.Name = "Tipo";
            cTbc.Width = 15;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.ToolTipText = "Manuale, Divulgazione, Fattura";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 9;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.000";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lia_prv";
            cTbc.Name = "Prezzo vendita";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 9;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "lia_ann";
            cCbc.Name = "Ann.";
            cCbc.Width = 25;
            dgv1.Columns.Add(cCbc);
        }

        private void FillTabs()
        {
            string p = "AnaFornitori";
            string s = "SELECT * FROM AnaFornitori ORDER BY for_des";
            DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
            DataRow x = t.NewRow();
            x["for_cod"] = "";
            x["for_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbArtFor.DataSource = t;
            cmbArtFor.DisplayMember = "for_des";
            cmbArtFor.ValueMember = "for_cod";
            cmbArtFor.SelectedValue = "";
        }

        private void cmbArtFor_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillDati();
        }

        private void FillDati()
        {
            string s = "";
            DataRow[] j;
            String sFor = cmbArtFor.SelectedValue.ToString();

            s = "SELECT * FROM (";
            s += "SELECT ";
            s += "ROW_NUMBER() OVER (PARTITION BY liv_art ORDER BY liv_art, liv_dti DESC) AS ROW, ";
            s += "liv_art, ";
            s += "liv_prv ";
            s += "FROM GesLisVendita ) AS A WHERE ROW = 1";
            DataTable tLiv = _clsFun.FillTabSql("", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tLiv.Columns["liv_art"];
            tLiv.PrimaryKey = keys;

            s = "SELECT ";
            s += "GesLisAcquisto_1.lia_dti, ";
            s += "GesLisAcquisto_1.lia_dtf, ";
            s += "GesLisAcquisto_1.lia_for, ";
            s += "GesLisAcquisto_1.lia_stf, ";
            s += "GesLisAcquisto_1.lia_tip, ";
            s += "GesLisAcquisto_1.lia_art, ";
            s += "GesLisAcquisto_1.lia_arf, ";
            s += "GesLisAcquisto_1.lia_pxc, ";
            s += "GesLisAcquisto_1.lia_cos, ";
            s += "GesLisAcquisto_1.lia_prv, ";
            s += "GesLisAcquisto_1.lia_ann, ";
            s += "AnaArticoli.art_sta, ";
            s += "AnaArticoli.art_des, ";
            s += "AnaArticoli.art_rep ";
            s += "FROM (";
            s += "SELECT lia_tip, lia_for, MAX(lia_dti) AS lia_dti, lia_art ";
            s += "FROM  GesLisAcquisto ";
            s += "WHERE lia_for = '" + sFor + "' AND lia_ann=0 ";
            s += "GROUP BY lia_tip, lia_for, lia_art) AS A ";
            s += "LEFT OUTER JOIN GesLisAcquisto AS GesLisAcquisto_1 ON ";
            s += "A.lia_tip = GesLisAcquisto_1.lia_tip AND ";
            s += "A.lia_for = GesLisAcquisto_1.lia_for AND ";
            s += "A.lia_art = GesLisAcquisto_1.lia_art ";
            s += "LEFT OUTER JOIN AnaArticoli ON GesLisAcquisto_1.lia_art = AnaArticoli.art_cod ";

            DataTable t = _clsFun.FillTabSql("TabAcq", s, false, _strConSql);

            foreach(DataRow y in t.Rows)
            {
                j = tLiv.Select("liv_art='" + y["lia_art"] + "'");
                if (j.Length > 0)
                    y["lia_prv"] = (decimal)j[0]["liv_prv"];
            }

            dgv1.DataSource = t;

            lblCnt.Text = t.Rows.Count.ToString();
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgv1 != null)
            {
                //Excel1();
                //Csv1();
                Excel2();
            }
        }

        private void Excel1()
        {
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "art_des", DataViewRowState.CurrentRows);
            t = v.Table;

            string s = "";
            //decimal d = 0;

            //foreach (DataRowView r in v)
            //{
            //    if ((string)r["inv_art"] == "0003984")
            //        Console.WriteLine("aaaa");

            //    j = t.Select("inv_art='" + r["inv_art"] + "'");
            //    if (j.Length == 0)
            //    {
            //        DataRow x = t.NewRow();
            //        x["inv_art"] = r["inv_art"];
            //        x["InvArd"] = r["InvArd"];
            //        x["InvL1c"] = r["InvL1c"];
            //        x["InvL2c"] = r["InvL2c"];
            //        x["InvL3c"] = r["InvL3c"];
            //        x["InvL1d"] = r["InvL1d"];
            //        x["InvL2d"] = r["InvL2d"];
            //        x["InvL3d"] = r["InvL3d"];
            //        x["InvUmi"] = r["InvUmi"];
            //        x["inv_qta"] = 0;
            //        x["inv_cos"] = 0;
            //        x["InvImp"] = 0;
            //        t.Rows.Add(x);
            //        j = t.Select("inv_art='" + r["inv_art"] + "'");
            //    }

            //    j[0]["inv_qta"] = (decimal)j[0]["inv_qta"] + (decimal)r["inv_qta"];
            //    j[0]["InvImp"] = (decimal)j[0]["InvImp"] + (decimal)r["InvImp"];
            //    j[0]["inv_cos"] = (decimal)j[0]["InvImp"] / (decimal)j[0]["inv_qta"];

            //    d += (decimal)r["InvImp"];
            //}

            string sTit = DateTime.Now.ToString("dd/MM/yyyy") + " - Estrazione articoli del fornitore " + cmbArtFor.Text;
            string sFil = "";

            s = "";
            s += "lia_tip, Tipo, 70, StringLiteral;";
            s += "lia_stf, Stato, 30, StringLiteral;";
            s += "lia_art, Codice, 90, StringLiteral;";
            s += "lia_arf, Art. fornitore, 70, StringLiteral;";
            s += "art_des, Descrizione, 190, StringLiteral;";
            s += "lia_pxc, Pxc, 40, Integer;";
            s += "lia_cos, P.costo, 40, Decimal3;";
            s += "lia_prv, P.vendita, 40, Decimal2;";
            //s += "lia_ann, Ann, 40, Decimal2;";

            for (int i = 0; i <= 100; i++)
            {
                sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_ForArticoli.xls";
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

        private void Csv1()
        {
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "art_des", DataViewRowState.CurrentRows);
            t = v.Table;

            string s = "";

            string sTit = DateTime.Now.ToString("dd/MM/yyyy") + " - Estrazione articoli del fornitore " + cmbArtFor.Text;
            string sFil = "AssFornitori";

            s = "";
            s += "lia_tip, Tipo, 70, StringLiteral;";
            s += "lia_stf, Stato, 30, StringLiteral;";
            s += "lia_art, Codice, 90, StringLiteral;";
            s += "lia_arf, Art. fornitore, 70, StringLiteral;";
            s += "art_des, Descrizione, 190, StringLiteral;";
            s += "lia_pxc, Pxc, 40, Integer;";
            s += "lia_cos, P.costo, 40, Decimal3;";
            s += "lia_prv, P.vendita, 40, Decimal2;";

            string sFld = s;

            //DataSet ds = new DataSet();

            //DataTable tTit = (new clsExcel()).FillTitStr(s, "TabTit");

            //ds.Tables.Add(tTit);
            //ds.Tables.Add(t);

            (new clsExcel()).exportToCsv1(sFld, t, sFil, sTit, "");

            //ds.Tables.Remove(t);
            //ds.Clear();
            //ds.Dispose();
        }

        private void Excel2()
        {
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "art_des", DataViewRowState.CurrentRows);
            t = v.Table;

            string s = "";

            string sTit = DateTime.Now.ToString("dd/MM/yyyy") + " - Estrazione articoli del fornitore " + cmbArtFor.Text;
            string sFil = "";
            string sFoo = "";

            string sFld = "";
            sFld += "art_sta, Stato, 70, StringLiteral;";
            //sFld += "lia_stf, Stato, 30, StringLiteral;";
            sFld += "lia_art, Codice, 90, StringLiteral;";
            sFld += "lia_arf, Art. fornitore, 70, StringLiteral;";
            sFld += "art_des, Descrizione, 190, StringLiteral;";
            sFld += "lia_pxc, Pxc, 40, Decimal;";
            sFld += "lia_cos, P.costo, 40, Decimal;";
            sFld += "lia_prv, P.vendita, 40, Decimal;";
            //s += "lia_ann, Ann, 40, Decimal2;";

            for (int i = 0; i <= 100; i++)
            {
                sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_ForArticoli.xls";
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

            sFil = "ForListino.xls";

            (new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);

            //DataSet ds = new DataSet();

            //DataTable tTit = (new clsExcel()).FillTitStr(s, "TabTit");

            //ds.Tables.Add(tTit);
            //ds.Tables.Add(t);

            //(new clsExcel()).exportToExcel(ds, sFil, sTit, "", true);

            //ds.Tables.Remove(t);
            //ds.Clear();
            //ds.Dispose();
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3)
            {
                string sArt = dgv1.Rows[e.RowIndex].Cells["Articolo"].Value.ToString();

                frmAnaArticolo f = new frmAnaArticolo();
                f._strArtCod = sArt;
                f.ShowDialog();
            }
        }

        private void btnAtt_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Confermi l'attivazione alle casse di tutti gli articoli del fornitore?", "ATTIVAZIONE ARTICOLI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                 Art2Pos();
        }

        private void Art2Pos()
        {
            string s = "";

            clsVariazioni clsVar = new clsVariazioni();

            DataTable t = (DataTable)dgv1.DataSource;

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                s = "UPDATE AnaArticoli SET art_sta='" + _clsDef.STAATT + "' WHERE art_cod='" + y["lia_art"] + "'";
                _clsFun.SqlWrite(s, _strConSql);

                clsVar.Variazioni((string)y["lia_art"], "Forza", "Attivazione da ass.timentofornitori", "ALL");

            }

        }

    }
}
