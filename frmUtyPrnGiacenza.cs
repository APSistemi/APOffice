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
    public partial class frmUtyPrnGiacenza : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private const string TABANAART = "AnaArticoli";
        private const string TABMOVFAT = "GesFatTestate";
        private const string TABSTAVEN = "GesNegVen";

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public frmUtyPrnGiacenza()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmUtyPrnGiacenza_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");

            //for (int i = DateTime.Today.Year - 4; i < DateTime.Today.Year + 1; i++)
            //    cmbYea.Items.Add(i.ToString());
            //cmbYea.SelectedIndex = 4;

            FillTab();

            SetDgv1();
        }

       private void frmUtyPrnGiacenza_KeyDown(object sender, KeyEventArgs e)
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
            this.Close();
        }

        private void btnEstrai_Click(object sender, EventArgs e)
        {
            if(cmbNeg.SelectedValue == null || cmbNeg.SelectedValue.ToString()  == "")
                MessageBox.Show("Negozio non selezionato", "CONTROLLO GIACENZA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
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
            cTbc.DataPropertyName = "art_rep";
            cTbc.Name = "Reparto";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_des";
            cTbc.Name = "Reparto descrizione";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_cod";
            cTbc.Name = "Articolo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "Articolo descrizione";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "gia_iqt";
            cTbc.Name = "Q.tà inventario";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "gia_aqt";
            cTbc.Name = "Q.tà acquistato";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "gia_vqt";
            cTbc.Name = "Q.tà venduta";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "gia_mqt";
            cTbc.Name = "Q.tà rettifiche";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_gia";
            cTbc.Name = "Q.tà rimanente";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.000";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_val";
            cTbc.Name = "Valore riga";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.000";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_sta";
            cTbc.Name = "Stato";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillTab()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabNegozi";
            s = "SELECT * FROM TabNegozi"; // WHERE tab_tip='L'";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_pos"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbNeg.DataSource = t;
            cmbNeg.DisplayMember = "tab_des";
            cmbNeg.ValueMember = "tab_cod";
            cmbNeg.SelectedValue = "001";
        }

        private void FillDati()
        {
            clsAggGiacenza cls = new clsAggGiacenza();
            cls._proBar = progressBar1;
            cls._lblMsg = lblMsg;
            cls.AggGiacenza();

            string s = "";
            DataRow[] j;
            DataRow x;
            decimal dVal = 0;
            string sNeg = cmbNeg.SelectedValue.ToString();

            s = "SELECT ";
            s += "AnaArtGiacenza.*, ";
            s += "AnaArticoli.art_cod, ";
            s += "AnaArticoli.art_des, ";
            s += "AnaArticoli.art_umi, ";
            s += "AnaArticoli.art_rep, ";
            s += "AnaArticoli.art_gia, ";
            s += "AnaArticoli.art_sta, ";
            s += "TabReparti.tab_des ";
            s += "FROM AnaArtGiacenza ";
            s += "LEFT OUTER JOIN AnaArticoli ON AnaArtGiacenza.gia_art = AnaArticoli.art_cod ";
            s += "LEFT JOIN TabReparti ON tab_cod = art_rep ";
            s += "WHERE ";
            s += "gia_neg='" + sNeg + "' ";
            if(chkArtAtt.Checked)
                s += " AND art_sta='" + _clsDef.STAATT + "' ";
            DataTable tGia = _clsFun.FillTabSql("AnaArtGiacenza", s, false, _strConSql);

            tGia.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_cos",
                Caption = "Costo",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            tGia.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_val",
                Caption = "Valore",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            foreach (DataRow y in tGia.Rows)
            {
                decimal dCos = 0;
                if (!DBNull.Value.Equals(y["gia_ico"]) && (decimal)y["gia_ico"] > 0)
                    dCos = (decimal)y["gia_ico"];
                if (!DBNull.Value.Equals(y["gia_aco"]) && (decimal)y["gia_aco"] > 0)
                    dCos = (decimal)y["gia_aco"];

                y["tmp_cos"] = dCos;

                if(dCos > 0)
                    y["tmp_val"] = (decimal)y["art_gia"] * dCos;

                dVal += (decimal)y["tmp_val"];
            }

            dgv1.DataSource = tGia;

            lblArts.Text = tGia.Rows.Count.ToString();
            lblCos.Text = dVal.ToString("#,###,##0.000");
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3 && e.RowIndex >= 0)
            {
                string sArt = dgv1.Rows[e.RowIndex].Cells["Articolo"].Value.ToString();
                frmAnaArticolo f = new frmAnaArticolo();
                f._strArtCod = sArt;
                f.ShowDialog();
            }
        }

        private DataTable GiaFor(string strFor, DataTable tabArt)
        {
            DataRow[] j;
            DataTable t = tabArt.Clone();

            string s = "SELECT DISTINCT lia_art FROM GesLisAcquisto WHERE lia_for='" + strFor + "' AND lia_ann=0";
            DataTable tLia = _clsFun.FillTabSql("GesLisAcquisto", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tLia.Columns["lia_art"];
            tLia.PrimaryKey = keys;

            foreach(DataRow y in tabArt.Rows)
            {
                j = tLia.Select("lia_art='" + y["tmp_art"] + "'");
                if (j.Length > 0)
                    t.ImportRow(y);
            }

            return t;
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel1();
        }

        private void csvToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Csv1();
        }

        private void Excel1()
        {
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "art_rep, art_des", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            string s = "";
            decimal d = 0;

            //t = (DataTable)dgv1.DataSource;

            foreach (DataRowView r in v)
            {
                t.ImportRow(r.Row);
            }

            string sTit = "Giacenza - etrazione dati al " + DateTime.Now.ToString("dd/MM/yyyy");
            string sFil = "";

            s = "";
            //s += "art_rep, Reparto, 110, StringLiteral;";
            s += "tab_des, Reparto, 110, StringLiteral;";
            s += "art_cod, Articolo, 40, StringLiteral;";
            s += "art_des, Descrizione, 200, StringLiteral;";
            s += "art_umi, UM, 50, StringLiteral;";
            s += "gia_iqt, Q.ta inventario, 45, Decimal2;";
            //s += "gia_ikg, Q.tà kg inventario, 45, Decimal2;";
            s += "gia_aqt, Q.ta acquistata, 45, Decimal2;";
            //s += "gia_akg, Q.tà kg acquistata, 45, Decimal2;";
            s += "gia_vqt, Q.ta venduta, 45, Decimal2;";
            //s += "gia_vkg, Q.tà kg venduta, 45, Decimal2;";
            s += "gia_mqt, Q.ta rettifiche, 45, Decimal2;";
            //s += "gia_mpz, Q.tà kg rettifiche, 45, Decimal2;";
            s += "art_gia, Q.tà giacenza, 45, Decimal2;";
            s += "tmp_cos, Costo, 45, Decimal3;";
            s += "tmp_val, Importo, 50, Decimal3;";
            s += "art_sta, Stato, 30, StringLiteral;";

            for (int i = 0; i <= 100; i++)
            {
                sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_Giacenza.xls";
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

            (new clsExcel()).exportToExcel(ds, sFil, sTit, "Valore " + lblCos.Text, true);

            ds.Tables.Remove(t);
            ds.Clear();
            ds.Dispose();
        }

        private void Csv1()
        {
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "art_rep, art_des", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            string s = "";
            decimal d = 0;

            //t = (DataTable)dgv1.DataSource;

            foreach (DataRowView r in v)
            {
                t.ImportRow(r.Row);
            }

            string sTit = "Giacenza - etrazione dati al " + DateTime.Now.ToString("dd/MM/yyyy");
            //string sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_Giacenza.csv";
            string sFil = "Giacenza.csv";

            s = "";
            //s += "art_rep, Reparto, 110, StringLiteral;";
            s += "tab_des, Reparto, 110, StringLiteral;";
            s += "art_cod, Articolo, 40, StringLiteral;";
            s += "art_des, Descrizione, 200, StringLiteral;";
            s += "art_umi, UM, 50, StringLiteral;";
            s += "gia_iqt, Q.ta inventario, 45, Decimal2;";
            //s += "gia_ikg, Q.tà kg inventario, 45, Decimal2;";
            s += "gia_aqt, Q.ta acquistata, 45, Decimal2;";
            //s += "gia_akg, Q.tà kg acquistata, 45, Decimal2;";
            s += "gia_vqt, Q.ta venduta, 45, Decimal2;";
            //s += "gia_vkg, Q.tà kg venduta, 45, Decimal2;";
            s += "gia_mqt, Q.ta rettifiche, 45, Decimal2;";
            //s += "gia_mpz, Q.tà kg rettifiche, 45, Decimal2;";
            s += "art_gia, Q.ta giacenza, 45, Decimal2;";
            s += "tmp_cos, Costo, 45, Decimal3;";
            s += "tmp_val, Importo, 50, Decimal3;";

            for (int i = 0; i <= 100; i++)
            {
                //sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_Giacenza.csv";
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

            //(new clsExcel()).exportToExcel(ds, sFil, sTit, "Valore " + lblCos.Text, true);
            (new clsExcel()).exportToCsv1(s, t, sFil, sTit, "Valore " + lblCos.Text);


            ds.Tables.Remove(t);
            ds.Clear();
            ds.Dispose();
        }

        private void excel2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "art_rep, art_des", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            string s = "";
            decimal d = 0;

            foreach (DataRowView r in v)
            {
                t.ImportRow(r.Row);
            }

            //string sTit = "Giacenza - etrazione dati dal " + dtpIni.Value.ToString("dd/MM/yyyy") + " al " + DateTime.Now.ToString("dd/MM/yyyy");
            string sTit = "Giacenza - etrazione dati al " + DateTime.Now.ToString("dd/MM/yyyy");
            string sFoo = "'Totale',,,,,,,,,," + lblCos.Text.Replace(".", "").Replace(",", ".") + ",";
            string sFil = "";

            string sFld = "";
            sFld += "tab_des, Reparto, 110, StringLiteral;";
            sFld += "art_cod, Articolo, 40, StringLiteral;";
            sFld += "art_des, Descrizione, 200, StringLiteral;";
            sFld += "art_umi, UM, 50, StringLiteral;";
            sFld += "gia_iqt, Q.ta inventario, 45, Decimal;";
            sFld += "gia_aqt, Q.ta acquistata, 45, Decimal;";
            sFld += "gia_vqt, Q.ta venduta, 45, Decimal;";
            sFld += "gia_mqt, Q.ta rettifiche, 45, Decimal;";
            sFld += "art_gia, Q.tà giacenza, 45, Decimal;";
            sFld += "tmp_cos, Costo, 45, Decimal;";
            sFld += "tmp_val, Importo, 50, Decimal;";
            sFld += "art_sta, Stato, 50, StringLiteral;";

            //sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_Giacenza.xls";
            sFil = "Giacenza.xls";

            //for (int i = 0; i <= 100; i++)
            //{
            //    //sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_Giacenza.xls";
            //    if (File.Exists(sFil))
            //    {
            //        try
            //        {
            //            File.Delete(sFil);
            //            break;
            //        }
            //        catch (Exception ex)
            //        {
            //            MessageBox.Show(ex.Message);
            //        }

            //    }
            //    else
            //        break;
            //}

            //DataSet ds = new DataSet();
            //DataTable tTit = (new clsExcel()).FillTitStr(s, "TabTit");
            //ds.Tables.Add(tTit);
            //ds.Tables.Add(t);
            //(new clsExcel()).exportToExcel(ds, sFil, sTit, "Valore " + lblCos.Text, true);
            //ds.Tables.Remove(t);
            //ds.Clear();
            //ds.Dispose();

            (new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);

        }
    }
}
