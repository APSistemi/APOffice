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
    public partial class frmGesStatPremi : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public DateTime _dayStaIni = new DateTime();
        public DateTime _dayStaFin = new DateTime();

        public frmGesStatPremi()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
        }

        private void frmGesStatPremi_Load(object sender, EventArgs e)
        {
            SetDgv1();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmGesStatPremi_KeyDown(object sender, KeyEventArgs e)
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

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_day";
            cTbc.Name = "Data";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            //cTbc.Visible = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_pos";
            cTbc.Name = "Pos";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_ora";
            cTbc.Name = "Ora";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_sco";
            cTbc.Name = "Scontrino";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_fid";
            cTbc.Name = "Tessera";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_val";
            cTbc.Name = "Importo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_imp";
            cTbc.Name = "Punti";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_ean";
            cTbc.Name = "Barcode";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ArtDes";
            cTbc.Name = "Descrizione";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillDati()
        {
            string s = "";
            DataRow[] j;

            s = "SELECT art_des, art_cod, ean_ean FROM AnaBarcode LEFT JOIN AnaArticoli ON AnaArticoli.art_cod = AnaBarcode.ean_art WHERE art_sta='P'";
            DataTable tArt = _clsFun.FillTabSql("tArt", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[2];
            keys[0] = tArt.Columns["art_cod"];
            keys[1] = tArt.Columns["ean_ean"];
            tArt.PrimaryKey = keys;

            s = "SELECT ";
            s += "GesNegVet.vet_day, ";
            s += "GesNegVet.vet_fid, ";
            s += "GesNegVep.vep_pos, ";
            s += "GesNegVep.vep_ora, ";
            s += "GesNegVep.vep_sco, ";
            s += "GesNegVep.vep_cod, ";
            s += "GesNegVep.vep_imp, ";
            s += "GesNegVep.vep_val, ";
            s += "GesNegVep.vep_ean ";
            s += "FROM GesNegVep LEFT OUTER JOIN ";
            s += "GesNegVet ON GesNegVep.vep_cau = GesNegVet.vet_cau AND ";
            s += "GesNegVep.vep_day = GesNegVet.vet_day AND ";
            s += "GesNegVep.vep_ora = GesNegVet.vet_ora AND ";
            s += "GesNegVep.vep_pos = GesNegVet.vet_pos AND ";
            s += "GesNegVep.vep_sco = GesNegVet.vet_sco ";
            s += "WHERE ";
            s += "GesNegVet.vet_day >= " + _clsFun.DaySql(_dayStaIni) + " AND ";
            s += "GesNegVet.vet_day <= " + _clsFun.DaySql(_dayStaFin) + " AND ";
            s += "(GesNegVep.vep_ean <> '') AND (GesNegVep.vep_cau <> 'PUN')";
            DataTable t = _clsFun.FillTabSql("tPre", s, false, _strConSqlSta);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "ArtDes",
                Caption = "Des",
                ReadOnly = false,
                DefaultValue = (String)""
            });

            foreach(DataRow y in t.Rows)
            {
                j = tArt.Select("ean_ean='" + y["vep_ean"] + "'");
                if (j.Length > 0)
                    y["ArtDes"] = j[0]["art_des"];
            }

            dgv1.DataSource = t;

        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Excel1();
            Csv1();
        }

        private void Excel1()
        {
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "vet_day, vet_fid", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            string s = "";
            decimal d = 0;

            foreach (DataRowView r in v)
            {
                t.ImportRow(r.Row);
            }

            string sTit = DateTime.Today.ToString("dd/MM/yyyy") + " - Estrazione premi dal " + _dayStaIni.ToString("dd/MM/yyyy") + " al " + _dayStaFin.ToString("dd/MM/yyyy");
            string sFil = "";

            s = "";
            s += "vet_day, Data, 60, StringLiteral;";
            s += "vep_pos, Pos, 50, StringLiteral;";
            s += "vep_ora, Ora, 50, StringLiteral;";
            s += "vep_sco, Scontrino, 50, StringLiteral;";
            s += "vet_fid, Tessera, 70, StringLiteral;";
            s += "vep_val, Importo, 50, Decimal2;";
            s += "vep_imp, Punti, 50, Integer;";
            s += "vep_ean, Barcode, 70, StringLiteral;";
            s += "ArtDes, Descrizione, 230, StringLiteral;";

            for (int i = 0; i <= 100; i++)
            {
                sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_FidPremi.xls";
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
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "vet_day, vet_fid", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            string s = "";
            decimal d = 0;

            foreach (DataRowView r in v)
            {
                t.ImportRow(r.Row);
            }

            string sTit = DateTime.Today.ToString("dd/MM/yyyy") + " - Estrazione premi dal " + _dayStaIni.ToString("dd/MM/yyyy") + " al " + _dayStaFin.ToString("dd/MM/yyyy");
            string sFil = "";

            s = "";
            s += "vet_day, Data, 60, StringLiteral;";
            s += "vep_pos, Pos, 50, StringLiteral;";
            s += "vep_ora, Ora, 50, StringLiteral;";
            s += "vep_sco, Scontrino, 50, StringLiteral;";
            s += "vet_fid, Tessera, 70, StringLiteral;";
            s += "vep_val, Importo, 50, Decimal2;";
            s += "vep_imp, Punti, 50, Integer;";
            s += "vep_ean, Barcode, 70, StringLiteral;";
            s += "ArtDes, Descrizione, 230, StringLiteral;";

            string sFld = s;

            (new clsExcel()).exportToCsv1(sFld, t,  sFil, sTit, "");

        }

    }
}
