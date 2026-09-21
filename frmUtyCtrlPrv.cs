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
    public partial class frmUtyCtrlPrv : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strConMdb = "";

        public frmUtyCtrlPrv()
        {
            InitializeComponent();
        }

        private void frmUtyCtrlPrv_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FillDati();
        }

        private void FillDati()
        {
            DataTable tTmp = new DataTable();

            string s = "SELECT art_rep, art_cod, art_des, art_pve, art_off, art_odi, art_odf FROM AnaArticoli";
            DataTable t = _clsFun.FillTabMdb("AnaArticoli", s, false, _strConMdb);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "apo_sta",
                Caption = "Stato",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "apo_ard",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "apo_pve",
                Caption = "Pve",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "apo_dif",
                Caption = "Pve",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "apo_off",
                Caption = "Off",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                //if (progressBar1.Value > 1000)
                //    break;

                tTmp = _clsQry.ArtSeek((string)y["art_cod"], "");
                if(tTmp.Rows.Count > 0)
                {
                    y["apo_sta"] = tTmp.Rows[0]["tmp_sta"];
                    y["apo_ard"] = tTmp.Rows[0]["tmp_ard"];
                    y["apo_pve"] = tTmp.Rows[0]["tmp_prv"];
                    y["apo_dif"] = (decimal)y["apo_pve"] - (decimal)y["art_pve"];
                    if ((string)y["art_off"] != "")
                        y["apo_off"] = (string)y["art_off"] + " " + ((DateTime)y["art_odi"]).ToString("dd/MM/yy") + "-" + ((DateTime)y["art_odf"]).ToString("dd/MM/yy");
                }
            }

            dgv1.DataSource = t;
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

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s = "";

            DataTable t = (DataTable)dgv1.DataSource;

            s = "apo_dif <> 0";
            DataView v = new DataView(t, s, "art_rep, art_des", DataViewRowState.CurrentRows);

            DataRow[] j;
            t = v.Table.Clone();

            s = "";
            decimal d = 0;

            foreach (DataRowView r in v)
            {
                //d += (decimal)r["MovImp"];
                t.ImportRow(r.Row);
            }

            string sTit = "Differenze";
            string sFil = "DiffPrv";

            string sFld = "";
            sFld += "art_rep, Reparto, 50, StringLiteral;";
            sFld += "art_cod, Codice, 50, StringLiteral;";
            sFld += "art_des, Descriz. cassa, 50, StringLiteral;";
            sFld += "art_pve, Pv cassa, 50, StringLiteral;";
            sFld += "apo_sta, Stato, 10, StringLiteral;";
            sFld += "apo_ard, Descriz. PC, 50, StringLiteral;";
            sFld += "apo_pve, Pv PC, 50, StringLiteral;";
            sFld += "apo_off, Offerta, 50, StringLiteral;";

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

    }
}
