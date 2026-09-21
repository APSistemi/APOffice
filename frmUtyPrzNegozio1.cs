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
    public partial class frmUtyPrzNegozio1 : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private const string TABANAART = "AnaArticoli";
        private const string TABMOVFAT = "GesFatTestate";
        private const string TABSTAVEN = "GesNegVen";

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public frmUtyPrzNegozio1()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmUtyPrzNegozio1_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
            SetDgv1();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmUtyPrzNegozio1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void btnFill_Click(object sender, EventArgs e)
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
            cTbc.DataPropertyName = "art_red";
            cTbc.Name = "Reparto";
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
            cTbc.DataPropertyName = "art_umi";
            cTbc.Name = "UM";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "001_day";
            cTbc.Name = "001 Data";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "001_prv";
            cTbc.Name = "001 prezzo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "002_day";
            cTbc.Name = "002 Data";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "002_prv";
            cTbc.Name = "002 prezzo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_gia";
            cTbc.Name = "Q.tà Giacenza";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##00.00";
        }

        private void FillDati()
        {
            string s = "";

            DataTable tTmp = new DataTable();
            DataRow x;
            DataRow[] j;

            s = "SELECT ";
            s += "AnaArticoli.art_sta, ";
            s += "AnaArticoli.art_cod, ";
            s += "AnaArticoli.art_des,  ";
            s += "AnaArticoli.art_umi, ";
            s += "AnaArticoli.art_rep ";
            s += "FROM AnaArticoli ";
            s += "ORDER BY art_des";
            DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);

            s = "SELECT * FROM TabReparti";
            DataTable tRep = _clsFun.FillTabSql("TabReparti", s, false, _strConSql);

            s = "SELECT * FROM AnaArtGiacenza";
            DataTable tGia = _clsFun.FillTabSql("TabReparti", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            progressBar1.Value = 0;
            progressBar1.Maximum = tArt.Rows.Count;
            progressBar1.Minimum = 0;

            DataTable t = new clsGenTabTmp().TabTmpPrzNeg("TabPrv");

            foreach(DataRow y in tArt.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                if (!chkAtt.Checked || (string)y["art_sta"] == "A")
                {
                    x = t.NewRow();

                    x["art_cod"] = y["art_cod"];
                    x["art_des"] = y["art_des"];
                    x["art_umi"] = y["art_umi"];
                    x["art_rep"] = y["art_rep"];

                    j = tRep.Select("tab_cod='" + y["art_rep"] + "'");
                    if (j.Length > 0)
                        x["art_red"] = (string)j[0]["tab_des"];

                    x["art_gia"] = 0;

                    s = "SELECT TOP 50 liv_prv, liv_lis, liv_dti, liv_dtf, liv_ann FROM GesLisVendita WHERE ";
                    s += "liv_art='" + y["art_cod"] + "' AND ";
                    s += "liv_dti <= " + _clsFun.DaySql(DateTime.Today) + " ";
                    //s += "liv_ann=0 ";
                    s += "ORDER BY liv_dti DESC";

                    tTmp = _clsFun.FillTabSql("Tmp", s, false, _strConSql);

                    if (tTmp.Rows.Count > 0)
                    {
                        string sReadLis = "";

                        foreach(DataRow yy in tTmp.Rows)
                        {
                            if(!sReadLis.Contains((string)yy["liv_lis"]))
                            {
                                string sLis = (string)yy["liv_lis"];

                                sReadLis += sLis;

                                if (!(Boolean)yy["liv_ann"])
                                {
                                    if (sLis == "001")
                                    {
                                        x["001_day"] = yy["liv_dti"];
                                        x["001_prv"] = yy["liv_prv"];
                                    }
                                    else if (sLis == "002")
                                    {
                                        x["002_day"] = yy["liv_dti"];
                                        x["002_prv"] = yy["liv_prv"];
                                    }
                                }
                            }
                        }
                    }

                    j = tGia.Select("gia_art='" + y["art_cod"] + "'");
                    if(j.Length > 0)
                    {
                        x["art_gia"] = (decimal)j[0]["gia_gia"];
                    }

                    t.Rows.Add(x);
                }

            }

            dgv1.DataSource = t;

        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Art2Xls();
        }

        private void Art2Xls()
        {
            DataTable t = (DataTable)dgv1.DataSource;

            string sTit = "Estrazione del " +DateTime.Today.ToString("dd/MM/yyyy");
            string sFil = "";
            string sFoo = "";

            string sFld = "";
            sFld += "art_red, Reparto, 50, StringLiteral;";
            sFld += "art_cod, Codice, 50, StringLiteral;";
            sFld += "art_des, Descrizione, 180, StringLiteral;";
            sFld += "art_umi, UM, 50, StringLiteral;";
            sFld += "001_day, Data, 70, DateTime;";
            sFld += "001_prv, Prezzo, 50, Decimal;";
            sFld += "002_day, Data, 70, DateTime;";
            sFld += "002_prv, Prezzo, 50, Decimal;";
            sFld += "art_gia, Giacenza, 70, Decimal;";

            for (int i = 0; i <= 100; i++)
            {
                sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_ArtPrezzi.xls";
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

            sFil = "ArtPrezzi.xls";

            (new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);

        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2 && e.RowIndex >= 0)
            {
                string sArt = dgv1.Rows[e.RowIndex].Cells["Articolo"].Value.ToString();
                frmAnaArticolo f = new frmAnaArticolo();
                f._strArtCod = sArt;
                f.ShowDialog();
            }
        }

    }
}
