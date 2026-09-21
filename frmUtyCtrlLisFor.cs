using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmUtyCtrlLisFor : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";

        public frmUtyCtrlLisFor()
        {
            InitializeComponent();
        }

        private void frmUtyCtrlLisFor_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private void frmUtyCtrlLisFor_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void btnFil_Click(object sender, EventArgs e)
        {
            FillDati();
            FillCtrl();
        }

        private void FillDati()
        {
            string s = "";

            s = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + txtLis.Text + ";Extended Properties=\"Excel 8.0;HDR={1};IMEX=1\"";
            OleDbConnection cn = new System.Data.OleDb.OleDbConnection(s);
            cn.Open();

            //DataTable sch = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            DataTable sch = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sTab = (string)sch.Rows[0][2];

            OleDbCommand cm = new OleDbCommand();
            cm.Connection = cn;
            //s = "SELECT * FROM [" + sTab + "] WHERE 'Cod# referenza' <> null AND 'Bar code' <> null";
            s = "SELECT * FROM [" + sTab + "]";

            cm.CommandText = s;
            OleDbDataReader dr = cm.ExecuteReader();
            DataTable tXls = new DataTable();
            tXls.TableName = "TabXls";
            tXls.Load(dr);

            DataColumn[] keys = new DataColumn[2];
            keys[0] = tXls.Columns["Cod# referenza"];
            keys[1] = tXls.Columns["Bar code"];
            tXls.PrimaryKey = keys;

            dgv1.DataSource = tXls;
        }

        private void FillCtrl()
        {
            DataRow x;
            DataRow[] j;

            string s = "";
            
            //s = "SELECT GesLisAcquisto.*, AnaArticoli.art_des FROM GesLisAcquisto ";
            //s += "LEFT OUTER JOIN AnaArticoli ON GesLisAcquisto.lia_art = AnaArticoli.art_cod ";
            //s += "ORDER BY lia_dti DESC";

            //DataTable tLia = _clsFun.FillTabSql("GesLisAcquisto", s, false, _strConSql);
            //DataColumn[] keys = new DataColumn[5];
            //keys[0] = tLia.Columns["lia_arf"];
            //keys[1] = tLia.Columns["lia_art"];
            //keys[2] = tLia.Columns["lia_tiP"];
            //keys[3] = tLia.Columns["lia_for"];
            //keys[4] = tLia.Columns["lia_dti"];
            //tLia.PrimaryKey = keys;

            s = "SELECT GesLisAcquisto.*, AnaArticoli.art_des FROM GesLisAcquisto ";
            s += "LEFT OUTER JOIN AnaArticoli ON GesLisAcquisto.lia_art = AnaArticoli.art_cod ";
            s += "WHERE lia_for='00001' AND lia_ann=0 ";
            s += "ORDER BY lia_dti DESC";

            DataTable tLia = _clsFun.FillTabSql("GesLisAcquisto", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[5];
            keys[0] = tLia.Columns["lia_arf"];
            keys[1] = tLia.Columns["lia_art"];
            keys[2] = tLia.Columns["lia_tiP"];
            keys[3] = tLia.Columns["lia_for"];
            keys[4] = tLia.Columns["lia_dti"];
            tLia.PrimaryKey = keys;

            DataTable tLis = (DataTable)dgv1.DataSource;

            DataTable t = TabTmpPrzNeg("Tab");

            progressBar1.Value = 0;
            progressBar1.Maximum = tLis.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in tLis.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                x = t.NewRow();
                x["arf_art"] = y["Codice articolo"];
                x["arf_ard"] = y["Descrizione articolo"];
                x["art_art"] = "";
                x["art_ard"] = "";

                j = tLia.Select("lia_arf='" + x["arf_art"] + "'");
                if(j.Length > 0)
                {
                    x["art_art"] = j[0]["lia_art"];
                    x["art_ard"] = j[0]["art_des"];
                }

                t.Rows.Add(x);
            }

            foreach(DataRow y in t.Rows)
            {
                if((string)y["art_art"] != "")
                {
                    j = t.Select("art_art='" + y["art_art"] + "'");
                    y["arf_cnt"] = j.Length;
                }
            }

            dgv2.DataSource = t;

        }

        public DataTable TabTmpPrzNeg(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {                
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "arf_art",
                        Caption = "Articolo",
                        MaxLength = 10,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "arf_ard",
                        Caption = "Descrizione",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "art_art",
                        Caption = "Articolo",
                        MaxLength = 7,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "art_ard",
                        Caption = "Descrizione",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "arf_cnt",
                    Caption = "Numero articoli",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                    },

                }
            };

            return t;
        }

        private void dgv2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                string sArt = dgv2.Rows[e.RowIndex].Cells["art_art"].Value.ToString();

                frmAnaArticolo f = new frmAnaArticolo();
                f._strArtCod = sArt;
                f.ShowDialog();
            }
        }



    }
}
