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
    public partial class frmUtyArtFillMerceologie : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private const string TABLV1 = "TabEcrLv1";
        private const string TABLV2 = "TabEcrLv2";
        private const string TABLV3 = "TabEcrLv3";
        private const string TABREP = "TabReparti";

        private string _strConSql = "";

        public frmUtyArtFillMerceologie()
        {
            InitializeComponent();
        }

        private void frmUtyArtFillMerceologie_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            SetDgv1();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmUtyArtFillMerceologie_KeyDown(object sender, KeyEventArgs e)
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
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 190;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_rep";
            cTbc.Name = "Reparto";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_red";
            cTbc.Name = "Descrizione";
            cTbc.Width = 190;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_l1c";
            cTbc.Name = "Liv 1";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_l1d";
            cTbc.Name = "Descrizione";
            cTbc.Width = 190;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_l2c";
            cTbc.Name = "Liv 2";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_l2d";
            cTbc.Name = "Descrizione";
            cTbc.Width = 190;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_l3c";
            cTbc.Name = "Liv 3";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_l3d";
            cTbc.Name = "Descrizione";
            cTbc.Width = 190;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillDati()
        {
            string s = "";
            DataRow x;
            DataRow[] j;

            s = "SELECT * FROM " + TABREP;
            DataTable tRep = _clsFun.FillTabSql(TABREP, s, false, _strConSql);

            s = "SELECT * FROM " + TABLV1;
            DataTable tLv1 = _clsFun.FillTabSql(TABLV1, s, false, _strConSql);

            s = "SELECT * FROM " + TABLV2;
            DataTable tLv2 = _clsFun.FillTabSql(TABLV2, s, false, _strConSql);

            s = "SELECT * FROM " + TABLV3;
            DataTable tLv3 = _clsFun.FillTabSql(TABLV3, s, false, _strConSql);

            s = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + lblPath.Text + ";Extended Properties=\"Excel 12.0;HDR={1};IMEX=1\"";

            OleDbConnection cn = new System.Data.OleDb.OleDbConnection(s);
            cn.Open();

            DataTable sch = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sTab = (string)sch.Rows[0][2];

            OleDbCommand cm = new OleDbCommand();
            cm.Connection = cn;

            //OleDbDataReader dr = cm.ExecuteReader();
            //DataTable sch = dr.GetSchemaTable();
            //s = sch.TableName[0].ToString();
            //s = "SELECT * FROM [Sheet1$]";

            s = "SELECT * FROM [" + sTab + "]";
            cm.CommandText = s;
            OleDbDataReader dr = cm.ExecuteReader();
            DataTable tXls = new DataTable();
            tXls.TableName = "TabXls";
            tXls.Load(dr);

            tXls.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_rep",
                Caption = "Reparto",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tXls.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_red",
                Caption = "Reparto descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tXls.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_l1c",
                Caption = "Settore",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tXls.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_l1d",
                Caption = "Settore descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tXls.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_l2c",
                Caption = "Famiglia",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tXls.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_l2d",
                Caption = "Fam descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tXls.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_l3c",
                Caption = "SottoFamiglia",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tXls.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_l3d",
                Caption = "SottooFam descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            progressBar1.Value = 0;
            progressBar1.Maximum = tXls.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in tXls.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                y["tmp_rep"] = Convert.ToDecimal(y["ETO"]).ToString("000");

                j = tRep.Select("tab_cod='" + y["tmp_rep"] + "'");
                if (j.Length > 0)
                    y["tmp_red"] = (string)j[0]["tab_des"];

                y["tmp_l1c"] = Convert.ToDecimal(y["art_ec1"]).ToString("000");
                y["tmp_l2c"] = Convert.ToDecimal(y["art_ec2"]).ToString("000");
                y["tmp_l3c"] = Convert.ToDecimal(y["art_ec3"]).ToString("000");

                j = tLv1.Select("tab_cod='" + y["tmp_l1c"] + "'");
                if (j.Length > 0)
                    y["tmp_l1d"] = (string)j[0]["tab_des"];

                j = tLv2.Select("tab_cod='" + y["tmp_l2c"] + "' AND tab_lv1='" + y["tmp_l1c"] + "'");
                if (j.Length > 0)
                    y["tmp_l2d"] = (string)j[0]["tab_des"];

                j = tLv3.Select("tab_cod='" + y["tmp_l3c"] + "' AND tab_lv2='" + y["tmp_l2c"] + "' AND tab_lv1='" + y["tmp_l1c"] + "'");
                if (j.Length > 0)
                    y["tmp_l3d"] = (string)j[0]["tab_des"];



            }

            dgv1.DataSource = tXls;
        }

        private void btnSalva_Click(object sender, EventArgs e)
        {
            Salva();
        }

        private void Salva()
        {
            string s = "";

            DataTable t = (DataTable)dgv1.DataSource;

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if (!DBNull.Value.Equals(y["art_cod"]) && Convert.ToString(y["art_cod"]).Trim() != "" && !DBNull.Value.Equals(y["tmp_rep"]) && ((string)y["tmp_rep"]).Trim() != "" && !DBNull.Value.Equals(y["tmp_l1c"]) && ((string)y["tmp_l1c"]).Trim() != "" && !DBNull.Value.Equals(y["tmp_l2c"]) && ((string)y["tmp_l2c"]).Trim() != "" && !DBNull.Value.Equals(y["tmp_l3c"]) && ((string)y["tmp_l3c"]).Trim() != "")
                {
                    string sArt = Convert.ToDecimal(y["art_cod"]).ToString("0000000");

                    s = "UPDATE AnaArticoli SET ";
                    s += "art_rep='" + y["tmp_rep"] + "', ";
                    s += "art_ec1='" + y["tmp_l1c"] + "', ";
                    s += "art_ec2='" + y["tmp_l2c"] + "', ";
                    s += "art_ec3='" + y["tmp_l3c"] + "' ";
                    s += "WHERE art_cod='" + sArt + "'";

                    _clsFun.SqlWrite(s, _strConSql);
                }
            }
        }
        
    }
}
