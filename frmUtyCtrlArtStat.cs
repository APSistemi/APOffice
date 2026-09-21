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
    public partial class frmUtyCtrlArtStat : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsVariazioni _clsVar = new clsVariazioni();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strConSqlSta = "";

        private const string TABANAART = "AnaArticoli";

        public frmUtyCtrlArtStat()
        {
            InitializeComponent();
        }

        private void frmUtyCtrlArts_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
            SetDgv1();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private void frmUtyCtrlArts_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
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
            cTbc.DataPropertyName = "art_sta";
            cTbc.Name = "Stato";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.000";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_prv";
            cTbc.Name = "Prezzo vendita";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_msg";
            cTbc.Name = "Fornitori";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillDati()
        {
            string s = "";
            DataRow[] j;

            s = "SELECT ven_art FROM GesNegVen GROUP BY ven_art ORDER BY ven_art";
            DataTable tSta = _clsFun.FillTabSql(TABANAART, s, false, _strConSqlSta);

            s = "SELECT art_sta, art_cod, art_des FROM " +  TABANAART + " WHERE art_sta='" + _clsDef.STANOA + "'";
            DataTable tArt = _clsFun.FillTabSql(TABANAART, s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            tArt.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_cos",
                Caption = "Costo",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            tArt.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_prv",
                Caption = "Vendita",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            tArt.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_msg",
                Caption = "Note",
                MaxLength = 200,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            DataTable t = tArt.Clone();

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = tSta.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in tSta.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                j = tArt.Select("art_cod='" + y["ven_art"] + "'");
                if (j.Length > 0)
                {
                    DataTable tTmp = _clsQry.ArtSeek((string)y["ven_art"], "");
                    if (tTmp.Rows.Count > 0)
                    {

                        j[0]["tmp_cos"] = tTmp.Rows[0]["tmp_cos"];
                        j[0]["tmp_prv"] = tTmp.Rows[0]["tmp_prv"];
                    }

                    tTmp = _clsQry.ArtFornitori((string)y["ven_art"]);

                    if (tTmp.Rows.Count > 0)
                    {
                        foreach (DataRow x in tTmp.Rows)
                        {
                            j[0]["tmp_msg"] += (string)x["for_des"] + " - ";
                        }
                    }

                    t.ImportRow(j[0]);

                }
            }

            dgv1.DataSource = t;

            lblCnt.Text = t.Rows.Count.ToString("###,##0");
        }

        private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = (string)x["art_sta"];
                    if (s == _clsDef.STAATT)
                        x["art_sta"] = _clsDef.STANOA;
                    else if (s == _clsDef.STANOA)
                        x["art_sta"] = _clsDef.STAATT;
                }
            }
            else if (e.ColumnIndex == 2)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = (string)x["art_cod"];
                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = s;
                    f.ShowDialog();
                }
            }

        }

        private void btnAgg_Click(object sender, EventArgs e)
        {
            Salva();
        }

        private void Salva()
        {
            //string s = "SELECT art_sta, art_cod, art_des FROM " + TABANAART + " WHERE art_sta='" + _clsDef.STAATT + "'";
            //DataTable tArt = _clsFun.FillTabSql(TABANAART, s, false, _strConSql);
            //DataColumn[] Key = new DataColumn[1] 
            //{ 
            //    tArt.Columns["art_cod"] 
            //};
            //tArt.PrimaryKey = Key;

            string s = "";

            DataTable t = (DataTable)dgv1.DataSource;

            foreach(DataRow y in t.Rows)
            {
                if ((string)y["art_sta"] == _clsDef.STANOA)
                {
                    s = "UPDATE " + TABANAART + " SET ";
                    s += "art_sta='" + _clsDef.STAATT + "', ";
                    s += "art_dtm=" + _clsFun.DaySql(DateTime.Today) + " ";
                    s += "WHERE art_cod='" + y["art_cod"] + "'";
                    _clsFun.SqlWrite(s, _strConSql);
                }
            }

            MessageBox.Show("Fine aggiornamento!");
        }

        private void recuperoStatoDaApShopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FillStaApShop();
        }

        private void FillStaApShop()
        {
            string sMdb = "C:\\ApProject\\ApShop\\DataBase\\DataBase.mdb";

            if (!File.Exists(sMdb))
                MessageBox.Show(sMdb, "MDB non presente!");
            else
            {
                DataRow[] j;

                string sCn = _clsFun.ConMdb(sMdb);

                string p = "AnaArticoli";
                string s = "SELECT art_cod FROM AnaArticoli ";
                DataTable tMdb = _clsFun.FillTabMdb(p, s, false, sCn);

                DataColumn[] keys = new DataColumn[1];
                keys[0] = tMdb.Columns["art_cod"];
                tMdb.PrimaryKey = keys;

                DataTable t = (DataTable)dgv1.DataSource;

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = t.Rows.Count;
                progressBar1.Minimum = 0;

                foreach(DataRow y in t.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    j = tMdb.Select("art_cod='" + (String)y["art_cod"] + "'");
                    if(j.Length == 0)
                        y["art_sta"] = _clsDef.STANOA;

                }


            }


        }
    }
}
