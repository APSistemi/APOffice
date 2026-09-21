using System;
using System.Collections;
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
    public partial class frmUtyCtrlArts : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsVariazioni _clsVar = new clsVariazioni();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strConSqlSta = "";

        private const string TABANAART = "AnaArticoli";

        public frmUtyCtrlArts()
        {
            InitializeComponent();
        }

        private void frmUtyCtrlArts_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");

            cmbSta.Items.Add("CANCELLATO");
            cmbSta.Items.Add("NON ATTIVO");

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
            if (cmbSta.Text == "")
                MessageBox.Show("Stato non impostato");
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
            //DataGridViewCheckBoxColumn cCbc;
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
            cTbc.DataPropertyName = "art_dti";
            cTbc.Name = "Dt inserimento articolo";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_pvd";
            cTbc.Name = "Dt prezzo vendita";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_dtm";
            cTbc.Name = "Dt ultimo movimento";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_dtv";
            cTbc.Name = "Dt ultima vendita";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_msg";
            cTbc.Name = "Fornitore";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillDati()
        {
            string s = "SELECT * FROM (";
            s += "SELECT ";
            s += "ROW_NUMBER() OVER (PARTITION BY liv_art ORDER BY liv_art, liv_dti DESC) AS ROW, ";
            s += "liv_art, ";
            s += "liv_dti, ";
            s += "liv_prv ";
            s += "FROM GesLisVendita ) AS A WHERE ROW = 1";
            DataTable tLiv = _clsFun.FillTabSql("GesLisVendita", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tLiv.Columns["liv_art"];
            tLiv.PrimaryKey = keys;

            s = "SELECT MAX(ven_day) AS ven_day, ven_art ";
            s += "FROM GesNegVen ";
            s += "GROUP BY ven_art";
            DataTable tVen = _clsFun.FillTabSql("GesNegVen", s, false, _strConSqlSta);
            keys = new DataColumn[1];
            keys[0] = tVen.Columns["ven_art"];
            tVen.PrimaryKey = keys;

            s = "SELECT * FROM (";
            s += "SELECT ";
            s += "ROW_NUMBER() OVER (PARTITION BY lia_art ORDER BY lia_art, lia_dti DESC) AS ROW, ";
            s += "lia_art, ";
            s += "lia_cos, ";
            s += "lia_for ";
            s += "FROM GesLisAcquisto ) AS A WHERE ROW = 1";
            DataTable tLia = _clsFun.FillTabSql("PCO", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tLia.Columns["lia_art"];
            tLia.PrimaryKey = keys;

            s = "SELECT for_cod, for_des FROM AnaFornitori";
            DataTable tFor = _clsFun.FillTabSql("AnaFornitori", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tFor.Columns["for_cod"];
            tFor.PrimaryKey = keys;

            s = "SELECT art_sta, art_cod, art_des, art_dti, art_dtm FROM AnaArticoli WHERE art_sta='" + _clsDef.STAATT + "'";
            DataTable t = _clsFun.FillTabSql(TABANAART, s, false, _strConSql);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_cos",
                Caption = "Costo",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_prv",
                Caption = "Vendita",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_msg",
                Caption = "Note",
                MaxLength = 200,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "tmp_pvd",
                Caption = "Data ultimo movimento di vendita",
                ReadOnly = false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "tmp_dtm",
                Caption = "Data ultimo movimento di magazzino",
                ReadOnly = false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "tmp_dtv",
                Caption = "Data ultima vendita",
                ReadOnly = false
            });

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                lblCnt.Text = progressBar1.Value.ToString();

                if ((string)y["art_cod"] == "0010740")
                {
                    Console.WriteLine("xxxxxxxxx");
                }

                /*
                DataTable tTmp = _clsQry.ArtSeek((string)y["art_cod"], "");
                if (tTmp.Rows.Count > 0)
                {
                    y["tmp_cos"] = tTmp.Rows[0]["tmp_cos"];
                    y["tmp_prv"] = tTmp.Rows[0]["tmp_prv"];
                }

                tTmp = _clsQry.ArtFornitori((string)y["art_cod"]);

                if (tTmp.Rows.Count > 0)
                {
                    foreach (DataRow x in tTmp.Rows)
                    {
                        y["tmp_msg"] += (string)x["for_des"] + " - ";
                    }
                }
                */

                //s = "SELECT liv_prv FROM GesLisVendita WHERE liv_art='" + y["art_cod"] + "' ORDER BY liv_dti DESC";
                //DataTable tTmp = _clsFun.FillTabSql("GesLisVendita", s, true, _strConSql);
                //if (tTmp.Rows.Count > 0)
                //    y["tmp_prv"] = tTmp.Rows[0]["liv_prv"];

                DataRow[] j = tLiv.Select("liv_art='" + y["art_cod"] + "'");
                if (j.Length > 0)
                {
                    y["tmp_pvd"] = (DateTime)j[0]["liv_dti"];
                    y["tmp_prv"] = (decimal)j[0]["liv_prv"];
                }
                //s = "SELECT lia_for, lia_cos FROM GesLisAcquisto WHERE lia_art='" + y["art_cod"] + "' ORDER BY lia_dti DESC";
                //DataTable tTmp = _clsFun.FillTabSql("GesLisAcquisto", s, true, _strConSql);
                //if (tTmp.Rows.Count > 0)
                //{
                //    y["tmp_prv"] = tTmp.Rows[0]["lia_cos"];

                //    j = tFor.Select("for_cod='" + tTmp.Rows[0]["lia_for"] + "'");
                //    if(j.Length > 0)
                //        y["tmp_msg"] += (string)j[0]["for_des"];
                //}

                j = tLia.Select("lia_art='" + y["art_cod"] + "'");
                if (j.Length > 0)
                {
                    y["tmp_cos"] = (decimal)j[0]["lia_cos"];
                    j = tFor.Select("for_cod='" + j[0]["lia_for"] + "'");
                    if (j.Length > 0)
                        y["tmp_msg"] += (string)j[0]["for_des"];
                }

                s = "SELECT mov_day FROM GesMovimenti WHERE mov_art='" + y["art_cod"] + "' ORDER BY mov_day DESC";
                DataTable tTmp = _clsFun.FillTabSql("GesMovimenti", s, true, _strConSql);
                if (tTmp.Rows.Count > 0)
                    y["tmp_dtm"] = tTmp.Rows[0]["mov_day"];

                //s = "SELECT ven_day FROM GesNegVen WHERE ven_art='" + y["art_cod"] + "' ORDER BY ven_day DESC";
                //tTmp = _clsFun.FillTabSql("GesNegVen", s, true, _strConSqlSta);
                //if (tTmp.Rows.Count > 0)
                //    y["tmp_dtv"] = tTmp.Rows[0]["ven_day"];

                j = tVen.Select("ven_art='" + y["art_cod"] + "'");
                if(j.Length > 0)
                    y["tmp_dtv"] = j[0]["ven_day"];


                //if (progressBar1.Value > 1000)
                //    break;
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

                    string sSta = cmbSta.Text.Substring(0, 1);

                    string s = (string)x["art_sta"];
                    if (s != sSta)
                        x["art_sta"] = sSta;
                    else if (s == sSta)
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

            string sPar016PathDivNegozi = _clsFun.ParGet(clsDefine.enuParametri.Par016PathDivNegozi, _strConSql);

            if (dgv1.DataSource == null)
                MessageBox.Show("Dati non presenti!");
            else if (cmbSta.Text == "")
                MessageBox.Show("Stato non impostato!");
            else
            {
                string sSta = cmbSta.Text.Substring(0, 1);

                string s = "";

                DataTable t = (DataTable)dgv1.DataSource;

                ArrayList a = new ArrayList();

                foreach (DataRow y in t.Rows)
                {
                    if ((string)y["art_sta"] == sSta)
                    {
                        s = "UPDATE AnaArticoli SET ";
                        s += "art_sta='" + sSta + "', ";
                        s += "art_dtm=" + _clsFun.DaySql(DateTime.Today) + " ";
                        s += "WHERE art_cod='" + y["art_cod"] + "'";
                        _clsFun.SqlWrite(s, _strConSql);

                        _clsVar.Variazioni((string)y["art_cod"], s, "Utility annullo articolo", _clsDef.VARPOS);

                        a.Add((string)y["art_cod"]);
                    }
                }

                _clsVar.DivNegArticoli(sPar016PathDivNegozi, a, null, "", "");

                MessageBox.Show("Fine aggiornamento!");
            }
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

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
