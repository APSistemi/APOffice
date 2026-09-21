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
    public partial class frmUtyArtEstrazioni : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsVariazioni _clsVar = new clsVariazioni();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";

        private const string TABANAART = "AnaArticoli";

        public frmUtyArtEstrazioni()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmUtyCtrlArts_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
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
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_cho";
            cTbc.Name = "Scelto";
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
            cTbc.DataPropertyName = "art_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.000";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_prv";
            cTbc.Name = "Prezzo vendita";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "RepDes";
            cTbc.Name = "Reparto";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillDati()
        {
            string sWhe = "";
            if (rdbArtWeb.Checked)
                sWhe = "art_web=1";
            else if (rdbArtCel.Checked)
                sWhe = "art_cel=1";
            else if (rdbArtIng.Checked)
                sWhe = "art_plu<>''";

            if (sWhe == "")
                MessageBox.Show("Nessuna selezione valida", "CONTROLLO PARAMETRI",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
            else
            {
                //string s = "SELECT art_sta, art_cod, art_des, art_cos, art_prv FROM AnaArticoli WHERE " + sWhe;

                string s = "SELECT ";
                s += "AnaArticoli.art_cod, ";
                s += "AnaArticoli.art_des, ";
                s += "AnaArticoli.art_umi, ";
                s += "AnaArticoli.art_rep, ";
                s += "AnaArticoli.art_cos,  ";
                s += "AnaArticoli.art_prv,  ";
                s += "AnaArticoli.art_plu,  ";
                s += "TabReparti.tab_des AS RepDes ";
                s += "FROM AnaArticoli LEFT OUTER JOIN TabReparti ON AnaArticoli.art_rep = TabReparti.tab_cod ";
                s += "WHERE " + sWhe;
                DataTable t = _clsFun.FillTabSql(TABANAART, s, false, _strConSql);
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "tmp_cho",
                    Caption = "Scelto",
                    ReadOnly = false,
                    DefaultValue = (String)"S"
                });

                dgv1.DataSource = t;
                lblCnt.Text = t.Rows.Count.ToString("###,##0");
            }
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = (string)x["tmp_cho"];
                    if (s == "S")
                        x["tmp_cho"] = "";
                    else if (s == "")
                        x["tmp_cho"] = "S";
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
            //_clsFun.ErrorLog("001", "passo");

            if(dgv1.DataSource == null)
                MessageBox.Show("Non sono presenti articoli estratti!!", "CONTROLLO PARAMETRI", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if (rdbArtWeb.Checked)
            {
                string sWebPath = _clsFun.ParGet(clsDefine.enuParametri.Par029ParPathWeb, _strConSql);
                if(sWebPath == "")
                    MessageBox.Show("Generazione divulgazione WEB non configurata!", "CONTROLLO PARAMETRI", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else
                    GenWebFile(sWebPath);
            }
            else if (rdbArtIng.Checked)
            {
                //_clsFun.ErrorLog("002", "passo");

                string sIngPath = _clsFun.ParGet(clsDefine.enuParametri.Par030ParPathIng, _strConSql);
                if (sIngPath == "")
                    MessageBox.Show("Generazione ingredienti WEB non configurata!", "CONTROLLO PARAMETRI", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else
                {
                    DataTable t = ((DataTable)dgv1.DataSource).Clone();

                    foreach (DataRow y in ((DataTable)dgv1.DataSource).Rows)
                    {
                        if ((string)y["tmp_cho"] == "S")
                            t.ImportRow(y);
                    }

                    if (t.Rows.Count > 0)
                    {
                        //_clsFun.ErrorLog("003", "passo ");

                        string[] a = sIngPath.Split(',');
                        if (a.Length > 1)
                        {
                            //_clsFun.ErrorLog("004", "passo");
                            new clsOpenXml1().xlsLibroIngredienti(t, a[1], chkPrnPrv.Checked);
                        }
                    }
                }
            }
        }

        private void GenWebFile(string strPth)
        {
            string sIni = "\"";
            string sFin = "\";";
            decimal d = 0;
            DataRow[] j;

            string s = "";

            string[] a = strPth.Split(',');

            if (a.Length > 1)
            {
                string sFil = a[1];           // +"ExportArticoliWeb.txt";

                if (File.Exists(sFil))
                    File.Delete(sFil);

                DataTable tEcr = _clsQry.tabEcr();

                s = "SELECT * FROM TabReparti";
                DataTable tRep = _clsFun.FillTabSql("TabReparti", s, false, _strConSql);

                DataTable t = (DataTable)dgv1.DataSource;

                StreamWriter sw = new StreamWriter(sFil, true);

                string sRig = "\"CodiceInterno\";\"Codice\";\"Descrizione\";\"Prezzo Ivato\";\"Codice Livello 1 Piano Merceologico\";\"Descrizione Livello 1 Piano Merceologico\";\"Codice Livello 2 Piano Merceologico\";\"Descrizione Livello 2 Piano Merceologico\";\"Codice Area\";\"Descrizione Area\";\"Codice EAN\";\"Um Vendita\";\"Plu\";\"Ingredienti\";\"Orario Aggiornamento\"";
                sw.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in t.Rows)
                {
                    if ((string)y["tmp_cho"] == "S")
                    {
                        s = (string)y["art_cod"];
                        DataTable tTmp = _clsQry.ArtSeek(s, "");
                        if (tTmp.Rows.Count > 0)
                        {
                            sRig = sIni + (string)y["art_cod"] + sFin;
                            sRig += sIni + (string)y["art_cod"] + sFin;
                            sRig += sIni + (string)y["art_des"] + sFin;

                            d = (decimal)tTmp.Rows[0]["tmp_prv"];
                            sRig += d.ToString() + ";";

                            s = (string)tTmp.Rows[0]["tmp_ecr"];
                            if (s.Length > 8)
                            {
                                string sLv1 = s.Substring(0, 3);
                                string sLv2 = s.Substring(3, 3);
                                string sLv3 = s.Substring(6, 3);

                                s = "l1c='" + sLv1 + "' AND l2c='" + sLv2 + "'AND l3c='" + sLv3 + "'";
                                j = tEcr.Select(s);
                                if (j.Length > 0)
                                {
                                    sRig += sIni + sLv1 + sFin;
                                    sRig += sIni + (string)j[0]["l1d"] + sFin;

                                    sRig += sIni + sLv2 + sFin;
                                    sRig += sIni + (string)j[0]["l2d"] + sFin;

                                    sRig += sIni + sLv3 + sFin;
                                    sRig += sIni + (string)j[0]["l3d"] + sFin;
                                }
                                else
                                    sRig += sIni + sFin + sIni + sFin + sIni + sFin + sIni + sFin + sIni + sFin + sIni + sFin;
                            }
                            else
                                sRig += sIni + sFin + sIni + sFin + sIni + sFin + sIni + sFin + sIni + sFin + sIni + sFin;


                            j = tRep.Select("tab_cod='" + (string)tTmp.Rows[0]["tmp_rep"] + "'");
                            if (j.Length > 0)
                            {
                                sRig += sIni + (string)j[0]["tab_cod"] + sFin;
                                sRig += sIni + (string)j[0]["tab_des"] + sFin;
                            }
                            else
                                sRig += sIni + sFin + sIni + sFin;

                            sRig += sIni + (string)tTmp.Rows[0]["tmp_ean"] + sFin;
                            sRig += sIni + (string)tTmp.Rows[0]["tmp_umi"] + sFin;
                            sRig += sIni + (string)tTmp.Rows[0]["tmp_plu"] + sFin;

                            string sIng = "";
                            s = _clsDef.PATHINGREDIENTI + "et01_" + y["art_cod"] + ".rtf";
                            if (File.Exists(s))
                            {
                                RichTextBox rch = new RichTextBox();
                                rch.LoadFile(s);
                                sIng = rch.Text;
                                //sIng = _clsFun.StrSplit(sIng, 40);
                            }

                            sRig += sIni + sIng + sFin;
                            sRig += sIni + DateTime.Now.ToString("dd/MM/yyyy HH.mm.ss") + sFin;

                            sw.Write(sRig + _clsDef.CRLF);
                        }
                    }
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();
            }
        }

        private void chkAll_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource != null)
            {
                string sSta = "";
                if (chkSelAll.Checked)
                    sSta = "S";
                DataTable t = (DataTable)dgv1.DataSource;
                foreach (DataRow y in t.Rows)
                    y["tmp_cho"] = sSta;
            }
        }

    }
}
