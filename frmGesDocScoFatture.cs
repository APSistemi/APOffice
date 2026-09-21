using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmGesDocScoFatture : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private const string TABSTAVET = "GesNegVet";
        private const string TABSTAVEN = "GesNegVen";
        private const string TABSTAVEP = "GesNegVep";

        private const string TABGESFAT = "GesFatTestate";
        private const string TABGESMOT = "GesMovTestate";
        private const string TABGESMOV = "GesMovimenti";
        private const string TABTABSTD = "TabStatoDivulgazioni";

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public string _strPar026FattFromCassa = "";

        public frmGesDocScoFatture()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmGesDocScoFatture_Load(object sender, EventArgs e)
        {
            dtpDti.Value = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);

            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
            SetDgv1();
            FillDati();
        }

        private void frmGesDocScoFatture_KeyDown(object sender, KeyEventArgs e)
        {   
            if(e.KeyCode == Keys.Escape)
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
            FillDati();
        }

        private void SetDgv1()
        {
            DataTable tSta = _clsFun.FillTabSql(TABTABSTD, "SELECT * FROM " + TABTABSTD + " WHERE tab_var=1", false, _strConSql);

            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            DataGridViewComboBoxColumn cCmb;

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "mov_idx";
            //cTbc.Name = "Idx";
            //cTbc.Width = 20;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = true;
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dgv1.Columns.Add(cTbc);

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "DocInv";
            cCmb.Name = "Stato";
            cCmb.Width = 60;
            cCmb.DataSource = tSta;
            cCmb.ValueMember = "tab_cod";
            cCmb.DisplayMember = "tab_des";
            cCmb.ReadOnly = true;
            cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            cCmb.DefaultCellStyle.Font = new Font("Microsoft Sans", 8.75F, GraphicsUnit.Pixel);
            dgv1.Columns.Add((DataGridViewColumn)cCmb);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "DocInv";
            //cTbc.Name = "Invio";
            //cTbc.Width = 60;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc); 
            
            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_day";
            cTbc.Name = "Data";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_cau";
            cTbc.Name = "Causale";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_ora";
            cTbc.Name = "Ora";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_pos";
            cTbc.Name = "POS";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_sco";
            cTbc.Name = "Scontrino";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_imp";
            cTbc.Name = "Importo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_fnu";
            cTbc.Name = "Documento";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "CliPiv";
            cTbc.Name = "P.Iva";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "CliDes";
            cTbc.Name = "Descrizione";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "CliCod";
            cTbc.Name = "Codice";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocErr";
            cTbc.Name = "Errori";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);
        }

        private void FillDati()
        {
            DataRow[] j;
            string s = "";

            s = "SELECT * FROM TabDocTipo";
            DataTable tTdo = _clsFun.FillTabSql("TabDocTipo", s, false, _strConSql);

            s = "SELECT * FROM TabMovCausali";
            DataTable tMca = _clsFun.FillTabSql("TabMovCausali", s, false, _strConSql);

            s = "SELECT * FROM AnaClienti";
            DataTable tCli = _clsFun.FillTabSql("AnaClienti", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tCli.Columns["cli_cod"];
            tCli.PrimaryKey = keys;

            s = "SELECT vet_idx, vet_day, vet_cau, vet_ora, vet_pos, vet_sco, vet_imp, vet_fnu, vet_fpi, vet_neg, vet_doc ";
            s += "FROM GesNegVet ";
            s += "WHERE ";
            s += "vet_day >= " + _clsFun.DaySql(dtpDti.Value) + " AND ";
            s += "vet_fnu <> ''";
            if(chkNoAll.Checked)
                s += " AND vet_sta = ''";
            DataTable t = _clsFun.FillTabSql("GesNegVet", s, false, _strConSqlSta);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "CliCod",
                Caption = "Codice",
                MaxLength = 5,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "CliPiv",
                Caption = "P.Iva",
                MaxLength = 16,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "CliDes",
                Caption = "Cld",
                MaxLength = 70,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "DocTip",
                Caption = "Fatt/Mov",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "DocTdo",
                Caption = "Tipo doc",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "DocTdd",
                Caption = "Descrizione doc",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "DocTpa",
                Caption = "Tipo pagamento",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "DocStd",
                Caption = "Stato pagamento",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "DocInv",
                Caption = "Invio",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)"0"
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "DocErr",
                Caption = "Errore",
                MaxLength = 150,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            foreach (DataRow y in t.Rows)
            {
                j = tCli.Select("cli_cod='" + (string)y["vet_fpi"] + "'");
                if(j.Length > 0)
                {
                    y["CliCod"] = (string)j[0]["cli_cod"];
                    y["CliPiv"] = (string)j[0]["cli_piv"];
                    y["CliDes"] = (string)j[0]["cli_des"];
                }

                string sTip = "";
                string sTdo = "";
                string sTpa = "";       //Tipo pagamento
                string sStd = "";       //Stato pagamento

                s = ((string)y["vet_doc"]).Trim();

                if (s != "")
                {
                    string[] a = s.Split('|');

                    Console.WriteLine("zzz");

                    if(a.Length > 1)
                    {
                        s = a[0];

                        string[] aa = s.Split('-');
                        if (aa.Length > 0)
                        {
                            sTip = aa[0];
                            if (sTip != "MOV")
                                sTip = "FAT";
                        }
                        sTdo = a[1];
                        if(a.Length > 2)
                            sTpa = a[2];
                        if (a.Length > 3)
                            sStd = a[3];
                    }

                    y["DocTip"] = sTip;
                    y["DocTdo"] = sTdo;
                    y["DocTpa"] = sTpa;
                    y["DocStd"] = sStd;

                    //y["DocTdd"] = sTdd;

                    if (sTip == "MOV")
                        j = tMca.Select("tab_cod='" + sTdo + "'");
                    else
                        j = tTdo.Select("tab_cod='" + sTdo + "'");
                        
                    if(j.Length > 0)
                        y["DocTdd"] = (string)j[0]["tab_des"];

                    if(sTip == "FAT")
                    {
                        string sYea = ((DateTime)y["vet_day"]).Year.ToString();
                        string sFnu = (string)y["vet_fnu"];

                        s = "SELECT * FROM GesFatTestate WHERE fat_yfa='" + sYea + "' AND fat_tpd='FV' AND fat_ndo='" + sFnu.PadLeft(10,Convert.ToChar("0")) + "'";
                        DataTable tTmp = _clsFun.FillTabSql("GesFatTestate", s, true, _strConSql);
                        if(tTmp.Rows.Count > 0)
                        {
                            y["DocErr"] = "Documento già presente";
                        }
                    }

                }
            }

            dgv1.DataSource = t;
        }

        private void btnImp_Click(object sender, EventArgs e)
        {
            ImpDocumenti();
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

                    string s = (string)x["DocInv"];
                    if (s == "5")
                        x["DocInv"] = "0";
                    else if (s == "0")
                        x["DocInv"] = "4";
                    else if (s == "4")
                        x["DocInv"] = "5";
                }
            }
        }

        private void dgv1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dgv1.Columns[e.ColumnIndex].Name == "Errori")
            {
                if (!DBNull.Value.Equals(dgv1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value))
                {
                    string s = (string)dgv1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                    if (s != "")
                        e.CellStyle.BackColor = Color.Red;
                }
            }
        }

        private void ImpDocumenti()
        {
            DataRow[] j;

            string s = "SELECT * FROM TabNegozi ORDER BY tab_cod";
            DataTable tNeg = _clsFun.FillTabSql("TabNegozi", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tNeg.Columns["tab_cod"];
            tNeg.PrimaryKey = keys;

            s = "SELECT * FROM TabIva ORDER BY tab_cod";
            DataTable tIva = _clsFun.FillTabSql("TabIva", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tIva.Columns["tab_cod"];
            tIva.PrimaryKey = keys;

            s = "SELECT * FROM TabMovCausali";
            DataTable tMca = _clsFun.FillTabSql("TabMovCausali", s, false, _strConSql);

            DataSet das = new DataSet();
            das.Tables.Add(tNeg);
            das.Tables.Add(tIva);

            DataTable t = (DataTable)dgv1.DataSource;
            foreach(DataRow y in t.Rows)
            {
                if ((string)y["DocInv"] == "0")
                {
                    string sTip = "";
                    string sCau = "";

                    s = ((string)y["vet_doc"]).Trim();
                    string[] a = s.Split('|');
                    if (a.Length > 3 && a[0].Length > 2 && a[0].Substring(0,3) == "MOV")    //Solo in caso di movimenti
                    {
                        //s = a[0];
                        //a = s.Split('-');
                        //if (a.Length > 2)
                        //{
                        sCau = a[1];
                        j = tMca.Select("tab_cod='" + sCau + "'");
                        if (j.Length > 0 && !DBNull.Value.Equals(j[0]["tab_tip"]))
                            sTip = (string)j[0]["tab_tip"];
                        //}
                    }

                    Vet2Doc(das, y, "", sTip);

                    s = ((string)y["vet_doc"]).Trim();
                    if (s != "")
                    {
                        a = s.Split('|');
                        if (a.Length > 0)
                        {
                            s = a[0];
                            a = s.Split('-');
                            if (a.Length > 2)
                            {
                                s = a[2];
                                if (s != "")
                                {
                                    string sTmpTdo = (string)y["DocTdo"];
                                    string sTmpNeg = (string)y["vet_neg"];
                                    y["DocTdo"] = s;

                                    j = tMca.Select("tab_cod='" + s + "'");
                                    if (j.Length > 0)
                                    {
                                        s = (string)j[0]["tab_mag"];
                                        j = tNeg.Select("tab_mag='" + s + "'");
                                        if (j.Length > 0)
                                        {
                                            y["vet_neg"] = (string)j[0]["tab_cod"];
                                            Vet2Doc(das, y, "2", "");
                                        }
                                    }
                                    y["DocTdo"] = sTmpTdo;
                                    y["vet_neg"] = sTmpNeg;
                                }
                            }
                        }
                    }

                    s = "UPDATE GesNegVet SET vet_sta='R' WHERE vet_idx=" + Convert.ToInt32(y["vet_idx"]);
                    _clsFun.SqlWrite(s, _strConSqlSta);
                }
                else if ((string)y["DocInv"] == "4")
                {
                    s = "UPDATE GesNegVet SET vet_sta='C' WHERE vet_idx=" + Convert.ToInt32(y["vet_idx"]);
                    _clsFun.SqlWrite(s, _strConSqlSta);
                }
            }

            FillDati();
        }

        private void Vet2Doc(DataSet dasTab, DataRow rowVet, string strPar, string strTip)
        {
            string s = "";
            DataRow x;
            DataRow[] j;

            DataTable tNeg = dasTab.Tables["TabNegozi"];
            DataTable tIva = dasTab.Tables["TabIva"];
            DataTable t = new DataTable("XXX");

            decimal d = 0;
            DateTime dDay = (DateTime)rowVet["vet_day"];
            string sOra = (string)rowVet["vet_ora"];
            string sPos = (string)rowVet["vet_pos"];
            string sNum = (string)rowVet["vet_sco"];
            decimal dImp = (decimal)rowVet["vet_imp"];
            string sTdo = (string)rowVet["DocTdo"];
            string sTip = (string)rowVet["DocTip"];
            string sCli = (string)rowVet["CliCod"];
            string sNdo = ((string)rowVet["vet_fnu"]).PadLeft(10, Convert.ToChar("0"));
            string sNeg = (string)rowVet["vet_neg"];
            string sTpg = (string)rowVet["DocTpa"];
            string sSta = (string)rowVet["DocStd"];

            string sUbi = _clsFun.FileIni("R", clsDefine.enuIni.Ini13Ubicazione, "");
            if (sUbi == "")
                sUbi = sNeg;

            if (strPar == "2")
                sNdo = "1" + sNdo.Substring(1);

            string sMag = "";
            j = tNeg.Select("tab_cod='" + sNeg + "'");
            if (j.Length > 0)
                sMag = (string)j[0]["tab_mag"];

            s = "SELECT * ";
            s += "FROM GesNegVet INNER JOIN GesNegVen ON ";
            s += "GesNegVet.vet_neg = GesNegVen.ven_neg AND ";
            s += "GesNegVet.vet_cau = GesNegVen.ven_cau AND ";
            s += "GesNegVet.vet_day = GesNegVen.ven_day AND ";
            s += "GesNegVet.vet_ora = GesNegVen.ven_ora AND ";
            s += "GesNegVet.vet_pos = GesNegVen.ven_pos AND ";
            s += "GesNegVet.vet_sco = GesNegVen.ven_sco ";
            s += "WHERE ";
            s += "(GesNegVet.vet_day = " + _clsFun.DaySql(dDay) + ") AND ";
            s += "(GesNegVet.vet_ora = '" + sOra + "') AND ";
            s += "(GesNegVet.vet_pos = '" + sPos + "') AND ";
            s += "(GesNegVet.vet_sco = '" + sNum + "') ";
            s += "ORDER BY ven_idx";
            DataTable tVen = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlSta);

            s = "SELECT * ";
            s += "FROM GesNegVet INNER JOIN GesNegVep ON ";
            s += "GesNegVet.vet_neg = GesNegVep.vep_neg AND ";
            s += "GesNegVet.vet_cau = GesNegVep.vep_cau AND ";
            s += "GesNegVet.vet_day = GesNegVep.vep_day AND ";
            s += "GesNegVet.vet_ora = GesNegVep.vep_ora AND ";
            s += "GesNegVet.vet_pos = GesNegVep.vep_pos AND ";
            s += "GesNegVet.vet_sco = GesNegVep.vep_sco ";
            s += "WHERE ";
            s += "(GesNegVet.vet_day = " + _clsFun.DaySql(dDay) + ") AND ";
            s += "(GesNegVet.vet_ora = '" + sOra + "') AND ";
            s += "(GesNegVet.vet_pos = '" + sPos + "') AND ";
            s += "(GesNegVet.vet_sco = '" + sNum + "')";
            DataTable tVep = _clsFun.FillTabSql(TABSTAVEP, s, false, _strConSqlSta);

            ArrayList aWhe = new ArrayList();
            ArrayList aExl = new ArrayList();

            if (tVen.Rows.Count > 0)
            {
                string sYea = dDay.Year.ToString();
                string sNfm = "";                       //Numero movimento fattura o movimento

                /* TESTATA FATTURA*/

                if (sTip == "FAT")
                {
                    s = "SELECT * FROM GesFatTestate WHERE fat_yfa='" + sYea + "' AND  fat_ndo='" + sNdo + "' AND fat_tpd='" + sTdo + "' AND fat_ubi='" + sUbi + "'";
                    t = _clsFun.FillTabSql(TABGESFAT, s, false, _strConSql);

                    x = t.NewRow();
                    x["fat_ubi"] = sUbi;
                    x["fat_yfa"] = sYea;
                    x["fat_tpd"] = sTdo;
                    x["fat_tdo"] = "FA";
                    //x["fat_nfa"] = sNfa;
                    x["fat_ndo"] = sNdo;
                    x["fat_ddo"] = dDay;
                    x["fat_cfo"] = sCli;
                    x["fat_no1"] = "";
                    x["fat_tpg"] = "";
                    x["fat_neg"] = (string)tVen.Rows[0]["vet_neg"];
                    x["fat_ann"] = false;
                    //y["fat_pvi"] = chkMftPvi.Checked;
                    x["fat_sta"] = "";
                    x["fat_des"] = "";                  //Destinazione
                    x["fat_neg"] = sNeg;                //Magazzino
                    x["fat_mag"] = sMag;                //Negozio
                    x["fat_tpg"] = sTpg;                //Tipo pagamento
                    x["fat_sta"] = sSta;                //Stato documento (Pagato)

                    if (t.Rows.Count == 0)
                    {
                        sNfm = _clsFun.NewNum(sYea, clsDefine.enuNumeratori.NumGesMovFatture, 6, _strConSql);
                        x["fat_nfa"] = sNfm;
                        s = _clsFun.SqlInsertRow(TABGESFAT, t, x);
                    }
                    else
                    {
                        sNfm = (string)t.Rows[0]["fat_nfa"];
                        aWhe = new ArrayList();
                        aWhe.Add("fat_ubi");
                        aWhe.Add("fat_yfa");
                        aWhe.Add("fat_nfa");
                        aWhe.Add("fat_tpd");
                        aExl = new ArrayList();
                        s = _clsFun.SqlUpdRow(TABGESFAT, t, t.Rows[0], x, aWhe, aExl);
                    }
                    if (s != "")
                    {
                        _clsFun.SqlWrite(s, _strConSql);
                        _clsFun.FileLog("GesFatTestate", sYea + "-" + sNdo, s);
                    }
                }
                else
                {
                    /* TESTATA MOVIMENTO */

                    s = "SELECT * FROM GesMovTestate WHERE mot_ymo='" + sYea + "' AND  mot_ndo='" + sNdo + "' AND mot_ubi='" + sUbi + "'";
                    t = _clsFun.FillTabSql("GesMovTestate", s, false, _strConSql);

                    x = t.NewRow();
                    x["mot_ymo"] = sUbi;
                    x["mot_ymo"] = sYea;
                    x["mot_day"] = DateTime.Today;
                    x["mot_cfo"] = sCli;
                    x["mot_cau"] = sTdo;
                    x["mot_ndo"] = sNdo;
                    x["mot_ddo"] = dDay;
                    x["mot_neg"] = sNeg;

                    x["mot_mag"] = sMag;

                    x["mot_no1"] = "";
                    x["mot_ann"] = false;
                    //y["mot_pvi"] = chkMftPvi.Checked;
                    x["mot_sta"] = "";
                    //y["mot_des"] = cmbMftDes.SelectedValue.ToString();

                    x["mot_des"] = "";              //Destinazione
                    x["mot_tpg"] = sTpg;            //Tipo pagamento
                    x["mot_sta"] = sSta;            //Stato documento (Pagato)

                    x["mot_yfa"] = _clsDef.COD04X;
                    x["mot_nfa"] = _clsDef.COD06X;

                    if (t.Rows.Count == 0)
                    {
                        sNfm = _clsFun.NewNum(sYea, clsDefine.enuNumeratori.NumGesMovimenti, 6, _strConSql);
                        x["mot_nmo"] = sNfm;
                        s = _clsFun.SqlInsertRow(TABGESMOT, t, x);
                    }
                    else
                    {
                        sNfm = (string)t.Rows[0]["mot_nmo"];
                        x["mot_nmo"] = (string)t.Rows[0]["mot_nmo"];
                        aWhe = new ArrayList();
                        aExl = new ArrayList();
                        aWhe.Add("mot_ubi");
                        aWhe.Add("mot_ymo");
                        aWhe.Add("mot_nmo");
                        s = _clsFun.SqlUpdRow(TABGESMOT, t, t.Rows[0], x, aWhe, aExl);
                    }
                    if (s != "")
                    {
                        _clsFun.SqlWrite(s, _strConSql);
                        //_clsFun.FileLog(TABGESMOT, sYea + "-" + (string)x["mot_nmo"], s);
                    }
                }

                /* RIGHE */

                if (sTip == "FAT")
                {
                    s = "SELECT * FROM GesMovimenti WHERE ";
                    s += "mov_ubi = '" + sUbi + "' AND ";
                    s += "mov_yfa = '" + sYea + "' AND ";
                    s += "mov_nfa = '" + sNfm + "' ";
                    s += "ORDER BY mov_rfa";
                }
                else
                {
                    s = "SELECT * FROM GesMovimenti WHERE ";
                    s += "mov_ubi = '" + sUbi + "' AND ";
                    s += "mov_ymo = '" + sYea + "' AND ";
                    s += "mov_nmo = '" + sNfm + "' ";
                    s += "ORDER BY mov_rfa";
                }
                DataTable tMov = _clsFun.FillTabSql(TABGESMOV, s, false, _strConSql);

                aExl = new ArrayList();

                //DataTable tTmp = ((DataView)dgv1.DataSource).ToTable();

                int iRig = 0;

                decimal dDelta = 0;
                decimal dScoTot = 0;
                decimal dSconto = 0;

                ////Calcolo differenze per arrotondamento
                decimal dDocTot = 0;
                //decimal dMovTot = 0;

                if (tVep != null && tVep.Rows.Count > 0)
                {
                    dDocTot = (decimal)tVep.Rows[0]["vet_imp"];

                    foreach (DataRow yy in tVep.Rows)
                    {
                        if (((string)yy["vep_tip"]).Substring(0, 2) == "SC")
                            dSconto += (decimal)yy["vep_imp"];
                        else if (((string)yy["vep_tip"]) == "PAG")
                            dScoTot += (decimal)yy["vep_imp"];
                        else if (((string)yy["vep_tip"]) == "RES")
                            dScoTot -= (decimal)yy["vep_imp"];
                    }

                    if (dSconto > 0)
                        dDelta = dSconto * 100 / (dScoTot + dSconto);
                }

                iRig++;

                string sMat = "";
                if ((string)tVep.Rows[0]["vet_cau"] != "MOV")
                {
                    s = "SELECT * FROM TabPos WHERE tab_ann=0";
                    t = _clsFun.FillTabSql("TabPos", s, false, _strConSql);
                    if (t.Rows.Count > 0 && !DBNull.Value.Equals(t.Rows[0]["tab_mat"]))
                    {
                        int iPoss = Convert.ToInt16(tVen.Rows[0]["vet_pos"]);

                        s = ((string)t.Rows[0]["tab_mat"]).Trim();

                        string[] a = s.Split(',');
                        if (a.Length >= iPoss && a[iPoss - 1] != "")
                            sMat = a[iPoss - 1];
                    }
                }

                x = tMov.NewRow();
                if (sTip == "FAT")
                {
                    x["mov_ubi"] = sUbi;
                    x["mov_yfa"] = sYea;
                    x["mov_nfa"] = sNfm;
                    x["mov_ymo"] = _clsDef.COD04X;
                    x["mov_nmo"] = _clsDef.COD06X;
                    x["mov_rfa"] = iRig.ToString("0000");
                    x["mov_rmo"] = _clsDef.COD04X;
                }
                else
                {
                    x["mov_ubi"] = sUbi;
                    x["mov_yfa"] = _clsDef.COD04X;
                    x["mov_nfa"] = _clsDef.COD06X;
                    x["mov_ymo"] = sYea;
                    x["mov_nmo"] = sNfm;
                    x["mov_rfa"] = _clsDef.COD04X;
                    x["mov_rmo"] = iRig.ToString("0000");
                }

                //s = "Scontrino N. " + sNum + " del " + dDay.ToShortDateString();
                s = "Documento gestionale N. " + sNum + " del " + dDay.ToShortDateString();
                if (sMat != "" && sTip == "FAT")
                    s += " Matr. " + sMat;
                else
                    s += " POS " + (string)tVen.Rows[0]["vet_pos"];

                x["mov_art"] = "";
                x["mov_ard"] = s;


                tMov.Rows.Add(x);

                s = _clsFun.SqlInsertRow(TABGESMOV, tMov, x);
                if (s != "")
                {
                    _clsFun.SqlWrite(s, _strConSql);
                    _clsFun.FileLog("GesMovimenti", (string)x["mov_art"], s);
                }

                foreach (DataRow xx in tVen.Rows)
                {
                    iRig++;

                    x = tMov.NewRow();
                    if (sTip == "FAT")
                    {
                        x["mov_ubi"] = sUbi;
                        x["mov_yfa"] = sYea;
                        x["mov_nfa"] = sNfm;
                        x["mov_ymo"] = _clsDef.COD04X;
                        x["mov_nmo"] = _clsDef.COD06X;
                        x["mov_rfa"] = iRig.ToString("0000");
                        x["mov_rmo"] = _clsDef.COD04X;
                    }
                    else
                    {
                        x["mov_ubi"] = sUbi;
                        x["mov_yfa"] = _clsDef.COD04X;
                        x["mov_nfa"] = _clsDef.COD06X;
                        x["mov_ymo"] = sYea;
                        x["mov_nmo"] = sNfm;
                        x["mov_rfa"] = _clsDef.COD04X;
                        x["mov_rmo"] = iRig.ToString("0000");
                    }

                    if (((string)xx["ven_art"]).Length == 7)
                        x["mov_art"] = xx["ven_art"];
                    else
                        x["mov_art"] = "";

                    x["mov_ard"] = xx["ven_ard"];
                    x["mov_iva"] = xx["ven_iva"];
                    x["mov_umi"] = xx["ven_umi"];
                    x["mov_qta"] = xx["ven_qta"];
                    x["mov_qkg"] = xx["ven_qkg"];
                    x["mov_cos"] = 0;
                    x["mov_ann"] = false;
                    x["mov_day"] = DateTime.Today;

                    if (((string)xx["ven_art"]).Length != 7)
                        x["mov_no1"] = xx["ven_art"];

                    x["mov_prv"] = xx["ven_prz"];
                    x["mov_imp"] = xx["ven_ven"];

                    //dMovTot += (decimal)x["mov_imp"];

                    /* SCONTO */

                    if ((decimal)x["mov_imp"] < 0)
                        Console.WriteLine("aaaa");

                    if ((string)x["mov_art"] == "0156195")
                        Console.WriteLine("aaaa");

                    if (dDelta > 0)
                    {
                        d = (decimal)x["mov_imp"];
                        d = _clsFun.MenoPer(d, dDelta);
                        d = Math.Round(d, 3, MidpointRounding.AwayFromZero);
                        s = "V" + ((decimal)x["mov_imp"] - d).ToString();
                        if (s.Length > 15)
                            s = s.Substring(0, 15);
                        x["mov_sco"] = s;
                        x["mov_imp"] = d;
                    }

                    /* SCORPORO IVA */

                    decimal dIva = 22;
                    j = tIva.Select("tab_cod='" + (string)x["mov_iva"] + "'");
                    if (j.Length > 0)
                        dIva = Convert.ToDecimal(j[0]["tab_ali"]);

                    d = _clsFun.MenoIva((decimal)x["mov_imp"], dIva);

                    if ((string)x["mov_umi"] == "KG")
                    {
                        if (d != 0 && (decimal)x["mov_qkg"] != 0)
                            x["mov_prv"] = d / (decimal)x["mov_qkg"];
                    }
                    else
                    {
                        if (d != 0 && (decimal)x["mov_qta"] != 0)
                            x["mov_prv"] = d / (decimal)x["mov_qta"];
                    }

                    x["mov_prv"] = Math.Round((decimal)x["mov_prv"], 3, MidpointRounding.ToEven);
                    x["mov_imp"] = Math.Round(d, 3);

                    if (strTip == "COS")
                    {
                        DataTable tArt = _clsQry.ArtSeek((string)xx["ven_art"], "SEEK");
                        if(tArt.Rows.Count > 0 && !DBNull.Value.Equals(tArt.Rows[0]["tmp_cos"]) && (decimal)tArt.Rows[0]["tmp_cos"] > 0)
                        {
                            d = (decimal)tArt.Rows[0]["tmp_cos"];
                            x["mov_prv"] = d;

                            if ((string)x["mov_umi"] == "KG")
                            {
                                if (d != 0 && (decimal)x["mov_qkg"] != 0)
                                    x["mov_imp"] = d * (decimal)x["mov_qkg"];
                            }
                            else
                            {
                                if (d != 0 && (decimal)x["mov_qta"] != 0)
                                    x["mov_imp"] = d * (decimal)x["mov_qta"];
                            }
                        }
                    }

                    s = (string)xx["ven_cau"] + ",";
                    s += ((DateTime)xx["ven_day"]).ToString("yyyyMMdd") + ",";
                    s += (string)xx["ven_ora"] + ",";
                    s += (string)xx["ven_pos"] + ",";
                    s += (string)xx["ven_sco"] + ",";
                    s += dDocTot.ToString().Replace(",", ".");

                    x["mov_ori"] = s;

                    aWhe = new ArrayList();
                    if (sTip == "FAT")
                    {
                        aWhe.Add("mov_ubi");
                        aWhe.Add("mov_yfa");
                        aWhe.Add("mov_nfa");
                        aWhe.Add("mov_rfa");
                    }
                    else
                    {
                        aWhe.Add("mov_ubi");
                        aWhe.Add("mov_ymo");
                        aWhe.Add("mov_nmo");
                        aWhe.Add("mov_rmo");
                    }

                    if (sTip == "FAT")
                        s = "mov_rfa='" + x["mov_rfa"] + "'";
                    else
                        s = "mov_rmo='" + x["mov_rmo"] + "'";

                    j = tMov.Select(s);
                    if (j.Length > 0)
                        s = _clsFun.SqlUpdRow(TABGESMOV, tMov, j[0], x, aWhe, aExl);
                    else
                        s = _clsFun.SqlInsertRow(TABGESMOV, tMov, x);

                    if (s != "")
                    {
                        _clsFun.SqlWrite(s, _strConSql);
                        _clsFun.FileLog("GesMovimenti", (string)x["mov_art"], s);
                    }
                }

                //if(dDocTot != dMovTot)
                //{
                //}

            }

        }

        private void dgv1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            string s = "";
            if (e.ColumnIndex == 7)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    int iIdx = (int)x["vet_idx"];
                    string sFnu = (string)x["vet_fnu"];

                    s = "UPDATE GesNegVet SET vet_fnu='" + sFnu + "' WHERE vet_idx=" + Convert.ToInt32(iIdx);
                    _clsFun.SqlWrite(s, _strConSqlSta);

                }
            }
        }

    }
}
