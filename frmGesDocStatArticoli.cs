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
    public partial class frmGesDocStatArticoli : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public DateTime _dayDti = DateTime.Today;
        public DateTime _dayDtf = DateTime.Today;

        private string _strConSql = "";
        private string _strConSqlSta = "";
        public string _strStatVisProf = "";

        public DataTable _tabCho = new DataTable("TabArt");

        public frmGesDocStatArticoli()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
            //_strStatVisProf = _clsFun.ParGet(clsDefine.enuParametri.ParStatVisProf, _strConSql);
        }

        private void frmGesStaArticoli_Load(object sender, EventArgs e)
        {
            this.Show();

            dtpDti.Value = _dayDti;
            dtpDtf.Value = _dayDtf;
            SetDgv1();
            FillTab();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmGesStaArticoli_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            dgv1.CurrentCell = dgv1[1, 0];

            DataView v = new DataView((DataTable)dgv1.DataSource, "tmp_cho", "", DataViewRowState.CurrentRows);

            _tabCho = v.ToTable();

            this.Close();
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

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.DataPropertyName = "tmp_cho";
            cCbc.Name = "Scelto";
            cCbc.Width = 40;
            cCbc.ValueType = typeof(string);
            cCbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_rep";
            cTbc.Name = "Reparto";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_red";
            cTbc.Name = "Descrizione";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_ec3";
            cTbc.Name = "Merceologia";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_ard";
            cTbc.Name = "Descrizione articolo";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_qta";
            cTbc.Name = "Q.tà";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_qkg";
            cTbc.Name = "Q.peso";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_ven";
            cTbc.Name = "Vendite";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_mav";
            cTbc.Name = "Margine Vendita";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_sco";
            cTbc.Name = "Scontrino";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_ora";
            cTbc.Name = "Ora";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_pos";
            cTbc.Name = "Cassa";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_day";
            cTbc.Name = "Data";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_iva";
            cTbc.Name = "IVA";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillTab()
        {
            string p = "TabNegozi";
            string s = "SELECT * FROM TabNegozi ORDER BY tab_cod";
            DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
            DataRow x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Tutti";
            t.Rows.InsertAt(x, 0);
            cmbNeg.DataSource = t;
            cmbNeg.DisplayMember = "tab_des";
            cmbNeg.ValueMember = "tab_cod";
            cmbNeg.SelectedValue = "";

            p = "TabReparti";
            s = "SELECT * FROM TabReparti ORDER BY tab_des";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Tutti";
            t.Rows.InsertAt(x, 0);
            cmbRep.DataSource = t;
            cmbRep.DisplayMember = "tab_des";
            cmbRep.ValueMember = "tab_cod";
            cmbRep.SelectedValue = "";
        }

        private void FillDati()
        {
            string s = "";
            DataRow[] j;
            string sScoKey = "";
            decimal d = 0;
            decimal dTotVen = 0;

            string sNeg = cmbNeg.SelectedValue.ToString();
            string sRep = cmbRep.SelectedValue.ToString();

            s = "SELECT * FROM TabReparti";
            DataTable tRep = _clsFun.FillTabSql("TabReparti", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tRep.Columns["tab_cod"];
            tRep.PrimaryKey = keys;

            s = "SELECT AnaArticoli.art_cod, AnaArticoli.art_des, AnaArticoli.art_iva, AnaArticoli.art_sfr, AnaArticoli.art_rep, TabReparti.tab_des AS RepDes, TabEcrLv3.tab_des AS EcrDes ";
            s += "FROM AnaArticoli ";
            s += "LEFT OUTER JOIN TabReparti ON AnaArticoli.art_rep = TabReparti.tab_cod ";
            s += "LEFT OUTER JOIN TabEcrLv3 ON AnaArticoli.art_ec1 = TabEcrLv3.tab_lv1 AND AnaArticoli.art_ec2 = TabEcrLv3.tab_lv2 AND AnaArticoli.art_ec3 = TabEcrLv3.tab_cod ";
            DataTable tArt = _clsFun.FillTabSql("Art", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            s = "SELECT ";
            s += "ven_neg, ";
            s += "ven_cau, ";
            s += "ven_day, ";
            s += "ven_ora, ";
            s += "ven_pos, ";
            s += "ven_sco, ";
            s += "ven_art, ";
            s += "ven_ard, ";
            s += "ven_iva, ";
            s += "ven_umi, ";
            s += "ven_rep, ";
            s += "ven_sct, ";
            s += "ven_ean, ";
            s += "ven_qta, ven_qkg, ven_prz, ven_ven, ";
            s += "CASE WHEN ven_qkg > 0 THEN ven_qkg ELSE ven_qta END * ven_cos AS ven_cos, ";
            s += "vet_imp, ";
            s += "vet_sct ";
            s += "FROM GesNegVen JOIN GesNegVet ON ";
            s += "GesNegVet.vet_neg = GesNegVen.ven_neg AND ";
            s += "GesNegVet.vet_cau = GesNegVen.ven_cau AND ";
            s += "GesNegVet.vet_day = GesNegVen.ven_day AND ";
            s += "GesNegVet.vet_ora = GesNegVen.ven_ora AND ";
            s += "GesNegVet.vet_pos = GesNegVen.ven_pos AND ";
            s += "GesNegVet.vet_sco = GesNegVen.ven_sco ";
            s += "WHERE (ven_day >= " + _clsFun.DaySql(dtpDti.Value) + ") AND (ven_day <= " + _clsFun.DaySql(dtpDtf.Value) + ") AND ";
            s += "(GesNegVen.ven_cau <> 'MOV') "; 
            if (sNeg != "")
                s += " AND ven_neg='" + sNeg + "' ";
            if (_strStatVisProf == "N")
                s += " AND GesNegVet.vet_cau<>'PRO' ";


            //s += " AND GesNegVet.vet_cau='SCO' AND vet_sco='00001' AND vet_pos='01' ";
            //s += "ORDER BY ven_sco";
            s += "ORDER BY ven_day,ven_cau,ven_neg,ven_pos,ven_ora,ven_sco";

            DataTable t = _clsFun.FillTabSql("Sta", s, false, _strConSqlSta);
            keys = new DataColumn[9];
            keys[0] = t.Columns["ven_day"];
            keys[1] = t.Columns["ven_cau"];
            keys[2] = t.Columns["ven_neg"];
            keys[3] = t.Columns["ven_pos"];
            keys[4] = t.Columns["ven_ora"];
            keys[5] = t.Columns["ven_sco"];
            keys[6] = t.Columns["ven_sct"];
            keys[7] = t.Columns["ven_ean"];
            keys[8] = t.Columns["ven_art"];
            t.PrimaryKey = keys;

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "tmp_cho",
                Caption = "Scelto",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ard",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_rep",
                Caption = "Reparto",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_red",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ec3",
                Caption = "Ecr",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });

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
                ColumnName = "tmp_mav",
                Caption = "Margine",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_iva",
                Caption = "IVA",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            /* Da togliere inizio 

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_arf",
                Caption = "Arf",
                MaxLength = 20,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_cov",
                Caption = "Costo Vega",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_mve",
                Caption = "Margine Vega",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_vve",
                Caption = "Vendita Vega",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });





            //t.Columns.Add(new DataColumn()
            //{
            //    DataType = Type.GetType("System.Decimal"),
            //    ColumnName = "tmp_coq",
            //    Caption = "costo * qta",
            //    ReadOnly = false,
            //    DefaultValue = (Decimal)0
            //});

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "ven_vni",
                Caption = "vendita no iva",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            //t.Columns.Add(new DataColumn()
            //{
            //    DataType = Type.GetType("System.Decimal"),
            //    ColumnName = "ven_viq",
            //    Caption = "Vendita no IVA per qta",
            //    ReadOnly = false,
            //    DefaultValue = (Decimal)0
            //});

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_cvq",
                Caption = "Costo vega per qta",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_vvi",
                Caption = "Ven netto IVA * qta",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });


            Da togliere inizio fine */

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if ((string)y["ven_cau"] == "SCO" && (string)y["ven_pos"] == "01" && (string)y["ven_sco"] == "00033")
                    Console.WriteLine("aaaa");

                s = ((DateTime)y["ven_day"]).ToString("yyyyMMdd") + (string)y["ven_cau"] + (string)y["ven_neg"] + (string)y["ven_pos"] + (string)y["ven_sco"];
                if (s != sScoKey)
                {
                    sScoKey = s;

                    decimal dScn = (decimal)y["vet_sct"];
                    decimal dLor = dScn + (decimal)y["vet_imp"];
                    decimal dDelta = 0;
                    if(dScn > 0 && dLor > 0)
                        dDelta = dScn * 100 / dLor;
                    decimal dd = 0;

                    s = "ven_neg = '" + (string)y["ven_neg"] + "' AND ";
                    s += "ven_cau = '" + (string)y["ven_cau"] + "' AND ";
                    s += "ven_day = " + _clsFun.DayMdb((DateTime)y["ven_day"]) + " AND ";
                    s += "ven_ora = '" + (string)y["ven_ora"] + "' AND ";
                    s += "ven_pos = '" + (string)y["ven_pos"] + "' AND ";
                    s += "ven_sco = '" + (string)y["ven_sco"] + "'";
                    j = t.Select(s);

                    if ((decimal)y["vet_sct"] > 0)
                    {
                        dScn = (decimal)y["vet_sct"];
                        dLor = dScn + (decimal)y["vet_imp"];
                        dDelta = dScn * 100 / dLor;
                        dd = 0;
                    }
                    /* 20171026 Seck serve nel caso ci sia il totale scontrino diverso dal dettaglio*/
                    else
                    {
                        d = 0;
                        for(int i = 0; i < j.Length; i++)
                        {
                            if ((string)j[i]["ven_sct"] == "RES")
                                d -= (decimal)j[i]["ven_ven"];
                            else
                                d += (decimal)j[i]["ven_ven"];
                        }
            
                        if(d != (decimal)y["vet_imp"])
                        {
                            dScn = d - (decimal)y["vet_imp"];
                            dLor = dScn + (decimal)y["vet_imp"];
                            dDelta = dScn * 100 / dLor;
                            dd = 0;
                        }
                    }
                    /**/

                    if (dDelta > 0)
                    {
                        if (j.Length > 0)
                        {
                            for (int i = 0; i < j.Length; i++)
                            {
                                if ((string)j[i]["ven_art"] == "0002282")
                                    Console.WriteLine("aaaa");

                                d = _clsFun.MenoPer((decimal)j[i]["ven_ven"], dDelta);

                                j[i]["ven_ven"] = d;

                                dd += d;

                                if (i == j.Length - 1)
                                {
                                    if (dd != (decimal)y["vet_imp"])
                                    {
                                        d = (decimal)y["vet_imp"] - dd;
                                        j[i]["ven_ven"] = (decimal)j[i]["ven_ven"] + d;
                                    }
                                }
                            }
                        }
                    }

                }

                j = tArt.Select("art_cod='" + y["ven_art"] +"'");
                if(j.Length > 0)
                {
                    y["tmp_rep"] = (string)j[0]["art_rep"];
                    y["tmp_ard"] = (string)j[0]["art_des"];
                    y["tmp_red"] = (string)j[0]["RepDes"];
                    y["tmp_ec3"] = (string)j[0]["EcrDes"];

                    decimal dIva = 0;
                    if (_clsFun.Numerico(j[0]["art_iva"],"0123456789"))
                        dIva = Convert.ToDecimal(j[0]["art_iva"]);

                    y["tmp_mav"] = _clsFun.Margine((decimal)y["ven_ven"], (decimal)y["ven_cos"], dIva,  (decimal)j[0]["art_sfr"], "P");
                    y["tmp_iva"] = dIva;
                }
                else
                {
                    y["tmp_rep"] = (string)y["ven_rep"];
                    y["tmp_ard"] = (string)y["ven_ard"];
                    y["tmp_red"] = "Non definito";
                    j = tRep.Select("tab_cod='" + (string)y["ven_rep"] + "'");
                    if(j.Length > 0)
                        y["tmp_red"] = (string)j[0]["tab_des"];
                    y["tmp_ec3"] = "";
                }

                if ((string)y["ven_sct"] == "RES")
                    dTotVen -= (decimal)y["ven_ven"];
                else
                    dTotVen += (decimal)y["ven_ven"];
            }

            if (sRep != "")
            {
                dTotVen = 0;

                DataTable t2 = t.Clone();

                foreach(DataRow y in t.Rows)
                {
                    if ((string)y["tmp_rep"] == sRep)
                    {
                        t2.ImportRow(y);
                        dTotVen += (decimal)y["ven_ven"];
                    }
                }

                t = t2.Copy();
            }

            /* Da togliere inizio fine
            t = FillArfAlba(t);
            */

            dgv1.DataSource = t;

            lblTotVen.Text = dTotVen.ToString("#,##0.00");
        }

        private DataTable FillArfAlba(DataTable t)
        {
            string s = "";
            decimal d = 0;
            DataRow[] j;

            s = "SELECT DISTINCT lia_art, lia_arf, lia_for, lia_cos, lia_prv ";
            s += "FROM GesLisAcquisto ";
            s += "WHERE (lia_for = '00001' OR lia_for = '00002')";
            DataTable tLia = _clsFun.FillTabSql("Art", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[4];
            keys[0] = tLia.Columns["lia_art"];
            keys[1] = tLia.Columns["lia_for"];
            keys[2] = tLia.Columns["lia_arf"];
            keys[3] = tLia.Columns["lia_cos"];
            tLia.PrimaryKey = keys;

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if ((string)y["ven_art"] == "0005056")
                    Console.WriteLine("zzzzzzzzz");

                decimal dIva = (decimal)y["tmp_iva"];

                //Costo * qta
                //y["tmp_coq"] = (decimal)y["ven_cos"] * (decimal)y["ven_qta"];

                //Vendita al netto IVA
                y["ven_vni"] = _clsFun.MenoIva((decimal)y["ven_ven"], dIva);

                ////Netto IVA * qta
                //y["ven_viq"] = (decimal)y["ven_vni"] * (decimal)y["ven_qta"];

                j = tLia.Select("lia_art='" + y["ven_art"] + "' AND lia_for='00001'");
                if (j.Length > 0)
                {
                    y["tmp_arf"] = j[0]["lia_arf"];
                }
                j = tLia.Select("lia_art='" + y["ven_art"] + "' AND lia_for='00002'");
                if (j.Length > 0)
                {
                    y["tmp_cov"] = j[0]["lia_cos"];
                    y["tmp_vve"] = j[0]["lia_prv"];

                    if (dIva > 0 && (decimal)y["tmp_cov"] > 0 && (decimal)y["tmp_vve"] > 0)
                    {
                        y["tmp_mve"] = _clsFun.Margine((decimal)y["tmp_vve"], (decimal)y["tmp_cov"], dIva, 0, "P");

                        y["tmp_cvq"] = (decimal)y["tmp_cov"] * (decimal)y["ven_qta"];


                        //Netto IVA * qta
                        y["tmp_vvi"] = _clsFun.MenoIva((decimal)y["tmp_vve"], dIva) * (decimal)y["ven_qta"];
                    }
                }
            }

            return t;
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 4)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = (string)x["ven_art"];
                    f.ShowDialog();
                }
            }
            else if(e.ColumnIndex == 10)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    string sNeg = (string)r.Row["ven_neg"];
                    string sCau = (string)r.Row["ven_cau"];
                    DateTime dDay = (DateTime)r.Row["ven_day"];
                    string sOra = (string)r.Row["ven_ora"];
                    string sPos = (string)r.Row["ven_pos"];
                    string sSco = (string)r.Row["ven_sco"];

                    frmGesStatScoDettaglio f = new frmGesStatScoDettaglio();
                    f._strNeg = sNeg;
                    f._strCau = sCau;
                    f._dayDay = dDay;
                    f._strOra = sOra;
                    f._strPos = sPos;
                    f._strSco = sSco;
                    f.ShowDialog();

                    if (f._tabVet != null && ((DataTable)f._tabVet).Rows.Count > 0)
                    {
                        r.Row["vet_pun"] = f._tabVet.Rows[0]["vet_pun"];
                        r.Row["vet_imp"] = f._tabVet.Rows[0]["vet_imp"];
                    }
                }
            }

        }

        private void txtSeek_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                string s = txtSeek.Text.Trim();

                if (s != "" && dgv1.RowCount > 0)
                {
                    Boolean b = false;
                    int iCur = dgv1.CurrentCell.RowIndex;
                    int i = 0;

                    foreach (DataGridViewRow row in dgv1.Rows)
                    {
                        i++;

                        if (i > iCur + 1)
                        {
                            if (row.Cells["Descrizione articolo"].Value.ToString().ToLower().Contains(s.ToLower()))
                            {
                                dgv1.Rows[row.Index].Selected = true;
                                dgv1.FirstDisplayedScrollingRowIndex = row.Index;
                                dgv1.CurrentCell = dgv1.Rows[row.Index].Cells[0];
                                b = true;
                                break;
                            }
                        }
                    }
                    if (!b)
                        dgv1.CurrentCell = dgv1.Rows[0].Cells[0];
                }
            }

        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "tmp_red, tmp_ard", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            string s = "";
            decimal d = 0;

            foreach (DataRowView r in v)
            {
                t.ImportRow(r.Row);
            }

            string sTit = DateTime.Today.ToString("dd/MM/yyyy") + " - Estrazione venduto dal " + dtpDti.Value.ToString("dd/MM/yyyy") + " al " + dtpDtf.Value.ToString("dd/MM/yyyy");
            string sFil = "";

            s = "";
            s += "tmp_red, Reparto, 100, StringLiteral;";
            s += "ven_art, Articolo, 50, StringLiteral;";
            s += "tmp_arf, Codice, 50, StringLiteral;";
            s += "tmp_ard, Descrizione, 50, StringLiteral;";
            s += "ven_qta, Q.ta, 50, StringLiteral;";
            s += "ven_qkg, Peso, 50, StringLiteral;";
            s += "ven_cos, Costo, 50, StringLiteral;";
            s += "ven_ven, Importo, 50, StringLiteral;";
            s += "tmp_mav, Margine, 50, StringLiteral;";
            s += "ven_day, Data, 50, StringLiteral;";

            string sFld = s;

            string sFoo = ";;;;Totale;" + lblTotVen.Text.Replace(",", ".");

            (new clsExcel()).exportToCsv1(sFld, t, sFil, sTit, sFoo);
        }

        private void ZZexcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /* da togliere estrazione confronto VEGA */

            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "tmp_red, tmp_ard", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            string s = "";
            decimal d = 0;

            foreach (DataRowView r in v)
            {
                //if ((string)r["ven_art"] == "0005056")
                //{
                // }
                   t.ImportRow(r.Row);
            }

            string sTit = DateTime.Today.ToString("dd/MM/yyyy") + " - Estrazione venduto dal " + dtpDti.Value.ToString("dd/MM/yyyy") + " al " + dtpDtf.Value.ToString("dd/MM/yyyy");
            string sFil = "";

            s = "";
            s += "ven_day, Data, 50, StringLiteral;";
            s += "tmp_red, Reparto, 100, StringLiteral;";
            s += "ven_art, Articolo, 50, StringLiteral;";
            s += "tmp_arf, Codice CF, 50, StringLiteral;";
            s += "tmp_ard, Descrizione, 50, StringLiteral;";

            s += "ven_iva, IVA, 50, StringLiteral;";

            s += "ven_qta, Q.ta, 50, StringLiteral;";
            //s += "ven_qkg, Peso, 50, StringLiteral;";
            s += "ven_cos, Costo CF, 50, StringLiteral;";

            //s += "ven_prz, Vendita CF, 50, StringLiteral;";
            s += "ven_vni, Vendita CF netto IVA, 50, StringLiteral;";
            s += "tmp_mav, Margine CF, 50, StringLiteral;";
            //s += "ven_viq, Vendita CF nIVAxQqta, 50, StringLiteral;";

            //VEGA

            //s += "tmp_cov, Costo B, 50, StringLiteral;";
            s += "tmp_cvq, Costo B, 50, StringLiteral;";
            s += "tmp_vvi, Vendita B netto IVA, 50, StringLiteral;";
            //s += "tmp_vve, Vendita B, 50, StringLiteral;";
            s += "tmp_mve, Margine B, StringLiteral;";

            string sFld = s;

            string sFoo = ";;;;Totale;" + lblTotVen.Text.Replace(",", ".");

            (new clsExcel()).exportToCsv1(sFld, t, sFil, sTit, sFoo);
        }

        private void cvsExcelPerMerceologiaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "tmp_ec3, tmp_ard", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            string s = "";
            decimal d = 0;

            foreach (DataRowView r in v)
            {
                t.ImportRow(r.Row);
            }

            string sTit = DateTime.Today.ToString("dd/MM/yyyy") + " - Estrazione venduto dal " + dtpDti.Value.ToString("dd/MM/yyyy") + " al " + dtpDtf.Value.ToString("dd/MM/yyyy");
            string sFil = "";

            s = "";
            s += "tmp_ec3, Merceologia, 100, StringLiteral;";
            s += "ven_art, Articolo, 50, StringLiteral;";
            s += "tmp_ard, Descrizione, 50, StringLiteral;";
            s += "ven_qta, Q.ta, 50, StringLiteral;";
            s += "ven_qkg, Peso, 50, StringLiteral;";
            s += "ven_ven, Importo, 50, StringLiteral;";

            string sFld = s;

            (new clsExcel()).exportToCsv1(sFld, t, sFil, sTit, "");

        }

        //private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        //{
        //    Console.WriteLine("zzz");
        //}

        private void dgv1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgv1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        //private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        //{
        //    if (dgv1.CurrentCell is DataGridViewCheckBoxCell)
        //    {
        //        dgv1.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
        //    }
        //}

    }
}
