using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmGesStatReparto : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public DateTime _dayDti = DateTime.Today;
        public DateTime _dayDtf = DateTime.Today;

        public string _strStatVisProf = "";

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public frmGesStatReparto()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
        }

        private void frmGesStaReparto_Load(object sender, EventArgs e)
        {
            this.Show();

            dtpDti.Value = _dayDti;
            dtpDtf.Value = _dayDtf;
            SetDgv1();
            FillTab();
            FillDati();
            progressBar1.Visible = false;
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmGesStaReparto_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
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
            cTbc.DataPropertyName = "tmp_cod";
            cTbc.Name = "Reparto";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 130;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_ils";
            cTbc.Name = "Vendite lordo sconti";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_imp";
            cTbc.Name = "Vendite";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "tmp_inc";
            //cTbc.Name = "Vendite";
            //cTbc.Width = 100;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //cTbc.DefaultCellStyle.Format = "##0";
            //cTbc.ReadOnly = false;
            //dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_inc";
            cTbc.Name = "Inc%";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_sco";
            cTbc.Name = "Scontrini";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_sci";
            cTbc.Name = "Inc%";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_scm";
            cTbc.Name = "Media venduto / scontrini";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_qta";
            cTbc.Name = "Q.tà";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_qti";
            cTbc.Name = "Inc%";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_qtm";
            cTbc.Name = "Media venduto / q.tà";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_niv";
            cTbc.Name = "Netto IVA";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_vcc";
            cTbc.Name = "Venduto con costo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_cdv";
            cTbc.Name = "Costo del venduto";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_mav";
            cTbc.Name = "Margine vendita";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0.00";
            cTbc.ReadOnly = false;
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
        }

        private void FillDati()
        {
            string s = "";
            DataRow x;
            DataRow[] j;
            DataRow[] j2;
            decimal d = 0;
            decimal dQta = 0;
            decimal dTot = 0;
            decimal dTls = 0;       //Venduto lordo sconti    
            decimal dNiv = 0;       //Venduto netto iVA
            decimal dTotNiv = 0;    //Somma venduto netto iVA
            decimal dTotCdv = 0;    //Somma del venduto

            string sNeg = cmbNeg.SelectedValue.ToString();
            string sRpd = _clsFun.ParGet(clsDefine.enuParametri.Par013CodRepxDefault, _strConSql);
            string sRep = "";
            string sRed = "";
            string sArd = "";
            string sMsg = "";

            ArrayList aScoRep = new ArrayList();

            decimal dSco = 0;
            decimal dImpMax = 0;
            if (_clsFun.Numerico(txtImpMax.Text, "0123456789"))
                dImpMax = Convert.ToDecimal(txtImpMax.Text);

            DataTable t = new clsGenTabTmp().TabTmpStaReparto("TabRep");
            DataColumn[] keys = new DataColumn[2];
            keys[0] = t.Columns["tab_cod"];
            keys[1] = t.Columns["tab_neg"];
            t.PrimaryKey = keys;

            s = "SELECT * FROM TabReparti";
            DataTable tRep = _clsFun.FillTabSql("TabRep", s, false, _strConSql);

            s = "SELECT AnaArticoli.art_cod, AnaArticoli.art_des, AnaArticoli.art_iva, AnaArticoli.art_rep, TabReparti.tab_des AS RepDes ";
            s += "FROM AnaArticoli ";
            s += "LEFT OUTER JOIN TabReparti ON AnaArticoli.art_rep = TabReparti.tab_cod";
            DataTable tArt = _clsFun.FillTabSql("Art", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            s = "SELECT ";
            s += "ven_day, ";
            s += "ven_neg, ";
            s += "ven_cau, ";
            s += "ven_pos, ";
            s += "ven_ora, ";
            s += "ven_sco, ";
            s += "ven_sct, ";
            s += "ven_art, ";
            s += "ven_ean, ";
            s += "ven_iva, ";
            s += "ven_rep, ";
            s += "ven_qta, ";
            s += "ven_qkg, ";
            s += "ven_cos, ";
            s += "ven_ven, ";
            s += "ven_ven AS VenVls, ";   //Importo lordo sconti
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
            if(dImpMax > 0)
                s += " AND GesNegVet.vet_imp >=" + dImpMax.ToString() + " " ;

            //s += "ORDER BY ven_sco";
            s += "ORDER BY ven_day,ven_cau,ven_neg,ven_pos,ven_sco";
            DataTable tVen = _clsFun.FillTabSql("VEN", s, false, _strConSqlSta);

            //tVen.Columns.Add(new DataColumn()
            //{
            //    DataType = Type.GetType("System.Decimal"),
            //    ColumnName = "VenVls",
            //    Caption = "Importo vendita lordo sconto ",
            //    ReadOnly = false,
            //    DefaultValue = (Decimal)0
            //});

            keys = new DataColumn[9];
            keys[0] = tVen.Columns["ven_day"];
            keys[1] = tVen.Columns["ven_cau"];
            keys[2] = tVen.Columns["ven_neg"];
            keys[3] = tVen.Columns["ven_pos"];
            keys[4] = tVen.Columns["ven_ora"];
            keys[5] = tVen.Columns["ven_sco"];
            keys[6] = tVen.Columns["ven_sct"];
            keys[7] = tVen.Columns["ven_ean"];
            keys[8] = tVen.Columns["ven_art"];
            tVen.PrimaryKey = keys;

            dTot = 0;
            dTls = 0;

            string sScoKey = "";

            progressBar1.Visible = true;
            progressBar1.Value = 0;
            progressBar1.Maximum = (int)tVen.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in tVen.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                s = ((DateTime)y["ven_day"]).ToString("yyyyMMdd") + (string)y["ven_cau"] + (string)y["ven_neg"] + (string)y["ven_pos"] + (string)y["ven_sco"];

                if (s != sScoKey)
                {
                    dSco += 1;
                    sScoKey = s; 
                    aScoRep.Clear();

                    //Ricalcolo importo in caso su sconto su totale

                    s = "ven_neg = '" + y["ven_neg"] + "' AND ";
                    s += "ven_cau = '" + y["ven_cau"] + "' AND ";
                    s += "ven_day = " + _clsFun.DayMdb((DateTime)y["ven_day"]) + " AND ";
                    s += "ven_ora = '" + y["ven_ora"] + "' AND ";
                    s += "ven_pos = '" + y["ven_pos"] + "' AND ";
                    s += "ven_sco = '" + y["ven_sco"] + "'";
                    j = tVen.Select(s);

                    decimal dScn = (decimal)y["vet_sct"];
                    decimal dLor = dScn + (decimal)y["vet_imp"];
                    decimal dDelta = 0;
                    if (dScn > 0 && dLor > 0)
                        dDelta = dScn * 100 / dLor;

                    decimal dd = 0;

                    if((decimal)y["vet_sct"] > 0)
                    {
                        dScn = (decimal)y["vet_sct"];
                        dLor = dScn + (decimal)y["vet_imp"];
                        dDelta = dScn * 100 / dLor;
                        dd = 0;
                    }
                    /* 20171026 Seck serve nel caso ci sia il totale scontrino diverso dal dettaglio */
                    else
                    {
                        if ((string)y["ven_cau"] == "SCO" && (string)y["ven_pos"] == "01" && (string)y["ven_sco"] == "00033")
                            Console.WriteLine("aaaa");

                        d = 0;
                        for (int i = 0; i < j.Length; i++)
                        {
                            if ((string)j[i]["ven_sct"] == "RES")
                                d -= (decimal)j[i]["ven_ven"];
                            else
                                d += (decimal)j[i]["ven_ven"];
                        }
                        if (d != (decimal)y["vet_imp"])
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
                                d = _clsFun.MenoPer((decimal)j[i]["ven_ven"], dDelta);

                                j[i]["ven_ven"] = d;

                                dd += d;

                                Console.WriteLine("aaaaaaaaaaa");

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

                if ((string)y["ven_art"] == "0003663")
                    Console.WriteLine("aaaaaaaaaa");

                sRep = sRpd;
                sArd = "";

                j2 = tArt.Select("art_cod='" + (string)y["ven_art"] + "'");
                if (j2.Length > 0)
                {
                    sRep = (string)j2[0]["art_rep"];
                    sArd = (string)j2[0]["art_des"];
                    y["ven_iva"] = (string)j2[0]["art_iva"];
                }
                else if (!DBNull.Value.Equals(y["ven_rep"]))
                    sRep = (string)y["ven_rep"];

                j = tRep.Select("tab_cod='" + sRep + "'");
                if (j.Length > 0)
                    sRed = (string)j[0]["tab_des"];

                if (sRep == "")
                {
                    sMsg += "Articolo " + y["ven_art"] + " - " + sArd + " con reparto non definito!" + _clsDef.CRLF;
                    sRed = "Non definito";
                }

                if(sNeg == "")
                    j = t.Select("tmp_cod='" + sRep + "'");
                else
                    j = t.Select("tmp_cod='" + sRep + "' AND tmp_neg='" + sNeg + "'");

                if (j.Length == 0)
                {
                    x = t.NewRow();
                    x["tmp_neg"] = sNeg;
                    x["tmp_cod"] = sRep;
                    x["tmp_des"] = sRed;
                    x["tmp_ils"] = 0;               //Importo al lordo degli sconti
                    x["tmp_imp"] = 0;
                    x["tmp_inc"] = 0;
                    x["tmp_niv"] = 0;
                    t.Rows.Add(x);

                    j = t.Select("tmp_cod='" + sRep + "'");
                }

                if (((string)y["ven_sct"]).Trim() == "RES")
                {
                    dQta -= (decimal)y["ven_qta"];
                    dTot -= (decimal)y["ven_ven"];
                    dTls -= (decimal)y["VenVls"];
                    j[0]["tmp_ils"] = (decimal)j[0]["tmp_ils"] + (decimal)y["VenVls"];
                    j[0]["tmp_imp"] = (decimal)j[0]["tmp_imp"] - (decimal)y["ven_ven"];
                }
                else
                {
                    dQta += (decimal)y["ven_qta"];
                    dTot += (decimal)y["ven_ven"];
                    dTls += (decimal)y["VenVls"];
                    j[0]["tmp_ils"] = (decimal)j[0]["tmp_ils"] + (decimal)y["VenVls"];
                    j[0]["tmp_imp"] = (decimal)j[0]["tmp_imp"] + (decimal)y["ven_ven"];
                    j[0]["tmp_qta"] = (decimal)j[0]["tmp_qta"] + (decimal)y["ven_qta"];

                    dNiv = 0;
                    if (_clsFun.Numerico(y["ven_iva"]))
                        dNiv = _clsFun.MenoIva((decimal)y["ven_ven"], Convert.ToDecimal(y["ven_iva"]));
                    else
                        dNiv = (decimal)y["ven_ven"];

                    j[0]["tmp_niv"] = (decimal)j[0]["tmp_niv"] + dNiv;

                    if (aScoRep.IndexOf((string)y["ven_sco"] + sRep) < 0)
                    {
                        aScoRep.Add((string)y["ven_sco"] + sRep);
                        j[0]["tmp_sco"] = (decimal)j[0]["tmp_sco"] + 1;
                    }

                    if((decimal)y["ven_cos"] > 0)
                    {
                        if ((decimal)y["ven_qkg"] > 0 )
                            d = (decimal)y["ven_cos"] * (decimal)y["ven_qkg"];
                        else
                            d = (decimal)y["ven_cos"] * (decimal)y["ven_qta"];

                        j[0]["tmp_cdv"] = (decimal)j[0]["tmp_cdv"] + d;
                        j[0]["tmp_vcc"] = (decimal)j[0]["tmp_vcc"] + dNiv;

                        _clsFun.FileLogDivFor("ST REP", (string)y["ven_art"], dNiv.ToString().Replace(",", "."));
                    }
                }
         
            }

            DataTable t2 = t.Clone();
            DataView v = new DataView(t, "", "tmp_cod", DataViewRowState.CurrentRows);

            foreach (DataRowView r in v)
            {
                r["tmp_inc"] = (decimal)r["tmp_imp"] / dTot * 100;
                r["tmp_sci"] = (decimal)r["tmp_sco"] / dSco * 100;
                r["tmp_scm"] = (decimal)r["tmp_imp"] / (decimal)r["tmp_sco"];
                r["tmp_qti"] = (decimal)r["tmp_qta"] / dQta * 100;
                if ((decimal)r["tmp_qta"] > 0)
                    r["tmp_qtm"] = (decimal)r["tmp_imp"] / (decimal)r["tmp_qta"];

                if ((decimal)r["tmp_vcc"] > 0 && (decimal)r["tmp_cdv"] > 0)
                    r["tmp_mav"] = Math.Round(((decimal)r["tmp_vcc"] - (decimal)r["tmp_cdv"]) / (decimal)r["tmp_vcc"], 2)*100;

                t2.ImportRow(r.Row);

                dTotCdv = dTotCdv + (decimal)r["tmp_cdv"];       //Costo del venduto netto iVA
                dTotNiv = dTotNiv + (decimal)r["tmp_vcc"];       //Venduto netto iVA
            }

            v = new DataView(t, "", "tmp_cod", DataViewRowState.CurrentRows);

            dgv1.DataSource = v;

            lblSco.Text = dSco.ToString("#,###,##0");
            lblVen.Text = dTot.ToString("#,###,##0.00");
            lblVls.Text = dTls.ToString("#,###,##0.00");
            lblQta.Text = dQta.ToString("#,###,##0");

            lblCdv.Text = dTotCdv.ToString("#,###,##0.00");
            lblNiv.Text = dTotNiv.ToString("#,###,##0.00");
            lblMar.Text = "0";
            if (dTotCdv != 0 && dTotNiv != 0)
                lblMar.Text = ((dTotNiv - dTotCdv) / dTotNiv * 100).ToString("#,###,##0.00");

            FatImporto();

            progressBar1.Visible = false;
        }

        private void FatImporto()
        {
            string sYea = dtpDti.Value.Year.ToString();

            string s = "";
            s = "SELECT ";
            s += "SUM(R.mov_imp + R.MovIva) AS MovImp ";
            s += "FROM GesFatTestate ";
            s += "LEFT OUTER JOIN AnaClienti ON GesFatTestate.fat_cfo = AnaClienti.cli_cod ";
            s += "LEFT OUTER JOIN TabStato ON GesFatTestate.fat_sta = TabStato.tab_cod ";
            s += "LEFT OUTER JOIN TabPagamenti ON GesFatTestate.fat_tpg = TabPagamenti.tab_cod ";
            s += "LEFT OUTER JOIN (";
            s += "SELECT ";
            s += "GesMovimenti.mov_nfa, ";
            s += "SUM(GesMovimenti.mov_imp) AS mov_imp, ";
            s += "SUM(CASE WHEN tab_ali IS NOT NULL AND tab_ali > 0 AND mov_imp > 0 THEN mov_imp * tab_ali / 100 ELSE 0 END) AS MovIva ";
            s += "FROM GesMovimenti ";
            s += "LEFT OUTER JOIN TabIva ON GesMovimenti.mov_iva = TabIva.tab_cod ";
            s += "WHERE (GesMovimenti.mov_ann = 0) AND (GesMovimenti.mov_yfa = '" + sYea + "') ";
            s += "GROUP BY GesMovimenti.mov_yfa, GesMovimenti.mov_nfa) AS R ON GesFatTestate.fat_nfa = R.mov_nfa ";
            s += "WHERE ";
            s += "(GesFatTestate.fat_tpd = 'FV') AND ";
            s += "(GesFatTestate.fat_ddo >= " + _clsFun.DaySql(dtpDti.Value) + ") AND ";
            s += "(GesFatTestate.fat_ddo <= " + _clsFun.DaySql(dtpDtf.Value) + ") AND ";
            s += "(GesFatTestate.fat_ann = 0)";

            DataTable t = _clsFun.FillTabSql("GesFatTestate", s, false, _strConSql);

            lblFat.Text = "0";
            if(t.Rows.Count > 0)
                lblFat.Text = ((decimal)t.Rows[0]["MovImp"]).ToString("#,###,##0.00");
        }

        private void stampaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s = "";

            DataTable t = ((DataView)dgv1.DataSource).ToTable();

            decimal dQta = Convert.ToDecimal(lblQta.Text);
            decimal dSco = Convert.ToDecimal(lblSco.Text);
            decimal dMar = Convert.ToDecimal(lblMar.Text);
            
            string sNeg = "";
            s = cmbNeg.SelectedValue.ToString();
            if (s != "")
                sNeg = cmbNeg.Text;

            clsGenPdfSta1 cls = new clsGenPdfSta1();
            cls._dayIni = dtpDti.Value;
            cls._dayFin = dtpDtf.Value;
            cls.PrnPdfRep("ven_rep", t, dSco, dQta, sNeg, dMar);
        }

     }
}
