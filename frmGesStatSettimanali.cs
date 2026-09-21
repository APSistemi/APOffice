using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.IO;
using System.Diagnostics;

namespace APOffice
{
    public partial class frmGesStatSettimanali : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public string _strStatVisProf = "";

        public frmGesStatSettimanali()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
        }

        private void frmGesStatSettimanali_Load(object sender, EventArgs e)
        {
            FillTab();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmGesStatSettimanali_KeyDown(object sender, KeyEventArgs e)
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
            if (rdbRep.Checked)
                FillDatiRep();
            else
                FillDatiFas();
        }

        private void dtpDay_ValueChanged(object sender, EventArgs e)
        {
            DateTime dDay = dtpDay.Value;
            DateTime dDti = dtpDay.Value;
            DateTime dDtf = dtpDay.Value;

            string s = "SELECT * FROM TabCalendario WHERE ";
            s += "tab_yea='" + dDay.Year.ToString() + "' AND ";
            s += "tab_tri='S' AND ";
            //WHERE        (tab_yea = '2019') AND (tab_tri = 'S') AND (tab_dti <= CONVERT(DATETIME, '2019-12-05', 102)) AND (tab_dtf >= CONVERT(DATETIME, '2019-12-05', 102))
            s += "tab_dti <= " + _clsFun.DaySql(dDay) + " AND ";
            s += "tab_dtf >= " + _clsFun.DaySql(dDay) + " ";
            //s += "ORDER BY tab_dti";
            DataTable tCal = _clsFun.FillTabSql("TabCalendario", s, false, _strConSql);
            if (tCal.Rows.Count > 0)
            {
                dtpDti.Value = (DateTime)tCal.Rows[0]["tab_dti"];
                dtpDtf.Value = (DateTime)tCal.Rows[0]["tab_dtf"];
            }
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

        private void FillDatiRep()
        {
            DateTime dDti = dtpDti.Value;
            DateTime dDtf = dtpDtf.Value;

            string s = "";
            DataRow x;
            DataRow[] j;
            DataRow[] j2;

            DataTable t = new clsGenTabTmp().TabTmpStaSettimana("TabSet");

            string sNeg = cmbNeg.SelectedValue.ToString();
            
            s = "SELECT * FROM TabReparti";
            DataTable tRep = _clsFun.FillTabSql("TabRep", s, false, _strConSql);

            s = "SELECT AnaArticoli.art_cod, AnaArticoli.art_des, AnaArticoli.art_iva, AnaArticoli.art_umi, AnaArticoli.art_rep, TabReparti.tab_des AS RepDes ";
            s += "FROM AnaArticoli ";
            s += "LEFT OUTER JOIN TabReparti ON AnaArticoli.art_rep = TabReparti.tab_cod";
            DataTable tArt = _clsFun.FillTabSql("Art", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            //lblDti.Text = dDti.ToShortDateString();
            //lblDtf.Text = dDtf.ToShortDateString();

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
            s += "ven_umi, ";
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
            s += "WHERE (ven_day >= " + _clsFun.DaySql(dDti) + ") AND (ven_day <= " + _clsFun.DaySql(dDtf) + ") AND ";
            s += "(GesNegVen.ven_cau <> 'MOV') ";
            if (sNeg != "")
                s += " AND ven_neg='" + sNeg + "' ";
            if (_strStatVisProf == "N")
                s += " AND GesNegVet.vet_cau<>'PRO' ";
            s += "ORDER BY ven_day,ven_cau,ven_neg,ven_pos,ven_sco";
            DataTable tVen = _clsFun.FillTabSql("VEN", s, false, _strConSqlSta);

            string sScoKey = "";
            decimal d = 0;
            decimal dSco = 0;
            ArrayList aScoRep = new ArrayList();

            x = t.NewRow();
            x["TmpRep"] = "TOT";
            x["TmpRed"] = "";
            x["TmpInc"] = 0;
            t.Rows.Add(x);

            progressBar1.Value = 0;
            progressBar1.Maximum = (int)tVen.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in tVen.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                s = ((DateTime)y["ven_day"]).ToString("yyyyMMdd") + (string)y["ven_cau"] + (string)y["ven_neg"] + (string)y["ven_pos"] + (string)y["ven_sco"];

                if (sScoKey.Length > 8 && sScoKey.Substring(0, 8) != s.Substring(0,8))
                    dSco = 0;

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

                    if ((decimal)y["vet_sct"] > 0)
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

                string sRep = "";
                string sRed = "";
                string sUmi = "";

                j2 = tArt.Select("art_cod='" + (string)y["ven_art"] + "'");
                if (j2.Length > 0)
                {
                    sRep = (string)j2[0]["art_rep"];
                    sUmi = (string)j2[0]["art_umi"];
                }
                else if (!DBNull.Value.Equals(y["ven_rep"]))
                {
                    sRep = (string)y["ven_rep"];
                    sUmi = (string)y["ven_umi"];
                }

                j = tRep.Select("tab_cod='" + sRep + "'");
                if (j.Length > 0)
                    sRed = (string)j[0]["tab_des"];

                if (sRep == "")
                    sRed = "Non definito";

                j = t.Select("TmpRep='" + sRep + "'");
                if (j.Length == 0)
                {
                    x = t.NewRow();
                    x["TmpDay"] = ((DateTime)y["ven_day"]).ToString("yyyyMMdd");
                    x["TmpRep"] = sRep;
                    x["TmpRed"] = sRed;
                    t.Rows.Add(x);
                }

                j = t.Select("TmpRep='" + sRep + "'");

                int ii = (int)((DateTime)y["ven_day"]).DayOfWeek;

                if (ii == 0)
                    ii = 7;

                string sDay = ii.ToString();
                string sFldv = "TmpD" + sDay + "v";     //Venduto
                string sFldi = "TmpD" + sDay + "i";     //Venduto con costo netto iVA
                string sFlda = "TmpD" + sDay + "a";     //Costo
                string sFldm = "TmpD" + sDay + "m";     //Margine
                string sFldc = "TmpD" + sDay + "c";     //Clienti

                decimal dVni = 0;
                decimal dAcq = 0;
                if (!DBNull.Value.Equals(y["ven_cos"]) && (decimal)y["ven_cos"] > 0)
                {
                    //dVni = _clsFun.MenoIva((decimal)y["ven_ven"], 1);
                    dVni = _clsFun.MenoIva((decimal)y["ven_ven"], Convert.ToDecimal(y["ven_iva"]));
                    if ((decimal)y["ven_qkg"] > 0)
                        dAcq = (decimal)y["ven_cos"] * (decimal)y["ven_qkg"];
                    else
                        dAcq = (decimal)y["ven_cos"] * (decimal)y["ven_qta"];
                }

                j[0][sFldv] = (decimal)j[0][sFldv] + (decimal)y["ven_ven"];
                j[0][sFldi] = (decimal)j[0][sFldi] + dVni;
                j[0][sFlda] = (decimal)j[0][sFlda] + dAcq;
                j[0][sFldm] = _clsFun.Margine((decimal)j[0][sFldi],(decimal)j[0][sFlda], 0, 0, "P");

                j[0]["TmpD0v"] = (decimal)j[0]["TmpD0v"] + (decimal)y["ven_ven"];
                j[0]["TmpD0i"] = (decimal)j[0]["TmpD0i"] + dVni;
                j[0]["TmpD0a"] = (decimal)j[0]["TmpD0a"] + dAcq;
                j[0]["TmpD0m"] = _clsFun.Margine((decimal)j[0]["TmpD0i"], (decimal)j[0]["TmpD0a"], 0, 0, "P");

                j = t.Select("TmpRep='TOT'");

                j[0][sFldv] = (decimal)j[0][sFldv] + (decimal)y["ven_ven"];
                j[0][sFldi] = (decimal)j[0][sFldi] + dVni;
                j[0][sFlda] = (decimal)j[0][sFlda] + dAcq;
                j[0][sFldm] = _clsFun.Margine((decimal)j[0][sFldi], (decimal)j[0][sFlda], 0, 0, "P");
                j[0][sFldc] = dSco;

                j[0]["TmpD0v"] = (decimal)j[0]["TmpD0v"] + (decimal)y["ven_ven"];
                j[0]["TmpD0i"] = (decimal)j[0]["TmpD0i"] + dVni;
                j[0]["TmpD0a"] = (decimal)j[0]["TmpD0a"] + dAcq;
                j[0]["TmpD0m"] = _clsFun.Margine((decimal)j[0]["TmpD0i"], (decimal)j[0]["TmpD0a"], 0, 0, "P");
                //j[0]["TmpD0c"] = (decimal)j[0]["TmpD0c"] + dSco;
                    
            }

            d = 0;
            j = t.Select("TmpRep='TOT'");
            if(j.Length > 0)
                d = (decimal)j[0]["TmpD0v"];

            d = 0;
            for (int ii = 1; ii <= 7; ii++)
            {
                string sFld = "TmpD" + ii.ToString() + "c";
                d += ((decimal)t.Rows[0][sFld]);
            }

            t.Rows[0]["TmpD0c"] = d;

            foreach(DataRow y in t.Rows)
            {
                if ((decimal)y["TmpD0v"] > 0 && d > 0)
                    y["TmpInc"] = (decimal)y["TmpD0v"] /d * 100;
            }

            dgv1.DataSource = t;
        }

        private void FillDatiFas()
        {
            DateTime dDti = dtpDti.Value;
            DateTime dDtf = dtpDtf.Value;

            string s = "";
            DataRow x;
            DataRow[] j;
            DataRow[] j2;

            DataTable t = new clsGenTabTmp().TabTmpStaSettimanaFascia("TabSet");

            string sNeg = cmbNeg.SelectedValue.ToString();
            //DateTime dDay = dtpDay.Value;
            //DateTime dDti = dtpDay.Value;
            //DateTime dDtf = dtpDay.Value;

            //s = "SELECT * FROM TabCalendario WHERE ";
            //s += "tab_yea='" + dDay.Year.ToString() + "' AND ";
            //s += "tab_tri='S' AND ";
            ////WHERE        (tab_yea = '2019') AND (tab_tri = 'S') AND (tab_dti <= CONVERT(DATETIME, '2019-12-05', 102)) AND (tab_dtf >= CONVERT(DATETIME, '2019-12-05', 102))
            //s += "tab_dti <= " + _clsFun.DaySql(dDay) + " AND ";
            //s += "tab_dtf >= " + _clsFun.DaySql(dDay) + " ";
            ////s += "ORDER BY tab_dti";
            //DataTable tCal = _clsFun.FillTabSql("TabCalendario", s, false, _strConSql);
            
            //if(tCal.Rows.Count > 0)
            //{
            s = "SELECT * FROM TabReparti";
            DataTable tRep = _clsFun.FillTabSql("TabRep", s, false, _strConSql);

            s = "SELECT AnaArticoli.art_cod, AnaArticoli.art_des, AnaArticoli.art_iva, AnaArticoli.art_umi, AnaArticoli.art_rep, TabReparti.tab_des AS RepDes ";
            s += "FROM AnaArticoli ";
            s += "LEFT OUTER JOIN TabReparti ON AnaArticoli.art_rep = TabReparti.tab_cod";
            DataTable tArt = _clsFun.FillTabSql("Art", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            //dDti = (DateTime)tCal.Rows[0]["tab_dti"];
            //dDtf = (DateTime)tCal.Rows[0]["tab_dtf"];

            lblDti.Text = dDti.ToShortDateString();
            lblDtf.Text = dDtf.ToShortDateString();

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
            s += "WHERE (ven_day >= " + _clsFun.DaySql(dDti) + ") AND (ven_day <= " + _clsFun.DaySql(dDtf) + ") AND ";
            s += "(GesNegVen.ven_cau <> 'MOV') ";
            if (sNeg != "")
                s += " AND ven_neg='" + sNeg + "' ";
            if (_strStatVisProf == "N")
                s += " AND GesNegVet.vet_cau<>'PRO' ";
            s += "ORDER BY ven_day,ven_cau,ven_neg,ven_pos,ven_sco";
            DataTable tVen = _clsFun.FillTabSql("VEN", s, false, _strConSqlSta);

            string sScoKey = "";
            decimal d = 0;
            decimal dSco = 0;
            ArrayList aScoRep = new ArrayList();

            x = t.NewRow();
            x["TmpOra"] = "TOT";
            t.Rows.Add(x);

            progressBar1.Value = 0;
            progressBar1.Maximum = (int)tVen.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in tVen.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                s = ((DateTime)y["ven_day"]).ToString("yyyyMMdd") + (string)y["ven_ora"] + (string)y["ven_cau"] + (string)y["ven_neg"] + (string)y["ven_pos"] + (string)y["ven_sco"];

                if (sScoKey.Length > 12 && sScoKey.Substring(8, 2) != s.Substring(8,2))
                    dSco = 0;

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

                    if ((decimal)y["vet_sct"] > 0)
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

                int ii = Convert.ToInt16( ((string)y["ven_ora"]).Substring(0, 2));

                string sOra = ii.ToString("00") + "-" + (ii + 1).ToString("00");

                string sUmi = "";

                j2 = tArt.Select("art_cod='" + (string)y["ven_art"] + "'");
                if (j2.Length > 0)
                    sUmi = (string)j2[0]["art_umi"];
                else if (!DBNull.Value.Equals(y["ven_rep"]))
                    sUmi = (string)y["ven_umi"];

                //j = tRep.Select("tab_cod='" + sRep + "'");
                //if (j.Length > 0)
                //    sRed = (string)j[0]["tab_des"];

                //if (sRep == "")
                //    sRed = "Non definito";

                j = t.Select("TmpOra='" + sOra + "'");
                if (j.Length == 0)
                {
                    x = t.NewRow();
                    x["TmpDay"] = ((DateTime)y["ven_day"]).ToString("yyyyMMdd");
                    x["TmpOra"] = sOra;
                    t.Rows.Add(x);
                }

                j = t.Select("TmpOra='" + sOra + "'");

                ii = (int)((DateTime)y["ven_day"]).DayOfWeek;

                if (ii == 0)
                    ii = 7;

                string sDay = ii.ToString();
                string sFldv = "TmpD" + sDay + "v";     //Venduto
                string sFldc = "TmpD" + sDay + "c";     //Clienti
                string sFldm = "TmpD" + sDay + "m";     //Media

                j[0][sFldv] = (decimal)j[0][sFldv] + (decimal)y["ven_ven"];
                j[0][sFldc] = dSco;
                j[0][sFldm] = Math.Round((decimal)j[0][sFldv] / (decimal)j[0][sFldc], 2);


                j[0]["TmpD0v"] = (decimal)j[0]["TmpD0v"] + (decimal)y["ven_ven"];

                j = t.Select("TmpOra='TOT'");

                j[0][sFldv] = (decimal)j[0][sFldv] + (decimal)y["ven_ven"];
                //j[0][sFldi] = (decimal)j[0][sFldi] + dVni;
                //j[0][sFlda] = (decimal)j[0][sFlda] + dAcq;
                //j[0][sFldm] = _clsFun.Margine((decimal)j[0][sFldi], (decimal)j[0][sFlda], 0, 0, "P");
                j[0][sFldc] = dSco;

                j[0]["TmpD0v"] = (decimal)j[0]["TmpD0v"] + (decimal)y["ven_ven"];
                //j[0]["TmpD0i"] = (decimal)j[0]["TmpD0i"] + dVni;
                //j[0]["TmpD0a"] = (decimal)j[0]["TmpD0a"] + dAcq;
                //j[0]["TmpD0m"] = _clsFun.Margine((decimal)j[0]["TmpD0i"], (decimal)j[0]["TmpD0a"], 0, 0, "P");
                ////j[0]["TmpD0c"] = (decimal)j[0]["TmpD0c"] + dSco;
                    
            }

            d = 0;
            j = t.Select("TmpOra='TOT'");
            if(j.Length > 0)
                d = (decimal)j[0]["TmpD0v"];

            d = 0;

            Console.WriteLine("zzz");

            DataView v = new DataView(t, "", "TmpOra", DataViewRowState.CurrentRows);

            t = v.ToTable();

            //Aggiorno clienti e media sui giorni

            string sFld = "";

            foreach(DataRow y in t.Rows)
            {
                d = 0;

                for (int ii = 1; ii <= 7; ii++)
                {
                    sFld = "TmpD" + ii.ToString() + "c";
                    d += ((decimal)y[sFld]);
                }
                y["TmpD0c"] = d;
                y["TmpD0m"] = Math.Round((decimal)y["TmpD0v"] / (decimal)y["TmpD0c"], 2);

            }

            //Aggiorno clienti e media su riga TOT

            Console.WriteLine("zzz");

            decimal dCli = 0;

            for (int ii = 1; ii <= 7; ii++)
            {
                sFld = "TmpD" + ii.ToString() + "c";
                d = 0;
                foreach (DataRow y in t.Rows)
                {
                    if ((string)y["TmpOra"] == "TOT")
                    {
                        y[sFld] = d;
                        dCli += d;

                        //y["TmpD0m"] = Math.Round((decimal)y["TmpD0v"] / d, 2);

                        if (d > 0)
                        {
                            //sFld = "TmpD" + ii.ToString() + "m";
                            y["TmpD" + ii.ToString() + "m"] = Math.Round((decimal)y["TmpD" + ii.ToString() + "v"] / d, 2);
                        }
                    }
                    else
                        d += ((decimal)y[sFld]);
                }

            }

            Console.WriteLine("zzz");


            //foreach(DataRow y in t.Rows)
            //{
            //    y["TmpInc"] = (decimal)y["TmpD0v"] /d * 100;
            //}

            j = t.Select("TmpOra='TOT'");
            j[0]["TmpD0c"] = dCli;
            j[0]["TmpD0m"] = Math.Round((decimal)j[0]["TmpD0v"] / dCli, 2);

            dgv1.DataSource = t;
        }


        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (rdbRep.Checked)
                PdfRep();
            else
                PdfOre();
        }

        private void PdfRep()
        {
            string s = "";
            DataRow[] j;
            int iRows = 14;
            int iRow = 0;

            DataView v = new DataView((DataTable)dgv1.DataSource, "", "TmpRep", DataViewRowState.CurrentRows);

            DataTable t = v.ToTable();

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 7, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 6, XFontStyle.Bold);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);
            XFont font4 = new XFont("Courier new", 8, XFontStyle.Bold);

            double Y = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            XPen pen = new XPen(XColors.Black, 1);

            //string sFam = ""; // (string)t.Rows[0]["InvL1c"] + (string)t.Rows[0]["InvL2c"] + (string)t.Rows[0]["InvL3c"];

            //decimal dRig = 0;
            //decimal dPez = 0;
            //decimal dVal = 0;

            Boolean bPiede = false;

            for (int i = 0; i <= t.Rows.Count - 1; i++)
            {
                if (iRow > iRows || iRow == 0 || bPiede)
                {
                    if (iRow > 0)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }

                    X = 20;
                    iRow = 0;

                    s = DateTime.Today.ToShortDateString() + "   VENDUTO SETTIMANALE PER REPARTO DAL " + dtpDti.Value.ToShortDateString() + " al " + dtpDtf.Value.ToShortDateString();
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    pen = new XPen(XColors.Black, 2);
                    gfx.DrawLine(pen, Y + 0, X + 7, Y + 580, X + 7);

                    X += 20;

                    if (bPiede)
                        s = "Lunedì     Martedì   Mercoledì     Giovedì     Venerdì     Sabato    Domenica      Totale";
                    else
                        s = "Lunedì     Martedì   Mercoledì     Giovedì     Venerdì     Sabato    Domenica      Totale    %Inc.";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 155, X+5, XStringFormats.Default);

                    DateTime dDay = dtpDti.Value;

                    //s = "";

                    //for (int ii = 0; ii < 7; ii++)
                    //{
                    //    s += dDay.AddDays(ii).ToShortDateString() + "    ";
                    //}
                    //gfx.DrawString(s, font2, XBrushes.Black, Y + 145, X + 10, XStringFormats.Default);

                    pen = new XPen(XColors.Black, 1);
                    gfx.DrawLine(pen, Y + 0, X + 17, Y + 580, X + 17);

                    X += 35;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    if ((string)t.Rows[i]["TmpRep"] == "TOT")
                        s = "TOTALI";
                    else
                    {
                        s = (string)t.Rows[i]["TmpRep"] + " " + (string)t.Rows[i]["TmpRed"];
                        if (s.Length > 15)
                            s = s.Substring(0, 15);
                    }
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X, XStringFormats.Default);

                    s = "Valore venduto";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 13, XStringFormats.Default);

                    s = "Valore in acquisto lordo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 23, XStringFormats.Default);

                    s = "Margine";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 33, XStringFormats.Default);

                    if (bPiede)
                    {
                        s = "Presenze";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 43, XStringFormats.Default);

                        s = "Scontrino medio";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 53, XStringFormats.Default);

                        s = "Media incassi settimanali";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 63, XStringFormats.Default);

                        s = "Media scontrino settimanale";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 73, XStringFormats.Default);

                        s = "Media presenze settimanali";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 83, XStringFormats.Default);
                    }

                    int iStep = 0;
                    string sFld = "";

                    for (int ii = 1; ii <= 7; ii++)
                    {
                        sFld = "TmpD" + ii.ToString() + "v";
                        s = ((decimal)t.Rows[i][sFld]).ToString("####,##0.00");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 13, frmDX);

                        sFld = "TmpD" + ii.ToString() + "a";
                        s = ((decimal)t.Rows[i][sFld]).ToString("####,##0.00");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 23, frmDX);

                        sFld = "TmpD" + ii.ToString() + "m";
                        s = ((decimal)t.Rows[i][sFld]).ToString("####,##0.00");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 33, frmDX);

                        if (bPiede)
                        {
                            Console.WriteLine("xxx");

                            sFld = "TmpD" + ii.ToString() + "c";
                            s = ((decimal)t.Rows[i][sFld]).ToString("####,##0");
                            gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 43, frmDX);

                            sFld = "TmpD" + ii.ToString() + "v";
                            decimal dImp = (decimal)t.Rows[i][sFld];
                            sFld = "TmpD" + ii.ToString() + "c";
                            decimal dCli = (decimal)t.Rows[i][sFld];

                            if (dImp > 0 && dCli > 0)
                            {
                                s = (dImp / dCli).ToString("####,##0.00");
                                gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 53, frmDX);
                            }

                        }

                        iStep += 50;
                    }

                    sFld = "TmpD" + 0.ToString() + "v";
                    s = ((decimal)t.Rows[i][sFld]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 13, frmDX);

                    sFld = "TmpD" + 0.ToString() + "a";
                    s = ((decimal)t.Rows[i][sFld]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 23, frmDX);

                    sFld = "TmpD" + 0.ToString() + "m";
                    s = ((decimal)t.Rows[i][sFld]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 33, frmDX);

                    if ((string)t.Rows[i]["TmpRep"] != "TOT")
                    {
                        s = ((decimal)t.Rows[i]["TmpInc"]).ToString("####,##0.00");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 210 + iStep, X + 13, frmDX);
                    }
                    else
                    {
                        sFld = "TmpD" + 0.ToString() + "c";
                        decimal dCli = (decimal)t.Rows[i][sFld];
                        s = dCli.ToString("####,##0");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 43, frmDX);

                        sFld = "TmpD" + 0.ToString() + "v";
                        decimal dImp = (decimal)t.Rows[i][sFld];
                        s = (dImp / 7).ToString("####,##0.00");
                        gfx.DrawString(s, font1, XBrushes.Black, Y - 170 + iStep, X + 63, frmDX);

                        //sFld = "TmpD" + 0.ToString() + "c";
                        //decimal d2 = d1 / (decimal)t.Rows[0][sFld];
                        s = (dImp / dCli).ToString("####,##0.00");
                        gfx.DrawString(s, font1, XBrushes.Black, Y - 170 + iStep, X + 73, frmDX);

                        //sFld = "TmpD" + 0.ToString() + "c";
                        //decimal d3 = (decimal)t.Rows[0][sFld] / 7;
                        s = (dCli / 7).ToString("####,##0.00");
                        gfx.DrawString(s, font1, XBrushes.Black, Y - 170 + iStep, X + 83, frmDX);
                    }

                    iStep += 50;

                    if (bPiede)
                    {
                        pen = new XPen(XColors.Black, 1);
                        gfx.DrawLine(pen, Y + 0, X + 100, Y + 580, X + 100);
                    }
                    else
                    {
                        pen = new XPen(XColors.Black, 1);
                        gfx.DrawLine(pen, Y + 0, X + 40, Y + 580, X + 40);
                    }
                }

                X += 50;

                if (iRow == t.Rows.Count - 1)
                {
                    bPiede = true;
                }

            }

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\StaSetRep_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch (Exception ex)
            {
                MessageBox.Show("File di stampa già aperto!");
            }

        }

        private void PdfOre()
        {
            string s = "";
            DataRow[] j;
            int iRows = 14;
            int iRow = 0;

            DataView v = new DataView((DataTable)dgv1.DataSource, "", "TmpOra", DataViewRowState.CurrentRows);

            DataTable t = v.ToTable();

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 7, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 6, XFontStyle.Bold);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);
            XFont font4 = new XFont("Courier new", 8, XFontStyle.Bold);

            double Y = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            XPen pen = new XPen(XColors.Black, 1);

            string sFam = ""; // (string)t.Rows[0]["InvL1c"] + (string)t.Rows[0]["InvL2c"] + (string)t.Rows[0]["InvL3c"];

            decimal dRig = 0;
            decimal dPez = 0;
            decimal dVal = 0;

            Boolean bPiede = false;

            for (int i = 0; i <= t.Rows.Count - 1; i++)
            {
                if (iRow > iRows || iRow == 0 || bPiede)
                {
                    if (iRow > 0)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }

                    X = 20;
                    iRow = 0;

                    s = DateTime.Today.ToShortDateString() + "   VENDUTO SETTIMANALE PER FASCIA ORARIA DAL " + lblDti.Text + " al " + lblDtf.Text;
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    pen = new XPen(XColors.Black, 2);
                    gfx.DrawLine(pen, Y + 0, X + 7, Y + 580, X + 7);

                    X += 20;

                    //if (bPiede)
                    //    s = "Lunedì     Martedì   Mercoledì     Giovedì     Venerdì     Sabato    Domenica      Totale";
                    //else
                    //s = "Lunedì     Martedì   Mercoledì     Giovedì     Venerdì     Sabato    Domenica      Totale    %Inc.";
                    s = "Lunedì     Martedì   Mercoledì     Giovedì     Venerdì     Sabato    Domenica      Totale";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 155, X + 5, XStringFormats.Default);

                    //DateTime dDay = Convert.ToDateTime(lblDti.Text);

                    //s = "";

                    //for (int ii = 0; ii < 7; ii++)
                    //{
                    //    s += dDay.AddDays(ii).ToShortDateString() + "    ";
                    //}
                    //gfx.DrawString(s, font2, XBrushes.Black, Y + 145, X + 10, XStringFormats.Default);

                    //pen = new XPen(XColors.Black, 1);
                    //gfx.DrawLine(pen, Y + 0, X + 17, Y + 580, X + 17);

                    X += 35;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    if ((string)t.Rows[i]["TmpOra"] == "TOT")
                        s = "TOTALI";
                    else
                    {
                        s = (string)t.Rows[i]["TmpOra"];
                        if (s.Length > 15)
                            s = s.Substring(0, 15);
                    }
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X, XStringFormats.Default);

                    s = "Valore venduto";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 13, XStringFormats.Default);

                    s = "Presenze";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 23, XStringFormats.Default);

                    s = "Spesa media";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 33, XStringFormats.Default);

                    //if (bPiede)
                    //{
                    //    s = "Presenze";
                    //    gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 43, XStringFormats.Default);

                    //    s = "Scontrino medio";
                    //    gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 53, XStringFormats.Default);

                    //    s = "Media incassi settimanali";
                    //    gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 63, XStringFormats.Default);

                    //    s = "Media scontrino settimanale";
                    //    gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 73, XStringFormats.Default);

                    //    s = "Media presenze settimanali";
                    //    gfx.DrawString(s, font1, XBrushes.Black, Y + 2, X + 83, XStringFormats.Default);
                    //}

                    int iStep = 0;
                    string sFld = "";

                    for (int ii = 1; ii <= 7; ii++)
                    {
                        sFld = "TmpD" + ii.ToString() + "v";
                        s = ((decimal)t.Rows[i][sFld]).ToString("####,##0.00");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 13, frmDX);

                        sFld = "TmpD" + ii.ToString() + "c";
                        s = ((decimal)t.Rows[i][sFld]).ToString("####,##0");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 23, frmDX);

                        sFld = "TmpD" + ii.ToString() + "m";
                        s = ((decimal)t.Rows[i][sFld]).ToString("####,##0.00");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 33, frmDX);

                        //if (bPiede)
                        //{
                        //    Console.WriteLine("xxx");

                        //    sFld = "TmpD" + ii.ToString() + "c";
                        //    s = ((decimal)t.Rows[i][sFld]).ToString("####,##0");
                        //    gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 43, frmDX);

                        //    sFld = "TmpD" + ii.ToString() + "v";
                        //    decimal dImp = (decimal)t.Rows[i][sFld];
                        //    sFld = "TmpD" + ii.ToString() + "c";
                        //    decimal dCli = (decimal)t.Rows[i][sFld];

                        //    if (dImp > 0 && dCli > 0)
                        //    {
                        //        s = (dImp / dCli).ToString("####,##0.00");
                        //        gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 53, frmDX);
                        //    }

                        //}

                        iStep += 50;
                    }

                    sFld = "TmpD" + 0.ToString() + "v";
                    s = ((decimal)t.Rows[i][sFld]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 13, frmDX);

                    sFld = "TmpD" + 0.ToString() + "c";
                    s = ((decimal)t.Rows[i][sFld]).ToString("####,##0");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 23, frmDX);

                    sFld = "TmpD" + 0.ToString() + "m";
                    s = ((decimal)t.Rows[i][sFld]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 33, frmDX);

                    //if ((string)t.Rows[i]["TmpOra"] != "TOT")
                    //{
                    //    s = ((decimal)t.Rows[i]["TmpInc"]).ToString("####,##0.00");
                    //    gfx.DrawString(s, font1, XBrushes.Black, Y + 210 + iStep, X + 13, frmDX);
                    //}
                    //else
                    //{
                    //    sFld = "TmpD" + 0.ToString() + "c";
                    //    decimal dCli = (decimal)t.Rows[i][sFld];
                    //    s = dCli.ToString("####,##0");
                    //    gfx.DrawString(s, font1, XBrushes.Black, Y + 180 + iStep, X + 43, frmDX);

                    //    sFld = "TmpD" + 0.ToString() + "v";
                    //    decimal dImp = (decimal)t.Rows[i][sFld];
                    //    s = (dImp / 7).ToString("####,##0.00");
                    //    gfx.DrawString(s, font1, XBrushes.Black, Y - 170 + iStep, X + 63, frmDX);

                    //    //sFld = "TmpD" + 0.ToString() + "c";
                    //    //decimal d2 = d1 / (decimal)t.Rows[0][sFld];
                    //    s = (dImp / dCli).ToString("####,##0.00");
                    //    gfx.DrawString(s, font1, XBrushes.Black, Y - 170 + iStep, X + 73, frmDX);

                    //    //sFld = "TmpD" + 0.ToString() + "c";
                    //    //decimal d3 = (decimal)t.Rows[0][sFld] / 7;
                    //    s = (dCli / 7).ToString("####,##0.00");
                    //    gfx.DrawString(s, font1, XBrushes.Black, Y - 170 + iStep, X + 83, frmDX);
                    //}

                    iStep += 50;

                    if (bPiede)
                    {
                        pen = new XPen(XColors.Black, 1);
                        gfx.DrawLine(pen, Y + 0, X + 100, Y + 580, X + 100);
                    }
                    else
                    {
                        pen = new XPen(XColors.Black, 1);
                        gfx.DrawLine(pen, Y + 0, X + 40, Y + 580, X + 40);
                    }
                }

                X += 50;

                if (iRow == t.Rows.Count - 1)
                {
                    bPiede = false;
                }

            }

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\StaSetOra_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch (Exception ex)
            {
                MessageBox.Show("File di stampa già aperto!");
            }

        }
    }
}
