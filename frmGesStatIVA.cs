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
using System.Diagnostics;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace APOffice
{
    public partial class frmGesStatIVA : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public DateTime _dayDti = DateTime.Today;
        public DateTime _dayDtf = DateTime.Today;

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public frmGesStatIVA()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
        }

        private void frmGesStaIVA_Load(object sender, EventArgs e)
        {
            dtpDti.Value = _dayDti;
            dtpDtf.Value = _dayDtf;


            FillTab();
            SetDgv1();
            dgv1.DataSource = new clsGenTabTmp().TabTmpIvaRiep("TabIva");

            //progressBar1.Visible = false;
        }
        private void frmGesStatIVA_KeyDown(object sender, KeyEventArgs e)
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
            string sNeg = cmbNeg.SelectedValue.ToString();
            DateTime dDti = dtpDti.Value;
            DateTime dDtf = dtpDtf.Value;

            FillDatiIva(sNeg, dDti, dDtf);
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
            cTbc.DataPropertyName = "iva_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 130;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_imp";
            cTbc.Name = "Imponibile";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_iva";
            cTbc.Name = "Imposta";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_tot";
            cTbc.Name = "Totale";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
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

        public DataView FillDatiIva(string strNeg, DateTime dayIni, DateTime dayFin)
        {
            string s = "";
            DataRow x;
            DataRow[] j;
            DataRow[] j2;
            decimal d = 0;
            decimal dQta = 0;
            decimal dTot = 0;
            decimal dNiv = 0;       //Venduto netto iVA
            decimal dTotImp = 0;    //Somma Imponibile
            decimal dTotIva = 0;    //Somma imposta

            //string sNeg = cmbNeg.SelectedValue.ToString();
            string sNeg = strNeg;
            string sIvd = _clsFun.ParGet(clsDefine.enuParametri.Par012CodIVAxDefault, _strConSql);  //Iva di default
            string sIva = "";       //IVA
            string sIde = "";       //Descrizione IVA
            string sArd = "";
            string sMsg = "";

            ArrayList aScoRep = new ArrayList();

            decimal dSco = 0;

            //DataTable t = new clsGenTabTmp().TabTmpStaReparto("TabRep");
            //DataColumn[] keys = new DataColumn[2];
            //keys[0] = t.Columns["tab_cod"];
            //keys[1] = t.Columns["tab_neg"];
            //t.PrimaryKey = keys;
            
            DataTable t = new clsGenTabTmp().TabTmpIvaRiep("TabRep");
            DataColumn[] keys = new DataColumn[2];
            keys[0] = t.Columns["iva_cod"];
            keys[1] = t.Columns["iva_neg"];
            t.PrimaryKey = keys;

            s = "SELECT * FROM TabIva";
            DataTable tIva = _clsFun.FillTabSql("TabIva", s, false, _strConSql);

            s = "SELECT AnaArticoli.art_cod, AnaArticoli.art_des, AnaArticoli.art_iva, TabReparti.tab_des AS RepDes ";
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
            s += "ven_qta, ";
            s += "ven_qkg, ";
            s += "ven_cos, ";
            s += "ven_ven, ";
            s += "vet_imp, ";
            s += "vet_sct ";
            s += "FROM GesNegVen JOIN GesNegVet ON ";
            s += "GesNegVet.vet_neg = GesNegVen.ven_neg AND ";
            s += "GesNegVet.vet_cau = GesNegVen.ven_cau AND ";
            s += "GesNegVet.vet_day = GesNegVen.ven_day AND ";
            s += "GesNegVet.vet_ora = GesNegVen.ven_ora AND ";
            s += "GesNegVet.vet_pos = GesNegVen.ven_pos AND ";
            s += "GesNegVet.vet_sco = GesNegVen.ven_sco ";
            s += "WHERE (ven_day >= " + _clsFun.DaySql(dayIni) + ") AND (ven_day <= " + _clsFun.DaySql(dayFin) + ") AND ";
            //s += "(GesNegVen.ven_cau <> 'MOV') ";
            s += "(GesNegVen.ven_cau = 'SCO') ";
            if (sNeg != "")
                s += " AND ven_neg='" + sNeg + "' ";
            //s += "ORDER BY ven_sco";
            s += "ORDER BY ven_day,ven_cau,ven_neg,ven_pos,ven_sco";
            DataTable tVen = _clsFun.FillTabSql("VEN", s, false, _strConSqlSta);

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

            string sScoKey = "";

            progressBar1.Visible = true;
            progressBar1.Value = 0;
            progressBar1.Maximum = (int)tVen.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in tVen.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if ((string)y["ven_art"] == "0021097")
                    Console.WriteLine("zzzzzz");

                //s = ((DateTime)y["ven_day"]).ToString("yyyyMMdd") + (string)y["ven_neg"] + (string)y["ven_pos"] + (string)y["ven_sco"];
                s = ((DateTime)y["ven_day"]).ToString("yyyyMMdd") + (string)y["ven_cau"] + (string)y["ven_neg"] + (string)y["ven_pos"] + (string)y["ven_sco"];

                if ((string)y["ven_sco"] == "00045")
                    Console.WriteLine("zzzzzz");

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

                if ((string)y["ven_art"] == "0021097")
                    Console.WriteLine("aaaaaaaaaa");

                sIva = "";
                sArd = "";
                decimal dAli = 0;

                j2 = tArt.Select("art_cod='" + (string)y["ven_art"] + "'");
                if (j2.Length > 0)
                {
                    sIva = (string)j2[0]["art_iva"];
                    sArd = (string)j2[0]["art_des"];
                }

                if(sIva == "")
                    sIva = (string)y["ven_iva"];

                if (sIva == "")
                    sIva = sIvd;

                j = tIva.Select("tab_cod='" + sIva + "'");
                if (j.Length > 0)
                {
                    sIde = (string)j[0]["tab_des"];
                    dAli = Convert.ToDecimal(j[0]["tab_ali"]);
                }
                if (sIva == "")
                {
                    sMsg += "Articolo " + y["ven_art"] + " - " + sArd + " con reparto non definito!" + _clsDef.CRLF;
                    sIde = "Non definito";
                }

                if (sNeg == "")
                    j = t.Select("iva_cod='" + sIva + "'");
                else
                    j = t.Select("iva_cod='" + sIva + "' AND iva_neg='" + sNeg + "'");

                //_clsFun.FileLogDivFor("IVA",(string)y["ven_art"], sIva);

                if (j.Length == 0)
                {
                    x = t.NewRow();
                    x["iva_neg"] = sNeg;
                    x["iva_cod"] = sIva;
                    x["iva_des"] = sIde;
                    x["iva_ali"] = dAli;
                    x["iva_imp"] = 0;
                    //x["tmp_niv"] = 0;
                    t.Rows.Add(x);

                    j = t.Select("iva_cod='" + sIva + "'");
                }

                if (((string)y["ven_sct"]).Trim() == "RES")
                {
                    dQta -= (decimal)y["ven_qta"];
                    dTot -= (decimal)y["ven_ven"];
                    j[0]["iva_imp"] = (decimal)j[0]["iva_imp"] - (decimal)y["ven_ven"];
                }
                else
                {
                    dQta += (decimal)y["ven_qta"];
                    dTot += (decimal)y["ven_ven"];
                    j[0]["iva_imp"] = (decimal)j[0]["iva_imp"] + (decimal)y["ven_ven"];
                    //j[0]["tmp_qta"] = (decimal)j[0]["tmp_qta"] + (decimal)y["ven_qta"];

                    //dNiv = 0;
                    //if (_clsFun.Numerico(y["ven_iva"]))
                    //    dNiv = _clsFun.MenoIva((decimal)y["ven_ven"], Convert.ToDecimal(y["ven_iva"]));
                    //else
                    //    dNiv = (decimal)y["ven_ven"];

                    //j[0]["tmp_niv"] = (decimal)j[0]["tmp_niv"] + dNiv;

                    //if (aScoRep.IndexOf((string)y["ven_sco"] + sIva) < 0)
                    //{
                    //    aScoRep.Add((string)y["ven_sco"] + sIva);
                    //    j[0]["tmp_sco"] = (decimal)j[0]["tmp_sco"] + 1;
                    //}

                    //if ((decimal)y["ven_cos"] > 0)
                    //{
                    //    if ((decimal)y["ven_qkg"] > 0)
                    //        d = (decimal)y["ven_cos"] * (decimal)y["ven_qkg"];
                    //    else
                    //        d = (decimal)y["ven_cos"] * (decimal)y["ven_qta"];

                    //    j[0]["tmp_cdv"] = (decimal)j[0]["tmp_cdv"] + d;
                    //    j[0]["tmp_vcc"] = (decimal)j[0]["tmp_vcc"] + dNiv;
                    //}
                }

            }

            DataTable t2 = t.Clone();
            DataView v = new DataView(t, "", "iva_cod", DataViewRowState.CurrentRows);

            foreach (DataRowView r in v)
            {
                //r["iva_iva"] = _clsFun.ValIva((decimal)r["iva_imp"], (decimal)r["iva_ali"]);

                r["iva_iva"] = Math.Round(_clsFun.ValIva((decimal)r["iva_imp"], (decimal)r["iva_ali"]), 2, MidpointRounding.ToEven);
                r["iva_iva"] = Math.Round((decimal)r["iva_iva"], 2, MidpointRounding.ToEven);
                //r["iva_tot"] = Math.Round((decimal)r["iva_imp"] + (decimal)r["iva_iva"], 2, MidpointRounding.ToEven);
                r["iva_tot"] = Math.Round((decimal)r["iva_imp"], 2, MidpointRounding.ToEven);
                r["iva_imp"] = (decimal)r["iva_tot"] - (decimal)r["iva_iva"];           //20180118 Seck

                t2.ImportRow(r.Row);

                dTotImp += (decimal)r["iva_imp"];
                dTotIva += (decimal)r["iva_iva"];
            }

            v = new DataView(t2, "", "iva_cod", DataViewRowState.CurrentRows);

            dgv1.DataSource = v;

            lblImp.Text = dTotImp.ToString("#,###,##0.00");
            lblIva.Text = dTotIva.ToString("#,###,##0.00");
            lblTot.Text = (dTotImp + dTotIva).ToString("#,###,##0.00");

            progressBar1.Visible = false;

            return v;
        }

        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string sNeg = cmbNeg.SelectedValue.ToString();

            DataTable tIva = ((DataView)dgv1.DataSource).ToTable();

            PrnPdfIva("ven_iva", tIva, sNeg);
        }

        public void PrnPdfIva(string strTip, DataTable tabTab, string strNeg)
        {
            string s = "";
            int iRows = 18;
            int iRow = 0;
            decimal dImp = 0;
            decimal dIva = 0;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Reparti";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 11, XFontStyle.Bold);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);
            XFont font4 = new XFont("Courier new", 10, XFontStyle.Bold);

            double Y = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            DataTable t = tabTab.Copy();

            for (int i = 0; i <= t.Rows.Count - 1; i++)
            {
                if (iRow > iRows || iRow == 0)
                {
                    if (iRow > 0)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }

                    X = 20;
                    iRow = 0;
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA VENDUTO PER ";
                    if (strTip == "ven_rep")
                        s += "Reparto";
                    else
                        s += "IVA";
                    s += " dal " + dtpDti.Value.ToString("dd/MM/yy") + " al " + dtpDtf.Value.ToString("dd/MM/yy");

                    if (strNeg != "")
                        s += " - PV " + strNeg;

                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    s = "Descrizione";

                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = "Imponibile";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 190, X, XStringFormats.Default);

                    s = "Imposta";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 305, X, XStringFormats.Default);

                    s = "Totale";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 410, X, XStringFormats.Default);

                    X += 35;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = (string)t.Rows[i]["iva_cod"] + " " + (string)t.Rows[i]["iva_des"];
                    gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = ((decimal)t.Rows[i]["iva_imp"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 260, X + 3, frmDX);
                    //dTot += (decimal)t.Rows[i]["tmp_ils"];

                    s = ((decimal)t.Rows[i]["iva_iva"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 355, X + 3, frmDX);

                    s = ((decimal)t.Rows[i]["iva_imp"] + (decimal)t.Rows[i]["iva_iva"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 450, X + 3, frmDX);

                    dImp += (decimal)t.Rows[i]["iva_imp"];
                    dIva += (decimal)t.Rows[i]["iva_iva"];

                    //if (strTip == "ven_rep")
                    //{
                    //    s = ((decimal)t.Rows[i]["tmp_inc"]).ToString("####,##0.00") + "%";
                    //    gfx.DrawString(s, font4, XBrushes.Black, Y + 340, X + 13, frmDX);
                    //    dInc += (decimal)t.Rows[i]["tmp_inc"];

                    //    s = ((decimal)t.Rows[i]["tmp_sco"]).ToString("###,##0");
                    //    gfx.DrawString(s, font1, XBrushes.Black, Y + 400, X + 3, frmDX);
                    //    dSco += (decimal)t.Rows[i]["tmp_sco"];

                    //    s = ((decimal)t.Rows[i]["tmp_sci"]).ToString("###,##0.00") + "%";
                    //    gfx.DrawString(s, font4, XBrushes.Black, Y + 400, X + 13, frmDX);

                    //    s = ((decimal)t.Rows[i]["tmp_scm"]).ToString("###,##0.00") + "m";
                    //    gfx.DrawString(s, font4, XBrushes.Black, Y + 400, X + 23, frmDX);


                    //    s = ((decimal)t.Rows[i]["tmp_qta"]).ToString("###,##0");
                    //    gfx.DrawString(s, font1, XBrushes.Black, Y + 470, X + 3, frmDX);
                    //    dQta += (decimal)t.Rows[i]["tmp_qta"];

                    //    s = ((decimal)t.Rows[i]["tmp_qti"]).ToString("###,##0.00") + "%";
                    //    gfx.DrawString(s, font4, XBrushes.Black, Y + 470, X + 13, frmDX);

                    //    s = ((decimal)t.Rows[i]["tmp_qtm"]).ToString("###,##0.00") + "m";
                    //    gfx.DrawString(s, font4, XBrushes.Black, Y + 470, X + 23, frmDX);

                    //    s = ((decimal)t.Rows[i]["tmp_mav"]).ToString("##0.00") + "%";
                    //    gfx.DrawString(s, font1, XBrushes.Black, Y + 540, X + 23, frmDX);
                    //}
                }

                X += 40;
                //break;
            }

            X += 15;

            s = "Totali    " + new string(' ', 17) + (dImp).ToString("####,##0.00") + new string(' ', 7) + dIva.ToString("####,##0.00") + new string(' ', 5) + (dImp + dIva).ToString("####,##0.00") + "  ";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 10, X, XStringFormats.Default);

            //if (strTip == "ven_rep")
            //{
            //    s = decSco.ToString("####,##0");
            //    gfx.DrawString(s, font1, XBrushes.Black, Y + 355, X, XStringFormats.Default);

            //    s = decQta.ToString("####,##0");
            //    gfx.DrawString(s, font1, XBrushes.Black, Y + 435, X, XStringFormats.Default);

            //    s = (dTot / decSco).ToString("###,##0.00") + "m";
            //    gfx.DrawString(s, font4, XBrushes.Black, Y + 376, X + 13, frmDX);

            //    s = decMar.ToString("###,##0.00") + "%"; ;
            //    gfx.DrawString(s, font1, XBrushes.Black, Y + 540, X + 3, frmDX);
            //}
            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\VenStaIva_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch (Exception ex)
            {
                //MessageBox.Show("File di stampa già aperto!");
                Console.WriteLine("");
            }
        }

    }
}
