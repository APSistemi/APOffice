using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace APOffice
{
    public partial class frmGesDocPrint : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        clsResize _form_resize;

        public DataSet _dasGen = new DataSet();

        private string _strFrmPar = "";

        public string _strMovFat = "";
        public string _strCauTpd = "";
        public string _strTipCod = "";
        public Boolean _bolMftPvi = false;           //Prezzo di vendita ivato
        public string _strQta;
        public string _strSuf = "";
        public string _strDes = "";
        public string _strPrnPeso = "";
        public string _strDocTip = "";
        public string _strCfo = "";

        private string _strDocX2 = "";

        private string _strConSql = "";
        private string _strPar022DocRaggruppArticoli = "";
        private string _strPar034ParamFattureDiff = "";

        public frmGesDocPrint()
        {
            InitializeComponent();

            _form_resize = new clsResize(this);
            this.Load += _Load;
            this.Resize += _Resize;

            new clsGesGraph().SetGraph(this, 0);
        }
        private void _Load(object sender, EventArgs e)
        {
            _form_resize._get_initial_size();
        }
        private void _Resize(object sender, EventArgs e)
        {
            _form_resize._resize();
        }
        private void frmGesDocPrint_Load(object sender, EventArgs e)
        {
            Video();

            _strConSql = _clsFun.ConSql("");

            _strDocX2 = _clsFun.ParGet(clsDefine.enuParametri.Par025DocX2Pdf, _strConSql);
            _strPar022DocRaggruppArticoli = _clsFun.ParGet(clsDefine.enuParametri.Par022DocRaggruppArticoli, _strConSql);
            _strPar034ParamFattureDiff = _clsFun.ParGet(clsDefine.enuParametri.Par034ParamFattureDiff, _strConSql);

            lblCopie.Visible = false;
            chkCopie.Visible = false;
            if (_strDocX2 == "S")
            {
                lblCopie.Visible = true;
                chkCopie.Visible = true;
                chkCopie.Checked = true;
            }

            if (_strMovFat != "M")
            {
                lblPrezzi.Visible = false;
                chkPrezzi.Visible = false;
            }

            txtNumCol.Text = _strQta;
            FillTabs();
        }
        private void Video()
        {
            _form_resize._get_initial_size();
            _strFrmPar = _clsQry.ParForm(this, "R", "");

            this.CenterToScreen();
        }
        private void frmGesDocPrint_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void Esci()
        {

            string sCau = cmbCauTra.SelectedValue.ToString();
            string sAsp = cmbAspBen.SelectedValue.ToString();

            string s = sCau + "-" + sAsp;

            _clsQry.ParForm(this, "W", s);
            this.Close();
        }

        private void FillTabs()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabCauTrasporto";
            s = "SELECT * FROM TabCauTrasporto WHERE tab_ann=0";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = _strTipCod;
            t.Rows.InsertAt(x, 0);
            cmbCauTra.DataSource = t;
            cmbCauTra.DisplayMember = "tab_des";
            cmbCauTra.ValueMember = "tab_cod";
            cmbCauTra.SelectedValue = TipDoc("CAU", _strFrmPar);

            p = "TabAspettoBeni";
            s = "SELECT * FROM " + p;
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "";
            t.Rows.InsertAt(x, 0);
            cmbAspBen.DataSource = t;
            cmbAspBen.DisplayMember = "tab_des";
            cmbAspBen.ValueMember = "tab_cod";
            cmbAspBen.SelectedValue = TipDoc("ASP", _strFrmPar);
        }

        private string TipDoc(string strTip, string strPar)
        {
            string sRes = "";
            if (strPar != "")
            {
                string[] a = strPar.Split('-');
                if (strTip == "CAU" && a.Length > 0)
                    sRes = a[0];
                if (strTip == "ASP" && a.Length > 1)
                    sRes = a[1];
            }
            return sRes;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            string s = "";

            if (_strMovFat == "F" && _strCauTpd == "FA")
                PdfFatAcq();
            else if (_strMovFat == "F" && ((_strCauTpd.Length > 0 && _strCauTpd.Substring(0,1)== "F") || _strCauTpd == "PR"))
                PdfFatVen();
            else if (_strMovFat == "M" || _strMovFat == "N")
            {

                s = _clsFun.ParGet(clsDefine.enuParametri.Par024DecimalPesoPrezzo, _strConSql);
                int iDecimalPesoPrezzo = 2;
                if (_clsFun.Numerico(s))
                    iDecimalPesoPrezzo = Convert.ToInt16(s); 
                clsGenPdfMovVen cls = new clsGenPdfMovVen();
                cls._strConSql = _strConSql;
                cls._dasGen = _dasGen;
                cls._strTipDoc = _strTipCod;
                cls._strAspBen = cmbAspBen.Text;
                cls._strCauTra = cmbCauTra.Text;
                cls._strNumCol = txtNumCol.Text;
                cls._strDocNot = txtDocNot.Text;
                cls._strCauTpd = _strCauTpd;
                cls._strMovFat = _strMovFat;
                cls._strDes = _strDes;
                cls._strCfo = _strCfo;
                cls._str022DocRaggruppArticoli = _strPar022DocRaggruppArticoli;
                cls._strPrnPeso = _strPrnPeso;
                cls._strDocTip = _strDocTip;
                cls._intPar024DecimalPesoPrezzo = iDecimalPesoPrezzo;
                cls._bolDocX2 = chkCopie.Checked;
                cls._bolPrezzi = chkPrezzi.Checked;
                string sMsg = cls.PrnPdfMovVen();
                if (sMsg != "")
                    MessageBox.Show(sMsg);
                cls = null;
            }

            //s = _clsFun.ParGet(clsDefine.enuParametri.Par039DivDocumenti, _strConSql);
            //if (s.Length > 1 && s.Substring(0, 1) == "S")
            //    DivDocs(s);
        }

        private void PdfFatAcq()
        {
            string s = "";
            int iRows = 37;
            int iRow = 0;

            DataTable tTes = _dasGen.Tables["DocTes"];
            DataTable t = _dasGen.Tables["GesMovimenti"];
            DataTable tIva = _dasGen.Tables["TabIva"];

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Acquisti";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 10, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 9, XFontStyle.Regular);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);

            double Y = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

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
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " ";
                    s += "Documento di acquisto di " + ((string)tTes.Rows[0]["tmp_rag"]).Trim();
                    s += " - Numero " + (string)tTes.Rows[0]["tmp_ndo"];
                    s += " del " + ((DateTime)tTes.Rows[0]["tmp_ddo"]).ToString("dd/MM/yyyy");
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 20;

                    s = "Codice";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = "Descrizione";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 50, X, XStringFormats.Default);

                    s = "IVA";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 320, X, XStringFormats.Default);

                    s = "Qta";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 360, X, XStringFormats.Default);

                    s = "Costo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 400, X, XStringFormats.Default);

                    s = "Importo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 450, X, XStringFormats.Default);

                    X += 15;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = (string)t.Rows[i]["mov_art"];
                    gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["mov_ard"];
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 50, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["mov_iva"];
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 320, X, XStringFormats.Default);

                    //s = (string)t.Rows[i]["tab_des"];
                    //gfx.DrawString(s, font2, XBrushes.Black, Y + 180, X, XStringFormats.Default);

                    s = ((decimal)t.Rows[i]["mov_qta"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 380, X + 3, frmDX);

                    s = ((decimal)t.Rows[i]["mov_cos"]).ToString("####,##0.0000");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 430, X + 3, frmDX);

                    s = ((decimal)t.Rows[i]["mov_imp"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 490, X + 3, frmDX);

                    //break;
                }

                X += 20;
                //break;

            }

            X += 15;

            /****/
            t = tIva.Copy();

            for (int i = 0; i <= t.Rows.Count - 1; i++)
            {
                if (iRow > iRows || iRow == 0 || i == 0)
                {
                    if (iRow > iRows)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                        X = 20;
                        iRow = 0;

                    }
                    s = "Codice";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = "Descrizione";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 50, X, XStringFormats.Default);

                    s = "Imponibile";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 120, X, XStringFormats.Default);

                    s = "IVA";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 210, X, XStringFormats.Default);

                    s = "Totale";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 255, X, XStringFormats.Default);

                    //s = "Importo";
                    //gfx.DrawString(s, font1, XBrushes.Black, Y + 450, X, XStringFormats.Default);

                    X += 10;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = (string)t.Rows[i]["iva_cod"];
                    gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["iva_des"];
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 50, X, XStringFormats.Default);

                    s = ((decimal)t.Rows[i]["iva_imp"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 180, X + 3, frmDX);

                    s = ((decimal)t.Rows[i]["iva_iva"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 230, X + 3, frmDX);

                    s = ((decimal)t.Rows[i]["iva_tot"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 290, X + 3, frmDX);
                    //break;
                }

                X += 10;
                //break;

            }

            /****/


            //s = "Saldo    " + txtTotTot.Text;
            //gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = _clsDef.PATHPDF + "FatAcq_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch (Exception ex)
            {
                MessageBox.Show("File di stampa già aperto!", ex.Message);
            }
        }

        private void PdfFatVen()
        {
            string s = "";
            decimal d = 0;
            DataRow[] j;
            //int iRows = 20;
            int iRows = 21;
            //int iRowDoc = 8;
            int iRow = 0;
            int iPag = 0;

            decimal dSco = 0;

            s = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
            if (s == "")
                s = "001";
            DataTable tCnf = new clsQuery().ConfAzienda(s);

            DataTable tTes = _dasGen.Tables["DocTes"];
            DataTable tTmp = _dasGen.Tables["GesMovimenti"].Copy();
            DataTable tIva = _dasGen.Tables["TabIva"];
            DataTable tNot = _dasGen.Tables["TabNote"];

            DataTable t = tTmp.Clone();

            if (_strPar022DocRaggruppArticoli == "S")
            {
                foreach (DataRow y in tTmp.Rows)
                {
                    if (!(Boolean)y["mov_ann"])
                    {
                        j = t.Select("mov_art='" + (string)y["mov_art"] + "' AND mov_nmo='" + (string)y["mov_nmo"] + "'");
                        if (j.Length == 0 || ((string)y["mov_art"]).Trim() == "")
                            t.ImportRow(y);
                        else
                        {
                            j[0]["mov_qta"] = (decimal)j[0]["mov_qta"] + (decimal)y["mov_qta"];
                            j[0]["mov_qkg"] = (decimal)j[0]["mov_qkg"] + (decimal)y["mov_qkg"];
                            j[0]["mov_imp"] = (decimal)j[0]["mov_imp"] + (decimal)y["mov_imp"];
                        }
                    }
                }
            }
            else
            {
                foreach (DataRow y in tTmp.Rows)
                {
                    if (!(Boolean)y["mov_ann"])
                        t.ImportRow(y);
                }
            }

            //Calcolo DDT da pagare
            DataTable tMotDaPagare = new DataTable();
            if (_strPar034ParamFattureDiff != "" && _strPar034ParamFattureDiff.Substring(0, 1) == "S")
            {
                string sYea = (string)t.Rows[0]["mov_yfa"];
                string sNfa = (string)t.Rows[0]["mov_nfa"];
                tMotDaPagare = _clsQry.DdtStato(sYea, sNfa, _clsDef.STADAC);
            }

            //foreach (DataRow y in tTmp.Rows)
            //{
            //    if (!(Boolean)y["mov_ann"])
            //        t.ImportRow(y);
            //}

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Fattura vendita";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 10, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 9, XFontStyle.Regular);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);
            XFont font4 = new XFont("Times New Roman", 28, XFontStyle.Bold);
            XFont font5 = new XFont("Times New Roman", 12, XFontStyle.Bold);
            XFont font6 = new XFont("Times New Roman", 10, XFontStyle.Bold);
            XFont font7 = new XFont("Times New Roman", 15, XFontStyle.Bold);
            XFont font8 = new XFont("Times New Roman", 6, XFontStyle.Regular);
            XFont font9 = new XFont("Times New Roman", 9, XFontStyle.Regular);

            double Y = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            for (int i = 0; i <= t.Rows.Count - 1; i++)
            {
                if (iRow > iRows || iRow == 0)
                {
                    if (iRow > 0)
                    {
                        PdfFatVenPiede(X, Y, frmDX, gfx, page, font1, font2, font5, font6, font8, font7, null, dSco, iRow, pd.PageCount);

                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }

                    X = PdfFatVenTestata(tTes, iRow, X, Y, pd, gfx, page, font1, font5, font6);

                    iRow = 0;
                    iPag ++;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    Boolean bCommento = false;
                    if (((string)t.Rows[i]["mov_iva"]).Trim() == "")
                        bCommento = true;

                    if (!bCommento)
                    {
                        s = (string)t.Rows[i]["mov_art"];
                        gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);
                    }

                    if ((string)t.Rows[i]["mov_art"] == "0157805")
                        Console.WriteLine("xxxxxx");

                    //s = (string)t.Rows[i]["mov_ard"];
                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["mov_ard"]))
                        s = (string)t.Rows[i]["mov_ard"];
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 40, X, XStringFormats.Default);
                    if (!bCommento)
                    {
                        s = "";
                        if ((decimal)t.Rows[i]["mov_qta"] > 0)
                        {
                            s = "NR";
                            if (!DBNull.Value.Equals(t.Rows[i]["mov_umi"]))
                                s = (string)t.Rows[i]["mov_umi"];
                            gfx.DrawString(s, font1, XBrushes.Black, Y + 330, X, XStringFormats.Default);
                        }

                        if ((decimal)t.Rows[i]["mov_qkg"] > 0)
                        {
                            s = ((decimal)t.Rows[i]["mov_qkg"]).ToString("####,##0.00");
                            gfx.DrawString(s, font1, XBrushes.Black, Y + 385, X + 3, frmDX);
                        }
                        else if ((decimal)t.Rows[i]["mov_qta"] > 0)
                        {
                            s = ((decimal)t.Rows[i]["mov_qta"]).ToString("####,##0.00");
                            gfx.DrawString(s, font1, XBrushes.Black, Y + 385, X + 3, frmDX);
                        }

                        if( (decimal)t.Rows[i]["mov_prv"] != 0)
                        {
                            d = (decimal)t.Rows[i]["mov_prv"];
                            j = tIva.Select("iva_cod='" + (string)t.Rows[i]["mov_iva"] + "'");
                            if (j.Length > 0 && _bolMftPvi)
                                d = _clsFun.MenoIva((decimal)t.Rows[i]["mov_prv"], (decimal)j[0]["iva_ali"]);
                            s = d.ToString("####,##0.00");
                            gfx.DrawString(s, font1, XBrushes.Black, Y + 443, X + 3, frmDX);

                            s = "";
                            if (!DBNull.Value.Equals(t.Rows[i]["mov_sco"]))
                                s = (string)t.Rows[i]["mov_sco"];
                            gfx.DrawString(s, font1, XBrushes.Black, Y + 450, X, XStringFormats.Default);

                            d = (decimal)t.Rows[i]["mov_imp"];
                            j = tIva.Select("iva_cod='" + (string)t.Rows[i]["mov_iva"] + "'");
                            if (j.Length > 0 && _bolMftPvi)
                            {
                                d = _clsFun.MenoIva(d, (decimal)j[0]["iva_ali"]);
                            }
                            s = d.ToString("####,##0.00");
                            gfx.DrawString(s, font1, XBrushes.Black, Y + 550, X + 3, frmDX);

                            s = (string)t.Rows[i]["mov_iva"];
                            gfx.DrawString(s, font1, XBrushes.Black, Y + 560, X, XStringFormats.Default);

                            s = "";
                            if (!DBNull.Value.Equals(t.Rows[i]["mov_sco"]))
                            {
                                s = (string)t.Rows[i]["mov_sco"];
                                if (s != "")
                                {
                                    if (s.Substring(0, 1) == "V")
                                    {
                                        d = Convert.ToDecimal(s.Substring(1));
                                        dSco += d;
                                    }
                                    else if(_clsFun.Numerico(s))
                                    {
                                        d = ((decimal)t.Rows[i]["mov_qta"] * (decimal)t.Rows[i]["mov_prv"]) - (decimal)t.Rows[i]["mov_imp"];
                                        dSco += d;
                                    }
                                }
                            }

                            if (tMotDaPagare.Rows.Count > 0)
                            {
                                s = (string)t.Rows[i]["mov_nmo"];

                                j = tMotDaPagare.Select("mot_nmo='" + s + "'");
                                if(j.Length > 0)
                                {
                                    d = (decimal)t.Rows[i]["mov_imp"];
                                    
                                    DataRow[] jj = tIva.Select("iva_cod='" + (string)t.Rows[i]["mov_iva"] + "'");
                                    if (jj.Length > 0)
                                        d = _clsFun.PiuIva((decimal)t.Rows[i]["mov_imp"], (decimal)jj[0]["iva_ali"]);

                                    j[0]["TmpImp"] = (decimal)j[0]["TmpImp"] + d;
                                }
                            }

                        }
                    }
                    //break;
                }

                X += 20;
                //break;
            }

            if (tMotDaPagare.Rows.Count > 0)
            {
                foreach (DataRow y in tMotDaPagare.Rows)
                {
                    DataRow x = tNot.NewRow();
                    x["tab_txt"] = "DDT " + (string)y["mot_ndo"] + " del " + ((DateTime)y["mot_ddo"]).ToString("dd/MM/yyy") + " con euro " + ((decimal)y["TmpImp"]).ToString("#0.00") + " da saldare";
                    tNot.Rows.Add(x);
                }
            }

            if (tNot.Rows.Count > 0)
            {
                int i = iRows - iRow - (tNot.Rows.Count / 2);

                if (i < 0)
                {
                    PdfFatVenPiede(X, Y, frmDX, gfx, page, font1, font2, font5, font6, font8, font7, null, dSco, iRow, pd.PageCount);

                    page = pd.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    X = 20;

                    X = PdfFatVenTestata(tTes, iRow, X, Y, pd, gfx, page, font1, font5, font6);

                    iRow = 0;
                    i = iRows - (tNot.Rows.Count / 2);
                }
                else
                    i = iRows - iRow - (tNot.Rows.Count / 2);

                for (int ii = 0; ii < i; ii++)
                {
                    X += 20;
                }

                foreach (DataRow y in tNot.Rows)
                {
                    s = (string)y["tab_txt"];
                    gfx.DrawString(s, font9, XBrushes.Black, Y + 10, X, XStringFormats.Default);
                    X += 10;
                    iRow++;
                }
            }

            PdfFatVenPiede(X, Y, frmDX, gfx, page, font1, font2, font5, font6, font8, font7, tIva, dSco, iRow, pd.PageCount);

            string sFil = "";

            try
            {
                s = ((string)tTes.Rows[0]["tmp_rag"]).Trim().Replace(".","");

                if(s.Length > 10)
                    s = s.Substring(0, 10);

                s = _clsFun.CtrlCrt(s, "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM");

                sFil = _clsDef.PATHPDF + ((DateTime)tTes.Rows[0]["tmp_ddo"]).ToString("yyyyMMdd") + "_";
                if (((string)tTes.Rows[0]["tmp_ndo"]).Trim() != "")
                {
                    if (((string)tTes.Rows[0]["tmp_ndo"]).Length >= 10)
                        sFil += ((string)tTes.Rows[0]["tmp_ndo"]).Substring(7, 3);
                    else
                        sFil += ((string)tTes.Rows[0]["tmp_ndo"]).Trim();
                    sFil += "_" + s + ".pdf";
                }
                else
                    sFil += s + ".pdf";

                //sFil = _clsDef.PATHPDF + ((DateTime)tTes.Rows[0]["tmp_ddo"]).ToString("yyyyMMdd") + "_doc.pdf";

                pd.Save(sFil);
                // ...and start a viewer.

                //if (_strDocX2 == "S" && iPag == 1)
                if (chkCopie.Checked)
                    sFil = new clsGenPdfMovVen().PdfDuplica(sFil);

                Process.Start(sFil);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " (" + sFil + ")", "Errore sul file!");
            }
        }

        //private double PdfFatVenTestata(DataTable tTes, int iRow, double X, double Y, PdfDocument pd, XGraphics gfx, PdfPage page, XFont font1, XFont font5, XFont font6)


        private void PdfFatVenPiede(double X, double Y, XStringFormat frmDX, XGraphics gfx, PdfPage page, XFont font1, XFont font2, XFont font5, XFont font6, XFont font8, XFont font7, DataTable tIva, decimal dSco, int iRow, int iPage)
        {
            string s = "";
            X = 700;
            gfx.DrawLine(XPens.Black, 0, X, page.Width, X);

            X += 10;
            s = "TOTALE DOCUMENTO EURO ";
            gfx.DrawString(s, font6, XBrushes.Black, Y + 10, X + 5, XStringFormats.Default);

            if (tIva != null)
            {
                s = ((decimal)tIva.Rows[tIva.Rows.Count - 1]["iva_tot"]).ToString("####,##0.00");
                gfx.DrawString(s, font7, XBrushes.Black, Y + 220, X + 9, frmDX);
            }
            if (dSco > 0)
            {
                s = "TOTALE SCONTI EURO ";
                gfx.DrawString(s, font6, XBrushes.Black, Y + 260, X + 5, XStringFormats.Default);

                s = (dSco).ToString("####,##0.00");
                gfx.DrawString(s, font7, XBrushes.Black, Y + 450, X + 9, frmDX);
            }

            X += 10;
            gfx.DrawRectangle(XPens.Black, 0, X, 250, 100);

            X += 15;

            decimal dTot = 0;
            //DataTable t = tIva.Copy();

            /*
             * da verificare di non andare su pagina nuova
             * 
             * 
            if (iRow > iRows+1)
            {
                page = pd.AddPage();
                gfx = XGraphics.FromPdfPage(page);
                X = 20;
                iRow = 0;
            }
            */

            //s = "Codice";
            //gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

            s = "Descrizione";
            gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

            //X += 10;
            //gfx.DrawLine(XPens.Black, 0, X, 250, X);

            s = "Imponibile";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 70, X, XStringFormats.Default);

            s = "IVA";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 160, X, XStringFormats.Default);

            s = "Totale";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 195, X, XStringFormats.Default);

            //s = "Importo";
            //gfx.DrawString(s, font1, XBrushes.Black, Y + 450, X, XStringFormats.Default);

            X += 10;
            gfx.DrawLine(XPens.Black, 0, X, 250, X);
            X += 6;

            if (tIva != null)
            {
                Console.WriteLine("aaaaaaaaaaa");


                for (int i = 0; i <= tIva.Rows.Count - 1; i++)
                {
                    iRow++;
                    X += 0;

                    if (tIva.Rows.Count - 1 == i)
                    {
                        gfx.DrawLine(XPens.Black, 0, X, 250, X);
                        X += 2;
                    }
                    X += 6;

                    if (tIva.Rows.Count - 1 >= i)
                    {
                        //s = (string)t.Rows[i]["iva_cod"];
                        //gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                        s = (string)tIva.Rows[i]["iva_des"];
                        gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                        s = ((decimal)tIva.Rows[i]["iva_imp"]).ToString("####,##0.00");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 120, X + 3, frmDX);

                        s = ((decimal)tIva.Rows[i]["iva_iva"]).ToString("####,##0.00");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 180, X + 3, frmDX);

                        s = ((decimal)tIva.Rows[i]["iva_tot"]).ToString("####,##0.00");
                        gfx.DrawString(s, font6, XBrushes.Black, Y + 240, X + 3, frmDX);
                        dTot = (decimal)tIva.Rows[i]["iva_tot"];
                        //break;
                        //gfx.DrawRectangle(XPens.Black, 0, X, 250, X);
                    }

                    X += 6;
                    //break;
                }
            }

            Console.WriteLine("aaaaaaaaa");

            X = 720;
            Y = 250;
            gfx.DrawRectangle(XPens.Black, Y, X, 350, 100);

            s = "CAUSALE DEL TRASPORTO                                   ASPETTO ESTERIONE DEI BENI                                                              N.COLLI";
            gfx.DrawString(s, font8, XBrushes.Black, Y + 3, X + 7, XStringFormats.Default);

            if (tIva != null)
            {
                s = cmbCauTra.Text;
                gfx.DrawString(s, font5, XBrushes.Black, Y + 3, X + 20, XStringFormats.Default);

                s = cmbAspBen.Text;
                gfx.DrawString(s, font5, XBrushes.Black, Y + 135, X + 20, XStringFormats.Default);

                s = txtNumCol.Text;
                gfx.DrawString(s, font5, XBrushes.Black, Y + 325, X + 20, XStringFormats.Default);
            }

            X += 25;
            gfx.DrawLine(XPens.Black, Y, X, 595, X);

            s = "ORA RITIRO                                                                  DATA RITIRO";
            gfx.DrawString(s, font8, XBrushes.Black, Y + 3, X + 7, XStringFormats.Default);

            X += 25;
            gfx.DrawLine(XPens.Black, Y, X, 595, X);

            s = "VETTORE                                                                     FIRMA DEL CONDUCENTE                                    FIRMA DEL DESTINATARIO";
            gfx.DrawString(s, font8, XBrushes.Black, Y + 3, X + 7, XStringFormats.Default);

            X += 25;
            //gfx.DrawLine(XPens.Black, Y, X, 595, X);
            X += 25;
            Y = 5;

            gfx.DrawLine(XPens.Black, Y, X, 595, X);

            s = "NOTE";
            gfx.DrawString(s, font8, XBrushes.Black, Y + 3, X + 7, XStringFormats.Default);

            s = txtDocNot.Text;
            gfx.DrawString(s, font2, XBrushes.Black, Y + 3, X + 15, XStringFormats.Default);

            s = "PAG. " + iPage.ToString();
            if (tIva == null)
                s += "  SEGUE ===> ";
            gfx.DrawString(s, font6, XBrushes.Black, Y + 570, X + 0, frmDX);
        }

        private double PdfFatVenTestata(DataTable tTes, int iRow, double X, double Y, PdfDocument pd, XGraphics gfx, PdfPage page, XFont font1, XFont font5, XFont font6)
        {
            string s = "";

            X = 20;
            gfx.DrawLine(XPens.Black, 0, X, page.Width, X);

            XImage img = XImage.FromFile(_clsDef.PATHLOGHI + "LogoFattura.png");
            gfx.DrawImage(img, X, Y + 20, 430, 110);

            iRow = 0;

            X += 105;
            X += 15;
            gfx.DrawLine(XPens.Black, 0, X, page.Width, X);

            X += 20;
            s = "Spett.le ";
            gfx.DrawString(s, font6, XBrushes.Black, Y + 250, X-5, XStringFormats.Default);

            X += 5;

            s = (string)tTes.Rows[0]["tmp_ndo"];
            if (_clsFun.Numerico(s))
                s = Convert.ToInt64(s).ToString();
            if (_strSuf != "")
                s = _strTipCod + " numero " + s + _strSuf;
            else if (_strTipCod.ToUpper() == "PREVENTIVO" || _strTipCod.ToUpper() == "PROFORMA")
                s = _strTipCod + "/" + ((DateTime)tTes.Rows[0]["tmp_ddo"]).Year.ToString();
            else
                s = _strTipCod + " numero " + s + "/" + ((DateTime)tTes.Rows[0]["tmp_ddo"]).Year.ToString();
            gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

            X += 15;
            s = "Data                    " + ((DateTime)tTes.Rows[0]["tmp_ddo"]).ToString("dd/MM/yyyy");
            gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

            X += 0;
            s = (string)tTes.Rows[0]["tmp_rag"];
            gfx.DrawString(s, font5, XBrushes.Black, Y + 250, X - 11, XStringFormats.Default);

            X += 0;
            s = (string)tTes.Rows[0]["tmp_ra2"];
            gfx.DrawString(s, font5, XBrushes.Black, Y + 250, X - 0, XStringFormats.Default);

            X += 15;
            s = "Partita IVA       " + (string)tTes.Rows[0]["tmp_piv"];
            gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

            X += 0;
            s = ((string)tTes.Rows[0]["tmp_ind"]) + " " + (string)tTes.Rows[0]["tmp_ncv"];
            gfx.DrawString(s, font5, XBrushes.Black, Y + 250, X - 2, XStringFormats.Default);

            X += 15;
            s = "Codice fiscale    " + (string)tTes.Rows[0]["tmp_cfi"];
            gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

            X += 0;
            s = (string)tTes.Rows[0]["tmp_loc"];
            gfx.DrawString(s, font5, XBrushes.Black, Y + 250, X - 2, XStringFormats.Default);

            X += 15;
            s = "Pagamento         " + (string)tTes.Rows[0]["tmp_tpg"];
            gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

            if (_strDes != "")
            {
                //X += 15;
                s = "Dest.: " + _strDes;
                gfx.DrawString(s, font6, XBrushes.Black, Y + 180, X, XStringFormats.Default);
            }

            X += 5;
            gfx.DrawLine(XPens.Black, 0, X, page.Width, X);

            X += 20;
            s = "Codice";
            gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

            s = "Descrizione";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 40, X, XStringFormats.Default);

            s = "UM";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 330, X, XStringFormats.Default);

            s = "Qta";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 365, X, XStringFormats.Default);

            s = "Prezzo";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 405, X, XStringFormats.Default);

            s = "Sconti";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 450, X, XStringFormats.Default);

            s = "Importo";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 505, X, XStringFormats.Default);

            s = "IVA";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 560, X, XStringFormats.Default);

            X += 15;

            return X;
        }

        private void DivDocs(string strPar)
        {
            //DataTable tTes = _dasGen.Tables["DocTes"];
            //DataTable tTmp = _dasGen.Tables["GesMovimenti"].Copy();
            //DataTable tIva = _dasGen.Tables["TabIva"];

            //string[] a = strPar.Split('-');

            //string sFil = a[1];

            //if(_strMovFat == "M" &&  _strDocTip == "D") 
            //    sFil += "DDT-";
            //else if (_strMovFat == "F")
            //    sFil += "FAT-";
            //else 
            //    sFil += "MOV-";

            //sFil += ((DateTime)tTes.Rows[0]["tmp_ddo"]).ToString("yyyyMMdd") + "-";
            //sFil += (string)tTes.Rows[0]["tmp_ndo"] + ".csv";


            //StreamWriter sw = new StreamWriter(sFil);

            //string sRig = "";






            //sw.Write(sRig + "\r\n");
            
            
            //((TextWriter)sw).Flush();
            //sw.Close();
            //sw.Dispose();






        }
    }
}
