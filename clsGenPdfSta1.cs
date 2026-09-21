using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using BarcodeLib;

namespace APOffice
{
    class clsGenPdfSta1
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public DateTime _dayIni = new DateTime();
        public DateTime _dayFin = new DateTime();

        public void PrnPdfRep(string strTip, DataTable tabTab, decimal decSco, decimal decQta, string strNeg, decimal decMar)
        {
            string s = "";
            int iRows = 18;
            int iRow = 0;
            decimal dTot = 0;
            decimal dTls = 0;
            decimal dSco = 0;
            decimal dQta = 0;
            decimal dInc = 0;
            decimal dAcq = 0;
            decimal dVen = 0;

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
                    s += " dal " + _dayIni.ToString("dd/MM/yy") + " al " + _dayFin.ToString("dd/MM/yy");

                    if (strNeg != "")
                        s += " - PV " + strNeg;

                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    s = "Reparto";
                    if (strTip == "ven_iva")
                        s = "IVA";

                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = "Venduto";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 220, X, XStringFormats.Default);

                    s = "Lordo sconti";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 185, X + 10, XStringFormats.Default);

                    s = "Venduto";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 290, X, XStringFormats.Default);

                    if (strTip == "ven_rep")
                    {
                        s = "Scontrini";
                        gfx.DrawString(s, font2, XBrushes.Black, Y + 344, X, XStringFormats.Default);

                        s = "Pezzi";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 430, X, XStringFormats.Default);

                        s = "Marg.";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 500, X, XStringFormats.Default);
                    }
                    X += 35;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = (string)t.Rows[i]["tmp_cod"] + " " + (string)t.Rows[i]["tmp_des"];
                    gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = ((decimal)t.Rows[i]["tmp_ils"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 260, X + 3, frmDX);
                    //dTot += (decimal)t.Rows[i]["tmp_ils"];

                    s = ((decimal)t.Rows[i]["tmp_imp"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 340, X + 3, frmDX);
                    dTot += (decimal)t.Rows[i]["tmp_imp"];

                    if (strTip == "ven_rep")
                    {
                        s = ((decimal)t.Rows[i]["tmp_inc"]).ToString("####,##0.00") + "%";
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 340, X + 13, frmDX);
                        dInc += (decimal)t.Rows[i]["tmp_inc"];

                        s = ((decimal)t.Rows[i]["tmp_sco"]).ToString("###,##0");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 400, X + 3, frmDX);
                        dSco += (decimal)t.Rows[i]["tmp_sco"];

                        s = ((decimal)t.Rows[i]["tmp_sci"]).ToString("###,##0.00") + "%";
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 400, X + 13, frmDX);

                        s = ((decimal)t.Rows[i]["tmp_scm"]).ToString("###,##0.00") + "m";
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 400, X + 23, frmDX);


                        s = ((decimal)t.Rows[i]["tmp_qta"]).ToString("###,##0");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 470, X + 3, frmDX);
                        dQta += (decimal)t.Rows[i]["tmp_qta"];

                        s = ((decimal)t.Rows[i]["tmp_qti"]).ToString("###,##0.00") + "%";
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 470, X + 13, frmDX);

                        s = ((decimal)t.Rows[i]["tmp_qtm"]).ToString("###,##0.00") + "m";
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 470, X + 23, frmDX);

                        s = ((decimal)t.Rows[i]["tmp_mav"]).ToString("##0.00") + "%";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 540, X + 23, frmDX);
                    }
                }

                X += 40;
                //break;
            }

            //if (iRow >= t.Rows.Count)
            if (iRow >= iRows)
            {
                page = pd.AddPage();
                gfx = XGraphics.FromPdfPage(page);


                X = 20;
                iRow = 0;
            }

            X += 15;

            s = "Totale    " + dTot.ToString("####,##0.00");
            gfx.DrawString(s, font1, XBrushes.Black, Y + 150, X, XStringFormats.Default);

            if (strTip == "ven_rep")
            {
                s = decSco.ToString("####,##0");
                gfx.DrawString(s, font1, XBrushes.Black, Y + 355, X, XStringFormats.Default);

                s = decQta.ToString("####,##0");
                gfx.DrawString(s, font1, XBrushes.Black, Y + 435, X, XStringFormats.Default);

                s = (dTot / decSco).ToString("###,##0.00") + "m";
                gfx.DrawString(s, font4, XBrushes.Black, Y + 376, X + 13, frmDX);

                s = decMar.ToString("###,##0.00") + "%"; ;
                gfx.DrawString(s, font1, XBrushes.Black, Y + 540, X + 3, frmDX);
            }
            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\VenStaRep_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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

        public void PrnPdfEcr(DataTable tabTab, decimal decSco, decimal decQta, string strNeg, decimal decMar)
        {
            string s = "";
            int iRows = 18;
            int iRow = 0;
            decimal dTot = 0;
            decimal dTls = 0;
            decimal dSco = 0;
            decimal dQta = 0;
            decimal dInc = 0;
            decimal dAcq = 0;
            decimal dVen = 0;

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

            string sKey = "";

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
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA VENDUTO PER MERCEOLOGIA";
                    s += " dal " + _dayIni.ToString("dd/MM/yy") + " al " + _dayFin.ToString("dd/MM/yy");

                    if (strNeg != "")
                        s += " - PV " + strNeg;

                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    s = "Merceologia";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = "Venduto";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 220, X, XStringFormats.Default);

                    s = "Lordo sconti";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 185, X + 10, XStringFormats.Default);

                    s = "Venduto";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 290, X, XStringFormats.Default);

                    s = "Scontrini";
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 344, X, XStringFormats.Default);

                    s = "Pezzi";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 430, X, XStringFormats.Default);

                    s = "Marg.";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 500, X, XStringFormats.Default);
                    
                    X += 35;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = (string)t.Rows[i]["tmp_ec1"] + (string)t.Rows[i]["tmp_ec2"];

                    if(sKey != s)
                    {
                        sKey = s;
                        s = " ------ " + (string)t.Rows[i]["tmp_ec1"] + " " + (string)t.Rows[i]["tmp_ed1"] + " - " + (string)t.Rows[i]["tmp_ec2"] + " " + (string)t.Rows[i]["tmp_ed2"];
                        gfx.DrawString(s, font3, XBrushes.Black, Y, X-10, XStringFormats.Default);
                    }

                    s = (string)t.Rows[i]["tmp_ec3"] + " " + (string)t.Rows[i]["tmp_ed3"];
                    gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = ((decimal)t.Rows[i]["tmp_ils"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 260, X + 3, frmDX);
                    //dTot += (decimal)t.Rows[i]["tmp_ils"];

                    s = ((decimal)t.Rows[i]["tmp_imp"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 340, X + 3, frmDX);
                    dTot += (decimal)t.Rows[i]["tmp_imp"];

                    s = ((decimal)t.Rows[i]["tmp_inc"]).ToString("####,##0.00") + "%";
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 340, X + 13, frmDX);
                    dInc += (decimal)t.Rows[i]["tmp_inc"];

                    s = ((decimal)t.Rows[i]["tmp_sco"]).ToString("###,##0");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 400, X + 3, frmDX);
                    dSco += (decimal)t.Rows[i]["tmp_sco"];

                    s = ((decimal)t.Rows[i]["tmp_sci"]).ToString("###,##0.00") + "%";
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 400, X + 13, frmDX);

                    s = ((decimal)t.Rows[i]["tmp_scm"]).ToString("###,##0.00") + "m";
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 400, X + 23, frmDX);

                    s = ((decimal)t.Rows[i]["tmp_qta"]).ToString("###,##0");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 470, X + 3, frmDX);
                    dQta += (decimal)t.Rows[i]["tmp_qta"];

                    s = ((decimal)t.Rows[i]["tmp_qti"]).ToString("###,##0.00") + "%";
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 470, X + 13, frmDX);

                    s = ((decimal)t.Rows[i]["tmp_qtm"]).ToString("###,##0.00") + "m";
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 470, X + 23, frmDX);

                    s = ((decimal)t.Rows[i]["tmp_mav"]).ToString("##0.00") + "%";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 540, X + 23, frmDX);
                }

                X += 40;
                //break;
            }

            X += 15;

            s = "Totale    " + dTot.ToString("####,##0.00");
            gfx.DrawString(s, font1, XBrushes.Black, Y + 150, X, XStringFormats.Default);

            s = decSco.ToString("####,##0");
            gfx.DrawString(s, font1, XBrushes.Black, Y + 355, X, XStringFormats.Default);

            s = decQta.ToString("####,##0");
            gfx.DrawString(s, font1, XBrushes.Black, Y + 435, X, XStringFormats.Default);

            s = (dTot / decSco).ToString("###,##0.00") + "m";
            gfx.DrawString(s, font4, XBrushes.Black, Y + 376, X + 13, frmDX);

            s = decMar.ToString("###,##0.00") + "%"; ;
            gfx.DrawString(s, font1, XBrushes.Black, Y + 540, X + 3, frmDX);
           
            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\VenStaEcr_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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

        public void PrnPdfStaArt(DataTable tabTab)
        {
            string s = "";
            int iRows = 50;
            int iRow = 0;
            decimal d = 0;
            decimal dInc = 0;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);

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
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA VENDUTO PER ARTICOLO";
                    s += " dal " + _dayIni.ToString("dd/MM/yy") + " al " + _dayFin.ToString("dd/MM/yy");

                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    s = "Codice";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = "Descrizione";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 50, X, XStringFormats.Default);

                    s = "Q.tà";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 240, X, XStringFormats.Default);

                    s = "Costo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 300, X, XStringFormats.Default);

                    s = "P.Pubbl";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 350, X, XStringFormats.Default);

                    s = "Barcode";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 420, X, XStringFormats.Default);

                    X += 20;
                }

                iRow++;

                //x["tmp_art"] = y["ven_art"];
                //x["tmp_ard"] = y["ven_ard"];
                //x["tmp_ean"] = "";
                //x["tmp_pco"] = 0;
                //x["tmp_pve"] = 0;
                //x["tmp_cos"] = 0;
                //x["tmp_ven"] = 0;

                if (t.Rows.Count - 1 >= i)
                {
                    s = (string)t.Rows[i]["tmp_ard"];
                    if (s.Length > 30)
                        s = s.Substring(0, 30);

                    s = (string)t.Rows[i]["tmp_art"] + " " + s;
                    gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = ((decimal)t.Rows[i]["tmp_qta"]).ToString("####,##0.00");
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 270, X + 3, frmDX);

                    s = ((decimal)t.Rows[i]["tmp_pco"]).ToString("####,##0.00");
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 330, X + 3, frmDX);

                    s = ((decimal)t.Rows[i]["tmp_pve"]).ToString("####,##0.00");
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 390, X + 3, frmDX);

                    s = (string)t.Rows[i]["tmp_ean"];
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 420, X, XStringFormats.Default);
                }

                X += 15;
                //break;
            }

            X += 15;

            //s = "Totale    " + d.ToString("####,##0.00");
            //gfx.DrawString(s, font1, XBrushes.Black, Y + 280, X, XStringFormats.Default);

            //s = dInc.ToString("####,##0.00");
            //gfx.DrawString(s, font1, XBrushes.Black, Y + 450, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\VenStaArt_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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

        public void PrnPdfStaPos(DataTable tabTab, string strTotInc, string strTotDif, string strPrnPar)
        {
            string s = "";
            int iRows = 50;
            int iRow = 0;
            decimal d = 0;

            string[] a = strPrnPar.Split(';');

            string sDivUsr = a[0];
            string sDivPos = a[1];

            string sNeg = a[2];
            string sUsr = a[3];

            string sPos = "";
            if (a[3] != "")
                sPos = "0" + a[4];


            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Cassa";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);
            XFont font4 = new XFont("Courier new", 11, XFontStyle.Bold);
            XFont font5 = new XFont("Courier new", 11, XFontStyle.Regular);

            double Y = 5;
            double X = 20;

            double y = 0;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            DataTable t = tabTab.Copy();

            Console.WriteLine("xxx");

            //int iColPos = 0;
            //int iColUsr = 0;
            //int iColDes = 0;
            //int iColImp = 0;

            //if (sPos == "" || sDivPos == "S")
            //    iColUsr = iColPos + 50;
            //if (sUsr == "" || sDivUsr == "S")
            //    iColDes = iColUsr + 150;
            //iColImp = iColDes + 200;

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
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA VENDUTO PER CASSA";
                    s += " dal " + _dayIni.ToString("dd/MM/yy") + " al " + _dayFin.ToString("dd/MM/yy");
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    iRow = 0;

                    s = "";
                    if (sNeg != "")
                    {
                        X += 20;
                        s = "Negozio: " + sNeg;
                        gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);
                    }
                    if (sUsr != "")
                    {
                        X += 20;
                        s = "Utente: " + sUsr;
                        gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);
                    }
                    if (sPos != "")
                    {
                        X += 20;
                        s = "Cassa: " + sPos;
                        gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);
                    }

                    X += 30;

                    y = 0;

                    if (sPos != "" || sDivPos == "S")
                    {
                        s = "Pos";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + y, X, XStringFormats.Default);
                    }

                    if (sUsr != "" || sDivUsr == "S")
                    {
                        if (sPos != "" || sDivPos == "S")
                            y += 50;

                        s = "Utente";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + y - 20, X, XStringFormats.Default);

                        y += 60;
                    }

                    s = "Descrizione";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + y, X, XStringFormats.Default);

                    s = "Pagamenti";
                    gfx.DrawString(s, font4, XBrushes.Black, Y + y + 110, X, XStringFormats.Default);

                    y += 67;
                    s = "Versamenti";
                    gfx.DrawString(s, font4, XBrushes.Black, Y + y + 130, X, XStringFormats.Default);

                    y += 67;
                    s = "Prelievi";
                    gfx.DrawString(s, font4, XBrushes.Black, Y + y + 150, X, XStringFormats.Default);

                    y += 67;
                    s = "Saldo";
                    gfx.DrawString(s, font4, XBrushes.Black, Y + y + 200, X, XStringFormats.Default);

                    X += 20;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    y = 0;

                    if (sPos != "" || sDivPos == "S")
                    {
                        s = (string)t.Rows[i]["vet_pos"];
                        gfx.DrawString(s, font2, XBrushes.Black, Y + y, X, XStringFormats.Default);
                    }

                    if (sUsr != "" || sDivUsr == "S")
                    {
                        if (sPos != "" || sDivPos == "S")
                            y += 50;

                        s = (string)t.Rows[i]["VepUsr"];

                        if (s.Length > 10)
                            s = s.Substring(0, 10);

                        gfx.DrawString(s, font5, XBrushes.Black, Y + y-20, X, XStringFormats.Default);

                        y += 60;

                    }

                    if (i == 0)
                        Console.WriteLine("xxxx");

                    s = (string)t.Rows[i]["PagDes"];
                    gfx.DrawString(s, font2, XBrushes.Black, Y + y, X, XStringFormats.Default);

                    //s = ((decimal)t.Rows[i]["vep_imp"]).ToString("####,##0.00");
                    //gfx.DrawString(s, font2, XBrushes.Black, Y + iColImp, X + 3, frmDX);

                    s = ((decimal)t.Rows[i]["TotSco"]).ToString("####,##0.00");
                    gfx.DrawString(s, font2, XBrushes.Black, Y + y + 170, X + 3, frmDX);

                    y += 60;
                    s = ((decimal)t.Rows[i]["TotVer"]).ToString("####,##0.00");
                    gfx.DrawString(s, font2, XBrushes.Black, Y + y + 200, X + 3, frmDX);

                    y += 60;
                    s = ((decimal)t.Rows[i]["TotPre"]).ToString("####,##0.00");
                    gfx.DrawString(s, font2, XBrushes.Black, Y + y + 220, X + 3, frmDX);

                    y += 60;
                    s = ((decimal)t.Rows[i]["TotDif"]).ToString("####,##0.00");
                    gfx.DrawString(s, font2, XBrushes.Black, Y + y + 255, X + 3, frmDX);
                }

                X += 15;
            }

            X += 15;

            s = "Totale incassi                   ";
            gfx.DrawString(s, font4, XBrushes.Black, Y + 0, X, XStringFormats.Default);
            s = strTotInc; 
            gfx.DrawString(s, font1, XBrushes.Black, Y + 270, X+3, frmDX);

            s = "Saldo";
            gfx.DrawString(s, font4, XBrushes.Black, Y + 0, X + 30, XStringFormats.Default);
            s = strTotDif;
            gfx.DrawString(s, font1, XBrushes.Black, Y + 270, X + 30 + 4, frmDX);

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\VenPos_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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

        public string PrnPdfStaPos2(DataTable tabTab, string strTotInc, string strTotDif, string strPrnPar)
        {
            string s = "";
            string sMsg = "";
            int iRows = 50;
            int iRow = 0;
            decimal d = 0;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Cassa";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);
            XFont font4 = new XFont("Courier new", 11, XFontStyle.Bold);

            double Y = 5;
            double Yy = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            string[] a = strPrnPar.Split(';');
            string sNeg = a[1];
            string sUsr = a[2];

            string sPos = "";
            if(a[3] != "")
                sPos = "0" + a[3];

            DataTable t = tabTab.Clone();
            foreach(DataRow y in tabTab.Rows)
            {
                Boolean b = true;
                if (sNeg != "" && (string)y["vet_pos"] != sNeg)
                    b = false;
                if (sUsr != "" && (string)y["vet_usr"] != sUsr)
                    b = false;
                if (sPos != "" && (string)y["vet_pos"] != sPos)
                    b = false;
                if (b)
                    t.ImportRow(y);
            }

            Console.WriteLine("xxx");

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
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA VENDUTO PER CASSA";
                    s += " dal " + _dayIni.ToString("dd/MM/yy") + " al " + _dayFin.ToString("dd/MM/yy");

                    gfx.DrawString(s, font4, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    Yy = Y;
                    s = "Negozio";
                    gfx.DrawString(s, font4, XBrushes.Black, Yy, X, XStringFormats.Default);

                    Yy += 50;
                    s = "Utente";
                    gfx.DrawString(s, font4, XBrushes.Black, Yy, X, XStringFormats.Default);

                    Yy += 50;
                    s = "Cassa";
                    gfx.DrawString(s, font4, XBrushes.Black, Yy, X, XStringFormats.Default);

                    Yy += 50;
                    s = "Descrizione";
                    gfx.DrawString(s, font4, XBrushes.Black, Yy, X, XStringFormats.Default);

                    Yy += 100;
                    s = "Importo";
                    gfx.DrawString(s, font4, XBrushes.Black, Yy, X, XStringFormats.Default);

                    X += 20;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    Yy = Y;

                    s = (string)t.Rows[i]["VetNeg"];
                    gfx.DrawString(s, font2, XBrushes.Black, Yy, X, XStringFormats.Default);

                    Yy += 50;
                    s = (string)t.Rows[i]["VepUsr"];
                    gfx.DrawString(s, font2, XBrushes.Black, Yy, X, XStringFormats.Default);

                    Yy += 50;
                    s = (string)t.Rows[i]["vet_pos"];
                    gfx.DrawString(s, font2, XBrushes.Black, Yy, X, XStringFormats.Default);

                    Yy += 50;
                    s = (string)t.Rows[i]["PagDes"];
                    gfx.DrawString(s, font2, XBrushes.Black, Yy, X, XStringFormats.Default);

                    Yy += 145;
                    s = ((decimal)t.Rows[i]["vep_imp"]).ToString("####,##0.00");
                    gfx.DrawString(s, font2, XBrushes.Black, Yy, X + 3, frmDX);
                }

                X += 15;
            }

            if (iRow > 0)
            {
                X += 15;

                s = "Totale incassi                   ";
                gfx.DrawString(s, font4, XBrushes.Black, Y + 0, X, XStringFormats.Default);
                s = strTotInc;
                gfx.DrawString(s, font1, XBrushes.Black, Y + 295, X + 3, frmDX);

                s = "Differenze contanti/movimenti";
                gfx.DrawString(s, font4, XBrushes.Black, Y + 0, X + 30, XStringFormats.Default);
                s = strTotDif;
                gfx.DrawString(s, font1, XBrushes.Black, Y + 295, X + 30 + 4, frmDX);

                // Save the document...
                try
                {
                    string sFil = "C:\\ApProject\\PDF\\VenPos_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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
            else
                sMsg = "Non sono presenti dati da stampare!";

            return sMsg;

        }

        public void PrnPdfStaPos3(DataTable tabTab, string strTotInc, string strTotDif, string strPrnPar, DataView viwIva)
        {
            string s = "";
            int iRows = 50;
            int iRow = 0;
            decimal d = 0;

            //string[] a = strPrnPar.Split(';');

            //string sDivUsr = a[0];
            //string sDivPos = a[1];

            //string sNeg = a[2];
            //string sUsr = a[3];

            //string sPos = "";
            //if (a[3] != "")
            //    sPos = "0" + a[4];


            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Cassa";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Bold);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);
            XFont font4 = new XFont("Courier new", 11, XFontStyle.Bold);
            XFont font5 = new XFont("Courier new", 10, XFontStyle.Bold);

            double Y = 5;
            double X = 20;

            double y = 0;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            DataTable t = tabTab.Copy();

            Console.WriteLine("xxx");

            string[] aPar = strPrnPar.Split('-');
            s = aPar[0];
            string[] aFld = s.Split(';');
            s = aPar[1];
            string[] aUsd = s.Split(';');

            //int iColPos = 0;
            //int iColUsr = 0;
            //int iColDes = 0;
            //int iColImp = 0;

            //if (sPos == "" || sDivPos == "S")
            //    iColUsr = iColPos + 50;
            //if (sUsr == "" || sDivUsr == "S")
            //    iColDes = iColUsr + 150;
            //iColImp = iColDes + 200;

            string[] a = strPrnPar.Split('-');
            s = a[1];
            string[] aa = s.Split(';');

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
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA VENDUTO PER CASSA";
                    s += " dal " + _dayIni.ToString("dd/MM/yy") + " al " + _dayFin.ToString("dd/MM/yy");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 5, X, XStringFormats.Default);

                    iRow = 0;

                    s = "";

                    X += 30;

                    y = 0;

                    s = "PAGAMENTI";

                    gfx.DrawString(s, font5, XBrushes.Black, Y + y + 10, X, XStringFormats.Default);

                    foreach(string ss in aUsd)
                    {
                        y += 90;
                        s = ss;
                        //gfx.DrawString(s, font5, XBrushes.Black, Y + y, X, XStringFormats.Default);

                        XTextFormatter tf = new XTextFormatter(gfx);
                        XRect rect = new XRect(y - 8, X-20, 80, 40);
                        gfx.DrawRectangle(XBrushes.Transparent, rect);
                        tf.Alignment = XParagraphAlignment.Right; 
                        tf.DrawString(s, font5, XBrushes.Black, rect, XStringFormats.TopLeft);
                    }

                    X += 20;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    y = 5;

                    if (i == 0)
                        Console.WriteLine("xxxx");

                    y = 10;

                    s = (string)t.Rows[i]["TmpDes"];
                    gfx.DrawString(s, font2, XBrushes.Black, Y + y, X, XStringFormats.Default);

                    y = 65;

                    foreach (string ss in aFld)
                    {
                        if (ss != "")
                        {
                            y += 90;

                            s = ((decimal)t.Rows[i][ss]).ToString("####,##0.00");
                            gfx.DrawString(s, font2, XBrushes.Black, Y + y, X + 3, frmDX);
                        }
                    }

                }

                X += 15;
            }

            X += 15;

            s = "Totale incassi                   ";
            gfx.DrawString(s, font4, XBrushes.Black, Y + 10, X, XStringFormats.Default);
            s = strTotInc;
            gfx.DrawString(s, font1, XBrushes.Black, Y + 270, X + 3, frmDX);

            //s = "Saldo";
            //gfx.DrawString(s, font4, XBrushes.Black, Y + 0, X + 30, XStringFormats.Default);
            //s = strTotDif;
            //gfx.DrawString(s, font1, XBrushes.Black, Y + 270, X + 30 + 4, frmDX);

            X += 15;
            Y = 15;

            s = "Riepilogo IVA";
            gfx.DrawString(s, font4, XBrushes.Black, Y + 0, X + 30, XStringFormats.Default);
            //s = strTotDif;
            //gfx.DrawString(s, font1, XBrushes.Black, Y + 270, X + 30 + 4, frmDX);

            //foreach (DataRowView r in viwIva)
            //{

            t = viwIva.ToTable();

            decimal dImp = 0;
            decimal dIva = 0;

            iRow = 0;

            for (int i = 0; i <= t.Rows.Count - 1; i++)
            {
                if (iRow > iRows || iRow == 0)
                {
                    if (iRow > 0)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }

                    X += 20;
                    iRow = 0;
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA VENDUTO PER ";
                    s += "IVA";
                    s += " dal " + _dayIni.ToString("dd/MM/yy") + " al " + _dayFin.ToString("dd/MM/yy");

                    //if (strNeg != "")
                    //    s += " - PV " + strNeg;

                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    s = "Descrizione";
                    gfx.DrawString(s, font5, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = "Imponibile";
                    gfx.DrawString(s, font5, XBrushes.Black, Y + 148, X, XStringFormats.Default);

                    s = "Imposta";
                    gfx.DrawString(s, font5, XBrushes.Black, Y + 254, X, XStringFormats.Default);

                    s = "Totale";
                    gfx.DrawString(s, font5, XBrushes.Black, Y + 358, X, XStringFormats.Default);

                    X += 35;
                }


                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = (string)t.Rows[i]["iva_cod"] + " " + (string)t.Rows[i]["iva_des"];
                    gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = ((decimal)t.Rows[i]["iva_imp"]).ToString("####,##0.00");
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 200, X + 3, frmDX);
                    //dTot += (decimal)t.Rows[i]["tmp_ils"];

                    s = ((decimal)t.Rows[i]["iva_iva"]).ToString("####,##0.00");
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 295, X + 3, frmDX);

                    s = ((decimal)t.Rows[i]["iva_imp"] + (decimal)t.Rows[i]["iva_iva"]).ToString("####,##0.00");
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 390, X + 3, frmDX);

                    dImp += (decimal)t.Rows[i]["iva_imp"];
                    dIva += (decimal)t.Rows[i]["iva_iva"];

                }

                X += 20;
                //break;
            }

            X += 15;

            s = "Totali    " + new string(' ', 10) + (dImp).ToString("####,##0.00") + new string(' ', 7) + dIva.ToString("####,##0.00") + new string(' ', 6) + (dImp + dIva).ToString("####,##0.00") + "  ";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 0, X, XStringFormats.Default);


            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\VenPos_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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


        public void PrnPdfRepSto(string strTip, DataTable tabTab, decimal decSco, decimal decQta, string strNeg, decimal decMar)
        {
            string s = "";
            int iRows = 18;
            int iRow = 0;
            decimal dTot = 0;
            decimal dTls = 0;
            decimal dSco = 0;
            decimal dQta = 0;
            decimal dInc = 0;
            decimal dAcq = 0;
            decimal dVen = 0;

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
                    s += "Reparto";
                    s += " dal " + _dayIni.ToString("dd/MM/yy") + " al " + _dayFin.ToString("dd/MM/yy");

                    if (strNeg != "")
                        s += " - PV " + strNeg;

                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    s = "Reparto";
                    if (strTip == "ven_iva")
                        s = "IVA";

                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    //s = "Venduto";
                    //gfx.DrawString(s, font1, XBrushes.Black, Y + 220, X, XStringFormats.Default);

                    //s = "Lordo sconti";
                    //gfx.DrawString(s, font1, XBrushes.Black, Y + 185, X + 10, XStringFormats.Default);

                    s = "Venduto";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 290, X, XStringFormats.Default);

                    if (strTip == "ven_rep")
                    {
                        s = "Scontrini";
                        gfx.DrawString(s, font2, XBrushes.Black, Y + 344, X, XStringFormats.Default);

                        s = "Pezzi";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 430, X, XStringFormats.Default);

                        s = "Marg.";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 500, X, XStringFormats.Default);
                    }
                    X += 35;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = (string)t.Rows[i]["tmp_cod"] + " " + (string)t.Rows[i]["tmp_des"];
                    gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                    //s = ((decimal)t.Rows[i]["tmp_ils"]).ToString("####,##0.00");
                    //gfx.DrawString(s, font1, XBrushes.Black, Y + 260, X + 3, frmDX);

                    s = ((decimal)t.Rows[i]["tmp_imp"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 340, X + 3, frmDX);
                    dTot += (decimal)t.Rows[i]["tmp_imp"];

                    if (strTip == "ven_rep")
                    {

                        //s = ((decimal)t.Rows[i]["tmp_inc"]).ToString("####,##0.00") + "%";
                        s = "xxxxxx%";
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 340, X + 13, frmDX);
                        dInc += (decimal)t.Rows[i]["tmp_inc"];

                        s = ((decimal)t.Rows[i]["tmp_sco"]).ToString("###,##0");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 400, X + 3, frmDX);
                        dSco += (decimal)t.Rows[i]["tmp_sco"];

                        s = ((decimal)t.Rows[i]["tmp_sci"]).ToString("###,##0.00") + "%";
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 400, X + 13, frmDX);

                        s = ((decimal)t.Rows[i]["tmp_scm"]).ToString("###,##0.00") + "m";
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 400, X + 23, frmDX);


                        s = ((decimal)t.Rows[i]["tmp_qta"]).ToString("###,##0");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 470, X + 3, frmDX);
                        dQta += (decimal)t.Rows[i]["tmp_qta"];

                        s = ((decimal)t.Rows[i]["tmp_qti"]).ToString("###,##0.00") + "%";
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 470, X + 13, frmDX);

                        s = ((decimal)t.Rows[i]["tmp_qtm"]).ToString("###,##0.00") + "m";
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 470, X + 23, frmDX);

                        s = ((decimal)t.Rows[i]["tmp_mav"]).ToString("##0.00") + "%";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 540, X + 23, frmDX);
                    }
                }

                X += 40;
                //break;
            }

            X += 15;

            s = "Totale    " + dTot.ToString("####,##0.00");
            gfx.DrawString(s, font1, XBrushes.Black, Y + 150, X, XStringFormats.Default);

            if (strTip == "ven_rep")
            {
                s = decSco.ToString("####,##0");
                gfx.DrawString(s, font1, XBrushes.Black, Y + 355, X, XStringFormats.Default);

                s = decQta.ToString("####,##0");
                gfx.DrawString(s, font1, XBrushes.Black, Y + 435, X, XStringFormats.Default);

                s = "";
                if(dTot > 0 && decSco > 0)
                    s = (dTot / decSco).ToString("###,##0.00") + "m";
                gfx.DrawString(s, font4, XBrushes.Black, Y + 376, X + 13, frmDX);

                s = decMar.ToString("###,##0.00") + "%"; ;
                gfx.DrawString(s, font1, XBrushes.Black, Y + 540, X + 3, frmDX);
            }
            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\VenStaRep_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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
