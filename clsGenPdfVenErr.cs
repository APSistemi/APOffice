using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Diagnostics;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using BarcodeLib;

namespace APOffice
{
    class clsGenPdfVenErr
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public clsGenPdfVenErr()
        {
            //_strConMdb = _clsFun.ConMdb(MDBPATH);
        }

        public void PrnPdfVenErr(DataTable tabMsg)
        {
            string s = "";
            int iRows = 37;
            int iRow = 0;
            decimal d = 0;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Reparti";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 14, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 18, XFontStyle.Regular);
            XFont font3 = new XFont("Courier new", 10, XFontStyle.Regular);

            double Y = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            DataTable t = tabMsg.Copy();

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
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA ERRORI SU LETTURA DEL VENDUTO CASSE";
                    //if (chkCli.Checked)
                    //    s += "CLIENTI ";
                    //if (chkFor.Checked)
                    //    s += "FORNITORI ";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    s = "Codice";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = "Messaggio";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 100, X, XStringFormats.Default);


                    X += 35;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = (string)t.Rows[i]["tmp_ean"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["tmp_msg"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y+100, X, XStringFormats.Default);

                    //s = ((decimal)t.Rows[i]["vre_ven"]).ToString("####,##0.00");
                    //gfx.DrawString(s, font1, XBrushes.Black, Y + 430, X + 3, frmDX);
                    //d += (decimal)t.Rows[i]["vre_ven"];
                    //break;
                }

                X += 20;
                //break;

            }

            X += 15;

            //s = "Totale    " + d.ToString();
            //gfx.DrawString(s, font1, XBrushes.Black, Y + 240, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\VenErr_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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

        public void PrnPdfOrfErr(DataTable tabMsg)
        {
            string s = "";
            int iRows = 37;
            int iRow = 0;
            decimal d = 0;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Reparti";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 14, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 8, XFontStyle.Regular);
            XFont font3 = new XFont("Courier new", 10, XFontStyle.Regular);

            double Y = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            DataTable t = tabMsg.Copy();

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
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA ERRORI SU LETTURA ORDINE";
                    //if (chkCli.Checked)
                    //    s += "CLIENTI ";
                    //if (chkFor.Checked)
                    //    s += "FORNITORI ";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    s = "Codice";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = "Messaggio";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 100, X, XStringFormats.Default);


                    X += 35;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = (string)t.Rows[i]["tmp_ean"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["tmp_msg"];
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 100, X, XStringFormats.Default);

                    //s = ((decimal)t.Rows[i]["vre_ven"]).ToString("####,##0.00");
                    //gfx.DrawString(s, font1, XBrushes.Black, Y + 430, X + 3, frmDX);
                    //d += (decimal)t.Rows[i]["vre_ven"];
                    //break;
                }

                X += 20;
                //break;

            }

            X += 15;

            //s = "Totale    " + d.ToString();
            //gfx.DrawString(s, font1, XBrushes.Black, Y + 240, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\VenErr_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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

        public void PrnPdfTrmDocErr(DataTable tabMsg)
        {
            string s = "";
            int iRows = 33;
            int iRow = 0;
            decimal d = 0;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Reparti";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 14, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 9, XFontStyle.Bold);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Bold);

            double Y = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            DataTable t = tabMsg.Copy();

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
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA CONTROLLO DOCUMENTO DA TERMINALINO";

                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    s = "Articolo";
                    gfx.DrawString(s, font2, XBrushes.Black, Y, X - 8, XStringFormats.Default);

                    s = "Barcode            Codice   Descrizione";
                    gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 15;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = (string)t.Rows[i]["msg_co1"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["msg_co2"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 100, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["msg_des"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 150, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["msg_msg"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y, X + 8, XStringFormats.Default);

                }

                X += 22;
                //break;

            }

            X += 15;

            //s = "Totale    " + d.ToString();
            //gfx.DrawString(s, font1, XBrushes.Black, Y + 240, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\VenErr_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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
