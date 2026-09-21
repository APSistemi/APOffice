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
using System.Globalization;

namespace APOffice
{
    class clsGenPdfArtEcr
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        //public DataSet _dasGen = new DataSet();

        public DataTable _tabArt = new DataTable();
        public string _strConSql = "";
        public clsGenPdfArtEcr()
        {}

        public void PrnPdfArtEcr()
        {
            string s = "";
            int iRows = 36;
            int iRow = 0;
            decimal d = 0;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 9, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 9, XFontStyle.Bold);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            double Y = 5;
            double X = 20;

            DataTable t = _tabArt.Copy();

            DataView vArt = new DataView(_tabArt, "", "EcrDe2, EcrDe3, RepDes, art_des", DataViewRowState.CurrentRows);

            string sKey = "";

            for (int i = 0; i <= vArt.Count - 1; i++)
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
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA ARTICOLI";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 20;

                    if (i == 0)
                        Console.WriteLine("zzzz");

                    s = "Codice" + new string(' ', 5) + "Descrizione" + new string(' ', 26) + "St" + new string(' ', 1) + "UM" + new string(' ', 1) + "Tg" + new string(' ', 1) + "Peso" + new string(' ', 5) + "PLU" + new string(' ', 2) + "Prezzo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 15;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    if (sKey != (string)vArt[i]["EcrDe1"] + (string)vArt[i]["EcrDe2"] + (string)vArt[i]["EcrDe3"] + (string)vArt[i]["RepDes"])
                    {
                        sKey = (string)vArt[i]["EcrDe1"] + (string)vArt[i]["EcrDe2"] + (string)vArt[i]["EcrDe3"] + (string)vArt[i]["RepDes"];

                        s = (string)vArt[i]["EcrDe1"] + " " + (string)vArt[i]["EcrDe2"] + " " + (string)vArt[i]["EcrDe3"] + " " + (string)vArt[i]["RepDes"];
                        gfx.DrawString(s, font2, XBrushes.Black, Y+30, X, XStringFormats.Default);

                        X += 20;
                    }

                    s = (string)vArt[i]["art_cod"] + " " + (string)vArt[i]["art_des"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = (string)vArt[i]["art_sta"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 260, X, XStringFormats.Default);

                    s = (string)vArt[i]["art_umi"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 275, X, XStringFormats.Default);

                    s = (string)vArt[i]["art_tgr"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 290, X, XStringFormats.Default);

                    s = ((decimal)vArt[i]["art_pne"]).ToString("#0.00");
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 335, X + 2, frmDX);

                    s = (string)vArt[i]["art_plu"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 345, X, XStringFormats.Default);

                    s = ((decimal)vArt[i]["art_prv"]).ToString("#0.00");
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 415, X + 2, frmDX);
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
                string sFil = "C:\\ApProject\\PDF\\ArtEcr_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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

        public void PrnPdfArtListino()
        {
            string s = "";
            int iRows = 36;
            int iRow = 0;
            decimal d = 0;

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
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Bold);
            XFont font3 = new XFont("Courier new", 12, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            double Y = 5;
            double X = 20;

            DataTable t = _tabArt.Copy();

            //DataView vArt = new DataView(_tabArt, "", "EcrDe2, EcrDe3, RepDes, art_des", DataViewRowState.CurrentRows);
            DataView vArt = new DataView(_tabArt, "", "RepDes, art_des", DataViewRowState.CurrentRows);

            string sKey = "";

            for (int i = 0; i <= vArt.Count - 1; i++)
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
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA ARTICOLI";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 20;

                    if (i == 0)
                        Console.WriteLine("zzzz");

                    s = "Codice" + new string(' ', 5) + "Descrizione" + new string(' ', 20) + "UM" + new string(' ', 2) + "Barcode" + new string(' ', 16) + "Prezzo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 15;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    if (sKey != (string)vArt[i]["RepDes"])
                    {
                        sKey = (string)vArt[i]["RepDes"];

                        s = (string)vArt[i]["RepDes"];
                        gfx.DrawString(s, font2, XBrushes.Black, Y + 30, X, XStringFormats.Default);

                        X += 20;
                    }
                    string sEan = "";
                    s = "SELECT * FROM AnaBarcode WHERE ean_art='" + vArt[i]["art_cod"] + "' ORDER BY ean_dti DESC";
                    DataTable tTmp = _clsFun.FillTabSql("AnaBarcode", s, true, _strConSql);
                    if(tTmp.Rows.Count > 0)
                        sEan = (string)tTmp.Rows[0]["ean_ean"];

                    s = (string)vArt[i]["art_des"];
                    if (s.Length > 25)
                        s = s.Substring(0, 25);

                    s = (string)vArt[i]["art_cod"] + " " + s;
                    gfx.DrawString(s, font3, XBrushes.Black, Y, X, XStringFormats.Default);

                    //s = (string)vArt[i]["art_sta"];
                    //gfx.DrawString(s, font3, XBrushes.Black, Y + 260, X, XStringFormats.Default);

                    s = (string)vArt[i]["art_umi"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 300, X, XStringFormats.Default);

                    //s = (string)vArt[i]["art_tgr"];
                    //gfx.DrawString(s, font3, XBrushes.Black, Y + 290, X, XStringFormats.Default);

                    //s = ((decimal)vArt[i]["art_pne"]).ToString("#0.00");
                    //gfx.DrawString(s, font3, XBrushes.Black, Y + 335, X + 2, frmDX);


                    s = sEan;
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 330, X, XStringFormats.Default);

                    s = ((decimal)vArt[i]["art_prv"]).ToString("#0.00");
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 550, X + 5, frmDX);

                    XPen pen = new XPen(XColors.Black, 1);
                    gfx.DrawLine(pen, Y, X + 5, Y + 580, X + 5);

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
                string sFil = "C:\\ApProject\\PDF\\ArtEcr_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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

        public void PrnPdfTessere(string strGru, string strTes, string strDes)
        {
            string s = "";
            int iRows = 36;
            int iRow = 0;
            decimal d = 0;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 845.0;
            page.Orientation = PdfSharp.PageOrientation.Portrait;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Bold);
            XFont font3 = new XFont("Courier new", 15, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;


            double Y = 5;
            double X = 5;

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            s = "Gruppo: " + strGru;
            gfx.DrawString(s, font3, XBrushes.Black, Y + 10, X + 30, XStringFormats.Default);

            s = strDes;
            gfx.DrawString(s, font3, XBrushes.Black, Y + 10, X + 50, XStringFormats.Default);

            s = strTes;
            gfx.DrawString(s, font3, XBrushes.Black, Y + 10, X + 70, XStringFormats.Default);

            ////XPen pen = new XPen(XColors.Black, 1);
            ////gfx.DrawLine(pen, Y, X + 5, Y + 500, X + 5);

            //s = strTes;

            Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
            MemoryStream ms = new MemoryStream();
            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
            gfx.DrawImage(img, Y+10, X + 100, 150, 55);

            //s = "Totale    " + d.ToString();
            //gfx.DrawString(s, font1, XBrushes.Black, Y + 240, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\Tessera_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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
