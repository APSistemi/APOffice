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
    class clsGenPdfEti2
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        public int StartRow = 1;
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        
        public clsGenPdfEti2()
        {
            _strConSql = _clsFun.ConSql("");
        }

        public string PrnPdfEti030(DataTable tabEti, string strPar)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            //page.Height = 845.0;

            page.Height = 842.0;
            page.Width = 595;
            page.Orientation = PdfSharp.PageOrientation.Landscape;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 6.5, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Arial", 10, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 33, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 10, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 9, XFontStyle.Regular);
            XFont font7 = new XFont("Courier new", 7.5, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 5;
            int iRows = 5;
            int i = 0;
            int X_Ini = -1;
            int X_Step = 113;
            int Y_Ini = 0;
            int Y_Step = 173;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {

                        page = pd.AddPage();

                        page.Height = 842.0;
                        page.Width = 595;
                        page.Orientation = PdfSharp.PageOrientation.Landscape;


                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            /*
                             * Descrizione articolo
                             */
                            double x = X;
                            double y = Y + (iCol * Y_Step);
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            if (s.Length > 45)
                                s = s.Substring(0, 45);

                            //string s2 = "";
                            //string[] a = s.Split(' ');

                            //if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                            //    Console.WriteLine("aaaa");

                            //for (int i2 = 0; i2 < a.Length; i2++)
                            //{
                            //    if (a[i2].Length > 20)
                            //        a[i2] = a[i2].Substring(0, 10) + " " + a[i2].Substring(10, 10) + " " + a[i2].Substring(20);

                            //    else if (a[i2].Length > 10)
                            //        a[i2] = a[i2].Substring(0, 10) + " " + a[i2].Substring(10);

                            //    s2 += a[i2] + " ";
                            //}

                            //s = s2;


                            //XRect rect = new XRect(40, 100, 250, 220);
                            //gfx.DrawRectangle(XBrushes.SeaShell, rect);
                            //tf.DrawString(text, font, XBrushes.Black, rect, XStringFormats.TopLeft);

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            XRect rect = new XRect(y - 5, x, 150, 40);

                            XTextFormatter tf = new XTextFormatter(gfx);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.Alignment = XParagraphAlignment.Right;
                            tf.DrawString(s, font3, XBrushes.Black, rect, XStringFormats.TopLeft);

                            /*
                             * Data
                             */
                            x = X + 28;
                            y = Y + (iCol * Y_Step) + 0;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             * Fod
                             */
                            x = X + 28;
                            y = Y + (iCol * Y_Step) + 128;
                            //s = "C " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            s = string.Format("C {0,3}", (decimal)tabEti.Rows[i]["eti_pxc"]);

                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);

                            //20200401 non vogliono il fornitore sull'etichetta
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                             * Riga
                             */
                            x = X + 34;
                            y = Y + (iCol * Y_Step) - 23;
                            XPen pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y + 20, x, y + 170, x);

                            /*
                             * Barcode
                             */
                            x = X + 38;
                            y = Y + (iCol * Y_Step) + 0;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                    else
                                    {
                                        s = s.PadLeft(13, Convert.ToChar('0'));

                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            /*
                             *  Tipo grammatura + Peso netto
                             */
                            x = X + 42;
                            y = Y + (iCol * Y_Step) + 120;
                            s = (string)tabEti.Rows[i]["eti_tgr"];
                            s += " " + Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]).ToString("#0.00");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             * Euro
                             */
                            //x = X + 75;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "€";
                            //gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            /*
                            * Prezzo
                            */
                            x = X + 82;
                            y = Y + (iCol * Y_Step) + 130;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            /*
                             * Codice articolo
                             */
                            x = X + 89;
                            y = Y + (iCol * Y_Step) + 0;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                             * Fornitore + ARF
                             */
                            x = X + 89;
                            y = Y + (iCol * Y_Step) + 30;
                            s = "";
                            if (strPar.Contains("FOR"))         //20200525 C'è chi vuole il fornitore e chi non lo vuole
                            {
                                s = (string)tabEti.Rows[i]["eti_fod"];
                                if (s.Length > 5)
                                    s = s.Substring(0, 5);
                                if (s != "")
                                    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            }
                            //s = ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                             * Barcode
                             */
                            x = X + 81;
                            y = Y + (iCol * Y_Step) + 0;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                             * Prezzo al KG
                             */
                            x = X + 90;
                            y = Y + (iCol * Y_Step) + 87;
                            if ((string)tabEti.Rows[i]["eti_tgr"] != "PZ" &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            {
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));
                                if (d > 0 && d <= 500m)
                                {
                                    s = d.ToString("#####0.00") + " € al kg/L ";
                                    gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                                }
                            }

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti031(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

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
            XFont font1 = new XFont("Courier new", 11, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 15, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Franklin Gothic Demi Cond", 20, XFontStyle.Regular);
            //XFont font4 = new XFont("Arial", 55, XFontStyle.BoldItalic);
            XFont font4 = new XFont("Franklin Gothic Demi Cond", 60, XFontStyle.Bold);

            XFont font5 = new XFont("Impact", 28, XFontStyle.BoldItalic);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);
            //XFont font7 = new XFont("Impact", 10, XFontStyle.Regular);
            XFont font7 = new XFont("Arial", 24, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 5;
            int i = 0;
            int X_Ini = 0;
            int X_Step = 158;
            int Y_Ini = 45;
            int Y_Step = 250;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            if (i == 1)
                                Console.WriteLine("aaaa");

                            int iLen = 15;

                            double x = 0;
                            double y = 0;

                            /*
                             *  For + ARF
                             */ 
                            //y = Y + (iCol * Y_Step) + 130;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             *  Descrizione articolo
                             */
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            string s2 = "";
                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            x = X + 17;
                            y = Y + (iCol * Y_Step) + 20;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            XRect rect = new XRect(y, x, 210, 77);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);


                            /*
                             *  Riga
                             */ 
                            x = X + 58;
                            y = Y + (iCol * Y_Step) + 5;
                            XPen pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y + 20, x, y + 230, x);


                            /*
                             *  UMI
                             */
                            x = X + 72;
                            y = Y + (iCol * Y_Step) + 200;
                            //if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            //    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));
                            //s = "€ al kg/L " + d.ToString("#####0.00");
                            s = "al pz.";
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s = "al Kg";
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             *  Euro
                             */ 
                            x = X + 120;
                            y = Y + (iCol * Y_Step) + 60;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);


                            /*
                             *  Prezzo
                             */ 
                            x = X + 130;
                            y = Y + (iCol * Y_Step) + 220;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);



                            /*
                             * PLU
                             */
                            //x = X + 129;
                            x = X + 80;
                            y = Y + (iCol * Y_Step) + 20;
                            s = ((string)tabEti.Rows[i]["eti_plu"]);
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             *  Codice articolo
                             */
                            x = X + 140;
                            y = Y + (iCol * Y_Step) + 20;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             *  Codice articolo fornitore
                             */
                            x = X + 140;
                            y = Y + (iCol * Y_Step) + 100;
                            s = ((string)tabEti.Rows[i]["eti_arf"]);
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             *  Data
                             */
                            x = X + 140;
                            y = Y + (iCol * Y_Step) + 180;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             * BARCODE
                             * 
                             */
                            //y = Y + (iCol * Y_Step) + 95;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            //x = X + 120;
                            //y = Y + (iCol * Y_Step) + 175;
                            //s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            //if (s.Trim().Length > 0)
                            //{
                            //    try
                            //    {
                            //        if (s.Length <= 8)
                            //        {
                            //            Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //        else
                            //        {
                            //            Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                            //    }
                            //}

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 105;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti032(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            //page.Height = 845.0;

            //page.Height = 842.0;
            //page.Width = 595;
            //page.Orientation = PdfSharp.PageOrientation.Landscape;

            page.Height = 845.0;
            page.Orientation = PdfSharp.PageOrientation.Portrait;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 6.5, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Impact", 15, XFontStyle.Regular);
            XFont font4 = new XFont("Impact", 50, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 10, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 9, XFontStyle.Regular);
            XFont font7 = new XFont("Courier new", 7.5, XFontStyle.Bold);
            XFont font8 = new XFont("Courier new", 10, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 4;
            int i = 0;
            int X_Ini = 70;
            int X_Step = 212;
            int Y_Ini = 10;
            int Y_Step = 297;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                if (DBNull.Value.Equals(y["eti_cam"]))
                    y["eti_cam"] = "";

                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = 0;
                            double y = 0;

                            /*** Godina***/
                            s = "C:\\ApProject\\Temp\\Img\\Etichette\\eti032Fidelity.jpg";
                            if ((string)tabEti.Rows[i]["eti_cam"] != "" && File.Exists(s))
                            {
                                x = X-10;
                                y = Y + (iCol * Y_Step-8);
                                
                                XImage img = XImage.FromFile(s);
                                XRect rect1 = new XRect(y, x, 100, 40);
                                gfx.DrawImage(img, rect1);
                            }

                            /*** RETTANGOLO PER IL BORDO ETICHETTA X GATTI ***/
                            x = X - 65;
                            y = Y + (iCol * Y_Step);
                            XPen pen = new XPen(XColors.Black, 1);
                            if (File.Exists("C:\\ApProject\\Temp\\Img\\Etichette\\eti032conbordo.flg"))
                            {
                                //gfx.DrawRectangle(pen, y - 10, x - 5, Y_Step, X_Step); Bordo completo

                                gfx.DrawLine(pen, y - 10, x + X_Step-1, y + Y_Step, x + X_Step-1);          //Riga

                                if(y < 100)
                                    gfx.DrawLine(pen, y + Y_Step - 10, x, y + Y_Step - 10, x + X_Step);           //Colonna
                            }

                            x = X - 60;
                            y = Y + (iCol * Y_Step) + 55;
                            s = "C:\\ApProject\\Temp\\Img\\Etichette\\eti032titolo.png";
                            if (File.Exists(s))
                            {
                                XImage img = XImage.FromFile(s);
                                XRect rect1 = new XRect(y, x, 200, 50);
                                gfx.DrawImage(img, rect1);
                            }

                            /*
                             * Descrizione articolo
                             */
                            x = X;
                            y = Y + (iCol * Y_Step);
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            if (s.Length > 45)
                                s = s.Substring(0, 45);

                            //string s2 = "";
                            //string[] a = s.Split(' ');

                            //if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                            //    Console.WriteLine("aaaa");

                            //for (int i2 = 0; i2 < a.Length; i2++)
                            //{
                            //    if (a[i2].Length > 20)
                            //        a[i2] = a[i2].Substring(0, 10) + " " + a[i2].Substring(10, 10) + " " + a[i2].Substring(20);

                            //    else if (a[i2].Length > 10)
                            //        a[i2] = a[i2].Substring(0, 10) + " " + a[i2].Substring(10);

                            //    s2 += a[i2] + " ";
                            //}

                            //s = s2;

                            //XRect rect = new XRect(40, 100, 250, 220);
                            //gfx.DrawRectangle(XBrushes.SeaShell, rect);
                            //tf.DrawString(text, font, XBrushes.Black, rect, XStringFormats.TopLeft);

                            x = X - 5;
                            y = Y + (iCol * Y_Step) + 100;
                            XRect rect = new XRect(y + 0, x, 150, 40);
                            XTextFormatter tf = new XTextFormatter(gfx);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.Alignment = XParagraphAlignment.Center;
                            tf.DrawString(s, font3, XBrushes.Black, rect, XStringFormats.TopLeft);

                            if (i == 0)
                                Console.WriteLine("aaaa");
                            /*
                             * Date offerta
                             */

                            if ((string)tabEti.Rows[i]["eti_oft"] == "PRO" || ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ))
                            {

                                if ((string)tabEti.Rows[i]["eti_cam"] == "" && !File.Exists(s))     //In caso di fidelity stampo le date sotto
                                {
                                    x = X + 0;
                                    y = Y + (iCol * Y_Step) + 15;
                                    s = "Offerta valida";
                                    gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);
                                    x = X + 10;
                                    y = Y + (iCol * Y_Step) + 15;
                                    s = "dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yyyy");
                                    gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);
                                    x = X + 20;
                                    y = Y + (iCol * Y_Step) + 15;
                                    s = "al  " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");
                                    gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);

                                }

                                d = 0;
                                if ((decimal)tabEti.Rows[i]["eti_prv"] != 0 && (decimal)tabEti.Rows[i]["eti_pve"] != 0)
                                    d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                x = X + 45;
                                y = Y + (iCol * Y_Step) + 110;
                                s = "da € " + Convert.ToDecimal(tabEti.Rows[i]["eti_pve"]).ToString("######0.00") + " - ";
                                s += "sconto " + d.ToString("#,##0.00") + " %";
                                gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);
                            }
                            else if ((string)tabEti.Rows[i]["eti_oft"] == "002")
                            {
                                /*
                                 * %
                                 */
                                if ((decimal)tabEti.Rows[i]["eti_pve"] > 0)
                                {
                                    x = X + 45;
                                    y = Y + (iCol * Y_Step) + 110;
                                    d = (decimal)tabEti.Rows[i]["eti_pve"];
                                    //s = "da € " + Convert.ToDecimal(tabEti.Rows[i]["eti_pve"]).ToString("######0.00") + " - ";
                                    s = "sconto " + d.ToString("#,##0.00") + " %";
                                    gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);
                                }
                            }

                            /*
                             * PxC
                             */
                            if ((decimal)tabEti.Rows[i]["eti_pxc"] > 0)
                            {
                                x = X + 45;
                                y = Y + (iCol * Y_Step) + 15;
                                s = "C " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                                gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            }


                            ///*
                            // * Data
                            // */
                            //x = X + 48;
                            //y = Y + (iCol * Y_Step) + 0;
                            //s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                             * Riga
                             */
                            x = X + 34;
                            y = Y + (iCol * Y_Step) + 83;
                            pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y + 20, x, y + 190, x);

                            /*
                             * Riga 2
                             */
                            x = X + 110;
                            y = Y + (iCol * Y_Step) + 83;
                            pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y + 20, x, y + 190, x);

                            /*
                             * Barcode
                             */
                            x = X + 50;
                            y = Y + (iCol * Y_Step) + 15;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y - 5, x, 70, 18);
                                    }
                                    else
                                    {
                                        s = s.PadLeft(13, Convert.ToChar('0'));

                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            /*
                             *  Tipo grammatura + Peso netto
                             */
                            if ((decimal)tabEti.Rows[i]["eti_pne"] > 0)
                            {
                                x = X + 35;
                                y = Y + (iCol * Y_Step) + 15;
                                s = (string)tabEti.Rows[i]["eti_tgr"];
                                s += " " + Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]).ToString("#0.00");
                                gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            }

                            /*
                             * Euro
                             */
                            //x = X + 75;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "€";
                            //gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            /*
                             * Prezzo
                             */
                            x = X + 110;
                            y = Y + (iCol * Y_Step) + 210;
                            //s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");

                            d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]);
                            if ((string)tabEti.Rows[i]["eti_oft"] == "002" && Convert.ToDecimal(tabEti.Rows[i]["eti_pve"]) > 0)
                                d = _clsFun.MenoPer(d, Convert.ToDecimal(tabEti.Rows[i]["eti_pve"]));
                            gfx.DrawString(d.ToString("######0.00"), font4, XBrushes.Black, y + 10, x, frmDX);

                            /*
                             * Codice articolo
                             */
                            x = X + 89;
                            y = Y + (iCol * Y_Step) + 15;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                             * Fornitore + ARF
                             */
                            x = X + 89;
                            y = Y + (iCol * Y_Step) + 45;
                            s = "";
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();

                            s = ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            if (s != "")
                                s = "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                             * Barcode
                             */
                            x = X + 75;
                            y = Y + (iCol * Y_Step) + 20;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                             * Prezzo al KG
                             */
                            x = X + 80;
                            y = Y + (iCol * Y_Step) + 240;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            {
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));
                                if (d > 0 && d <= 500m)
                                {
                                    s = "€ " + d.ToString("#####0.00");
                                    gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                                    x = X + 90;
                                    y = Y + (iCol * Y_Step) + 240;
                                    s = "al kg/L ";
                                    gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                                }
                            }
                            //break;

                            s = "C:\\ApProject\\Temp\\Img\\Etichette\\eti032Fidelity.jpg";
                            if ((string)tabEti.Rows[i]["eti_cam"] != "" && File.Exists(s))
                            {
                                x = X + 120;
                                y = Y + (iCol * Y_Step) + 5;

                                s = "Offerta valida dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yyyy") + " al " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");
                                gfx.DrawString(s, font8, XBrushes.Black, y, x, XStringFormats.Default);
                                //x = X + 10;
                                //y = Y + (iCol * Y_Step) + 15;
                                //s = "dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yyyy");
                                //gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);
                                //x = X + 20;
                                //y = Y + (iCol * Y_Step) + 15;
                                //s = "al  " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");
                                //gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);

                                //d = 0;
                                //if ((decimal)tabEti.Rows[i]["eti_prv"] != 0 && (decimal)tabEti.Rows[i]["eti_pve"] != 0)
                                //    d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                //x = X + 45;
                                //y = Y + (iCol * Y_Step) + 110;
                                //s = "da € " + Convert.ToDecimal(tabEti.Rows[i]["eti_pve"]).ToString("######0.00") + " - ";
                                //s += "sconto " + d.ToString("#,##0.00") + " %";
                                //gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);
                            }
                            else
                            {
                                s = "C:\\ApProject\\Temp\\Img\\Etichette\\eti032offerta.jpg";
                                if (File.Exists(s) && (string)tabEti.Rows[i]["eti_oft"] == "001")
                                {
                                    x = X + 112;
                                    y = Y + (iCol * Y_Step) + 45;

                                    XImage img = XImage.FromFile(s);
                                    XRect rect1 = new XRect(y, x, 200, 25);
                                    gfx.DrawImage(img, rect1);
                                }
                            }
                            s = "C:\\ApProject\\Temp\\Img\\Etichette\\eti032MxN.png";
                            if (File.Exists(s) && (string)tabEti.Rows[i]["eti_oft"] == "003")
                            {
                                x = X + -12;
                                y = Y + (iCol * Y_Step) + 15;

                                XImage img = XImage.FromFile(s);
                                XRect rect1 = new XRect(y, x, 40, 40);
                                gfx.DrawImage(img, rect1);
                            }

                            i++;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti033(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();

            page.Height = 842.0;
            page.Width = 595;
            page.Orientation = PdfSharp.PageOrientation.Landscape;

            //System.Drawing.Size size = PdfPage.PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Orientation = PageOrientation.Portrait;
            //pdfPage.Width = size.Width;
            //pdfPage.Height = size.Height;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 9, XFontStyle.Regular);
            XFont font2 = new XFont("Arial", 15, XFontStyle.Bold);
            //XFont font3 = new XFont("Impact", 35, XFontStyle.Regular);
            XFont font3 = new XFont("Arial", 32, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 90, XFontStyle.Bold);
            XFont font5 = new XFont("Impact", 15, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 10, XFontStyle.Regular);

            XFont font7 = new XFont("Arial", 11, XFontStyle.Bold);
            XFont font8 = new XFont("Arial", 25, XFontStyle.Bold);
            XFont font9 = new XFont("Impact", 20, XFontStyle.Regular);
            XFont font10 = new XFont("Arial", 80, XFontStyle.Bold);
            XFont font12 = new XFont("Courier new", 20, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 2;
            int i = 0;
            int X_Ini = 0;
            int X_Step = 288;
            int Y_Ini = 5;
            int Y_Step = 440;
            int X_Pie = 260;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        page.Orientation = PdfSharp.PageOrientation.Landscape;

                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            int iLen = 15;
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            //if (s.Length > 45)
                            //    s = s.Substring(0, 45);

                            string s2 = "";

                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;

                            double x = X + 30;
                            double y = Y + (iCol * Y_Step) + 10;

                            XRect rect = new XRect(y, x, 180, 140);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            x = X + 30;
                            y = Y + (iCol * Y_Step) + 240;
                            rect = new XRect(y, x, 140, 235);
                            XPen pen = new XPen(XColors.Black, 1.5);
                            gfx.DrawRectangle(pen, rect);

                            for (int ii = 1; ii <= 3; ii++)
                            {
                                x += 58;
                                pen = new XPen(XColors.Black, 1.5);
                                if(ii == 2)
                                    gfx.DrawLine(pen, y -250, x, y + 140, x);
                                else
                                    gfx.DrawLine(pen, y + 0, x, y + 140, x);
                            }

                            x = X + 50;
                            y = Y + (iCol * Y_Step) + 280;

                            if ((string)tabEti.Rows[i]["eti_bor"] != "")
                            {
                                s = "Origine";
                                gfx.DrawString(s, font2, XBrushes.Black, y, x + 0, XStringFormats.Default);
                                s = (string)tabEti.Rows[i]["eti_bor"];
                                gfx.DrawString(s, font5, XBrushes.Black, y, x + 20, XStringFormats.Default);
                            }
                            if ((string)tabEti.Rows[i]["eti_bcl"] != "")
                            {
                                s = "Calibro";
                                gfx.DrawString(s, font2, XBrushes.Black, y, x + 62, XStringFormats.Default);
                                s = (string)tabEti.Rows[i]["eti_bcl"];
                                gfx.DrawString(s, font5, XBrushes.Black, y + 20, x + 82, XStringFormats.Default);
                            }
                            if ((string)tabEti.Rows[i]["eti_bct"] != "")
                            {
                                s = "Categoria";
                                gfx.DrawString(s, font2, XBrushes.Black, y, x + 125, XStringFormats.Default);
                                s = (string)tabEti.Rows[i]["eti_bct"];
                                gfx.DrawString(s, font5, XBrushes.Black, y, x + 145, XStringFormats.Default);
                            }

                            if (i == 0)
                                Console.WriteLine("aaaa");


                            //s = ((string)tabEti.Rows[i]["eti_art"]).Substring(1);
                            s = ((string)tabEti.Rows[i]["eti_art"]);
                            if (_clsFun.Numerico(s))
                            {
                                s = Convert.ToInt32(s).ToString();
                                gfx.DrawString(s, font7, XBrushes.Black, y - 10, x + 197, XStringFormats.Default);
                            }

                            s = ((string)tabEti.Rows[i]["eti_arf"]);
                            if (_clsFun.Numerico(s))
                            {
                                //s = Convert.ToInt32(s).ToString();
                                gfx.DrawString(s, font7, XBrushes.Black, y + 40, x + 197, XStringFormats.Default);
                            }

                            s = ((string)tabEti.Rows[i]["eti_plu"]);
                            if (_clsFun.Numerico(s))
                            {
                                s = Convert.ToInt32(s).ToString();
                                gfx.DrawString(s, font7, XBrushes.Black, y, x + 210, XStringFormats.Default);
                            }

                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font7, XBrushes.Black, y + 40, x + 210, XStringFormats.Default);

                            
                            /**** PLU FINE ****/

                            //if ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ)
                            //{
                            //    s = "OFFERTA";

                            //    //XTextFormatter tf = new XTextFormatter(gfx);
                            //    //tf.Alignment = XParagraphAlignment.Center;

                            //    //rect = new XRect(y, x - 50, 500, 400);
                            //    //gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //    //tf.DrawString(s, font3, XBrushes.Black, rect);

                            //    x = X + 100;
                            //    y = Y + (iCol * Y_Step) - 10; // +(iCol * Y_Step) + 190;
                            //    tf = new XTextFormatter(gfx);
                            //    rect = new XRect(y, x, 400, 40);

                            //    gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                            //    //tf.Alignment = XParagraphAlignment.Default; //tf.Alignment = ParagraphAlignment.Left; 
                            //    //tf.DrawString(s, font7, XBrushes.Black, rect, XStringFormats.TopLeft);

                            //    if ((decimal)tabEti.Rows[i]["eti_pve"] > 0 && (decimal)tabEti.Rows[i]["eti_prv"] > 0)
                            //        d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                            //    if (d != 0)
                            //    {
                            //        s = ((decimal)tabEti.Rows[i]["eti_pve"]).ToString("#,##0.00");
                            //        gfx.DrawString(s, font8, XBrushes.Black, y + 20, x + 30, XStringFormats.Default);

                            //        XPen pen = new XPen(XColors.Black, 3);
                            //        gfx.DrawLine(pen, y + 17, x + 30, y + 100, x + 15);
                            //    }

                            //    s = "OFFERTA";
                            //    gfx.DrawString(s, font9, XBrushes.White, y + 150, x + 35, XStringFormats.Default);

                            //    if (d != 0)
                            //    {
                            //        s = "Sc." + d.ToString("#,##0.00") + " %";
                            //        gfx.DrawString(s, font7, XBrushes.Black, y + 320, x + 30, XStringFormats.Default);
                            //    }
                            //}
                            /******/

                            //x = X + 250;
                            //y = Y + (iCol * Y_Step) + 10;
                            //s = "€";
                            //gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            /*
                             * PREZZO
                             */
                            x = X + 250;
                            y = Y + (iCol * Y_Step) + 210;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y, x, frmDX);

                            /*
                             * AL KG
                             */
                            x = X + 140;
                            y = Y + (iCol * Y_Step) + 90;
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s = "al Kg.";
                            else
                                s = "al Pz.";
                            gfx.DrawString(s, font12, XBrushes.Black, y-80, x + 122, XStringFormats.Default);

                            //x = X + 240;
                            //y = Y + (iCol * Y_Step) + 5;
                            //if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            //    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            //s = "€ al kg/L " + d.ToString("#####0.00");
                            //gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            ///*
                            // * Codice articolo
                            // */
                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 280;
                            //s = (string)tabEti.Rows[i]["eti_art"];
                            //gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 220;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            x = X + 210;
                            y = Y + (iCol * Y_Step) + 270;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 95, 25);
                                    }
                                    else
                                    {
                                        s = s.PadLeft(13, Convert.ToChar('0'));
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 95, 25);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 305;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 300;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti034(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 845.0;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Regular);
            XFont font2 = new XFont("Impact", 70, XFontStyle.Regular);
            XFont font3 = new XFont("Arial", 50, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 180, XFontStyle.Regular);
            XFont font5 = new XFont("Arial", 80, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 20, XFontStyle.Regular);
            XFont font7 = new XFont("Arial", 11, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 1;
            int iRows = 1;
            int i = 0;
            int X_Ini = 100;
            int X_Step = 100;
            int Y_Ini = 55;
            int Y_Step = 200;
            int X_Pie = 720;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = X;
                            double y = Y + (iCol * Y_Step);

                            /*
                             * TITOLO
                             */
                            s = "PREZZO SPECIALE";
                            gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                             * DESCRIZIONE
                             */
                            int iLen = 15;
                            s = (string)tabEti.Rows[i]["eti_ard"];

                            string s2 = "";

                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;

                            x = X + 50;
                            y = Y + (iCol * Y_Step);

                            XRect rect = new XRect(y, x, 500, 400);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            /*
                            * PREZZO
                            */
                            x = X + 550;
                            y = Y + (iCol * Y_Step) + 500;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);


                            /*
                            * EURO
                            */
                            x = X + 550;
                            y = Y + (iCol * Y_Step) - 20;
                            s = "€.";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);



                            /*
                            * PREZZO IN CHIARO
                            */
                            x = X + 580;
                            y = Y + (iCol * Y_Step) + 370;
                            d = 0;
                            //Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG" &&
                            if (
                                Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }


                            /*
                            * AL PZ
                            */
                            x = X + X_Pie - 40;
                            y = Y + (iCol * Y_Step) - 20;
                            s = "al pz.";
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s = "al Kg";
                            gfx.DrawString(s, font6, XBrushes.Black, y, x - 5, XStringFormats.Default);


                            /*
                            * TIPO GRAMMATURA
                            */
                            x = X + X_Pie - 43;
                            y = Y + (iCol * Y_Step) + 380;
                            s = (string)tabEti.Rows[i]["eti_tgr"];
                            s += " " + Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]).ToString("#0.00");
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             * LINEA
                             */
                            x = X + X_Pie -40;
                            y = Y + (iCol * Y_Step) - 50;
                            XPen pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y, x, y + 580, x);


                            /*
                             * PIEDE DESCRIZIONE
                             */
                            x = X + X_Pie - 20;
                            y = Y + (iCol * Y_Step) - 40;
                            s = "Codice articolo      PLU          Codice a Barre                  Pxc               Data di Stampa";
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * ARTICOLO
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) - 15;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * PLU
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 50;
                            s = (string)tabEti.Rows[i]["eti_plu"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * BARCODE
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 100;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * PXC
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 230;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                            * DATA
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 310;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            ///*
                            //* FORNITORE
                            //*/
                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 220;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * BARCODE
                            */
                            x = X + X_Pie - 25;
                            y = Y + (iCol * Y_Step) + 420;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                    else
                                    {
                                        s = s.PadLeft(13, Convert.ToChar('0'));
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }


                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 305;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti035(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            //page.Height = 845.0;

            page.Height = 842.0;
            page.Width = 595;
            page.Orientation = PdfSharp.PageOrientation.Landscape;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Regular);
            XFont font2 = new XFont("Verdana", 80, XFontStyle.Bold);
            XFont font3 = new XFont("Arial", 60, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 180, XFontStyle.Regular);
            XFont font5 = new XFont("Arial", 80, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 20, XFontStyle.Regular);
            XFont font7 = new XFont("Arial", 11, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 1;
            int iRows = 1;
            int i = 0;
            int X_Ini = 100;
            int X_Step = 100;
            int Y_Ini = 55;
            int Y_Step = 200;
            int X_Pie = 480;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = X;
                            double y = Y;

                            /*
                             * TITOLO
                             */
                            x = X;
                            y = Y + (iCol * Y_Step) - 45;
                            s = "PREZZO SPECIALE";
                            gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             * DESCRIZIONE
                             */
                            int iLen = 15;
                            s = (string)tabEti.Rows[i]["eti_ard"];

                            string s2 = "";

                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;

                            x = X + 50;
                            y = Y + (iCol * Y_Step);

                            XRect rect = new XRect(y, x, 750, 600);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            /*
                             * PREZZO
                             */
                            x = X + 390;
                            y = Y + (iCol * Y_Step) + 500;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);


                            /*
                             * EURO
                             */
                            x = X + 360;
                            y = Y + (iCol * Y_Step) - 20;
                            s = "€.";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);



                            /*
                            * PREZZO IN CHIARO
                            */
                            x = X + 370;
                            y = Y + (iCol * Y_Step) + 550;
                            d = 0;
                            //Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG" &&
                            if (
                                Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }


                            /*
                             * AL PZ
                             */
                            x = X + X_Pie - 40;
                            y = Y + (iCol * Y_Step) - 20;
                            s = "al pz.";
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s = "al Kg";
                            gfx.DrawString(s, font6, XBrushes.Black, y, x - 5, XStringFormats.Default);


                            /*
                             * TIPO GRAMMATURA
                             */
                            x = X + X_Pie - 43;
                            y = Y + (iCol * Y_Step) + 680;
                            s = (string)tabEti.Rows[i]["eti_tgr"];
                            s += " " + Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]).ToString("#0.00");
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             * LINEA
                             */
                            x = X + X_Pie - 40;
                            y = Y + (iCol * Y_Step) - 50;
                            XPen pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y, x, y + 850, x);


                            /*
                             * PIEDE DESCRIZIONE
                             */
                            x = X + X_Pie - 20;
                            y = Y + (iCol * Y_Step) - 40;
                            s = "Codice articolo                  PLU                             Codice a Barre                                       Pxc                                   Data di Stampa";
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * ARTICOLO
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) - 30;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * PLU
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 85;
                            s = (string)tabEti.Rows[i]["eti_plu"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * BARCODE
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 190;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * PXC
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 380;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                            * DATA
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 520;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            ///*
                            //* FORNITORE
                            //*/
                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 220;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * BARCODE
                            */
                            x = X + X_Pie - 25;
                            y = Y + (iCol * Y_Step) + 700;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                    else
                                    {
                                        s = s.PadLeft(13, Convert.ToChar('0'));
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }


                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 305;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti036(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

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
            XFont font1 = new XFont("Arial", 8, XFontStyle.Regular);
            XFont font2 = new XFont("Courier new", 15, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Arial", 15, XFontStyle.Regular);
            //XFont font4 = new XFont("Arial", 55, XFontStyle.BoldItalic);
            XFont font4 = new XFont("Franklin Gothic Demi Cond", 50, XFontStyle.Bold);

            XFont font5 = new XFont("Impact", 28, XFontStyle.BoldItalic);
            XFont font6 = new XFont("Arial", 10, XFontStyle.Bold);
            XFont font7 = new XFont("Impact", 10, XFontStyle.Regular);

            XFont font8 = new XFont("Arial", 10, XFontStyle.Regular);
            XFont font9 = new XFont("Arial", 9, XFontStyle.Italic);
            XFont font10 = new XFont("Arial", 10, XFontStyle.Bold);
            XFont font11 = new XFont("Arial", 8, XFontStyle.Bold);
            XFont font12 = new XFont("Arial", 24, XFontStyle.Bold);


            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 5;
            int i = 0;
            //int X_Ini = 70;
            int X_Ini = 63;
            int X_Step = 160;
            int Y_Ini = 50;
            int Y_Step = 245;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            s = "SELECT * FROM AnaIttico ORDER BY itt_art";
            DataTable tItt = _clsFun.FillTabSql("AnaIttico", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tItt.Columns["itt_art"];
            tItt.PrimaryKey = keys;

            DataRow[] jItt;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            if (i == 0)
                                Console.WriteLine("aaaa");

                            int iLen = 15;

                            double x = 0;
                            double y = 0;

                            /*
                             *  For + ARF
                             */
                            //y = Y + (iCol * Y_Step) + 130;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            string sDesArt = (string)tabEti.Rows[i]["eti_ard"];
                            string sDesCom = "";
                            string sDesLat = "";
                            string sDesPes = "";

                            jItt = tItt.Select("itt_art='" + (string)tabEti.Rows[i]["eti_art"] + "'");
                            if(jItt.Length > 0)
                            {
                                s = ((string)jItt[0]["itt_006"]).Trim();
                                if(s.Length > 40)
                                    sDesCom = s.Substring(40).Trim();

                                s = ((string)jItt[0]["itt_005"]).Trim();
                                if (s.Length > 40)
                                    sDesLat = s.Substring(40).Trim();

                                sDesPes = "";

                                s = ((string)jItt[0]["itt_002"]).Trim();
                                if (s.Length > 40)
                                    sDesPes += s.Substring(40).Trim();

                                //s = ((string)jItt[0]["itt_003"]).Trim();
                                //if (s.Length > 40)
                                //    sDesPes += s.Substring(40).Trim();

                                s = ((string)jItt[0]["itt_011"]).Trim();
                                if (s.Length > 40)
                                    sDesPes += " IN FAO " + s.Substring(40).Trim();
                                else
                                {
                                    s = ((string)jItt[0]["itt_004"]).Trim();
                                    if (s.Length > 40)
                                        sDesPes += " IN " + s.Substring(40).Trim();
                                }

                                s = ((string)jItt[0]["itt_003"]).Trim();
                                if (s.Length > 40)
                                    sDesPes += " CON " + s.Substring(40).Trim();
                                sDesPes = sDesPes.ToUpper();
                            }


                            /*
                             *  Descrizione articolo
                             */
                            //s = (string)tabEti.Rows[i]["eti_ard"];
                            //string s2 = "";
                            //if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                            //    Console.WriteLine("aaaa");
                            //s2 = s;
                            //for (int i3 = 0; i3 < 4; i3++)
                            //{
                            //    string[] a = s2.Split(' ');
                            //    s2 = "";

                            //    for (int i2 = 0; i2 < a.Length; i2++)
                            //    {
                            //        if (a[i2].Length > iLen)
                            //            a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                            //        s2 += a[i2] + " ";
                            //    }
                            //}

                            //if (s2 != "")
                            //    s = s2;


                            x = X + 3;
                            y = Y + (iCol * Y_Step) + 20;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            XRect rect = new XRect(y, x, 210, 77);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(sDesArt, font3, XBrushes.Black, rect);

                            x = X + 35;
                            y = Y + (iCol * Y_Step) + 20;
                            tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            rect = new XRect(y, x, 210, 77);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(sDesCom, font8, XBrushes.Black, rect);

                            x = X + 50;
                            y = Y + (iCol * Y_Step) + 20;
                            tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            rect = new XRect(y, x, 210, 77);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(sDesLat, font9, XBrushes.Black, rect);

                            x = X + 65;
                            y = Y + (iCol * Y_Step) + 20;
                            tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            rect = new XRect(y, x, 210, 77);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(sDesPes, font10, XBrushes.Black, rect);

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            /*
                             *  Riga
                             */
                            //x = X + 95;
                            x = X + 90;
                            y = Y + (iCol * Y_Step) - 20;
                            XPen pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y + 20, x, y + 270, x);

                            /*
                             *  UMI
                             */
                            //x = X + 108;
                            x = X + 103;
                            y = Y + (iCol * Y_Step) + 210;
                            //if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            //    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));
                            //s = "€ al kg/L " + d.ToString("#####0.00");
                            s = "al pz.";
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s = "al Kg";
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            ///*
                            // *  Euro
                            // */
                            //x = X + 132;
                            //y = Y + (iCol * Y_Step) + 60;
                            //s = "€";
                            //gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            /*
                             *  Prezzo
                             */
                            //x = X + 142;
                            x = X + 137;
                            y = Y + (iCol * Y_Step) + 195;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            /*
                             * BARCODE
                             * 
                            //x = X + 128;
                            x = X + 123;
                            y = Y + (iCol * Y_Step) + 15;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font11, XBrushes.Black, y, x, XStringFormats.Default);
                             */

                            /*
                             *  Codice articolo
                            //x = X + 138;
                            x = X + 133;
                            y = Y + (iCol * Y_Step) + 15;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                             */

                            /*
                             * PLU
                             */
                            //x = X + 148;
                            x = X + 110;
                            y = Y + (iCol * Y_Step) + 15;
                            s = ((string)tabEti.Rows[i]["eti_plu"]).Trim();
                            if (s != "" && _clsFun.Numerico(s))
                                s = Convert.ToInt16(tabEti.Rows[i]["eti_plu"]).ToString();
                            else
                                s = "";
                            gfx.DrawString(s, font12, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                             * ARF
                             */
                            //x = X + 148;
                            x = X + 143;
                            y = Y + (iCol * Y_Step) + 15;
                            s = ((string)tabEti.Rows[i]["eti_arf"]);
                            gfx.DrawString(s, font10, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             *  Data
                             */
                            x = X + 143;
                            y = Y + (iCol * Y_Step) + 200;
                            s = DateTime.Today.ToString("dd/MM/yy");
                            gfx.DrawString(s, font10, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                             * BARCODE
                             * 
                            //x = X + 128;
                            //x = X + 123;
                            //y = Y + (iCol * Y_Step) + 15;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font11, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 98;
                            x = X + 93;
                            y = Y + (iCol * Y_Step) + 14;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 21);
                                    }
                                    else
                                    {
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 21);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }
                             */

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 105;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti037(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 845.0;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 6.5, XFontStyle.Regular);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Impact", 60, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 110, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 60, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 20, XFontStyle.Regular);
            XFont font7 = new XFont("Arial", 25, XFontStyle.Bold);
            XFont font8 = new XFont("Arial", 30, XFontStyle.Bold);
            XFont font9 = new XFont("Impact", 50, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 1;
            int iRows = 2;
            int i = 0;
            int X_Ini = 80;
            int X_Step = 410;
            int Y_Ini = 30;
            int Y_Step = 200;
            int X_Pie = 330;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = X;
                            double y = Y + (iCol * Y_Step);
                            int iLen = 15;
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            //if (s.Length > 45)
                            //    s = s.Substring(0, 45);

                            string s2 = "";

                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;

                            XBrush brush = XBrushes.Transparent;

                            //XRect rect = new XRect(y, x - 50, 500, 100);
                            XRect rect = new XRect(y, x, 500, 200);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            if (false)
                            {
                                x = X + 90;
                                y = Y - 25; // +(iCol * Y_Step) + 190;
                                tf = new XTextFormatter(gfx);
                                rect = new XRect(y, x, 585, 60);
                                gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);

                                s = ((decimal)tabEti.Rows[i]["eti_pve"]).ToString("#,##0.00");
                                gfx.DrawString(s, font8, XBrushes.Black, y + 30, x + 50, XStringFormats.Default);

                                XPen pen = new XPen(XColors.Black, 2);
                                gfx.DrawLine(pen, y + 20, x + 60, y + 100, x + 25);

                                s = "OFFERTA";
                                gfx.DrawString(s, font9, XBrushes.White, y + 190, x + 50, XStringFormats.Default);

                                d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                s = "Sc." + d.ToString("#,##0.00") + " %";
                                gfx.DrawString(s, font7, XBrushes.Black, y + 450, x + 50, XStringFormats.Default);
                            }
                            //break;
                            /***/


                            x = X + 240;
                            y = Y + (iCol * Y_Step);
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 260;
                            y = Y + (iCol * Y_Step) + 400;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            x = X + 280;
                            y = Y + (iCol * Y_Step) + 370;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0 && d <= 500m)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }


                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 200;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 250;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie - 25;
                            y = Y + (iCol * Y_Step) + 420;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                    else
                                    {
                                        s = s.PadLeft(13, Convert.ToChar('0'));
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 5;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 58;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 431;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /***/

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti038(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

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
            XFont font1 = new XFont("Courier new", 11, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 15, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Franklin Gothic Demi Cond", 18, XFontStyle.Regular);
            //XFont font4 = new XFont("Arial", 55, XFontStyle.BoldItalic);
            XFont font4 = new XFont("Franklin Gothic Demi Cond", 60, XFontStyle.Bold);

            XFont font5 = new XFont("Impact", 28, XFontStyle.BoldItalic);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);
            //XFont font7 = new XFont("Impact", 10, XFontStyle.Regular);
            XFont font7 = new XFont("Arial", 24, XFontStyle.Bold);
            XFont font8 = new XFont("Courier new", 9, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 5;
            int i = 0;
            int X_Ini = 0;
            int X_Step = 158;
            int Y_Ini = 45;
            int Y_Step = 250;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            if (i == 1)
                                Console.WriteLine("aaaa");

                            int iLen = 15;

                            double x = 0;
                            double y = 0;

                            /*
                             *  For + ARF
                             */
                            //y = Y + (iCol * Y_Step) + 130;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             *  Descrizione articolo
                             */
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            string s2 = "";
                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;


                            x = X + 17;
                            y = Y + (iCol * Y_Step) + 20;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            XRect rect = new XRect(y, x, 210, 77);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);


                            /*
                             *  Riga
                             */
                            x = X + 58;
                            y = Y + (iCol * Y_Step) + 0;
                            XPen pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y + 20, x, y + 235, x);


                            /*
                             * Prezzo in chiaro
                             */

                            x = X + 72;
                            y = Y + (iCol * Y_Step) + 20;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0 && d <= 500m)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }


                            /*
                             *  UMI
                             */
                            x = X + 72;
                            y = Y + (iCol * Y_Step) + 200;
                            //if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            //    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));
                            //s = "€ al kg/L " + d.ToString("#####0.00");
                            s = "al pz.";
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s = "al Kg";
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             *  Euro
                             */
                            x = X + 110;
                            y = Y + (iCol * Y_Step) + 60;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);


                            /*
                             *  Prezzo
                             */
                            x = X + 130;
                            y = Y + (iCol * Y_Step) + 220;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);


                            /*
                             * Barcode
                             */
                            //x = X + 129;
                            //x = X + 80;
                            //y = Y + (iCol * Y_Step) + 20;
                            //s = ((string)tabEti.Rows[i]["eti_ean"]);
                            //gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            /*
                             * BARCODE
                             * 
                             */
                            x = X + 130;
                            y = Y + (iCol * Y_Step) + 20;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font8, XBrushes.Black, y, x, XStringFormats.Default);
                            x = X + 110;
                            y = Y + (iCol * Y_Step) + 20;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 10);
                                    }
                                    else
                                    {
                                        s = s.PadLeft(13, Convert.ToChar('0'));
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 10);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            /*
                             *  Codice articolo
                             */
                            x = X + 140;
                            y = Y + (iCol * Y_Step) + 20;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             *  Codice articolo fornitore
                             */
                            x = X + 140;
                            y = Y + (iCol * Y_Step) + 100;
                            s = ((string)tabEti.Rows[i]["eti_arf"]);
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             *  Data
                             */
                            x = X + 140;
                            y = Y + (iCol * Y_Step) + 180;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);



                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 105;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti039(DataTable tabEti, string strPar)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 845.0;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 6.5, XFontStyle.Regular);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Arial", 9, XFontStyle.Bold);
            XFont font4 = new XFont("Times New Roman", 30, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 10, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 8, XFontStyle.Regular);
            XFont font7 = new XFont("Courier new", 7.5, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 4;
            int iRows = 11;
            int i = 0;
            int X_Ini = 5;
            int X_Step = 75;
            int Y_Ini = 15;
            int Y_Step = 140;
            int i_Step = 0;

            if (strPar != "")
            {
                string[] a = strPar.Split(';');
                foreach (string ss in a)
                {
                    string[] aa = ss.Split('=');
                    if (aa.Length > 1)
                    {
                        if (_clsFun.Numerico(aa[1], "0123456789"))
                        {
                            i = Convert.ToInt16(aa[1]);
                            if (i > 0)
                            {
                                if (aa[0] == "xIni")
                                    X_Ini = i;
                                if (aa[0] == "xStep")
                                    X_Step = i;
                                if (aa[0] == "yIni")
                                    Y_Ini = i;
                                if (aa[0] == "yStep")
                                    Y_Step = i;
                                if (aa[0] == "iCols")
                                    iCols = i;
                                if (aa[0] == "iRows")
                                    iRows = i;
                            }
                        }
                    }
                }
            }

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;
            i = 0;

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step) + i_Step;
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = X;
                            double y = Y + (iCol * Y_Step);
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            if (s.Length > 45)
                                s = s.Substring(0, 45);

                            //string s2 = "";
                            //string[] a = s.Split(' ');

                            //if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                            //    Console.WriteLine("aaaa");

                            //for (int i2 = 0; i2 < a.Length; i2++)
                            //{
                            //    if (a[i2].Length > 20)
                            //        a[i2] = a[i2].Substring(0, 10) + " " + a[i2].Substring(10, 10) + " " + a[i2].Substring(20);

                            //    else if (a[i2].Length > 10)
                            //        a[i2] = a[i2].Substring(0, 10) + " " + a[i2].Substring(10);

                            //    s2 += a[i2] + " ";
                            //}

                            //s = s2;
                            if (i == 1)
                                Console.WriteLine("aaaa");

                            //gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);
                            //XRect rect = new XRect(40, 100, 250, 220);

                            /*
                             * Descrizione
                             */

                            XTextFormatter tf = new XTextFormatter(gfx);
                            XRect rect = new XRect(y + 20, x, 120, 35);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //tf.Alignment = ParagraphAlignment.Left; 
                            tf.DrawString(s, font3, XBrushes.Black, rect, XStringFormats.TopLeft);


                            /*
                             * Barcode
                             */

                            x = X + 35;
                            y = Y + (iCol * Y_Step) + 5;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 60, 18);
                                    }
                                    else
                                    {
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 60, 18);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            /*
                             * Barcode
                             */

                            x = X + 60;
                            y = Y + (iCol * Y_Step) + 5;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             * Data
                             */

                            x = X + 30;
                            y = Y + (iCol * Y_Step) + 5;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             * Codice articolo
                             */

                            //x = X + 56;
                            y = Y + (iCol * Y_Step) + 40;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                            * Euro

                            x = X + 35;
                            y = Y + (iCol * Y_Step) + 75;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);
                            */


                            /*
                             * Prezzo
                             */

                            x = X + 65;
                            y = Y + (iCol * Y_Step) + 130;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            /*
                             * Prezzo in chiaro
                             */

                            x = X + 30;
                            y = Y + (iCol * Y_Step) + 75;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0 && d <= 500m)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }




                            /*
                             * PxC

                            x = X + 81;
                            y = Y + (iCol * Y_Step) + 0;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                             */


                            /*
                             * Fornitore

                            //x = X + 83;
                            y = Y + (iCol * Y_Step) + 30;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);
                             */




                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                pd.Close();
                pd.Dispose();
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti040(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

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
            //XFont font1 = new XFont("Lucida Calligraphy", 22, XFontStyle.Italic);
            XFont font1 = new XFont("A.C.M.E. Secret Agent", 22, XFontStyle.Italic);

            XFont font2 = new XFont("Lucida Calligraphy", 26, XFontStyle.Italic);
            XFont font3 = new XFont("Cooper Black", 20, XFontStyle.Regular);
            //XFont font4 = new XFont("Arial", 55, XFontStyle.BoldItalic);
            XFont font4 = new XFont("Impact", 30, XFontStyle.Bold);

            XFont font5 = new XFont("Impact", 28, XFontStyle.BoldItalic);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);
            //XFont font7 = new XFont("Impact", 10, XFontStyle.Regular);
            XFont font7 = new XFont("Arial", 16, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 5;
            int i = 0;
            int X_Ini = 38;
            int X_Step = 144;
            int Y_Ini = 33;
            int Y_Step = 265;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            XPen pen = new XPen(XColors.Black, 0.8);

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {

                            int iLen = 15;

                            double x = X + 14;
                            double y = Y + (iCol * Y_Step) + 2;

                            //gfx.DrawRectangle(pen, Y + (iCol * Y_Step), x - 10, Y + (iCol * Y_Step) + 100, x + 50);

                            //if (i <= 1)
                                gfx.DrawRectangle(pen, y, x, Y_Step, X_Step);

                            /* Immagine logo */
                            s = "C:\\ApProject\\Temp\\Img\\Etichette\\eti040.png";
                            XImage img = XImage.FromFile(s);
                            x = X + 40;
                            y = Y + (iCol * Y_Step) + 130;
                            XRect rect1 = new XRect(y, x, 145, 105);
                            //XRect rect1 = new XRect(y, x, 145, 105);
                            gfx.DrawImage(img, rect1);

                            x = X + 14;
                            y = Y + (iCol * Y_Step) + 2;

                            //gfx.DrawRectangle(pen, Y + (iCol * Y_Step), x - 10, Y + (iCol * Y_Step) + 100, x + 50);

                            //if (i <= 1)
                            gfx.DrawRectangle(pen, y, x, Y_Step, X_Step);



                            /*
                             *  For + ARF
                             */
                            //y = Y + (iCol * Y_Step) + 130;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             *  Descrizione articolo
                             */


                            s = (string)tabEti.Rows[i]["eti_ard"];

                            //s = s.ToLowerInvariant();

                            s = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLower());

                            /*
                            string s2 = "";
                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            s = s.Trim();

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            x = X + 14;
                            y = Y + (iCol * Y_Step) + 20;

                            string[] aa = s.Split(' ');

                            if (aa.Length <= 2 && s.Length < 19)
                                x += 30;
                            */

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            //font1.Height = 25.0;


                            //XStringFormat format = new XStringFormat();
                            //format.Alignment = XStringAlignment.Center;

                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Left;
                            //XRect rect = new XRect(y, x, 300, 100);
                            XRect rect = new XRect(y, x+20, 275, 100);
                            //gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font1, XBrushes.Black, rect);


                            //rect = new XRect(y, x, 210, 77);
                            //gfx.DrawRectangle(XBrushes.SeaShell, rect);
                            //tf.Alignment = XParagraphAlignment.Center;
                            //tf.DrawString(s, font1, XBrushes.Black, rect, XStringFormats.TopLeft);


                            /*
                             *  Riga
                            x = X + 58;
                            y = Y + (iCol * Y_Step) + 5;
                            XPen pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y + 20, x, y + 230, x);
                             */


                            /*
                             *  Euro
                            x = X + 100;
                            y = Y + (iCol * Y_Step) + 60;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);
                             */


                            /*
                             *  Prezzo
                             */
                            x = X + 132;
                            y = Y + (iCol * Y_Step) + 70;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font2, XBrushes.Black, y + 10, x, frmDX);


                            /*
                             *  UMI
                             */
                            x = X + 125;
                            y = Y + (iCol * Y_Step) + 82;
                            //if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            //    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));
                            //s = "€ al kg/L " + d.ToString("#####0.00");
                            s = "al Pz.";
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s = "al Kg";
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             * PLU
                            //x = X + 129;
                            x = X + 80;
                            y = Y + (iCol * Y_Step) + 20;
                            s = ((string)tabEti.Rows[i]["eti_plu"]);
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);
                             */


                            /*
                             *  Codice articolo
                            x = X + 103;
                            y = Y + (iCol * Y_Step) + 20;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                             */


                            /*
                             *  Codice articolo fornitore
                            //x = X + 140;
                            y = Y + (iCol * Y_Step) + 100;
                            s = ((string)tabEti.Rows[i]["eti_arf"]);
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                             */


                            /*
                             *  Data
                            //x = X + 140;
                            y = Y + (iCol * Y_Step) + 175;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                             */


                            /*
                             * BARCODE
                             * 
                             */
                            //y = Y + (iCol * Y_Step) + 95;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            //x = X + 120;
                            //y = Y + (iCol * Y_Step) + 175;
                            //s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            //if (s.Trim().Length > 0)
                            //{
                            //    try
                            //    {
                            //        if (s.Length <= 8)
                            //        {
                            //            Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //        else
                            //        {
                            //            Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                            //    }
                            //}

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 105;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti041(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            //page.Height = 845.0;

            //page.Height = 842.0;
            //page.Width = 595;
            //page.Orientation = PdfSharp.PageOrientation.Landscape;

            page.Height = 845.0;
            page.Orientation = PdfSharp.PageOrientation.Portrait;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Amatic", 30, XFontStyle.Bold);
            XFont font2 = new XFont("Amatic", 30, XFontStyle.Bold);
            XFont font3 = new XFont("Amatic", 70, XFontStyle.Bold);
            XFont font4 = new XFont("Amatic", 25, XFontStyle.Bold);
            XFont font5 = new XFont("Amatic", 20, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            //int iCols = 2;
            //int iRows = 4;
            //int i = 0;
            //int X_Ini = 70;
            //int X_Step = 210;
            //int Y_Ini = 10;
            //int Y_Step = 300;

            int iCols = 2;
            int iRows = 4;
            int i = 0;
            int X_Ini = 0;
            int X_Step = 210;
            int Y_Ini = 0;
            int Y_Step = 300;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //if (DBNull.Value.Equals(tabEti.Rows[i]["eti_cam"]))
                //    tabEti.Rows[i]["eti_cam"] = "";

                if (tabEti.Rows.Count > 0)
                {
                    if (iRow >= iRows)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            page = pd.AddPage();
                            gfx = XGraphics.FromPdfPage(page);
                        }
                        X = X_Ini;
                        iRow = 0;
                    }
                }
                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            if (i == 0)
                                Console.WriteLine("zzzz");

                            /*
                             * Immagine di sfondo 
                             */
                            double x = X;
                            double y = Y + (iCol * Y_Step);

                            s = "C:\\ApProject\\Temp\\Img\\Godina\\eti041fidelity.png";
                            if ((string)tabEti.Rows[i]["eti_cam"] != "003" || !File.Exists(s))
                            //if ((string)tabEti.Rows[i]["eti_cam"] != "001" || !File.Exists(s))
                                    //    s = "C:\\ApProject\\Temp\\Img\\Godina\\eti041fidelity.jpg";
                            //else
                                s = "C:\\ApProject\\Temp\\Img\\Godina\\eti041.jpg";

                            XImage img = XImage.FromFile(s);
                            x = X;
                            y = Y + (iCol * Y_Step) + 0;
                            XRect rect1 = new XRect(y, x, 300, 210);
                            gfx.DrawImage(img, rect1);

                            s = "C:\\ApProject\\Temp\\Img\\Godina\\eti041tessera.png";
                            if ((string)tabEti.Rows[i]["eti_cam"] == "003" && File.Exists(s))
                            //if ((string)tabEti.Rows[i]["eti_cam"] == "001" && File.Exists(s))
                            {
                                img = XImage.FromFile(s);
                                x = X;
                                y = Y + (iCol * Y_Step) +150;
                                rect1 = new XRect(y, x, 150, 65);
                                gfx.DrawImage(img, rect1);
                            }

                            /*
                             * Descrizione articolo
                             */
                            x = X;
                            y = Y + (iCol * Y_Step);
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            if (s.Length > 45)
                                s = s.Substring(0, 45);

                            //string s2 = "";
                            //string[] a = s.Split(' ');

                            //if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                            //    Console.WriteLine("aaaa");

                            //for (int i2 = 0; i2 < a.Length; i2++)
                            //{
                            //    if (a[i2].Length > 20)
                            //        a[i2] = a[i2].Substring(0, 10) + " " + a[i2].Substring(10, 10) + " " + a[i2].Substring(20);

                            //    else if (a[i2].Length > 10)
                            //        a[i2] = a[i2].Substring(0, 10) + " " + a[i2].Substring(10);

                            //    s2 += a[i2] + " ";
                            //}

                            //s = s2;

                            //XRect rect = new XRect(40, 100, 250, 220);
                            //gfx.DrawRectangle(XBrushes.SeaShell, rect);
                            //tf.DrawString(text, font, XBrushes.Black, rect, XStringFormats.TopLeft);

                            x = X + 70;
                            y = Y + (iCol * Y_Step) + 90;
                            XRect rect = new XRect(y + 0, x, 210, 80);
                            XTextFormatter tf = new XTextFormatter(gfx);
                            //gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.Alignment = XParagraphAlignment.Center;
                            tf.DrawString(s, font1, XBrushes.Black, rect, XStringFormats.TopLeft);

                            /*
                             * Date offerta

                            if ((string)tabEti.Rows[i]["eti_oft"] == "PRO" || ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ))
                            //if ((string)tabEti.Rows[i]["eti_off"] != "" && ((string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ || (string)tabEti.Rows[i]["eti_oft"] == "PRO"))
                            {
                                x = X + 0;
                                y = Y + (iCol * Y_Step) + 15;
                                s = "Offerta valida";
                                gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);
                                x = X + 10;
                                y = Y + (iCol * Y_Step) + 15;
                                s = "dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yyyy");
                                gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);
                                x = X + 20;
                                y = Y + (iCol * Y_Step) + 15;
                                s = "al  " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");
                                gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);

                                d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                x = X + 45;
                                y = Y + (iCol * Y_Step) + 110;
                                s = "da € " + Convert.ToDecimal(tabEti.Rows[i]["eti_pve"]).ToString("######0.00") + " - ";
                                s += "sconto " + d.ToString("#,##0.00") + " %";
                                gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);
                            }
                             */

                            if ((string)tabEti.Rows[i]["eti_oft"] == "002")
                            {
                                /*
                                 * %
                                 */
                                if ((decimal)tabEti.Rows[i]["eti_pve"] > 0)
                                {
                                    x = X + 45;
                                    y = Y + (iCol * Y_Step) + 200;
                                    d = (decimal)tabEti.Rows[i]["eti_pve"];
                                    //s = "da € " + Convert.ToDecimal(tabEti.Rows[i]["eti_pve"]).ToString("######0.00") + " - ";
                                    s = "Sconto " + d.ToString("#,##0.00") + " %";
                                    gfx.DrawString(s, font4, XBrushes.White, y, x, XStringFormats.Default);
                                }
                            }
                            else if ((string)tabEti.Rows[i]["eti_oft"] == "001")
                            {
                                /*
                                 * %
                                 */
                                if ((decimal)tabEti.Rows[i]["eti_pve"] > 0)
                                {
                                    x = X + 80;
                                    y = Y + (iCol * Y_Step) + 230;
                                    s = ((decimal)tabEti.Rows[i]["eti_pve"]).ToString("#,##0.00");
                                    gfx.DrawString(s, font5, XBrushes.Black, y + 60, x + 50, frmDX);

                                    XPen pen = new XPen(XColors.Black, 2);
                                    gfx.DrawLine(pen, y + 35, x + 50, y + 65, x + 35);

                                    //s = "OFFERTA";
                                    //gfx.DrawString(s, font4, XBrushes.White, y + 190, x + 50, XStringFormats.Default);

                                    d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                    s = "-" + d.ToString("#,##0") + " %";
                                    gfx.DrawString(s, font5, XBrushes.Black, y + 60, x + 70, frmDX);

                                }
                            }

                            /*
                             * PxC
                            if ((decimal)tabEti.Rows[i]["eti_pxc"] > 0)
                            {
                                x = X + 45;
                                y = Y + (iCol * Y_Step) + 15;
                                s = "C " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                                gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            }
                             */

                            ///*
                            // * Data
                            // */
                            //x = X + 48;
                            //y = Y + (iCol * Y_Step) + 0;
                            //s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                             * Riga
                            x = X + 34;
                            y = Y + (iCol * Y_Step) + 83;
                            XPen pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y + 20, x, y + 190, x);
                             */

                            /*
                             * Riga 2
                            x = X + 110;
                            y = Y + (iCol * Y_Step) + 83;
                            pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y + 20, x, y + 190, x);
                             */

                            /*
                             * Barcode
                            x = X + 50;
                            y = Y + (iCol * Y_Step) + 15;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y - 5, x, 70, 18);
                                    }
                                    else
                                    {
                                        s = s.PadLeft(13, Convert.ToChar('0'));

                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }
                             */

                            /*
                             *  Tipo grammatura + Peso netto
                            if ((decimal)tabEti.Rows[i]["eti_pne"] > 0)
                            {
                                x = X + 35;
                                y = Y + (iCol * Y_Step) + 15;
                                s = (string)tabEti.Rows[i]["eti_tgr"];
                                s += " " + Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]).ToString("#0.00");
                                gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            }
                             */

                            /*
                             * Euro
                             */
                            //x = X + 75;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "€";
                            //gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            /*
                             * Prezzo
                             */
                            x = X + 200;
                            y = Y + (iCol * Y_Step) + 230;

                            d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]);
                            if ((string)tabEti.Rows[i]["eti_oft"] == "002" && Convert.ToDecimal(tabEti.Rows[i]["eti_pve"]) > 0)
                                d = _clsFun.MenoPer(d, Convert.ToDecimal(tabEti.Rows[i]["eti_pve"]));
                            s = "€" + d.ToString("######0.00");
                            gfx.DrawString(s, font3, XBrushes.Black, y + 10, x, frmDX);

                            /*
                             * Codice articolo
                            x = X + 89;
                            y = Y + (iCol * Y_Step) + 15;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                             */

                            /*
                             * Fornitore + ARF
                            x = X + 89;
                            y = Y + (iCol * Y_Step) + 45;
                            s = "";
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();

                            s = ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            if (s != "")
                                s = "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);
                             */

                            /*
                             * Barcode
                            x = X + 75;
                            y = Y + (iCol * Y_Step) + 20;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                             */

                            /*
                             * Prezzo al KG
                            x = X + 80;
                            y = Y + (iCol * Y_Step) + 240;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            {
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));
                                s = "€ " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            }
                            x = X + 90;
                            y = Y + (iCol * Y_Step) + 240;
                            s = "al kg/L ";
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            i++;
                             */

                            /*
                             * Peso 
                             */
                            x = X + 160;
                            y = Y + (iCol * Y_Step) + 65;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            {
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]);
                                s = (string)tabEti.Rows[i]["eti_tgr"];
                                s += " " + d.ToString("#####0");
                                gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);
                            }

                            i++;

                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti042(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            //Barcode _ean13 = new BarcodeLib.Barcode();
            //BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            //_ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            //Barcode _ean08 = new BarcodeLib.Barcode();
            //BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            //_ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 842.0;
            page.Width = 595;
            page.Orientation = PdfSharp.PageOrientation.Landscape;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Amatic", 100, XFontStyle.Bold);
            XFont font2 = new XFont("Amatic", 120, XFontStyle.Bold);
            XFont font3 = new XFont("Amatic", 70, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 1;
            int iRows = 1;
            int i = 0;
            int X_Ini = 0;
            int X_Step = 100;
            int Y_Ini = 0;
            int Y_Step = 200;
            int X_Pie = 735;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            s = "C:\\ApProject\\Temp\\Img\\Godina\\eti042.jpg";

                            XImage img = XImage.FromFile(s);
                            double x = X;
                            double y = Y + (iCol * Y_Step) + 0;
                            XRect rect1 = new XRect(y, x, 830, 595);
                            gfx.DrawImage(img, rect1);

                            x = X;
                            y = Y + (iCol * Y_Step);
                            int iLen = 15;
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            //if (s.Length > 45)
                            //    s = s.Substring(0, 45);

                            string sDes = "";

                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            sDes = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = sDes.Split(' ');
                                sDes = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    sDes += a[i2] + " ";
                                }
                            }

                            if (sDes != "")
                                s = sDes;

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            //x = X - 100;
                            //y = Y + (iCol * Y_Step) - 55;

                            //XRect rect = new XRect(y, x, page.Width, 80);
                            //gfx.DrawRectangle(XBrushes.Red, rect);

                            //s = "Migross conviene";
                            //s = strMsg;
                            //gfx.DrawString(s, font7, XBrushes.White, y + 7, x + 60, XStringFormats.Default);

                            //s = "M3G";
                            //gfx.DrawString(s, font7, XBrushes.White, y + 500, x + 60, XStringFormats.Default);

                            s = "C:\\ApProject\\Img\\" + (string)tabEti.Rows[i]["eti_art"] + ".jpg";
                            s = (string)tabEti.Rows[i]["eti_art"] + "*.jpg";

                            string[] aa = Directory.GetFiles("C:\\ApProject\\Img\\", s);

                            //if (File.Exists(s))
                            if (aa.Length > 0)
                            {
                                s = aa[0];

                                if ((string)tabEti.Rows[i]["eti_art"] == "0005933")
                                    Console.WriteLine("aaaa");
                                if (i == 0)
                                    Console.WriteLine("aaaa");

                                img = XImage.FromFile(s);
                                x = X + 45;
                                y = Y + (iCol * Y_Step) + 30;
                                rect1 = new XRect(y, x, 420, 420);
                                gfx.DrawImage(img, rect1);
                            }

                            x = X + 180;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            XRect rect = new XRect(y, x, 300, 400);
                            //gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(sDes, font1, XBrushes.White, rect);

                            //x = X + 410;
                            //y = Y + (iCol * Y_Step) - 35;

                            //if ((string)tabEti.Rows[i]["eti_bcl"] != "")
                            //{
                            //    s = "Calibro: " + (string)tabEti.Rows[i]["eti_bcl"];
                            //    gfx.DrawString(s, font9, XBrushes.Black, y, x, XStringFormats.Default);
                            //}
                            //if ((string)tabEti.Rows[i]["eti_bct"] != "")
                            //{
                            //    s = "Categoria: " + (string)tabEti.Rows[i]["eti_bct"];
                            //    gfx.DrawString(s, font9, XBrushes.Black, y, x + 25, XStringFormats.Default);
                            //}
                            //if ((string)tabEti.Rows[i]["eti_bor"] != "")
                            //{
                            //    s = "Origine: " + (string)tabEti.Rows[i]["eti_bor"];
                            //    gfx.DrawString(s, font9, XBrushes.Black, y, x + 50, XStringFormats.Default);
                            //}

                            //x = X + 470;
                            //y = Y + (iCol * Y_Step) - 35;
                            //gfx.DrawRoundedRectangle(XBrushes.Gold, y, x, 550, 240, 30, 20);

                            x = X + 550;
                            y = Y + (iCol * Y_Step) + 190;
                            s = "" + Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font2, XBrushes.White, y + 10, x, frmDX);

                            //x = X + 630;
                            //y = Y + (iCol * Y_Step) - 20;
                            //s = "€";
                            //gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            //x = X + 675;
                            //y = Y + (iCol * Y_Step) + 360;
                            //d = 0;
                            //if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG" &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            //    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            //if (d > 0)
                            //{
                            //    s = "€ al kg/L " + d.ToString("#####0.00");
                            //    gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            //}

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 180;
                            //s = (string)tabEti.Rows[i]["eti_art"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 220;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 5;

                            ////s = "";
                            ////if ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ)
                            ////    s = "Dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yyyy") + " al " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");
                            ////gfx.DrawString(s, font9, XBrushes.White, y - 50, x - 15, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 58;
                            //s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            ////x = X + X_Pie;
                            ////y = Y + (iCol * Y_Step) + 305;
                            ////s = (string)tabEti.Rows[i]["eti_ean"];
                            ////gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 470;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti044(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";
            DataRow[] j;

            DataTable tEcr = _clsQry.tabEcr();

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            //page.Height = 845.0;

            //page.Height = 842.0;
            //page.Width = 595;
            //page.Orientation = PdfSharp.PageOrientation.Landscape;

            //page.Height = 845.0;
            //page.Orientation = PdfSharp.PageOrientation.Portrait;

            page.Height = 845.0;
            page.Orientation = PdfSharp.PageOrientation.Portrait;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Arial", 18, XFontStyle.Bold);
            XFont font2 = new XFont("Calibri", 14, XFontStyle.Regular);
            XFont font3 = new XFont("Arimo", 36, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 8, XFontStyle.Regular);
            XFont font5 = new XFont("Arimo", 18, XFontStyle.Bold);
            XFont font6 = new XFont("Arimo", 20, XFontStyle.Bold);
            XFont font7 = new XFont("Amatic", 26, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 1;
            int iRows = 4;
            int i = 0;
            int X_Ini = 50;
            int X_Step = 185;
            int Y_Ini = 75;
            int Y_Step = 300;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {

                            if (i == 0)
                                Console.WriteLine("zzzz");

                            /*
                             * Immagine di sfondo 
                             */
                            double x = X;
                            double y = Y + (iCol * Y_Step);

                            s = "C:\\ApProject\\Temp\\Img\\Godina\\eti044.jpg";

                            XImage img = XImage.FromFile(s);
                            x = X;
                            y = Y + (iCol * Y_Step) + 0;
                            XRect rect1 = new XRect(y, x, 470, 180);
                            gfx.DrawImage(img, rect1);

                            /*
                             * Descrizione articolo
                             */
                            x = X;
                            y = Y + (iCol * Y_Step);
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            if (s.Length > 45)
                                s = s.Substring(0, 45);

                            //string s2 = "";
                            //string[] a = s.Split(' ');

                            //if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                            //    Console.WriteLine("aaaa");

                            //for (int i2 = 0; i2 < a.Length; i2++)
                            //{
                            //    if (a[i2].Length > 20)
                            //        a[i2] = a[i2].Substring(0, 10) + " " + a[i2].Substring(10, 10) + " " + a[i2].Substring(20);

                            //    else if (a[i2].Length > 10)
                            //        a[i2] = a[i2].Substring(0, 10) + " " + a[i2].Substring(10);

                            //    s2 += a[i2] + " ";
                            //}

                            //s = s2;

                            //XRect rect = new XRect(40, 100, 250, 220);
                            //gfx.DrawRectangle(XBrushes.SeaShell, rect);
                            //tf.DrawString(text, font, XBrushes.Black, rect, XStringFormats.TopLeft);
                            if (i == 0)
                                Console.WriteLine("aaaa");

                            x = X + 15;
                            y = Y + (iCol * Y_Step) + 130;
                            XRect rect = new XRect(y + 0, x, 300, 80);
                            XTextFormatter tf = new XTextFormatter(gfx);
                            //gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.Alignment = XParagraphAlignment.Center;
                            tf.DrawString(s, font1, XBrushes.Red, rect, XStringFormats.TopLeft);

                            /* Riga */
                            x = X + 65;
                            y = Y + (iCol * Y_Step) -15;
                            XPen pen = new XPen(XColors.Red, 1);
                            gfx.DrawLine(pen, y + 20, x, y + 480, x);

                            /* dicitura P R E Z Z O */
                            x = X + 80;
                            y = Y + (iCol * Y_Step) + 5;
                            s = "P   R   E   Z   Z   O";
                            gfx.DrawString(s, font2, XBrushes.White, y, x, XStringFormats.Default);

                            /*
                             * Prezzo
                             */
                            x = X + 140;
                            y = Y + (iCol * Y_Step) - 5;
                            //s = "€" + Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");

                            s = (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / 10).ToString("##0.00");
                            if ((string)tabEti.Rows[i]["eti_umi"] == "NR")
                                s = (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"])).ToString("##0.00");
                            string s1 = s.Substring(0, s.IndexOf(",") + 1);
                            gfx.DrawString(s1, font3, XBrushes.White, y + 10, x, XStringFormats.Default);
                            
                            y = y + 18 + (21 * (s1.Length-1));

                            string s2 = s.Substring(s.IndexOf(",") + 1);
                            gfx.DrawString(s2, font6, XBrushes.White, y, x-10, XStringFormats.Default);

                            s = "€\\";
                            if ((string)tabEti.Rows[i]["eti_umi"] == "NR")
                                s += "Pz";
                            else
                                s += "Hg";
                            gfx.DrawString(s, font5, XBrushes.White, y + 30, x, XStringFormats.Default);

                            if (i == 0)
                                Console.WriteLine("zzzz");

                            /*
                             * Ingredienti
                             */
                            s = _clsQry.GetFileIngredienti((string)tabEti.Rows[i]["eti_art"]);
                            if (s.Length > 0)
                            {
                                x = X + 70;
                                y = Y + (iCol * Y_Step) + 100;
                                rect = new XRect(y + 0, x, 350, 80);
                                tf = new XTextFormatter(gfx);
                                //gfx.DrawRectangle(XBrushes.Transparent, rect);
                                tf.Alignment = XParagraphAlignment.Right;
                                tf.DrawString(s, font4, XBrushes.White, rect, XStringFormats.TopLeft);
                            }

                            /*
                             * ECR
                             */

                            x = X + 170;
                            y = Y + (iCol * Y_Step) + 5;

                            s = ((string)tabEti.Rows[i]["eti_ecr"]).Trim();

                            if (s != "" && s.Length > 8)
                            {
                                string sLv1 = s.Substring(0, 3);
                                string sLv2 = s.Substring(3, 3);
                                string sLv3 = s.Substring(6, 3);

                                s = "l1c='" + sLv1 + "' AND l2c='" + sLv2 + "'AND l3c='" + sLv3 + "'";
                                j = tEcr.Select(s);
                                if (j.Length > 0)
                                {
                                    s = ((string)(string)j[0]["l3d"]).ToUpper();
                                    gfx.DrawString(s, font7, XBrushes.Red, y, x, XStringFormats.Default);
                                }   
                            }

                            /*
                             * Date offerta

                            if ((string)tabEti.Rows[i]["eti_oft"] == "PRO" || ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ))
                            //if ((string)tabEti.Rows[i]["eti_off"] != "" && ((string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ || (string)tabEti.Rows[i]["eti_oft"] == "PRO"))
                            {
                                x = X + 0;
                                y = Y + (iCol * Y_Step) + 15;
                                s = "Offerta valida";
                                gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);
                                x = X + 10;
                                y = Y + (iCol * Y_Step) + 15;
                                s = "dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yyyy");
                                gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);
                                x = X + 20;
                                y = Y + (iCol * Y_Step) + 15;
                                s = "al  " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");
                                gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);

                                d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                x = X + 45;
                                y = Y + (iCol * Y_Step) + 110;
                                s = "da € " + Convert.ToDecimal(tabEti.Rows[i]["eti_pve"]).ToString("######0.00") + " - ";
                                s += "sconto " + d.ToString("#,##0.00") + " %";
                                gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);
                            }
                             */

                            /*
                             * PxC
                            if ((decimal)tabEti.Rows[i]["eti_pxc"] > 0)
                            {
                                x = X + 45;
                                y = Y + (iCol * Y_Step) + 15;
                                s = "C " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                                gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            }
                             */

                            ///*
                            // * Data
                            // */
                            //x = X + 48;
                            //y = Y + (iCol * Y_Step) + 0;
                            //s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                             * Riga
                            x = X + 34;
                            y = Y + (iCol * Y_Step) + 83;
                            XPen pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y + 20, x, y + 190, x);
                             */

                            /*
                             * Riga 2
                            x = X + 110;
                            y = Y + (iCol * Y_Step) + 83;
                            pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y + 20, x, y + 190, x);
                             */

                            /*
                             * Barcode
                            x = X + 50;
                            y = Y + (iCol * Y_Step) + 15;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y - 5, x, 70, 18);
                                    }
                                    else
                                    {
                                        s = s.PadLeft(13, Convert.ToChar('0'));

                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }
                             */

                            /*
                             *  Tipo grammatura + Peso netto
                            if ((decimal)tabEti.Rows[i]["eti_pne"] > 0)
                            {
                                x = X + 35;
                                y = Y + (iCol * Y_Step) + 15;
                                s = (string)tabEti.Rows[i]["eti_tgr"];
                                s += " " + Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]).ToString("#0.00");
                                gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            }
                             */

                            /*
                             * Euro
                             */
                            //x = X + 75;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "€";
                            //gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);


                            /*
                             * Codice articolo
                            x = X + 89;
                            y = Y + (iCol * Y_Step) + 15;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                             */

                            /*
                             * Fornitore + ARF
                            x = X + 89;
                            y = Y + (iCol * Y_Step) + 45;
                            s = "";
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();

                            s = ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            if (s != "")
                                s = "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);
                             */

                            /*
                             * Barcode
                            x = X + 75;
                            y = Y + (iCol * Y_Step) + 20;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                             */

                            /*
                             * Prezzo al KG
                            x = X + 80;
                            y = Y + (iCol * Y_Step) + 240;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            {
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));
                                s = "€ " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            }
                            x = X + 90;
                            y = Y + (iCol * Y_Step) + 240;
                            s = "al kg/L ";
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            i++;
                             */

                            /*
                             * Peso 
                            x = X + 160;
                            y = Y + (iCol * Y_Step) + 65;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            {
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]);
                                s = (string)tabEti.Rows[i]["eti_tgr"];
                                s += " " + d.ToString("#####0");
                                gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);
                            }
                             */

                            i++;

                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti045(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            //page.Height = 845.0;

            page.Height = 842.0;
            page.Width = 595;
            page.Orientation = PdfSharp.PageOrientation.Landscape;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 9.5, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 15, XFontStyle.Regular);
            //XFont font3 = new XFont("Verdana", 18, XFontStyle.Bold);
            XFont font3 = new XFont("Times New Roman", 22, XFontStyle.Bold);
            XFont font4 = new XFont("Times New Roman", 46, XFontStyle.Bold);
            XFont font5 = new XFont("Times New Roman", 20, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 3;
            int i = 0;
            int X_Ini = 0;
            int X_Step = 161;
            int Y_Ini = 0;
            int Y_Step = 370;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = 0;
                            double y = 0;
                            int iLen = 15;
                            //double x = X;
                            //double y = Y + (iCol * Y_Step);
                            //int iLen = 15;
                            //x = X + 10;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 20;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = (string)tabEti.Rows[i]["eti_art"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);



                            x = X + 64;
                            y = Y + (iCol * Y_Step) + 62;
                            XRect rect = new XRect(y, x, 336, 118);
                            XPen pen = new XPen(XColors.Green, 8.9);
                            gfx.DrawRectangle(pen, rect);

                            //break;
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            string s2 = "";
                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            //x = X + 42;
                            //y = Y + (iCol * Y_Step) + 0;
                            x = X + 70;
                            y = Y + (iCol * Y_Step) + 68;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            rect = new XRect(y, x, 300, 100);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            x = X + 150;
                            y = Y + (iCol * Y_Step) + 200;
                            s = "€/";
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s += "KG";
                            else
                                s += "PZ";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 160;
                            y = Y + (iCol * Y_Step) + 380;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y, x, frmDX);


                            x = X + 179;
                            y = Y + (iCol * Y_Step) + 390;
                            s = ((string)tabEti.Rows[i]["eti_plu"]);
                            gfx.DrawString(s, font5, XBrushes.Black, y, x, frmDX);

                            if (i == 0)
                                Console.WriteLine("aaaa");


                            s = "C:\\ApProject\\Temp\\Img\\Godina\\eti045q.jpg";

                            XImage img = XImage.FromFile(s);
                            x = X + 130;
                            y = Y + (iCol * Y_Step) + 75;
                            XRect rect1 = new XRect(y, x, 100, 40);
                            gfx.DrawImage(img, rect1);



                            //x = X + 178;
                            //y = Y + (iCol * Y_Step) + 205;
                            //s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            //if (s.Trim().Length > 0)
                            //{
                            //    try
                            //    {
                            //        if (s.Length <= 8)
                            //        {
                            //            Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //        else
                            //        {
                            //            Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                            //    }
                            //}

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 205;
                            y = Y + (iCol * Y_Step);
                            //iLen = 15;
                            //x = X + 10;
                            //y = Y + (iCol * Y_Step) + 5;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 50;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);



                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 100;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            {
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                                if (d > 0 && d <= 500m)
                                {
                                    s = "€ al kg/L " + d.ToString("#####0.00");
                                }
                                //s = "al kg "; // +d.ToString("#####0.00");
                                //gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }



                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 200;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti046(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            //page.Height = 845.0;

            //page.Height = 842.0;
            //page.Width = 595;
            //page.Orientation = PdfSharp.PageOrientation.Landscape;

            page.Height = 845.0;
            page.Orientation = PdfSharp.PageOrientation.Portrait;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 9.5, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 15, XFontStyle.Regular);
            //XFont font3 = new XFont("Verdana", 18, XFontStyle.Bold);
            XFont font3 = new XFont("Times New Roman", 12, XFontStyle.Bold);
            XFont font4 = new XFont("Arial Black", 46, XFontStyle.Bold);
            XFont font5 = new XFont("Times New Roman", 12, XFontStyle.Bold);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 5;
            int i = 0;
            int X_Ini = 0;
            int X_Step = 105;
            int Y_Ini = 0;
            int Y_Step = 287;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = 0;
                            double y = 0;
                            int iLen = 15;
                            //double x = X;
                            //double y = Y + (iCol * Y_Step);
                            //int iLen = 15;
                            //x = X + 10;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 20;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = (string)tabEti.Rows[i]["eti_art"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /* Immagine logo */
                            s = "C:\\ApProject\\Temp\\Img\\Godina\\eti046.jpg";
                            XImage img = XImage.FromFile(s);
                            x = X + 70;
                            y = Y + (iCol * Y_Step) + 38;
                            XRect rect1 = new XRect(y, x, 95, 35);
                            gfx.DrawImage(img, rect1);

                            /* Rettangolo numero 1 */
                            x = X + 22;
                            y = Y + (iCol * Y_Step) + 26;
                            XRect rect = new XRect(y, x, 273, 88);
                            XPen pen = new XPen(XColors.Tomato, 8.9);
                            gfx.DrawRectangle(pen, rect);

                            /* Rettangolo numero 2 */
                            x = X + 29;
                            y = Y + (iCol * Y_Step) + 32;
                            rect = new XRect(y, x, 260, 74);
                            pen = new XPen(XColors.Green, 4);
                            gfx.DrawRectangle(pen, rect);

                            /* DESCRIZIONE */
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            string s2 = "";
                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            //x = X + 42;
                            //y = Y + (iCol * Y_Step) + 0;
                            x = X + 35;
                            y = Y + (iCol * Y_Step) + 35;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            rect = new XRect(y, x, 250, 30);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            /* UMI */
                            x = X + 65;
                            y = Y + (iCol * Y_Step) + 90;
                            s = "€/";
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s += "KG";
                            else
                                s += "PZ";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);

                            /* PREZZO */
                            x = X + 102;
                            y = Y + (iCol * Y_Step) + 260;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y, x, frmDX);

                            /* PLU */
                            x = X + 102;
                            y = Y + (iCol * Y_Step) + 288;
                            s = ((string)tabEti.Rows[i]["eti_plu"]);
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, frmDX);

                            if (i == 0)
                                Console.WriteLine("aaaa");


                            //x = X + 178;
                            //y = Y + (iCol * Y_Step) + 205;
                            //s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            //if (s.Trim().Length > 0)
                            //{
                            //    try
                            //    {
                            //        if (s.Length <= 8)
                            //        {
                            //            Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //        else
                            //        {
                            //            Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                            //    }
                            //}

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 205;
                            y = Y + (iCol * Y_Step);
                            //iLen = 15;
                            //x = X + 10;
                            //y = Y + (iCol * Y_Step) + 5;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 50;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);



                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 100;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            {
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                                if (d > 0 && d <= 500m)
                                {
                                    s = "€ al kg/L " + d.ToString("#####0.00");
                                }
                                //s = "al kg "; // +d.ToString("#####0.00");
                                //gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }



                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 200;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti047(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            //page.Height = 845.0;

            //page.Height = 842.0;
            //page.Width = 595;
            //page.Orientation = PdfSharp.PageOrientation.Landscape;

            page.Height = 845.0;
            page.Orientation = PdfSharp.PageOrientation.Portrait;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 9.5, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 15, XFontStyle.Regular);
            //XFont font3 = new XFont("Verdana", 18, XFontStyle.Bold);
            XFont font3 = new XFont("Times New Roman", 22, XFontStyle.Bold);
            XFont font4 = new XFont("Times New Roman", 56, XFontStyle.Bold);
            XFont font5 = new XFont("Times New Roman", 20, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 3;
            int i = 0;
            int X_Ini = 0;
            int X_Step = 210;
            int Y_Ini = 0;
            int Y_Step = 285;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = 0;
                            double y = 0;
                            int iLen = 15;
                            //double x = X;
                            //double y = Y + (iCol * Y_Step);
                            //int iLen = 15;
                            //x = X + 10;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 20;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = (string)tabEti.Rows[i]["eti_art"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            x = X + 41;
                            y = Y + (iCol * Y_Step) + 28;
                            XRect rect = new XRect(y, x, 274, 178);
                            XPen pen = new XPen(XColors.Green, 8.9);
                            gfx.DrawRectangle(pen, rect);

                            /* DESCRIZIONE */
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            string s2 = "";
                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            x = X + 55;
                            y = Y + (iCol * Y_Step) + 35;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            rect = new XRect(y, x, 250, 100);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            /* UMI 
                            x = X + 150;
                            y = Y + (iCol * Y_Step) + 200;
                            s = "€/";
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s += "KG";
                            else
                                s += "PZ";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);
                            */

                            /* PREZZO */
                            x = X + 160;
                            y = Y + (iCol * Y_Step) + 250;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00") + "€";
                            gfx.DrawString(s, font4, XBrushes.Black, y, x, frmDX);

                            /* Euro */
                            //x = X + 132;
                            //y = Y + (iCol * Y_Step) + 60;
                            //s = "€";
                            //gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            /* PLU */
                            x = X + 200;
                            y = Y + (iCol * Y_Step) + 285;
                            s = ((string)tabEti.Rows[i]["eti_plu"]);
                            gfx.DrawString(s, font5, XBrushes.Black, y, x, frmDX);

                            if (i == 0)
                                Console.WriteLine("aaaa");


                            s = "C:\\ApProject\\Temp\\Img\\Godina\\eti046.jpg";
                            XImage img = XImage.FromFile(s);
                            x = X + 174;
                            y = Y + (iCol * Y_Step) + 40;
                            XRect rect1 = new XRect(y, x, 100, 40);
                            gfx.DrawImage(img, rect1);



                            //x = X + 178;
                            //y = Y + (iCol * Y_Step) + 205;
                            //s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            //if (s.Trim().Length > 0)
                            //{
                            //    try
                            //    {
                            //        if (s.Length <= 8)
                            //        {
                            //            Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //        else
                            //        {
                            //            Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                            //    }
                            //}

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 205;
                            y = Y + (iCol * Y_Step);
                            //iLen = 15;
                            //x = X + 10;
                            //y = Y + (iCol * Y_Step) + 5;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 50;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);



                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 100;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            {
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                                if (d > 0 && d <= 500m)
                                {
                                    s = "€ al kg/L " + d.ToString("#####0.00");
                                }
                                //s = "al kg "; // +d.ToString("#####0.00");
                                //gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }



                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 200;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti048(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();

            page.Height = 842.0;
            page.Width = 595;
            page.Orientation = PdfSharp.PageOrientation.Landscape;

            //System.Drawing.Size size = PdfPage.PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Orientation = PageOrientation.Portrait;
            //pdfPage.Width = size.Width;
            //pdfPage.Height = size.Height;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 9, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Arial", 32, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 80, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 30, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 10, XFontStyle.Regular);

            XFont font7 = new XFont("Arial", 12, XFontStyle.Bold);
            XFont font8 = new XFont("Arial", 20, XFontStyle.Bold);
            XFont font9 = new XFont("Arial", 14, XFontStyle.Bold);
            XFont font10 = new XFont("Impact", 80, XFontStyle.Bold);
            XFont font11 = new XFont("Courier new", 11, XFontStyle.Bold);
            XFont font12 = new XFont("Courier new", 20, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 2;
            int i = 0;
            int X_Ini = 0;
            int X_Step = 315;
            int Y_Ini = 0;
            int Y_Step = 445;
            int X_Pie = 260;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        page.Orientation = PdfSharp.PageOrientation.Landscape;

                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            int iLen = 15;
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            //if (s.Length > 45)
                            //    s = s.Substring(0, 45);

                            string s2 = "";

                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            double x = X+20;
                            double y = Y + (iCol * Y_Step) - 10;

                            XRect rect = new XRect(y, x, 300, 170);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);


                            x = X + 130;
                            y = Y + (iCol * Y_Step) + 270;

                            if ((string)tabEti.Rows[i]["eti_bor"] != "")
                            {
                                s = "Origine: " + (string)tabEti.Rows[i]["eti_bor"];
                                gfx.DrawString(s, font9, XBrushes.Black, y, x + 0, XStringFormats.Default);
                            }
                            if ((string)tabEti.Rows[i]["eti_bct"] != "")
                            {
                                s = "Categoria: " + (string)tabEti.Rows[i]["eti_bct"];
                                gfx.DrawString(s, font9, XBrushes.Black, y, x + 20, XStringFormats.Default);
                            }
                            if ((string)tabEti.Rows[i]["eti_bcl"] != "")
                            {
                                s = "Calibro: " + (string)tabEti.Rows[i]["eti_bcl"];
                                gfx.DrawString(s, font9, XBrushes.Black, y, x + 40, XStringFormats.Default);
                            }

                            if (i == 0)
                                Console.WriteLine("zzzzzz");

                            /**** PLU INIZIO ****/
                            if (((string)tabEti.Rows[i]["eti_tas"]).Trim() != "")
                            {
                                x = X + 22;
                                y = Y + (iCol * Y_Step) + 280;
                                s = "TASTO BILANCIA";
                                gfx.DrawString(s, font11, XBrushes.Black, y, x - 5, XStringFormats.Default);

                                x = X + 20;
                                y = Y + (iCol * Y_Step) + 270;
                                tf = new XTextFormatter(gfx);
                                rect = new XRect(y, x, 120, 90);

                                XPen pen = new XPen(XColors.Black, 3);
                                gfx.DrawLine(pen, y, x, y + 120, x + 90);

                                //s = ((string)tabEti.Rows[i]["eti_tas"]).Substring(1);
                                s = (string)tabEti.Rows[i]["eti_tas"];
                                if (_clsFun.Numerico(s))
                                {
                                    s = Convert.ToInt16(s).ToString();
                                }

                                //gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);
                                //XPen pen = new XPen(XColors.RoyalBlue, Math.PI);
                                gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                                tf.Alignment = XParagraphAlignment.Center; //tf.Alignment = ParagraphAlignment.Left; 
                                tf.DrawString(s, font10, XBrushes.Black, rect, XStringFormats.TopLeft);
                            }
                            /**** PLU FINE ****/

                            //if ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ)
                            //{
                            //    s = "OFFERTA";

                            //    //XTextFormatter tf = new XTextFormatter(gfx);
                            //    //tf.Alignment = XParagraphAlignment.Center;

                            //    //rect = new XRect(y, x - 50, 500, 400);
                            //    //gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //    //tf.DrawString(s, font3, XBrushes.Black, rect);

                            //    x = X + 100;
                            //    y = Y + (iCol * Y_Step) - 10; // +(iCol * Y_Step) + 190;
                            //    tf = new XTextFormatter(gfx);
                            //    rect = new XRect(y, x, 400, 40);

                            //    gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                            //    //tf.Alignment = XParagraphAlignment.Default; //tf.Alignment = ParagraphAlignment.Left; 
                            //    //tf.DrawString(s, font7, XBrushes.Black, rect, XStringFormats.TopLeft);

                            //    if ((decimal)tabEti.Rows[i]["eti_pve"] > 0 && (decimal)tabEti.Rows[i]["eti_prv"] > 0)
                            //        d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                            //    if (d != 0)
                            //    {
                            //        s = ((decimal)tabEti.Rows[i]["eti_pve"]).ToString("#,##0.00");
                            //        gfx.DrawString(s, font8, XBrushes.Black, y + 20, x + 30, XStringFormats.Default);

                            //        XPen pen = new XPen(XColors.Black, 3);
                            //        gfx.DrawLine(pen, y + 17, x + 30, y + 100, x + 15);
                            //    }

                            //    s = "OFFERTA";
                            //    gfx.DrawString(s, font9, XBrushes.White, y + 150, x + 35, XStringFormats.Default);

                            //    if (d != 0)
                            //    {
                            //        s = "Sc." + d.ToString("#,##0.00") + " %";
                            //        gfx.DrawString(s, font7, XBrushes.Black, y + 320, x + 30, XStringFormats.Default);
                            //    }
                            //}


                            /******/

                            //x = X + 250;
                            //y = Y + (iCol * Y_Step) + 10;
                            //s = "€";
                            //gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 270;
                            y = Y + (iCol * Y_Step) + 270;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y, x, frmDX);

                            x = X + 250;
                            y = Y + (iCol * Y_Step) + 10;
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s = "AL KG";
                            else
                                s = "AL NR";
                            gfx.DrawString(s, font12, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            if (i == 0)
                                Console.WriteLine("zzzzzz");

                            if (((string)tabEti.Rows[i]["eti_plu"]).Trim() != "")
                            {
                                x = X + 230;
                                y = Y + (iCol * Y_Step) + 275;
                                s = "PLU " + (string)tabEti.Rows[i]["eti_plu"];
                                gfx.DrawString(s, font8, XBrushes.Black, y, x - 5, XStringFormats.Default);
                            }

                            x = X + 250;
                            y = Y + (iCol * Y_Step) + 275;
                            s = "EUR ";
                            gfx.DrawString(s, font8, XBrushes.Black, y, x - 5, XStringFormats.Default);


                            //x = X + 240;
                            //y = Y + (iCol * Y_Step) + 5;
                            //if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            //    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            //s = "€ al kg/L " + d.ToString("#####0.00");
                            //gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);


                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 280;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 220;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            //x = X + X_Pie - 35;
                            //y = Y + (iCol * Y_Step) + 290;
                            //s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            //if (s.Trim().Length > 0)
                            //{
                            //    try
                            //    {
                            //        if (s.Length <= 8)
                            //        {
                            //            Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 95, 25);
                            //        }
                            //        else
                            //        {
                            //            Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 95, 25);
                            //        }
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                            //    }
                            //}

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 10;
                            s = DateTime.Today.ToString("dd/MM/yyyy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 305;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 300;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti049(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 845.0;
            page.Orientation = PdfSharp.PageOrientation.Portrait;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Arial", 10, XFontStyle.Bold);
            XFont font2 = new XFont("Arial", 8, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Arial", 14, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 48, XFontStyle.Regular);
            XFont font5 = new XFont("Arial", 35, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 9, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 5;
            int i = 0;
            int X_Ini = 20;
            int X_Step = 163;
            int Y_Ini = 15;
            int Y_Step = 240;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            if (i == 1)
                                Console.WriteLine("aaaa");

                            double x = X;
                            double y = Y + (iCol * Y_Step);
                            int iLen = 15;


                            //y = Y + (iCol * Y_Step) + 130;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            s = (string)tabEti.Rows[i]["eti_ard"];
                            string s2 = "";
                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            /* DESCRIZIONE */
                            x = X + 12;
                            y = Y + (iCol * Y_Step) + 5;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            XRect rect = new XRect(y, x, 250, 77);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            //x = X + 107;
                            //y = Y + (iCol * Y_Step) + 11;
                            //s = "€";
                            //gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            /* PREZZO */
                            x = X + 110;
                            y = Y + (iCol * Y_Step) + 150;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            /* UMI */                            
                            x = X + 80;
                            y = Y + (iCol * Y_Step) + 5;
                            //if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            //    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));
                            //s = "€ al kg/L " + d.ToString("#####0.00");

                            s = "al Pezzo "; // +d.ToString("#####0.00");
                            if ((string)tabEti.Rows[i]["eti_umi"] == "KG")
                                if (d > 0 && d <= 500m)
                                {
                                    s = "al kg "; // +d.ToString("#####0.00");
                                    gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                                }


                            ////x = X + 142;
                            //y = Y + (iCol * Y_Step) + 95;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 120;
                            //y = Y + (iCol * Y_Step) + 175;
                            //s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            //if (s.Trim().Length > 0)
                            //{
                            //    try
                            //    {
                            //        if (s.Length <= 8)
                            //        {
                            //            Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //        else
                            //        {
                            //            Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                            //    }
                            //}

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 105;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            x = X + 70;
                            y = Y + (iCol * Y_Step) + 185;
                            s = "PLU";
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 80;
                            y = Y + (iCol * Y_Step) + 185;
                            s = (string)tabEti.Rows[i]["eti_plu"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 100;
                            y = Y + (iCol * Y_Step) + 185;
                            s = "EUR";
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 110;
                            y = Y + (iCol * Y_Step) + 150;
                            s = "CodArt. " + (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 110;
                            y = Y + (iCol * Y_Step) + 5;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti050(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

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
            //XFont font1 = new XFont("Arial", 18, XFontStyle.Regular);
            XFont font1 = new XFont("Lucida Calligraphy", 18, XFontStyle.Regular);

            XFont font2 = new XFont("Arial", 30, XFontStyle.Regular);
            XFont font3 = new XFont("Cooper Black", 20, XFontStyle.Regular);
            //XFont font4 = new XFont("Arial", 55, XFontStyle.BoldItalic);
            //XFont font4 = new XFont("Arial", 12, XFontStyle.Regular);
            XFont font4 = new XFont("Lucida Calligraphy", 12, XFontStyle.Regular);

            XFont font5 = new XFont("Impact", 28, XFontStyle.BoldItalic);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);
            //XFont font7 = new XFont("Impact", 10, XFontStyle.Regular);
            XFont font7 = new XFont("Arial", 16, XFontStyle.Bold);

            XFont font8 = new XFont("Lucida Calligraphy", 36, XFontStyle.Italic);


            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 5;
            int i = 0;
            int X_Ini = 38;
            int X_Step = 144;
            int Y_Ini = 33;
            int Y_Step = 265;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            XPen pen = new XPen(XColors.Black, 0.8);

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {

                            int iLen = 15;

                            double x = X + 14;
                            double y = Y + (iCol * Y_Step) + 2;

                            //gfx.DrawRectangle(pen, Y + (iCol * Y_Step), x - 10, Y + (iCol * Y_Step) + 100, x + 50);

                            //if (i <= 1)
                            gfx.DrawRectangle(pen, y, x, Y_Step, X_Step);

                            /* Immagine logo
                            s = "C:\\ApProject\\Temp\\Img\\Etichette\\eti040.png";
                            XImage img = XImage.FromFile(s);
                            x = X + 40;
                            y = Y + (iCol * Y_Step) + 130;
                            XRect rect1 = new XRect(y, x, 145, 105);
                            //XRect rect1 = new XRect(y, x, 145, 105);
                            gfx.DrawImage(img, rect1);
                            */

                            x = X + 14;
                            y = Y + (iCol * Y_Step) + 2;

                            //gfx.DrawRectangle(pen, Y + (iCol * Y_Step), x - 10, Y + (iCol * Y_Step) + 100, x + 50);

                            //if (i <= 1)
                            gfx.DrawRectangle(pen, y, x, Y_Step, X_Step);

                            /*
                             *  Descrizione articolo
                             */
                            s = (string)tabEti.Rows[i]["eti_ard"];

                            s = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLower());

                            //font1.Height = 25.0;


                            //XStringFormat format = new XStringFormat();
                            //format.Alignment = XStringAlignment.Center;

                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            XRect rect = new XRect(y, x + 5, 265, 100);
                            //gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font1, XBrushes.Black, rect);

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            /*
                             *  Euro
                             */ 
                            x = X + 93;
                            y = Y + (iCol * Y_Step) + 90;
                            
                            s = "€";
                            // gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                             *  Prezzo
                             */
                            x = X + 100;
                            y = Y + (iCol * Y_Step) + 170;
                            s = "€  " + Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font8, XBrushes.Black, y + 10, x, frmDX);


                            /*
                             *  UMI
                             
                            x = X + 110;
                            y = Y + (iCol * Y_Step) + 130;
                            //if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            //    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));
                            //s = "€ al kg/L " + d.ToString("#####0.00");
                            s = "al Pz.";
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s = "al Kg";
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            */




                            /*
                             * Ingredienti
                             */
                            s = _clsQry.GetFileIngredienti((string)tabEti.Rows[i]["eti_art"]);
                            if (s.Length > 0)
                            {
                                x = X + 105;
                                y = Y + (iCol * Y_Step) + 10;
                                rect = new XRect(y + 0, x, 250, 80);
                                tf = new XTextFormatter(gfx);
                                //gfx.DrawRectangle(XBrushes.Transparent, rect);
                                tf.Alignment = XParagraphAlignment.Center;
                                tf.DrawString(s, font4, XBrushes.Black, rect, XStringFormats.TopLeft);
                            }

                            /*
                             * PLU
                            //x = X + 129;
                            x = X + 80;
                            y = Y + (iCol * Y_Step) + 20;
                            s = ((string)tabEti.Rows[i]["eti_plu"]);
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);
                             */


                            /*
                             *  Codice articolo
                            x = X + 103;
                            y = Y + (iCol * Y_Step) + 20;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                             */


                            /*
                             *  Codice articolo fornitore
                            //x = X + 140;
                            y = Y + (iCol * Y_Step) + 100;
                            s = ((string)tabEti.Rows[i]["eti_arf"]);
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                             */


                            /*
                             *  Data
                            //x = X + 140;
                            y = Y + (iCol * Y_Step) + 175;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                             */


                            /*
                             * BARCODE
                             * 
                             */
                            //y = Y + (iCol * Y_Step) + 95;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            //x = X + 120;
                            //y = Y + (iCol * Y_Step) + 175;
                            //s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            //if (s.Trim().Length > 0)
                            //{
                            //    try
                            //    {
                            //        if (s.Length <= 8)
                            //        {
                            //            Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //        else
                            //        {
                            //            Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                            //    }
                            //}

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 105;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti051(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 845.0;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 15, XFontStyle.Regular);
            XFont font3 = new XFont("Verdana", 18, XFontStyle.Bold);
            //XFont font4 = new XFont("Arial Black", 50, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 55, XFontStyle.Bold);     //Verdana
            XFont font5 = new XFont("Arial", 30, XFontStyle.Bold);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);
            XFont font7 = new XFont("Arial", 11, XFontStyle.Bold);
            XFont font8 = new XFont("Verdana", 16, XFontStyle.Bold);
            XFont font10 = new XFont("Arial", 12, XFontStyle.Bold);
            XFont font11 = new XFont("Courier new", 11, XFontStyle.Bold);
            XFont font12 = new XFont("Impact", 35, XFontStyle.Bold);     //Verdana
            XFont font13 = new XFont("Arial", 40, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 1;
            int iRows = 4;
            int i = 0;
            int X_Ini = -10;
            int X_Step = 218;
            int Y_Ini = 15;
            int Y_Step = 300;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = 0;
                            double y = 0;
                            int iLen = 15;

                            XTextFormatter tf = new XTextFormatter(gfx);
                            XRect rect = new XRect(y, x, 260, 100);

                            /**** PLU INIZIO 1 ****/
                            if (((string)tabEti.Rows[i]["eti_plu"]).Trim() != "")
                            {

                                x = X + 70;
                                y = Y + (iCol * Y_Step) + 0;
                                s = ((string)tabEti.Rows[i]["eti_plu"]);
                                if (_clsFun.Numerico(s))
                                    s = Convert.ToInt16(s).ToString();
                                s = s = "plu " + s;
                                gfx.DrawString(s, font13, XBrushes.Black, y, x - 5, XStringFormats.Default);

                                //x = X + 35;
                                //y = Y + (iCol * Y_Step) + 70;
                                //tf = new XTextFormatter(gfx);
                                //rect = new XRect(y, x, 120, 50);

                                //XPen pen = new XPen(XColors.Black, 3);
                                //gfx.DrawLine(pen, y, x, y + 120, x + 90);


                                ////gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);
                                ////XPen pen = new XPen(XColors.RoyalBlue, Math.PI);
                                //gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                                //tf.Alignment = XParagraphAlignment.Center; //tf.Alignment = ParagraphAlignment.Left; 
                                //tf.DrawString(s, font5, XBrushes.Black, rect, XStringFormats.TopLeft);
                            }
                            /**** PLU FINE 1 ****/


                            /**** DESCRIZIONE 1 INIZIO ****/
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            string s2 = "";
                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            x = X + 80;
                            y = Y + (iCol * Y_Step);
                            tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            rect = new XRect(y, x, 260, 100);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font8, XBrushes.Black, rect);
                            /**** DESCRIZIONE 1 FINE ****/

                            /**** ARTICOLO INIZIO 1 ****/
                            x = X + 160;
                            y = Y + (iCol * Y_Step) + 0;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /**** ARTICOLO FINE 1 ****/

                            /*** PxC INIZIO ***/
                            x = X + 180;
                            y = Y + (iCol * Y_Step - 0);
                            s = "PxC " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            /*** PxC FINE ***/

                            /**** ARTICOLO INIZIO 2 ****
                            x = X + 120;
                            y = Y + (iCol * Y_Step) + 300;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            **** ARTICOLO FINE 2 ****/

                            /**** DESCRIZIONE 2 INIZIO ****/
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            s2 = "";
                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            x = X + 42;
                            y = Y + (iCol * Y_Step) + 290;
                            tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            rect = new XRect(y, x, 260, 100);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);
                            /**** DESCRIZIONE 2 FINE ****/

                            /**** EURO INIZIO 1 ****/
                            x = X + 180;
                            y = Y + (iCol * Y_Step) + 100;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);
                            /**** EURO FINE 1 ****/

                            /*** UMI INIZIO 1 ***/
                            x = X + 190;
                            y = Y + (iCol * Y_Step) + 122;
                            s = "al ";
                            if ((string)tabEti.Rows[i]["eti_umi"] == "NR")
                                s += "PZ";
                            else
                                s += (string)tabEti.Rows[i]["eti_umi"];
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, frmDX);
                            /*** UMI FINE 1 ***/

                            /**** PREZZO INIZIO 1 ****/
                            x = X + 180;
                            y = Y + (iCol * Y_Step) + 250;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font12, XBrushes.Black, y + 10, x, frmDX);
                            /**** PREZZO FINE 1 ****/

                            /*** PLU INIZIO 2 ***
                            x = X + 150;
                            y = Y + (iCol * Y_Step) + 340;
                            s = "PLU " + (string)tabEti.Rows[i]["eti_plu"];
                            gfx.DrawString(s, font7, XBrushes.Black, y + 10, x, frmDX);
                            *** PLU FINE 2 ***/

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            /*
                             * Ingredienti 
                             
                            s = _clsQry.GetFileIngredienti((string)tabEti.Rows[i]["eti_art"]);
                            if (s.Length > 0)
                            {
                                //x = X + 70;
                                //y = Y + (iCol * Y_Step) + 100;

                                x = X + 120;
                                y = Y + (iCol * Y_Step) + 300;

                                rect = new XRect(y + 0, x, 350, 80);
                                tf = new XTextFormatter(gfx);
                                //gfx.DrawRectangle(XBrushes.Transparent, rect);
                                tf.Alignment = XParagraphAlignment.Right;
                                tf.DrawString(s, font4, XBrushes.White, rect, XStringFormats.TopLeft);
                            }
                            */

                            /*
                             * Ingredienti
                             */
                            s = _clsQry.GetFileIngredienti((string)tabEti.Rows[i]["eti_art"]);
                            if (s.Length > 0)
                            {
                                x = X + 100;
                                y = Y + (iCol * Y_Step) + 300;
                                rect = new XRect(y + 0, x, 250, 80);
                                tf = new XTextFormatter(gfx);
                                //gfx.DrawRectangle(XBrushes.Transparent, rect);
                                tf.Alignment = XParagraphAlignment.Center;
                                tf.DrawString(s, font6, XBrushes.Black, rect, XStringFormats.TopLeft);
                            }

                            /**** EURO INIZIO 2 ****/
                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 380;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);
                            /**** EURO FINE 2 ****/

                            /*** UMI INIZIO 2 ***/
                            x = X + 200;
                            y = Y + (iCol * Y_Step) + 360;
                            s = "al ";
                            if ((string)tabEti.Rows[i]["eti_umi"] == "NR")
                                s += "PZ";
                            else
                                s += (string)tabEti.Rows[i]["eti_umi"];
                            gfx.DrawString(s, font7, XBrushes.Black, y + 10, x, frmDX);
                            /*** UMI FINE 2 ***/

                            /**** PREZZO INIZIO 2 ****/
                            x = X + 210;
                            y = Y + (iCol * Y_Step) + 550;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);
                            /**** PREZZO FINE 2 ****/

                            /**** BARCODE INIZIO ****/
                            x = X + 190;
                            y = Y + (iCol * Y_Step) + 205;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 15);
                                    }
                                    else
                                    {
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 15);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }
                            /**** BARCODE FINE ****/

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            /*** Barcode su arf di DADO **
                            x = X + 190;
                            y = Y + (iCol * Y_Step) - 20 + 500;
                            s = ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    //if (s.Length <= 8)
                                    //{
                                    //    //s = "0" + s;
                                    //    //s += new clsCtrlCodici().FindMod10Digit(s);

                                    //    Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                    //    MemoryStream ms = new MemoryStream();
                                    //    img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                    //    gfx.DrawImage(img, y, x, 70, 14);
                                    //}
                                    //else
                                    //{

                                    s = "7901" + s.PadLeft(8, Convert.ToChar('0'));
                                    s = s + new clsCtrlCodici().FindMod10Digit(s);

                                    //s = "7901078492014";
                                    Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                    MemoryStream ms = new MemoryStream();
                                    img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                    gfx.DrawImage(img, y, x, 70, 14);

                                    //}
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }
                            */

                            /**** DATA INIZIO 1 ****/
                            x = X + 205;
                            y = Y + (iCol * Y_Step);
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /**** DATA FINE 1 ****/

                            /**** DATA INIZIO 2 ****
                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 300;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            **** DATA FINE 2 ****/

                            ///**** PREZZO IN CHIARO INIZIO ****/
                            //x = X + 205;
                            //y = Y + (iCol * Y_Step) + 100;
                            //if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            //{
                            //    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            //    s = "€ al kg/L " + d.ToString("#####0.00");
                            //    //s = "al kg "; // +d.ToString("#####0.00");
                            //    gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            //}
                            ///**** PREZZO IN CHIARO FINE ****/

                            /**** EAN IN CIFRE INIZIO ****/
                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 100;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /**** EAN IN CIFRE FINE ****/

                            /**** EAN IN CIFRE INIZIO 2 ****
                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 420;
                            s = ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            **** EAN IN CIFRE FINE 2 ****/

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti052(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 845.0;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 7, XFontStyle.Bold);
            XFont font2 = new XFont("Arial", 6, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Arial", 9, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 36, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 20, XFontStyle.Bold);
            XFont font6 = new XFont("Arial", 6.5, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 3;
            int iRows = 8;
            int i = 0;
            int X_Ini = 115;
            int X_Step = 87; // 108;       //92
            int Y_Ini = 22;
            int Y_Step = 200;       //196

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = X + 3;
                            double y = Y + (iCol * Y_Step - 10);
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            if (s.Length > 45)
                                s = s.Substring(0, 45);

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            XTextFormatter tf = new XTextFormatter(gfx);

                            /*** Descrizione ***/
                            XRect rect = new XRect(y, x, 180, 40);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //tf.Alignment = ParagraphAlignment.Left; 
                            tf.DrawString(s, font3, XBrushes.Black, rect, XStringFormats.TopLeft);

                            /*** codice articolo ***/
                            x = X + 34;
                            y = Y + (iCol * Y_Step) - 20;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*** PxC ***/
                            x = X + 34;
                            y = Y + (iCol * Y_Step) + 10;
                            s = "Pxc " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);


                            /*** Fornitore ***/
                            x = X + 34;
                            y = Y + (iCol * Y_Step) + 35;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);


                            /*** Data ***/
                            x = X + 34;
                            y = Y + (iCol * Y_Step) + 90;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);



                            /*** Barcode in cifre ***/
                            x = X + 34;
                            y = Y + (iCol * Y_Step) + 118;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);




                            /*** Prezzo in chiaro ***/
                            x = X + 65;
                            y = Y + (iCol * Y_Step) - 20; ;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)

                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0 && d <= 500m)
                            {
                                s ="Al kg/L € " + d.ToString("#####0.00");
                                gfx.DrawString(s.Trim(), font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }

                            /*** Simbolo euro ***/
                            x = X + 71;
                            y = Y + (iCol * Y_Step) + 50;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            /*** Prezzo ***/
                            x = X + 73;
                            y = Y + (iCol * Y_Step) + 155;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            /*** Contenuto ***/
                            x = X + 55;
                            y = Y + (iCol * Y_Step) - 20;
                            s = "Cont. " + (string)tabEti.Rows[i]["eti_tgr"] + " " + ((decimal)tabEti.Rows[i]["eti_pne"]).ToString("#####0.00");
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** Grammatura ***/
                            x = X + 72;
                            y = Y + (iCol * Y_Step) + 150;
                            s = "al ";
                            if ((string)tabEti.Rows[i]["eti_umi"] == "NR")
                                s += "PZ";
                            else
                                s += (string)tabEti.Rows[i]["eti_umi"];

                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** Barcode su arf di DADO 
                            x = X + 57;
                            y = Y + (iCol * Y_Step) - 20;
                            s = ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    //if (s.Length <= 8)
                                    //{
                                    //    //s = "0" + s;
                                    //    //s += new clsCtrlCodici().FindMod10Digit(s);

                                    //    Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                    //    MemoryStream ms = new MemoryStream();
                                    //    img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                    //    gfx.DrawImage(img, y, x, 70, 14);
                                    //}
                                    //else
                                    //{

                                    s = "7901" + s.PadLeft(8, Convert.ToChar('0'));
                                    s = s + new clsCtrlCodici().FindMod10Digit(s);

                                    //s = "7901078492014";
                                    Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                    MemoryStream ms = new MemoryStream();
                                    img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                    gfx.DrawImage(img, y, x, 70, 14);

                                    //}
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }
                            ***/

                            /*** Barcode ***/
                            x = X + 18;
                            y = Y + (iCol * Y_Step) + 95;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 10);
                                    }
                                    else
                                    {
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 10);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti053(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

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
            XFont font1 = new XFont("Arial", 8, XFontStyle.Regular);
            XFont font2 = new XFont("Courier new", 15, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Arial", 15, XFontStyle.Regular);
            //XFont font4 = new XFont("Arial", 55, XFontStyle.BoldItalic);

            XFont font4 = new XFont("Impact", 15, XFontStyle.Italic);
            XFont font5 = new XFont("Franklin Gothic Demi Cond", 20, XFontStyle.Bold);

            XFont font6 = new XFont("Arial", 10, XFontStyle.Bold);
            XFont font7 = new XFont("Impact", 10, XFontStyle.Regular);

            XFont font8 = new XFont("Arial", 10, XFontStyle.Regular);
            XFont font9 = new XFont("Arial", 9, XFontStyle.Italic);
            XFont font10 = new XFont("Arial", 10, XFontStyle.Bold);
            XFont font11 = new XFont("Arial", 7, XFontStyle.Bold);
            XFont font12 = new XFont("Arial", 24, XFontStyle.Bold);


            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 6;
            int iRows = 12;
            int i = 0;
            //int X_Ini = 70;
            int X_Ini = 32;
            int X_Step = 68;
            int Y_Ini = 0;
            int Y_Step = 100;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            //s = "SELECT * FROM AnaIttico ORDER BY itt_art";
            //DataTable tItt = _clsFun.FillTabSql("AnaIttico", s, false, _strConSql);
            //DataColumn[] keys = new DataColumn[1];
            //keys[0] = tItt.Columns["itt_art"];
            //tItt.PrimaryKey = keys;

            //DataRow[] jItt;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            if (i == 0)
                                Console.WriteLine("aaaa");

                            int iLen = 15;

                            double x = 0;
                            double y = 0;

                            /*
                             *  For + ARF
                             */
                            //y = Y + (iCol * Y_Step) + 130;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //string sDesArt = (string)tabEti.Rows[i]["eti_ard"];
                            //string sDesCom = "";
                            //string sDesLat = "";
                            //string sDesPes = "";

                            //jItt = tItt.Select("itt_art='" + (string)tabEti.Rows[i]["eti_art"] + "'");
                            //if (jItt.Length > 0)
                            //{
                            //    s = ((string)jItt[0]["itt_006"]).Trim();
                            //    if (s.Length > 40)
                            //        sDesCom = s.Substring(40).Trim();

                            //    s = ((string)jItt[0]["itt_005"]).Trim();
                            //    if (s.Length > 40)
                            //        sDesLat = s.Substring(40).Trim();

                            //    sDesPes = "";

                            //    s = ((string)jItt[0]["itt_002"]).Trim();
                            //    if (s.Length > 40)
                            //        sDesPes += s.Substring(40).Trim();

                            //    //s = ((string)jItt[0]["itt_003"]).Trim();
                            //    //if (s.Length > 40)
                            //    //    sDesPes += s.Substring(40).Trim();

                            //    s = ((string)jItt[0]["itt_011"]).Trim();
                            //    if (s.Length > 40)
                            //        sDesPes += " IN FAO " + s.Substring(40).Trim();
                            //    else
                            //    {
                            //        s = ((string)jItt[0]["itt_004"]).Trim();
                            //        if (s.Length > 40)
                            //            sDesPes += " IN " + s.Substring(40).Trim();
                            //    }

                            //    s = ((string)jItt[0]["itt_003"]).Trim();
                            //    if (s.Length > 40)
                            //        sDesPes += " CON " + s.Substring(40).Trim();
                            //    sDesPes = sDesPes.ToUpper();
                            //}


                            /*
                             *  Descrizione articolo
                             */
                            //s = (string)tabEti.Rows[i]["eti_ard"];
                            //string s2 = "";
                            //if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                            //    Console.WriteLine("aaaa");
                            //s2 = s;
                            //for (int i3 = 0; i3 < 4; i3++)
                            //{
                            //    string[] a = s2.Split(' ');
                            //    s2 = "";

                            //    for (int i2 = 0; i2 < a.Length; i2++)
                            //    {
                            //        if (a[i2].Length > iLen)
                            //            a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                            //        s2 += a[i2] + " ";
                            //    }
                            //}

                            //if (s2 != "")
                            //    s = s2;


                            //x = X + 3;
                            //y = Y + (iCol * Y_Step) + 20;
                            //XTextFormatter tf = new XTextFormatter(gfx);
                            //tf.Alignment = XParagraphAlignment.Center;
                            ////XRect rect = new XRect(y, x, 300, 100);
                            //XRect rect = new XRect(y, x, 210, 77);
                            //gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //tf.DrawString(sDesArt, font3, XBrushes.Black, rect);

                            //x = X + 35;
                            //y = Y + (iCol * Y_Step) + 20;
                            //tf = new XTextFormatter(gfx);
                            //tf.Alignment = XParagraphAlignment.Center;
                            ////XRect rect = new XRect(y, x, 300, 100);
                            //rect = new XRect(y, x, 210, 77);
                            //gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //tf.DrawString(sDesCom, font8, XBrushes.Black, rect);

                            //x = X + 50;
                            //y = Y + (iCol * Y_Step) + 20;
                            //tf = new XTextFormatter(gfx);
                            //tf.Alignment = XParagraphAlignment.Center;
                            ////XRect rect = new XRect(y, x, 300, 100);
                            //rect = new XRect(y, x, 210, 77);
                            //gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //tf.DrawString(sDesLat, font9, XBrushes.Black, rect);

                            //x = X + 65;
                            //y = Y + (iCol * Y_Step) + 20;
                            //tf = new XTextFormatter(gfx);
                            //tf.Alignment = XParagraphAlignment.Center;
                            ////XRect rect = new XRect(y, x, 300, 100);
                            //rect = new XRect(y, x, 210, 77);
                            //gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //tf.DrawString(sDesPes, font10, XBrushes.Black, rect);

                            //if (i == 1)
                            //    Console.WriteLine("aaaa");

                            ///*
                            // *  Riga
                            // */
                            ////x = X + 95;
                            //x = X + 90;
                            //y = Y + (iCol * Y_Step) - 20;
                            //XPen pen = new XPen(XColors.Black, 1);
                            //gfx.DrawLine(pen, y + 20, x, y + 270, x);

                            ///*
                            // *  UMI
                            // */
                            ////x = X + 108;
                            //x = X + 103;
                            //y = Y + (iCol * Y_Step) + 210;
                            ////if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            ////    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            ////    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            ////    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));
                            ////s = "€ al kg/L " + d.ToString("#####0.00");
                            //s = "al pz.";
                            //if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                            //    s = "al Kg";
                            //gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);


                            ///*
                            // * BARCODE
                            // * 
                            ////x = X + 128;
                            //x = X + 123;
                            //y = Y + (iCol * Y_Step) + 15;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font11, XBrushes.Black, y, x, XStringFormats.Default);
                            // */

                            ///*
                            // *  Codice articolo
                            ////x = X + 138;
                            //x = X + 133;
                            //y = Y + (iCol * Y_Step) + 15;
                            //s = (string)tabEti.Rows[i]["eti_art"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            // */

                            ///*
                            // * PLU
                            // */
                            ////x = X + 148;
                            //x = X + 110;
                            //y = Y + (iCol * Y_Step) + 15;
                            //s = ((string)tabEti.Rows[i]["eti_plu"]).Trim();
                            //if (s != "" && _clsFun.Numerico(s))
                            //    s = Convert.ToInt16(tabEti.Rows[i]["eti_plu"]).ToString();
                            //else
                            //    s = "";
                            //gfx.DrawString(s, font12, XBrushes.Black, y, x, XStringFormats.Default);

                            ///*
                            // * ARF
                            // */
                            ////x = X + 148;
                            //x = X + 143;
                            //y = Y + (iCol * Y_Step) + 15;
                            //s = ((string)tabEti.Rows[i]["eti_arf"]);
                            //gfx.DrawString(s, font10, XBrushes.Black, y, x, XStringFormats.Default);


                            ///*
                            // *  Data
                            // */
                            //x = X + 143;
                            //y = Y + (iCol * Y_Step) + 200;
                            //s = DateTime.Today.ToString("dd/MM/yy");
                            //gfx.DrawString(s, font10, XBrushes.Black, y, x, XStringFormats.Default);

                            if (i == 0)
                                Console.WriteLine("aaaa");


                            /*
                             *  Euro
                             */
                            x = X + 45;
                            y = Y + (iCol * Y_Step) + 15;
                            s = "€";
                            gfx.DrawString(s, font4, XBrushes.Black, y, x - 5, XStringFormats.Default);


                            /*
                             *  Prezzo
                             */
                            //x = X + 142;
                            x = X + 45;
                            y = Y + (iCol * Y_Step) + 65;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font5, XBrushes.Black, y + 10, x, frmDX);

                            /*
                             * BARCODE
                             * 
                             */
                            x = X + 22;
                            y = Y + (iCol * Y_Step) + 15;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font11, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 0;
                            y = Y + (iCol * Y_Step) + 14;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 15);
                                    }
                                    else
                                    {
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 15);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }
                             

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 105;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti054(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 845.0;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 15, XFontStyle.Regular);
            XFont font3 = new XFont("Verdana", 18, XFontStyle.Bold);
            //XFont font4 = new XFont("Arial Black", 50, XFontStyle.Bold);
            //XFont font4 = new XFont("Impact", 55, XFontStyle.Bold);     //Verdana
            XFont font4 = new XFont("Arial", 44, XFontStyle.Regular);     //Verdana
            //XFont font5 = new XFont("Arial", 30, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 22, XFontStyle.Bold);
            XFont font6 = new XFont("Arial", 10, XFontStyle.Bold);
            XFont font7 = new XFont("Arial", 11, XFontStyle.Bold);
            XFont font8 = new XFont("Verdana", 16, XFontStyle.Bold);
            XFont font10 = new XFont("Arial", 12, XFontStyle.Bold);
            XFont font11 = new XFont("Courier new", 11, XFontStyle.Bold);
            XFont font12 = new XFont("Arial Black", 50, XFontStyle.Regular);     //Verdana
            XFont font13 = new XFont("Arial Black", 20, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 1;
            int iRows = 4;
            int i = 0;
            int X_Ini = -10;
            int X_Step = 218;
            int Y_Ini = 15;
            int Y_Step = 300;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            /***   
            //Filtro per stampare solo etichette con ingredienti
            DataTable t = tabEti.Clone();
            foreach(DataRow y in tabEti.Rows)
            {
                s = _clsQry.GetFileIngredienti((string)y["eti_art"]);
                if (s.Trim() != "")
                    t.ImportRow(y);
            }
            tabEti = t.Copy();
            ***/

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = 0;
                            double y = 0;
                            int iLen = 15;

                            string sIng = _clsQry.GetFileIngredienti((string)tabEti.Rows[i]["eti_art"]);

                            XTextFormatter tf = new XTextFormatter(gfx);
                            XRect rect = new XRect(y, x, 260, 100);


                            /**** DESCRIZIONE 1 INIZIO ****/
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            string s2 = "";
                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            x = X + 35;
                            y = Y + (iCol * Y_Step);
                            tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            rect = new XRect(y, x, 260, 100);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font5, XBrushes.Black, rect);
                            /**** DESCRIZIONE 1 FINE ****/

                            /**** ARTICOLO INIZIO 1 ****/
                            //x = X + 160;
                            //y = Y + (iCol * Y_Step) + 0;
                            //s = (string)tabEti.Rows[i]["eti_art"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /**** ARTICOLO FINE 1 ****/

                            /*** PxC INIZIO ***/
                            //x = X + 180;
                            //y = Y + (iCol * Y_Step - 10);
                            //s = "PxC " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            /*** PxC FINE ***/

                            /**** ARTICOLO INIZIO 2 ****/
                            //x = X + 120;
                            //y = Y + (iCol * Y_Step) + 300;
                            //s = (string)tabEti.Rows[i]["eti_art"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /**** ARTICOLO FINE 2 ****/

                            /**** DESCRIZIONE 2 INIZIO ****/
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            s2 = "";
                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            x = X + 35;
                            y = Y + (iCol * Y_Step) + 290;
                            tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            rect = new XRect(y, x, 260, 100);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //tf.DrawString(s, font5, XBrushes.Black, rect);
                            tf.DrawString(s, font5, XBrushes.Black, rect);
                            /**** DESCRIZIONE 2 FINE ****/

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            /*
                             * Ingredienti
                             */
                            if (s.Length > 0)
                            {
                                x = X + 105;
                                y = Y + (iCol * Y_Step) + 5;
                                rect = new XRect(y + 0, x, 280, 80);
                                tf = new XTextFormatter(gfx);
                                //gfx.DrawRectangle(XBrushes.Transparent, rect);
                                tf.Alignment = XParagraphAlignment.Left;
                                tf.DrawString(sIng, font6, XBrushes.Black, rect, XStringFormats.TopLeft);
                            }



                            /**** PLU INIZIO 1 ****/
                            if (((string)tabEti.Rows[i]["eti_plu"]).Trim() != "")
                            {

                                x = X + 125;
                                y = Y + (iCol * Y_Step) + 260;
                                s = ((string)tabEti.Rows[i]["eti_plu"]);
                                //if (_clsFun.Numerico(s))
                                //    s = Convert.ToInt32(s).ToString();
                                s = "PLU: " + s;
                                gfx.DrawString(s, font13, XBrushes.Black, y+30, x - 5, XStringFormats.Default);

                                //x = X + 35;
                                //y = Y + (iCol * Y_Step) + 70;
                                //tf = new XTextFormatter(gfx);
                                //rect = new XRect(y, x, 120, 50);

                                //XPen pen = new XPen(XColors.Black, 3);
                                //gfx.DrawLine(pen, y, x, y + 120, x + 90);


                                ////gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);
                                ////XPen pen = new XPen(XColors.RoyalBlue, Math.PI);
                                //gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                                //tf.Alignment = XParagraphAlignment.Center; //tf.Alignment = ParagraphAlignment.Left; 
                                //tf.DrawString(s, font5, XBrushes.Black, rect, XStringFormats.TopLeft);
                            }
                            /**** PLU FINE 1 ****/





                            /**** EURO INIZIO 1 ****
                            x = X + 180;
                            y = Y + (iCol * Y_Step) + 100;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);
                            **** EURO FINE 1 ****/

                            /*** UMI INIZIO 1 ***/
                            x = X + 175;
                            y = Y + (iCol * Y_Step) + 50;
                            s = "Euro al ";
                            if ((string)tabEti.Rows[i]["eti_umi"] == "NR")
                                s += "PZ";
                            else
                                s += (string)tabEti.Rows[i]["eti_umi"];
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, frmDX);
                            /*** UMI FINE 1 ***/

                            /**** PREZZO INIZIO 1 ****/
                            x = X + 195;
                            y = Y + (iCol * Y_Step) + 190;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font12, XBrushes.Black, y + 10, x, frmDX);
                            /**** PREZZO FINE 1 ****/

                            /*** PLU INIZIO 2 ***/
                            //x = X + 150;
                            //y = Y + (iCol * Y_Step) + 340;
                            //s = "PLU " + (string)tabEti.Rows[i]["eti_plu"];
                            //gfx.DrawString(s, font7, XBrushes.Black, y + 10, x, frmDX);
                            /*** PLU FINE 2 ***/

                            /**** EURO INIZIO 2 
                            x = X + 180;
                            y = Y + (iCol * Y_Step) + 370;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);
                            *** EURO FINE 2 ****/

                            /*** UMI INIZIO 2 ***/
                            x = X + 175;
                            y = Y + (iCol * Y_Step) + 320;
                            s = "al ";
                            if ((string)tabEti.Rows[i]["eti_umi"] == "NR")
                                s += "PZ";
                            else
                                s += (string)tabEti.Rows[i]["eti_umi"];
                            gfx.DrawString(s, font7, XBrushes.Black, y + 10, x, frmDX);
                            /*** UMI FINE 2 ***/

                            /**** PREZZO INIZIO 2 ****/
                            x = X + 185;
                            y = Y + (iCol * Y_Step) + 450;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);
                            /**** PREZZO FINE 2 ****/

                            /**** BARCODE INIZIO ****/
                            x = X + 90;
                            y = Y + (iCol * Y_Step) + 450;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 100, 30);
                                    }
                                    else
                                    {
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 100, 30);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }
                            /**** BARCODE FINE ****/

                            if (i == 0)
                                Console.WriteLine("aaaa");



                            /*** Barcode su arf di DADO ***/
                            //x = X + 190;
                            //y = Y + (iCol * Y_Step) - 20 + 500;
                            //s = ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //if (s.Trim().Length > 0)
                            //{
                            //    try
                            //    {
                            //        //if (s.Length <= 8)
                            //        //{
                            //        //    //s = "0" + s;
                            //        //    //s += new clsCtrlCodici().FindMod10Digit(s);

                            //        //    Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                            //        //    MemoryStream ms = new MemoryStream();
                            //        //    img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //        //    gfx.DrawImage(img, y, x, 70, 14);
                            //        //}
                            //        //else
                            //        //{

                            //        s = "7901" + s.PadLeft(8, Convert.ToChar('0'));
                            //        s = s + new clsCtrlCodici().FindMod10Digit(s);

                            //        //s = "7901078492014";
                            //        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                            //        MemoryStream ms = new MemoryStream();
                            //        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //        gfx.DrawImage(img, y, x, 70, 14);

                            //        //}
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                            //    }
                            //}





                            /**** DATA INIZIO 1 ****/
                            //x = X + 205;
                            //y = Y + (iCol * Y_Step);
                            //s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /**** DATA FINE 1 ****/

                            /**** DATA INIZIO 2 ****/
                            //x = X + 205;
                            //y = Y + (iCol * Y_Step) + 300;
                            //s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /**** DATA FINE 2 ****/

                            ///**** PREZZO IN CHIARO INIZIO ****/
                            //x = X + 205;
                            //y = Y + (iCol * Y_Step) + 100;
                            //if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            //{
                            //    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            //    s = "€ al kg/L " + d.ToString("#####0.00");
                            //    //s = "al kg "; // +d.ToString("#####0.00");
                            //    gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            //}
                            ///**** PREZZO IN CHIARO FINE ****/

                            /**** EAN IN CIFRE INIZIO ****/
                            x = X + 130;
                            y = Y + (iCol * Y_Step) + 450;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /**** EAN IN CIFRE FINE ****/

                            /**** EAN IN CIFRE INIZIO ****/
                            //x = X + 205;
                            //y = Y + (iCol * Y_Step) + 420;
                            //s = ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /**** EAN IN CIFRE FINE ****/

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti055(DataTable tabEti, string strImgTip)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            //page.Height = 845.0;

            page.Height = 842.0;
            page.Width = 595;
            page.Orientation = PdfSharp.PageOrientation.Landscape;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 6.5, XFontStyle.Regular);
            XFont font2 = new XFont("Arial", 40, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            //XFont font3 = new XFont("Impact", 70, XFontStyle.Bold);
            XFont font3 = new XFont("Arial", 60, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 180, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 80, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 20, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 1;
            int iRows = 1;
            int i = 0;
            int X_Ini = 100;
            int X_Step = 100;
            int Y_Ini = 55;
            int Y_Step = 200;
            int X_Pie = 710;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        page.Height = 842.0;
                        page.Width = 595;
                        page.Orientation = PdfSharp.PageOrientation.Landscape;

                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = 0;
                            double y = 0;

                            /*** DATA OFFERTA inizio ***/
                            if ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ)
                            {
                                x = X + 160;
                                y = Y + (iCol * Y_Step) + 600;
                                //s = "Dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yyyy") + " al " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");
                                s = "Dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd");
                                s += " al " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd") + " ";
                                s += ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("MMMM") + " ";
                                s += ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("yyyy");
                                gfx.DrawString(s, font6, XBrushes.Black, y - 50, x - 15, XStringFormats.Default);
                            }
                            /*** DATA OFFERTA fine ***/

                            /*** DESCRIZIONE inizio ***/
                            x = X + -50;
                            y = Y + (iCol * Y_Step) - 50;
                            int iLen = 15;
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            //if (s.Length > 45)
                            //    s = s.Substring(0, 45);

                            string s2 = "";

                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;

                            XRect rect = new XRect(y, x, 800, 400);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            /*** DESCRIZIONE fine ***/

                            ///*** EURO inizio ***/
                            //x = X + 340;
                            //y = Y + (iCol * Y_Step) + 250;
                            //s = "€";
                            //gfx.DrawString(s, font4, XBrushes.Black, y, x - 5, XStringFormats.Default);
                            ///*** EURO fine ***/

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            x = X + 120;
                            y = Y + (iCol * Y_Step) - 15;

                            s = "";

                            if (strImgTip == "1")
                                s = "C:\\ApProject\\Temp\\Img\\Etichette\\eti055sottoCosto.png";
                            if (strImgTip == "2")
                                s = "C:\\ApProject\\Temp\\Img\\Etichette\\eti055unoPiuUno.png";
                            if (strImgTip == "3")
                                s = "C:\\ApProject\\Temp\\Img\\Etichette\\eti055prezzoShock.png";
                            if (strImgTip == "4")
                                s = "C:\\ApProject\\Temp\\Img\\Etichette\\eti055piaceriRisparmio.png";
                            if (strImgTip == "5")
                                s = "C:\\ApProject\\Temp\\Img\\Etichette\\eti055superOfferta.png";

                            if (File.Exists(s))
                            {
                                XImage img = XImage.FromFile(s);
                                XRect rect1 = new XRect(y, x, 250, 250);
                                gfx.DrawImage(img, rect1);
                            }

                            /*** PREZZO inizio ***/
                            x = X + 350;
                            y = Y + (iCol * Y_Step) + 780;
                            s = "€ " + Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y, x, frmDX);
                            /*** PREZZO fine ***/

                            /*** OFFERTA inizio ***/
                            x = X + 430;
                            y = Y + (iCol * Y_Step) - 20;

                            s = "€  " + ((decimal)tabEti.Rows[i]["eti_pve"]).ToString("#,##0.00");
                            gfx.DrawString(s, font3, XBrushes.Black, y, x, XStringFormats.Default);

                            XPen pen = new XPen(XColors.Black, 4);
                            gfx.DrawLine(pen, y + 50, x, y + 160, x + -35);

                            d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;
                            s = "- " + d.ToString("#,##0.00") + " %";
                            gfx.DrawString(s, font3, XBrushes.Black, y + 210, x, XStringFormats.Default);

                            /*** OFFERTA fine ***/

                            /*** ARTICOLO inizio ***/
                            x = X + 470;
                            y = Y + (iCol * Y_Step) - 10;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            /*** ARTICOLO fine ***/

                            /*** GRAMMATURA inizio ***/
                            x = X + 470;
                            y = Y + (iCol * Y_Step) + 730;
                            d = 0;
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG" &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0)
                                s = "€/KG"; ///L " + d.ToString("#####0.00");
                            else
                                s = "€/PZ";
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            /*** GRAMMATURA fine ***/

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 220;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie - 25;
                            //y = Y + (iCol * Y_Step) + 420;
                            //s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            //if (s.Trim().Length > 0)
                            //{
                            //    try
                            //    {
                            //        if (s.Length <= 8)
                            //        {
                            //            Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //        else
                            //        {
                            //            Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                            //            MemoryStream ms = new MemoryStream();
                            //            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            //            gfx.DrawImage(img, y, x, 70, 18);
                            //        }
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                            //    }
                            //}

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 58;
                            //s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            ////x = X + X_Pie;
                            ////y = Y + (iCol * Y_Step) + 305;
                            ////s = (string)tabEti.Rows[i]["eti_ean"];
                            ////gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 430;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti056(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 845.0;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 6.5, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Verdana", 11, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 38, XFontStyle.Bold);
            XFont font5 = new XFont("Impact", 40, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 8, XFontStyle.Regular);
            XFont font7 = new XFont("Times New Roman", 35, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 8;
            int i = 0;
            int X_Ini = -13;
            int Y_Ini = 15;
            //int X_Step = 108;
            //int Y_Step = 310;
            int X_Step = 108;
            int Y_Step = 255;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = X;
                            double y = Y + (iCol * Y_Step);

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            /*** DESCRIZIONE ***/
                            x = X + 10;
                            y = Y + (iCol * Y_Step);
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            //if (s.Length > 45)
                            //    s = s.Substring(0, 45);
                            XTextFormatter tf = new XTextFormatter(gfx);
                            XRect rect = new XRect(y, x, 250 + 00, 55);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect, XStringFormats.TopLeft);

                            /*** EURO ***/
                            x = X + 32 + 47;
                            y = Y + (iCol * Y_Step) + 10 + 60;
                            s = "€";
                            gfx.DrawString(s, font7, XBrushes.Black, y, x - 5, XStringFormats.Default);


                            /*** PREZZO ***/
                            //x = X + 32 + 55;
                            y = Y + (iCol * Y_Step) + 100 + 90;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);


                            /*** PESO NETTO ***/
                            x = X + 50;
                            y = Y + (iCol * Y_Step) + 0;
                            s = (string)tabEti.Rows[i]["eti_tgr"];
                            s += " " + Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]).ToString("0");
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);


                            /*** PREZZO IN CHIARO ***/
                            x = X + 63;
                            y = Y + (iCol * Y_Step) + 0;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0 && d <= 500m)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }

                            /*** BARCODE ***/
                            x = X + 58 + 10;
                            y = Y + (iCol * Y_Step) + 0;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                    else
                                    {
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            if (i == 0)
                                Console.WriteLine("aaaa");


                            s = "C:\\ApProject\\Temp\\Img\\Loghi\\dPiu.png";

                            if (File.Exists(s))
                            {
                                x = X + 30;
                                y = Y + (iCol * Y_Step) + 200;

                                XImage img = XImage.FromFile(s);
                                gfx.DrawImage(img, y, x + 20, 60, 40);
                            }


                            /*** BARCODE ***/
                            x = X + 95;
                            y = Y + (iCol * Y_Step) + 0;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** ARTICOLO FOR INIZIO ***/
                            x = X + 95;
                            y = Y + (iCol * Y_Step) + 70;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** ARTICOLO  ***/
                            x = X + 95;
                            y = Y + (iCol * Y_Step) + 130;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** PXC ***/
                            x = X + 95;
                            y = Y + (iCol * Y_Step) + 170;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 95;
                            y = Y + (iCol * Y_Step) + 200;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti057(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 845.0;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 6.5, XFontStyle.Regular);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Verdana", 10, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 30, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 10, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 9, XFontStyle.Regular);
            XFont font7 = new XFont("Courier new", 7.5, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 7;
            int i = 0;
            int X_Ini = 45;
            int X_Step = 112;
            int Y_Ini = 70;
            int Y_Step = 230;
            int i_Step = 0;

            //if (strPar != "")
            //{
            //    string[] a = strPar.Split(';');
            //    foreach (string ss in a)
            //    {
            //        string[] aa = ss.Split('=');
            //        if (aa.Length > 1)
            //        {
            //            if (_clsFun.Numerico(aa[1], "0123456789"))
            //            {
            //                i = Convert.ToInt16(aa[1]);
            //                if (i > 0)
            //                {
            //                    if (aa[0] == "xIni")
            //                        X_Ini = i;
            //                    if (aa[0] == "xStep")
            //                        X_Step = i;
            //                    if (aa[0] == "yIni")
            //                        Y_Ini = i;
            //                    if (aa[0] == "yStep")
            //                        Y_Step = i;
            //                    if (aa[0] == "iCols")
            //                        iCols = i;
            //                    if (aa[0] == "iRows")
            //                        iRows = i;
            //                }
            //            }
            //        }
            //    }
            //}

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;
            i = 0;

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step) + i_Step;
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = X;
                            double y = Y + (iCol * Y_Step);
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            if (s.Length > 45)
                                s = s.Substring(0, 45);

                            //string s2 = "";
                            //string[] a = s.Split(' ');

                            //if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                            //    Console.WriteLine("aaaa");

                            //for (int i2 = 0; i2 < a.Length; i2++)
                            //{
                            //    if (a[i2].Length > 20)
                            //        a[i2] = a[i2].Substring(0, 10) + " " + a[i2].Substring(10, 10) + " " + a[i2].Substring(20);

                            //    else if (a[i2].Length > 10)
                            //        a[i2] = a[i2].Substring(0, 10) + " " + a[i2].Substring(10);

                            //    s2 += a[i2] + " ";
                            //}

                            //s = s2;
                            if (i == 1)
                                Console.WriteLine("aaaa");

                            //gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);
                            //XRect rect = new XRect(40, 100, 250, 220);

                            XTextFormatter tf = new XTextFormatter(gfx);
                            XRect rect = new XRect(y + 20, x, 120, 40);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //tf.Alignment = ParagraphAlignment.Left; 
                            tf.DrawString(s, font3, XBrushes.Black, rect, XStringFormats.TopLeft);

                            x = X + 56;
                            //y = Y + (iCol * Y_Step) + 95;
                            y = Y + (iCol * Y_Step) + 145;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 56;
                            //y = Y + (iCol * Y_Step) + 135;
                            y = Y + (iCol * Y_Step) + 185;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 75;
                            y = Y + (iCol * Y_Step) + 5;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 78;
                            y = Y + (iCol * Y_Step) + 80;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            x = X + 48;
                            //y = Y + (iCol * Y_Step) + 105;
                            y = Y + (iCol * Y_Step) + 155;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0 && d <= 500m)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }

                            x = X + 58;
                            //y = Y + (iCol * Y_Step) + 95;
                            y = Y + (iCol * Y_Step) + 145;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                    else
                                    {
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            x = X + 81;
                            y = Y + (iCol * Y_Step) + 0;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 83;
                            y = Y + (iCol * Y_Step) + 30;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 81;
                            //y = Y + (iCol * Y_Step) + 105;
                            y = Y + (iCol * Y_Step) + 155;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                pd.Close();
                pd.Dispose();
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        /* Etichette per NDO con possibilità di impostare la descrizione principale 
         SPENDO MENO
         PREZZO SHOCK
         GRANDE OFFERTA
        */
        public string PrnPdfEti058(DataTable tabEti, string strTit)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            //page.Height = 845.0;

            page.Height = 842.0;
            page.Width = 595;
            page.Orientation = PdfSharp.PageOrientation.Landscape;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Regular);
            XFont font2 = new XFont("Verdana", 80, XFontStyle.Bold);
            XFont font3 = new XFont("Arial", 60, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 180, XFontStyle.Regular);
            XFont font5 = new XFont("Arial", 80, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 20, XFontStyle.Regular);
            XFont font7 = new XFont("Arial", 11, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 1;
            int iRows = 1;
            int i = 0;
            int X_Ini = 100;
            int X_Step = 100;
            int Y_Ini = 55;
            int Y_Step = 200;
            int X_Pie = 480;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        page.Orientation = PdfSharp.PageOrientation.Landscape;
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = X;
                            double y = Y;

                            /*
                             * TITOLO
                             */
                            x = X;
                            y = Y + (iCol * Y_Step) - 45;
                            s = strTit; // "PREZZO SPECIALE";
                            gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             * DESCRIZIONE
                             */
                            int iLen = 15;
                            s = (string)tabEti.Rows[i]["eti_ard"];

                            string s2 = "";

                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;

                            x = X + 50;
                            y = Y + (iCol * Y_Step);

                            XRect rect = new XRect(y, x, 750, 600);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            /*
                             * PREZZO
                             */
                            x = X + 390;
                            y = Y + (iCol * Y_Step) + 500;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);


                            /*
                             * EURO
                             */
                            x = X + 360;
                            y = Y + (iCol * Y_Step) - 20;
                            s = "€.";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);



                            /*
                            * PREZZO IN CHIARO
                            */
                            x = X + 370;
                            y = Y + (iCol * Y_Step) + 550;
                            d = 0;
                            //Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG" &&
                            if (
                                Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }


                            /*
                             * AL PZ
                             */
                            x = X + X_Pie - 40;
                            y = Y + (iCol * Y_Step) - 20;
                            s = "al pz.";
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s = "al Kg";
                            gfx.DrawString(s, font6, XBrushes.Black, y, x - 5, XStringFormats.Default);


                            /*
                             * TIPO GRAMMATURA
                             */
                            x = X + X_Pie - 43;
                            y = Y + (iCol * Y_Step) + 680;
                            s = (string)tabEti.Rows[i]["eti_tgr"];
                            s += " " + Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]).ToString("#0.00");
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             * LINEA
                             */
                            x = X + X_Pie - 40;
                            y = Y + (iCol * Y_Step) - 50;
                            XPen pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y, x, y + 850, x);


                            /*
                             * PIEDE DESCRIZIONE
                             */
                            x = X + X_Pie - 20;
                            y = Y + (iCol * Y_Step) - 40;
                            s = "Codice articolo                  PLU                             Codice a Barre                                       Pxc                                   Data di Stampa";
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * ARTICOLO
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) - 30;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * PLU
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 85;
                            s = (string)tabEti.Rows[i]["eti_plu"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * BARCODE
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 190;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * PXC
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 380;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                            * DATA
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 520;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            ///*
                            //* FORNITORE
                            //*/
                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 220;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * BARCODE
                            */
                            x = X + X_Pie - 25;
                            y = Y + (iCol * Y_Step) + 700;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                    else
                                    {
                                        s = s.PadLeft(13, Convert.ToChar('0'));
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }


                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 305;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti059(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 845.0;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 9.5, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 15, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Verdana", 15, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 50, XFontStyle.BoldItalic);
            XFont font5 = new XFont("Arial", 18, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 4;
            int i = 0;
            int X_Ini = 88;
            int X_Step = 175;
            int Y_Ini = 70;
            int Y_Step = 230;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            if (i == 1)
                                Console.WriteLine("aaaa");

                            double x = X;
                            double y = Y + (iCol * Y_Step);
                            int iLen = 15;
                            x = X + 10;
                            y = Y + (iCol * Y_Step) + 5;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 20;
                            y = Y + (iCol * Y_Step) + 5;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            s = (string)tabEti.Rows[i]["eti_ard"];
                            string s2 = "";
                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;


                            x = X + 42;
                            y = Y + (iCol * Y_Step) + 0;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            XRect rect = new XRect(y, x, 220, 90);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            x = X + 137;
                            y = Y + (iCol * Y_Step) + 6;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 150;
                            y = Y + (iCol * Y_Step) + 180;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            x = X + 130;
                            y = Y + (iCol * Y_Step) + 20;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            //s = "€ al kg/L " + d.ToString("#####0.00");
                            if (d > 0 && d <= 500m)
                            {
                                s = "al kg "; // +d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }

                            x = X + 150;
                            y = Y + (iCol * Y_Step) + 130;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 10);
                                    }
                                    else
                                    {
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 10);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 160;
                            y = Y + (iCol * Y_Step) + 50;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string Prova_PrnPdfEti059(DataTable tabEti)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            //Barcode _ean13 = new BarcodeLib.Barcode();
            //BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            //_ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean128 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan128 = BarcodeLib.TYPE.CODE128;
            _ean128.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 845.0;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 9.5, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 15, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Verdana", 15, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 50, XFontStyle.BoldItalic);
            XFont font5 = new XFont("Arial", 18, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 4;
            int i = 0;
            int X_Ini = 88;
            int X_Step = 175;
            int Y_Ini = 70;
            int Y_Step = 230;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            if (i == 1)
                                Console.WriteLine("aaaa");

                            string sQta = "15";
                            string sLot = "E-00042-2020";

                            double x = X;
                            double y = Y + (iCol * Y_Step);
                            int iLen = 15;
                            x = X + 10;
                            y = Y + (iCol * Y_Step) + 5;
                            string sDay = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString("Data: " + sDay, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 20;
                            y = Y + (iCol * Y_Step) + 5;
                            //string sArt = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString("Q.tà " + sQta, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 30;
                            y = Y + (iCol * Y_Step) + 5;
                            //string sArt = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString("Lotto " + sLot, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            s = (string)tabEti.Rows[i]["eti_ard"];
                            string s2 = "";
                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            string sArd = s;

                            x = X + 42;
                            y = Y + (iCol * Y_Step) + 0;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            XRect rect = new XRect(y, x, 220, 90);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(sArd, font3, XBrushes.Black, rect);

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            //x = X + 137;
                            //y = Y + (iCol * Y_Step) + 6;
                            //s = "€";
                            //gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            //x = X + 150;
                            //y = Y + (iCol * Y_Step) + 180;
                            //s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            //gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            //x = X + 130;
                            //y = Y + (iCol * Y_Step) + 20;
                            //if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            //    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            ////s = "€ al kg/L " + d.ToString("#####0.00");
                            //s = "al kg "; // +d.ToString("#####0.00");
                            //gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 100;
                            y = Y + (iCol * Y_Step) + 1;

                            s = "123";

                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length == 0)
                            {
                                try
                                {
                                    //if (s.Length <= 8)
                                    //{

                                    s = "15 15/02/2019 E-00042-2020";
                                    s = sQta + " " + sDay + " " + sLot;

                                    Image img = _ean128.Encode(tpEan128, s, Color.Black, Color.White, 800, 400);
                                    MemoryStream ms = new MemoryStream();
                                    img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                    gfx.DrawImage(img, y, x, 270, 50);
                                    //}
                                    //else
                                    //{
                                    //    Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                    //    MemoryStream ms = new MemoryStream();
                                    //    img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                    //    gfx.DrawImage(img, y, x, 70, 10);
                                    //}
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 160;
                            y = Y + (iCol * Y_Step) + 50;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti060(DataTable tabEti, string strImgTip)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 845.0;

            //System.Drawing.Size size = PageSizeConverter.ToSize(PdfSharp.PageSize.A4);
            //page.Width = size.Width;
            //page.Height = size.Height;


            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 6.5, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Verdana", 12, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 38, XFontStyle.Bold);
            XFont font5 = new XFont("Impact", 42, XFontStyle.Bold);
            XFont font6 = new XFont("Arial", 8, XFontStyle.Regular);
            XFont font7 = new XFont("Times New Roman", 35, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 8;
            int i = 0;
            int X_Ini = -13;
            int Y_Ini = 15;
            //int X_Step = 108;
            //int X_Step = 100;
            int X_Step = 108;
            int Y_Step = 310;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = X;
                            double y = Y + (iCol * Y_Step);

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            /*** DESCRIZIONE ***/
                            x = X + 10;
                            y = Y + (iCol * Y_Step);
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            //if (s.Length > 45)
                            //    s = s.Substring(0, 45);
                            XTextFormatter tf = new XTextFormatter(gfx);
                            XRect rect = new XRect(y, x, 200 + 00, 60);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect, XStringFormats.TopLeft);

                            /*** EURO ***/
                            x = X + 32 + 47;
                            y = Y + (iCol * Y_Step) + 10 + 60;
                            s = "€";
                            gfx.DrawString(s, font7, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            /*** PREZZO ***/
                            //x = X + 32 + 55;
                            y = Y + (iCol * Y_Step) + 100 + 70;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("#0.00");
                            gfx.DrawString(s, font5, XBrushes.Black, y + 10, x+10, frmDX);

                            /*** IMMAGINE ***/
                            s = "";
                            if (strImgTip == "1")
                                s = "C:\\ApProject\\Temp\\Img\\Etichette\\eti060PrezziBassi.png";
                            if (strImgTip == "2")
                                s = "C:\\ApProject\\Temp\\Img\\Etichette\\eti060Novita.png";

                            if (s != "" && File.Exists(s))
                            {
                                x = X + 0;
                                y = Y + (iCol * Y_Step) + 190;

                                XImage img = XImage.FromFile(s);
                                gfx.DrawImage(img, y, x + 20, 65, 65);
                            }

                            /*** PESO NETTO ***/
                            x = X + 50;
                            y = Y + (iCol * Y_Step) + 0;
                            s = (string)tabEti.Rows[i]["eti_tgr"];
                            s += " " + Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]).ToString("0");
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);


                            /*** PREZZO IN CHIARO ***/
                            x = X + 63;
                            y = Y + (iCol * Y_Step) + 0;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0 && d <= 500m)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }

                            /*** BARCODE ***/
                            x = X + 58 + 10;
                            y = Y + (iCol * Y_Step) + 0;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                    else
                                    {
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            s = "C:\\ApProject\\Temp\\Img\\Loghi\\dPiu.png";
                            if (File.Exists(s))
                            {
                                x = X + 30;
                                y = Y + (iCol * Y_Step) + 200;

                                XImage img = XImage.FromFile(s);
                                gfx.DrawImage(img, y, x + 20, 60, 40);
                            }

                            /*** BARCODE ***/
                            x = X + 95;
                            y = Y + (iCol * Y_Step) + 0;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** ARTICOLO FOR INIZIO ***/
                            x = X + 95;
                            y = Y + (iCol * Y_Step) + 60;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** ARTICOLO  ***/
                            x = X + 95;
                            y = Y + (iCol * Y_Step) + 120;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** PXC ***/
                            x = X + 95;
                            y = Y + (iCol * Y_Step) + 150;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            if ((string)tabEti.Rows[i]["eti_oft"] == "PRO" || ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ))
                            {

                                //if ((string)tabEti.Rows[i]["eti_cam"] == "" && !File.Exists(s))     //In caso di fidelity stampo le date sotto
                                //{
                                //    x = X + 0;
                                //    y = Y + (iCol * Y_Step) + 15;
                                //    s = "Offerta valida";
                                //    gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);
                                //    x = X + 10;
                                //    y = Y + (iCol * Y_Step) + 15;
                                //    s = "dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yyyy");
                                //    gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);
                                //    x = X + 20;
                                //    y = Y + (iCol * Y_Step) + 15;
                                //    s = "al  " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");
                                //    gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);

                                //}

                                //d = 0;
                                //if ((decimal)tabEti.Rows[i]["eti_prv"] != 0 && (decimal)tabEti.Rows[i]["eti_pve"] != 0)
                                //    d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                //x = X + 45;
                                //y = Y + (iCol * Y_Step) + 110;
                                //s = "da € " + Convert.ToDecimal(tabEti.Rows[i]["eti_pve"]).ToString("######0.00") + " - ";
                                //s += "sconto " + d.ToString("#,##0.00") + " %";
                                //gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);

                                //    s = "al  " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");
                                //    gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);

                                x = X + 95;
                                y = Y + (iCol * Y_Step) + 190;
                                s = "FINO AL " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yy");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }

        public string PrnPdfEti061(DataTable tabEti, string strTit)
        {
            string s = "";
            decimal d = 0;
            string sMsg = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();
            //page.Height = 845.0;

            page.Height = 842.0;
            page.Width = 595;
            page.Orientation = PdfSharp.PageOrientation.Landscape;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 10, XFontStyle.Regular);
            XFont font2 = new XFont("Verdana", 40, XFontStyle.Bold);
            XFont font3 = new XFont("Arial", 30, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 90, XFontStyle.Regular);
            XFont font5 = new XFont("Arial", 40, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 15, XFontStyle.Regular);
            XFont font7 = new XFont("Arial", 11, XFontStyle.Regular);
            XFont font8 = new XFont("Arial", 20, XFontStyle.Bold);
            XFont font9 = new XFont("Impact", 30, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 1;
            int i = 0;
            int X_Ini = 100;
            int X_Step = 100;
            int Y_Ini = 45;
            int Y_Step = 425;
            int X_Pie = 480;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                    t.ImportRow(y);
            }
            tabEti = t.Copy();

            while (true)
            {
                //for (int iRow = 0; iRow <= iRows - 1; iRow++)
                //{
                if (iRow >= iRows)
                {
                    if (i < tabEti.Rows.Count)
                    {
                        page = pd.AddPage();
                        page.Orientation = PdfSharp.PageOrientation.Landscape;
                        gfx = XGraphics.FromPdfPage(page);
                    }
                    X = X_Ini;
                    iRow = 0;
                }

                if (i < tabEti.Rows.Count)
                {
                    Y = Y_Ini;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = X;
                            double y = Y;

                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            XPen pen = new XPen(XColors.Black, 1);

                            /*
                             * TITOLO
                             */
                            //x = X;
                            //y = Y + (iCol * Y_Step) - 45;

                            x = X -70;
                            y = Y + (iCol * Y_Step);

                            s = strTit; // "PREZZO SPECIALE";
                            //gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

                            XRect rect = new XRect(y, x, 375, 600);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font2, XBrushes.Black, rect);


                            /*
                             * DESCRIZIONE
                             */
                            int iLen = 15;
                            s = (string)tabEti.Rows[i]["eti_ard"];

                            string s2 = "";

                            if ((string)tabEti.Rows[i]["eti_art"] == "0003316")
                                Console.WriteLine("aaaa");
                            s2 = s;
                            for (int i3 = 0; i3 < 4; i3++)
                            {
                                string[] a = s2.Split(' ');
                                s2 = "";

                                for (int i2 = 0; i2 < a.Length; i2++)
                                {
                                    if (a[i2].Length > iLen)
                                        a[i2] = a[i2].Substring(0, iLen) + " " + a[i2].Substring(iLen);

                                    s2 += a[i2] + " ";
                                }
                            }

                            if (s2 != "")
                                s = s2;

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            x = X + 80;
                            y = Y + (iCol * Y_Step);

                            rect = new XRect(y, x, 375, 600);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            /*
                            * OFFERTA
                            */
                            x = X + 190;
                            //y = Y - 25; // +(iCol * Y_Step) + 190;
                            y = Y + (iCol * Y_Step) - 40;
                            tf = new XTextFormatter(gfx);
                            rect = new XRect(y, x+10, 400, 60);
                            if (true)
                            {
                                gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                                //tf.Alignment = XParagraphAlignment.Default; //tf.Alignment = ParagraphAlignment.Left; 
                                //tf.DrawString(s, font7, XBrushes.Black, rect, XStringFormats.TopLeft);

                                d = 0;
                                if ((decimal)tabEti.Rows[i]["eti_pve"] > 0 && (decimal)tabEti.Rows[i]["eti_prv"] > 0)
                                    d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                if (d != 0)
                                {
                                    s = ((decimal)tabEti.Rows[i]["eti_pve"]).ToString("#,##0.00");
                                    gfx.DrawString(s, font8, XBrushes.Black, y + 30, x + 50, XStringFormats.Default);

                                    pen = new XPen(XColors.Black, 2);
                                    gfx.DrawLine(pen, y + 20, x + 60, y + 100, x + 25);
                                }

                                s = "OFFERTA";
                                gfx.DrawString(s, font9, XBrushes.Black, y + 150, x + 50, XStringFormats.Default);

                                if (d != 0)
                                {
                                    s = "Sc." + d.ToString("#,##0.00") + " %";
                                    gfx.DrawString(s, font8, XBrushes.Black, y + 295, x + 50, XStringFormats.Default);
                                }
                            }

                            /*
                             * PREZZO
                             */
                            x = X + 390;
                            y = Y + (iCol * Y_Step) + 250;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);


                            /*
                             * EURO
                             */
                            x = X + 360;
                            y = Y + (iCol * Y_Step) - 20;
                            s = "€.";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);



                            /*
                            * PREZZO IN CHIARO
                            */
                            x = X + 370;
                            y = Y + (iCol * Y_Step) + 275;
                            d = 0;
                            //Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG" &&
                            if (
                                Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }


                            /*
                             * AL PZ
                             */
                            x = X + X_Pie - 40;
                            y = Y + (iCol * Y_Step) - 20;
                            s = "al pz.";
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s = "al Kg";
                            gfx.DrawString(s, font6, XBrushes.Black, y, x - 5, XStringFormats.Default);


                            /*
                             * TIPO GRAMMATURA
                             */
                            x = X + X_Pie - 43;
                            y = Y + (iCol * Y_Step) + 290;
                            s = (string)tabEti.Rows[i]["eti_tgr"];
                            s += " " + Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]).ToString("#0.00");
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                             * LINEA
                             */
                            x = X + X_Pie - 40;
                            y = Y + (iCol * Y_Step) - 50;
                            pen = new XPen(XColors.Black, 1);
                            gfx.DrawLine(pen, y, x, y + 850, x);


                            /*
                             * PIEDE DESCRIZIONE
                             */
                            x = X + X_Pie - 20;
                            y = Y + (iCol * Y_Step) - 40;
                            s = "Codice articolo    PLU        Codice a Barre       Pxc        Data";
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * ARTICOLO
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) - 30;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * PLU
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 38;
                            s = (string)tabEti.Rows[i]["eti_plu"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * BARCODE
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 87;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * PXC
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 173;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*
                            * DATA
                            */
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 225;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            ///*
                            //* FORNITORE
                            //*/
                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 220;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            /*
                            * BARCODE
                            */
                            x = X + X_Pie - 25;
                            y = Y + (iCol * Y_Step) + 290;
                            s = ((string)tabEti.Rows[i]["eti_ean"]).Trim();
                            if (s.Trim().Length > 0)
                            {
                                try
                                {
                                    if (s.Length <= 8)
                                    {
                                        Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                    else
                                    {
                                        s = s.PadLeft(13, Convert.ToChar('0'));
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 18);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }


                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 305;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            i++;
                            //break;
                        }
                    }
                    iRow++;
                    //break;
                }

                //X += X_Step;
                //}
                if (i >= tabEti.Rows.Count)
                    break;
            }
            //X += 15;

            // Save the document...
            try
            {
                string sFil = "C:\\APproject\\PDF\\Eti_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch
            {
                sMsg += "Stampa già aperta.";
            }

            return sMsg;
        }



    }
}
