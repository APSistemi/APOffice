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
    class clsGenPdfEti
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        //public DataTable _tabEti = new DataTable("tabTmp");
        private string _strConSql = "";
        public int StartRow = 1;

        public clsGenPdfEti()
        {
            _strConSql = _clsFun.ConSql("");
        }

        public string PrnPdfEti001(DataTable tabEti, string strPar)
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

            int iCols = 3;
            int iRows = 8;
            int i = 0;
            int X_Ini = 5;
            int X_Step = 108;
            int Y_Ini = 15;
            int Y_Step = 200;
            int i_Step = 0;

            if(strPar != "")
            {
                string[] a = strPar.Split(';');
                foreach (string ss in a)
                {
                    string[] aa = ss.Split('=');
                    if(aa.Length > 1)
                    {
                        if(_clsFun.Numerico(aa[1],"0123456789"))
                        {
                            i = Convert.ToInt16(aa[1]);
                            if(i > 0)
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
                            if(i == 1)
                                Console.WriteLine("aaaa");

                            //gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);
                            //XRect rect = new XRect(40, 100, 250, 220);

                            XTextFormatter tf = new XTextFormatter(gfx); 
                            XRect rect = new XRect(y+20, x, 120, 40);
                            gfx.DrawRectangle(XBrushes.Transparent, rect); 
                            //tf.Alignment = ParagraphAlignment.Left; 
                            tf.DrawString(s, font3, XBrushes.Black, rect, XStringFormats.TopLeft);

                            x = X + 56;
                            y = Y + (iCol * Y_Step) + 95;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 56;
                            y = Y + (iCol * Y_Step)+135;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 75;
                            y = Y + (iCol * Y_Step) + 5;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x-5, XStringFormats.Default);

                            x = X + 78;
                            y = Y + (iCol * Y_Step) + 80;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y+10, x, frmDX);

                            x = X + 48;
                            y = Y + (iCol * Y_Step) + 105;
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
                            y = Y + (iCol * Y_Step) + 105;
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

        public string PrnPdfEti002(DataTable tabEti)
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
            XFont font2 = new XFont("Arial", 8, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Verdana", 10, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 30, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 12, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 9, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 3;
            int iRows = 8;
            int i = 0;
            int X_Ini = 0;
            //int X_Step = 108;
            int X_Step = 100;
            int Y_Ini = 15;
            int Y_Step = 200;

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

                            XRect rect = new XRect(y, x, 120, 40);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //tf.Alignment = ParagraphAlignment.Left; 
                            tf.DrawString(s, font3, XBrushes.Black, rect, XStringFormats.TopLeft);

                            x = X + 56;
                            y = Y + (iCol * Y_Step) + 95;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 56;
                            y = Y + (iCol * Y_Step) + 135;
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
                            y = Y + (iCol * Y_Step) + 105;
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

                            x = X + 83;
                            y = Y + (iCol * Y_Step) + 3;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 83;
                            y = Y + (iCol * Y_Step) + 26;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 83;
                            y = Y + (iCol * Y_Step) + 102;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

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

        public string PrnPdfEti003(DataTable tabEti)
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
            XFont font3 = new XFont("Impact", 70, XFontStyle.Bold);
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

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            x = X +30;

                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;

                            XRect rect = new XRect(y, x, 500, 400);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            x = X + 610;
                            y = Y + (iCol * Y_Step) + 500;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            x = X + 650;
                            y = Y + (iCol * Y_Step) - 20;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 670;
                            y = Y + (iCol * Y_Step) - 20;
                            s = "al pz.";
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s = "al Kg";
                            gfx.DrawString(s, font6, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 630;
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

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 180;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 220;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie-25;
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

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 305;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 330;
                            s = (string)tabEti.Rows[i]["eti_tgr"];
                            s += " " + Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]).ToString("#0.00");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 430;
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

        public string PrnPdfEti004(DataTable tabEti)
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
            XFont font3 = new XFont("Verdana", 18, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 60, XFontStyle.BoldItalic);
            XFont font5 = new XFont("Arial", 18, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 4;
            int i = 0;
            int X_Ini = 30;
            int X_Step = 210;
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

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            x = X + 42;
                            y = Y + (iCol * Y_Step) + 0;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            XRect rect = new XRect(y, x, 260, 100);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            x = X + 137;
                            y = Y + (iCol * Y_Step) + 6;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 180;
                            y = Y + (iCol * Y_Step) + 220;
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

                            x = X + 178;
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

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 200;
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

        public string PrnPdfEti005(DataTable tabEti)
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
            XFont font5 = new XFont("Arial", 12, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 9, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 3;
            int iRows = 7;
            int i = 0;
            int X_Ini = 0;
            int Y_Ini = 15;
            //int X_Step = 108;
            //int X_Step = 100;
            int X_Step = 112;
            int Y_Step = 200;

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

                            XRect rect = new XRect(y, x, 120, 40);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //tf.Alignment = ParagraphAlignment.Left; 
                            tf.DrawString(s, font3, XBrushes.Black, rect, XStringFormats.TopLeft);

                            x = X + 56;
                            y = Y + (iCol * Y_Step) + 135;
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
                            y = Y + (iCol * Y_Step) + 105;
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
                            y = Y + (iCol * Y_Step) + 5;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 81;
                            y = Y + (iCol * Y_Step) + 58;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 81;
                            y = Y + (iCol * Y_Step) + 105;
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

        public string PrnPdfEti006(DataTable tabEti)
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
            XFont font4 = new XFont("Impact", 25, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 12, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 9, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 3;
            int iRows = 7;
            int i = 0;
            int X_Ini = 0;
            int Y_Ini = 15;
            //int X_Step = 108;
            //int X_Step = 100;
            int X_Step = 112;
            int Y_Step = 200;

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

                            x = X + 25;
                            y = Y + (iCol * Y_Step) + 100;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 27;
                            y = Y + (iCol * Y_Step) + 160;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            x = X + 25;
                            y = Y + (iCol * Y_Step);
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            if (s.Length > 45)
                                s = s.Substring(0, 45);
                            XTextFormatter tf = new XTextFormatter(gfx);
                            XRect rect = new XRect(y, x, 190, 40);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect, XStringFormats.TopLeft);

                            x = X + 56;
                            y = Y + (iCol * Y_Step) + 135;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 60;
                            y = Y + (iCol * Y_Step) + 2;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0 && d <= 500m)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }

                            x = X + 70;
                            y = Y + (iCol * Y_Step) + 2;
                            s = (string)tabEti.Rows[i]["eti_tgr"];
                            s += " " + Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]).ToString();
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 58;
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
                            y = Y + (iCol * Y_Step) + 2;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 81;
                            y = Y + (iCol * Y_Step) + 58;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 81;
                            y = Y + (iCol * Y_Step) + 105;
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

        public string PrnPdfEti007(DataTable tabEti)
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
            XFont font3 = new XFont("Verdana", 18, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 60, XFontStyle.BoldItalic);
            XFont font5 = new XFont("Arial", 18, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);
            XFont font7 = new XFont("Verdana", 28, XFontStyle.Bold);

            XFont font8 = new XFont("Arial", 20, XFontStyle.Bold);
            XFont font9 = new XFont("Impact", 25, XFontStyle.Bold);
            XFont font10 = new XFont("Arial", 12, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 4;
            int i = 0;
            int X_Ini = 0;
            int X_Step = 218;
            int Y_Ini = 10;
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
                            double x = X;
                            double y = Y + (iCol * Y_Step);
                            int iLen = 15;
                            //x = X + 10;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            // 20191024 disattivato plu per Venfri

                            x = X + 48;
                            y = Y + (iCol * Y_Step) + 195;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            XRect rect = new XRect(y, x, 85, 40);
                            if (((string)tabEti.Rows[i]["eti_plu"]).Length > 0 && ((string)tabEti.Rows[i]["eti_plu"]).Length < 4)
                            {
                                s = (string)tabEti.Rows[i]["eti_plu"];
                                if (_clsFun.Numerico(s))
                                {
                                    s = Convert.ToInt32(s).ToString();
                                }

                                //gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);
                                //XPen pen = new XPen(XColors.RoyalBlue, Math.PI);
                                gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                                tf.Alignment = XParagraphAlignment.Center; //tf.Alignment = ParagraphAlignment.Left; 
                                tf.DrawString(s, font7, XBrushes.Black, rect, XStringFormats.TopLeft);
                            }

                            //x = X + 20;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = (string)tabEti.Rows[i]["eti_art"];
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

                            x = X + 20;
                            y = Y + (iCol * Y_Step) + 15;
                            tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            rect = new XRect(y, x, 260, 100);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            if ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ)
                            {
                                //s = "OFFERTA";

                                x = X + 80;
                                y = Y + (iCol * Y_Step) - 10; // +(iCol * Y_Step) + 190;
                                tf = new XTextFormatter(gfx);
                                rect = new XRect(y, x, 290, 30);

                                gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);

                                if ((decimal)tabEti.Rows[i]["eti_pve"] > 0 && (decimal)tabEti.Rows[i]["eti_prv"] > 0)
                                    d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                if (d != 0)
                                {
                                    s = ((decimal)tabEti.Rows[i]["eti_pve"]).ToString("#,##0.00");
                                    gfx.DrawString(s, font8, XBrushes.Black, y + 20, x + 25, XStringFormats.Default);

                                    XPen pen = new XPen(XColors.Black, 3);
                                    gfx.DrawLine(pen, y + 17, x + 25, y + 60, x + 15);
                                }

                                s = "OFFERTA";
                                gfx.DrawString(s, font9, XBrushes.White, y + 95, x + 25, XStringFormats.Default);

                                if (d != 0)
                                {
                                    s = "Sc." + d.ToString("#,##0.00") + " %";
                                    gfx.DrawString(s, font10, XBrushes.Black, y + 220, x + 25, XStringFormats.Default);
                                }
                            }

                            x = X + 160;
                            y = Y + (iCol * Y_Step) + 15;
                            s = "€";
                            gfx.DrawString(s, font7, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 170;
                            y = Y + (iCol * Y_Step) + 220;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            x = X + 185;
                            y = Y + (iCol * Y_Step) + 5;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 185;
                            y = Y + (iCol * Y_Step) + 65;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 185;
                            d = 0;
                            y = Y + (iCol * Y_Step) + 120;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            {
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                                //s = "€ al kg/L " + d.ToString("#####0.00");

                                if(Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) != d)
                                {
                                    if (d > 0 && d <= 500m)
                                    {
                                        s = "al Kg/Lt € " + d.ToString("#####0.00");
                                        gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                                    }
                                }
                            }
                            x = X + 168;
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

        public string PrnPdfEti008(DataTable tabEti)
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
            XFont font3 = new XFont("Impact", 70, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 180, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 80, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 20, XFontStyle.Regular);
            XFont font7 = new XFont("Arial", 25, XFontStyle.Bold);
            XFont font8 = new XFont("Arial", 50, XFontStyle.Bold);
            XFont font9 = new XFont("Impact", 80, XFontStyle.Bold);
            XFont font10 = new XFont("Courier new", 12, XFontStyle.Bold);

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

                            XPen pen = new XPen(XColors.Black, 4);
                            XRect rect = new XRect();

                            x = X;
                            y = Y + (iCol * Y_Step);
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

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;


                            rect = new XRect(y, x-50, 500, 400);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            x = X + 300;
                            y = Y-50; // +(iCol * Y_Step) + 190;
                            tf = new XTextFormatter(gfx);
                            rect = new XRect(y, x, 585, 80);
                            if (true)
                            {
                                s = "OFFERTA";

                                //gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                                gfx.DrawRectangle(new SolidBrush(Color.White), rect);
                                //tf.Alignment = XParagraphAlignment.Default; //tf.Alignment = ParagraphAlignment.Left; 
                                //tf.DrawString(s, font7, XBrushes.Black, rect, XStringFormats.TopLeft);

                                //d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;
                                if ((decimal)tabEti.Rows[i]["eti_pve"] > 0 && (decimal)tabEti.Rows[i]["eti_prv"] > 0)
                                    d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;


                                if(d != 0)
                                {
                                    s = ((decimal)tabEti.Rows[i]["eti_pve"]).ToString("#,##0.00");
                                    gfx.DrawString(s, font8, XBrushes.Black, y + 20, x + 60, XStringFormats.Default);

                                    pen = new XPen(XColors.Black, 4);
                                    gfx.DrawLine(pen, y + 20, x + 60, y + 120, x + 25);
                                }

                                s = "OFFERTA";
                                //gfx.DrawString(s, font9, XBrushes.White, y + 190, x + 65, XStringFormats.Default);
                                gfx.DrawString(s, font9, XBrushes.Red, y + 160, x + 65, XStringFormats.Default);

                                if (d != 0)
                                {
                                    s = "Sc." + d.ToString("#,##0.00") + " %";
                                    gfx.DrawString(s, font7, XBrushes.Black, y + 445, x + 60, XStringFormats.Default);
                                }
                            }

                            x = X + 650;
                            y = Y + (iCol * Y_Step) - 10;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 610;
                            y = Y + (iCol * Y_Step) + 500;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            x = X + 630;
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
                            y = Y + (iCol * Y_Step) + 180;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 220;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 220;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
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


                            if ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ)
                            {
                                x = X + X_Pie-10;
                                y = Y + (iCol * Y_Step) + 5;
                                s = "Dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yyyy") + " al " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");
                                gfx.DrawString(s, font10, XBrushes.Black, y, x, XStringFormats.Default);
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
                            y = Y + (iCol * Y_Step) + 430;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 30;
                            //y = Y + (iCol * Y_Step) + 240;
                            rect = new XRect(10, 10, 575, 810);
                            pen = new XPen(XColors.Red, 2.0);
                            gfx.DrawRectangle(pen, rect);

                            rect = new XRect(10, 390, 575, 90);
                            pen = new XPen(XColors.Red, 1.5);
                            gfx.DrawRectangle(pen, rect);


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

        public string PrnPdfEti009(DataTable tabEti)
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
            XFont font3 = new XFont("Impact", 50, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 90, XFontStyle.Bold);
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

                            //XStringFormat form = new XStringFormat();
                            //form = XStringFormat.Center;
                            //form = 

                            //XStringFormat format = new XStringFormat();
                            //format.Alignment = XStringAlignment.Center;

                            //XRect rect = new XRect(y, x - 50, 500, 100);
                            //gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //tf.DrawString(s, font3, XBrushes.Black, rect);

                            //XFont font = new XFont("Verdana", 10);
                            XBrush brush = XBrushes.Transparent;
                            //XStringFormat format = new XStringFormat();
                            //format.Alignment = XStringAlignment.Far;
                            //format.LineAlignment = XLineAlignment.Center;

                            XRect rect = new XRect(y-20, x - 50, 600, 200);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            //gfx.DrawString(s, font3, XBrushes.Black, rect, format);

                            x = X + 90;
                            y = Y-25; // +(iCol * Y_Step) + 190;
                            tf = new XTextFormatter(gfx);
                            rect = new XRect(y, x, 585, 60);
                            if (true)
                            {
                                gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                                //tf.Alignment = XParagraphAlignment.Default; //tf.Alignment = ParagraphAlignment.Left; 
                                //tf.DrawString(s, font7, XBrushes.Black, rect, XStringFormats.TopLeft);

                                d = 0;
                                if((decimal)tabEti.Rows[i]["eti_pve"] > 0 && (decimal)tabEti.Rows[i]["eti_prv"] > 0)
                                    d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                if (d != 0)
                                {
                                    s = ((decimal)tabEti.Rows[i]["eti_pve"]).ToString("#,##0.00");
                                    gfx.DrawString(s, font8, XBrushes.Black, y + 30, x + 50, XStringFormats.Default);

                                    XPen pen = new XPen(XColors.Black, 2);
                                    gfx.DrawLine(pen, y + 20, x + 60, y + 100, x + 25);
                                }

                                s = "OFFERTA";
                                gfx.DrawString(s, font9, XBrushes.White, y + 190, x + 50, XStringFormats.Default);

                                if (d != 0)
                                {
                                    s = "Sc." + d.ToString("#,##0.00") + " %";
                                    gfx.DrawString(s, font7, XBrushes.Black, y + 450, x + 50, XStringFormats.Default);
                                }
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

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 250;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

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

                            if ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ)
                            {
                                //x = X + X_Pie - 30;
                                x = X + 280;
                                y = Y + (iCol * Y_Step) + 5;
                                s = "Dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yyyy") + " al " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");
                                //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
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

        public string PrnPdfEti010(DataTable tabEti)
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
            XFont font1 = new XFont("Courier new", 9.5, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 15, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Verdana", 17, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 55, XFontStyle.BoldItalic);
            XFont font5 = new XFont("Arial", 35, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 5;
            int i = 0;
            int X_Ini = 20;
            int X_Step = 165;
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
                            if (i == 1)
                                Console.WriteLine("aaaa");

                            double x = X;
                            double y = Y + (iCol * Y_Step);
                            int iLen = 15;
                            x = X + 13;
                            y = Y + (iCol * Y_Step) + 20;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 10;
                            y = Y + (iCol * Y_Step) + 80;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            y = Y + (iCol * Y_Step) + 130;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
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

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            x = X + 17;
                            y = Y + (iCol * Y_Step) + 5;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            XRect rect = new XRect(y, x, 250, 77);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            x = X + 107;
                            y = Y + (iCol * Y_Step) + 11;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 130;
                            y = Y + (iCol * Y_Step) + 220;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            x = X + 139;
                            y = Y + (iCol * Y_Step) + 15;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0 && d <= 500m)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                            }
                            //s = "al kg "; // +d.ToString("#####0.00");
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);


                            //x = X + 142;
                            y = Y + (iCol * Y_Step) + 95;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 120;
                            y = Y + (iCol * Y_Step) + 175;
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

        public string PrnPdfEti011(DataTable tabEti)
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

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 250;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

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

        public string PrnPdfEti012(DataTable tabEti)
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
            XFont font1 = new XFont("Courier new", 6.5, XFontStyle.Regular);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Impact", 35, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 100, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 50, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 10, XFontStyle.Regular);

            XFont font7 = new XFont("Arial", 12, XFontStyle.Bold);
            XFont font8 = new XFont("Arial", 25, XFontStyle.Bold);
            XFont font9 = new XFont("Impact", 35, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 2;
            int i = 0;
            int X_Ini = 25;
            int X_Step = 290;
            int Y_Ini = 25;
            int Y_Step = 420;
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

                            double x = X;
                            double y = Y + (iCol * Y_Step)-10;

                            XRect rect = new XRect(y, x, 400, 170);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            /******/

                            if ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ)
                            {
                                s = "OFFERTA";

                                //XTextFormatter tf = new XTextFormatter(gfx);
                                //tf.Alignment = XParagraphAlignment.Center;

                                //rect = new XRect(y, x - 50, 500, 400);
                                //gfx.DrawRectangle(XBrushes.Transparent, rect);
                                //tf.DrawString(s, font3, XBrushes.Black, rect);

                                x = X + 100;
                                y = Y + (iCol * Y_Step) - 10; // +(iCol * Y_Step) + 190;
                                tf = new XTextFormatter(gfx);
                                rect = new XRect(y, x, 400, 40);

                                gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                                //tf.Alignment = XParagraphAlignment.Default; //tf.Alignment = ParagraphAlignment.Left; 
                                //tf.DrawString(s, font7, XBrushes.Black, rect, XStringFormats.TopLeft);

                                if ((decimal)tabEti.Rows[i]["eti_pve"] > 0 && (decimal)tabEti.Rows[i]["eti_prv"] > 0)
                                    d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                if (d != 0)
                                {
                                    s = ((decimal)tabEti.Rows[i]["eti_pve"]).ToString("#,##0.00");
                                    gfx.DrawString(s, font8, XBrushes.Black, y + 20, x + 30, XStringFormats.Default);

                                    XPen pen = new XPen(XColors.Black, 3);
                                    gfx.DrawLine(pen, y + 17, x + 30, y + 100, x + 15);
                                }

                                s = "OFFERTA";
                                gfx.DrawString(s, font9, XBrushes.White, y + 150, x + 35, XStringFormats.Default);

                                if (d != 0)
                                {
                                    s = "Sc." + d.ToString("#,##0.00") + " %";
                                    gfx.DrawString(s, font7, XBrushes.Black, y + 320, x + 30, XStringFormats.Default);
                                }

                                if ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ)
                                {
                                    s = "OFFERTA dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yyyy") + " al " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");

                                    x = X + 247;
                                    y = Y + (iCol * Y_Step) + 5;
                                    gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                                }
                            }

                            /******/

                            x = X + 200;
                            y = Y + (iCol * Y_Step) + 10;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 240;
                            y = Y + (iCol * Y_Step) + 350;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y, x, frmDX);

                            x = X + 230;
                            y = Y + (iCol * Y_Step) + 5;
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
                            y = Y + (iCol * Y_Step) + 180;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 220;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            x = X + X_Pie - 35;
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
                                        gfx.DrawImage(img, y, x, 95, 25);
                                    }
                                    else
                                    {
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


                            if (i == 0)
                                Console.WriteLine("aaaa");

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 5;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 50;
                            s = "C.netto " + (string)tabEti.Rows[i]["eti_tgr"] +" "+ ((decimal)tabEti.Rows[i]["eti_pne"]).ToString("0");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 150;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 305;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 300;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

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

        public string PrnPdfEti013(DataTable tabEti)
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
            XFont font5 = new XFont("Arial", 12, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 8, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 7;
            int i = 0;
            int X_Ini = 30;
            int Y_Ini = 15;
            //int X_Step = 108;
            //int X_Step = 100;
            int X_Step = 112;
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

                            x = X + 32 + 30;
                            y = Y + (iCol * Y_Step) + 160 + 85;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            x = X + 15;
                            y = Y + (iCol * Y_Step);
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            //if (s.Length > 45)
                            //    s = s.Substring(0, 45);
                            XTextFormatter tf = new XTextFormatter(gfx);
                            XRect rect = new XRect(y, x, 170 + 00, 60);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect, XStringFormats.TopLeft);

                            x = X + 25 + 35;
                            y = Y + (iCol * Y_Step) + 100 + 70;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 60 + 5;
                            y = Y + (iCol * Y_Step) + 2;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0 && d <= 500m)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }

                            x = X + 70 + 5;
                            y = Y + (iCol * Y_Step) + 0;
                            s = (string)tabEti.Rows[i]["eti_tgr"];
                            s += " " + Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]).ToString();
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 58 + 10;
                            y = Y + (iCol * Y_Step) + 95 + 90;
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

                            x = X + 81 + 2;
                            y = Y + (iCol * Y_Step) + 2;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 81 + 20;
                            y = Y + (iCol * Y_Step) + 2;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 81 + 10;
                            y = Y + (iCol * Y_Step) + 2;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 81 + 10;
                            y = Y + (iCol * Y_Step) + 105 + 90;
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

        public string PrnPdfEti014(DataTable tabEti)
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
            XFont font3 = new XFont("Verdana", 18, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 60, XFontStyle.BoldItalic);
            XFont font5 = new XFont("Arial", 18, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 15, XFontStyle.BoldItalic);
            XFont font7 = new XFont("Verdana", 28, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 4;
            int i = 0;
            int X_Ini = 5;
            int X_Step = 214;
            int Y_Ini = 10;
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
                            double x = X;
                            double y = Y + (iCol * Y_Step);
                            int iLen = 15;
                            x = X + 10;
                            y = Y + (iCol * Y_Step) + 5;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 1;
                            y = Y + (iCol * Y_Step) + 190;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            XRect rect = new XRect(y, x, 85, 40);
                            if (((string)tabEti.Rows[i]["eti_plu"]).Trim() != "")
                            {
                                s = (string)tabEti.Rows[i]["eti_plu"];
                                if (_clsFun.Numerico(s))
                                {
                                    s = Convert.ToInt32(s).ToString();
                                }

                                //gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);
                                //XPen pen = new XPen(XColors.RoyalBlue, Math.PI);
                                gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                                tf.Alignment = XParagraphAlignment.Center; //tf.Alignment = ParagraphAlignment.Left; 
                                tf.DrawString(s, font7, XBrushes.Black, rect, XStringFormats.TopLeft);
                            }

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

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            x = X + 42;
                            y = Y + (iCol * Y_Step) + 0;
                            tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            rect = new XRect(y, x, 260, 100);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            x = X + 137;
                            y = Y + (iCol * Y_Step) + 6;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 180;
                            y = Y + (iCol * Y_Step) + 220;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            x = X + 130;
                            y = Y + (iCol * Y_Step) + 20;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            //s = "€ al kg/L " + d.ToString("#####0.00");
                            //s = "al kg "; // +d.ToString("#####0.00");
                            //gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 150;
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

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 180;
                            //y = Y + (iCol * Y_Step) + 105;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 190;
                            y = Y + (iCol * Y_Step) + 5;
                            s = "QUALITA' AL PREZZO PIU' BASSO"; // +d.ToString("#####0.00");
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

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

        public string PrnPdfEti015(DataTable tabEti, string IntRowIni)
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
            XFont font2 = new XFont("Arial", 8, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Verdana", 10, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 30, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 12, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 9, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 3;
            int iRows = 8;
            int i = 0;
            int X_Ini = 0;
            //int X_Step = 108;
            int X_Step = 108;
            int Y_Ini = 1;
            int Y_Step = 200;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            int iIniRow = Convert.ToInt16(IntRowIni);

            if (iIniRow > 1)
                iRow = iIniRow - 1;

            tabEti.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "EtiIng",
                Caption = "Ingredienti",
                MaxLength = 700,
                ReadOnly = false,
                DefaultValue = (string)""
            });

            //DataTable t = tabEti.Clone();
            //foreach (DataRow y in tabEti.Rows)
            //{
            //    s = "SELECT * FROM AnaArtIngredienti WHERE ing_art='" + y["eti_art"] + "' AND ing_ann=0 ORDER BY ing_row";
            //    DataTable tIng = _clsFun.FillTabSql("Ing", s, false, _strConSql);
            //    s = "";
            //    foreach (DataRow x in tIng.Rows)
            //        s += (string)x["ing_txt"] + _clsDef.CRLF;
            //    if (s != "")
            //    {
            //        y["EtiIng"] = s;
            //        for (int n = 0; n < (decimal)y["eti_qta"]; n++)
            //            t.ImportRow(y);
            //    }
            //}
            //tabEti = t.Copy();

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                s = "SELECT * FROM AnaArtIngredienti WHERE ing_art='" + y["eti_art"] + "' AND ing_ann=0 ORDER BY ing_row";
                DataTable tIng = _clsFun.FillTabSql("Ing", s, false, _strConSql);
                s = "";
                foreach (DataRow x in tIng.Rows)
                    s += (string)x["ing_crt"] + ";" + (string)x["ing_txt"] + "|";
                if (s != "")
                {
                    if (s.Length > 700)
                        s = s.Substring(0, 700);

                    y["EtiIng"] = s;
                    for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                        t.ImportRow(y);
                }
            }
            tabEti = t.Copy();



            while (true)
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
                            //s = (string)tabEti.Rows[i]["eti_ard"];
                            s = (string)tabEti.Rows[i]["EtiIng"];
                            //if (s.Length > 45)
                            //    s = s.Substring(0, 45);

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

                            //XTextFormatter tf = new XTextFormatter(gfx);
                            //XRect rect = new XRect(y, x, 120, 40);
                            //gfx.DrawRectangle(XBrushes.Transparent, rect);
                            ////tf.Alignment = ParagraphAlignment.Left; 
                            //tf.DrawString(s, font1, XBrushes.Black, rect, XStringFormats.TopLeft);

                            string[] a1 = s.Split('|');

                            x = X + 0;
                            y = Y + (iCol * Y_Step) + 4;

                            for (int n = 0; n < a1.Length; n++)
                            {
                                if (a1[n].Trim() != "")
                                {
                                    s = a1[n];
                                    string[] a2 = s.Split(';');

                                    s = a2[1];

                                    if (a2[0] == "G")
                                        gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                                    else
                                        gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);
                                }
                                x += 9;
                            }

                            //x = X + 56;
                            //y = Y + (iCol * Y_Step) + 95;
                            //s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 56;
                            //y = Y + (iCol * Y_Step) + 135;
                            //s = (string)tabEti.Rows[i]["eti_art"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 75;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "€";
                            //gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            //x = X + 78;
                            //y = Y + (iCol * Y_Step) + 80;
                            //s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            //gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            //x = X + 48;
                            //y = Y + (iCol * Y_Step) + 105;
                            //if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            //    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            //s = "€ al kg/L " + d.ToString("#####0.00");
                            //gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 58;
                            //y = Y + (iCol * Y_Step) + 95;
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

                            //x = X + 83;
                            //y = Y + (iCol * Y_Step) + 3;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 83;
                            //y = Y + (iCol * Y_Step) + 26;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 83;
                            //y = Y + (iCol * Y_Step) + 102;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

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

        public string PrnPdfEti016(DataTable tabEti, string IntRowIni)
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
            XFont font1 = new XFont("Courier new", 9, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 7.0, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Verdana", 10, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 30, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 12, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 9, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 8;
            int i = 0;
            int X_Ini = 7;
            //int X_Step = 108;
            int X_Step = 106;
            int Y_Ini = 1;
            //int Y_Step = 200;
            int Y_Step = 303;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = StartRow - 1;

            int iIniRow = Convert.ToInt16(IntRowIni);

            if (iIniRow > 1)
                iRow = iIniRow - 1;

            tabEti.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "EtiIng",
                Caption = "Ingredienti",
                MaxLength = 700,
                ReadOnly = false,
                DefaultValue = (string)""
            });

            DataTable t = tabEti.Clone();
            foreach (DataRow y in tabEti.Rows)
            {
                s = "SELECT * FROM AnaArtIngredienti WHERE ing_art='" + y["eti_art"] + "' AND ing_ann=0 ORDER BY ing_row";
                DataTable tIng = _clsFun.FillTabSql("Ing", s, false, _strConSql);
                s = "";
                foreach (DataRow x in tIng.Rows)
                    s += (string)x["ing_crt"] + ";" + (string)x["ing_txt"] + "|";
                if (s != "")
                {
                    if (s.Length > 700)
                        s = s.Substring(0, 700);

                    y["EtiIng"] = s;
                    for (int n = 0; n < (decimal)y["eti_qta"]; n++)
                        t.ImportRow(y);
                }
            }
            tabEti = t.Copy();

            while (true)
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
                            //s = (string)tabEti.Rows[i]["eti_ard"];
                            s = (string)tabEti.Rows[i]["EtiIng"];
                            //if (s.Length > 45)
                            //    s = s.Substring(0, 45);

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
                            if (i == 0)
                                Console.WriteLine("aaaa");

                            //gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);
                            //XRect rect = new XRect(40, 100, 250, 220);

                            //XTextFormatter tf = new XTextFormatter(gfx);
                            //XRect rect = new XRect(y, x, 180, 40);
                            //gfx.DrawRectangle(XBrushes.Transparent, rect);
                            ////tf.Alignment = ParagraphAlignment.Left; 
                            //tf.DrawString(s, font1, XBrushes.Black, rect, XStringFormats.TopLeft);

                            string[] a1 = s.Split('|');

                            x = X + 0;
                            y = Y + (iCol * Y_Step) + 4;

                            for (int n = 0; n < a1.Length; n++ )
                            {
                                if (a1[n].Trim() != "")
                                {
                                    s = a1[n];
                                    string[] a2 = s.Split(';');

                                    s = a2[1];

                                    if (a2[0] == "G")
                                        gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                                    else
                                        gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);
                                }
                                x += 9;
                            }

                            //x = X + 56;
                            //y = Y + (iCol * Y_Step) + 135;
                            //s = (string)tabEti.Rows[i]["eti_art"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 75;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = "€";
                            //gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            //x = X + 78;
                            //y = Y + (iCol * Y_Step) + 80;
                            //s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            //gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            //x = X + 48;
                            //y = Y + (iCol * Y_Step) + 105;
                            //if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                            //    Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            //    d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            //s = "€ al kg/L " + d.ToString("#####0.00");
                            //gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 58;
                            //y = Y + (iCol * Y_Step) + 95;
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

                            //x = X + 83;
                            //y = Y + (iCol * Y_Step) + 3;
                            //s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            //gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 83;
                            //y = Y + (iCol * Y_Step) + 26;
                            //s = (string)tabEti.Rows[i]["eti_fod"];
                            //if (s.Length > 5)
                            //    s = s.Substring(0, 5);
                            //if (s != "")
                            //    s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            //gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + 83;
                            //y = Y + (iCol * Y_Step) + 102;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

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

        public string PrnPdfEti017(DataTable tabEti)
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
            XFont font5 = new XFont("Arial", 12, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 9, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 3;
            int iRows = 7;
            int i = 0;
            int X_Ini = 0;
            int Y_Ini = 15;
            //int X_Step = 108;
            //int X_Step = 100;
            //int X_Step = 112;
            int X_Step = 124;
            int Y_Step = 200;

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

                            XRect rect = new XRect(y, x, 120, 40);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //tf.Alignment = ParagraphAlignment.Left; 
                            tf.DrawString(s, font3, XBrushes.Black, rect, XStringFormats.TopLeft);

                            x = X + 56;
                            y = Y + (iCol * Y_Step) + 135;
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
                            y = Y + (iCol * Y_Step) + 105;
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
                            y = Y + (iCol * Y_Step) + 5;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 81;
                            y = Y + (iCol * Y_Step) + 58;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 81;
                            y = Y + (iCol * Y_Step) + 105;
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

        public string PrnPdfEti018Vert(DataTable tabEti)    //Gatti 2x2 immagine
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

            //page.Height = 842.0;
            //page.Width = 595;
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
            XFont font1 = new XFont("Courier new", 6.5, XFontStyle.Regular);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            //XFont font3 = new XFont("Impact", 12, XFontStyle.Bold);
            XFont font3 = new XFont("Courier new", 12, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 40, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 20, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 10, XFontStyle.Regular);

            XFont font7 = new XFont("Arial", 12, XFontStyle.Bold);
            XFont font8 = new XFont("Arial", 25, XFontStyle.Bold);
            XFont font9 = new XFont("Impact", 35, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 2;
            int i = 0;
            int X_Ini = 25;
            int X_Step = 400;
            int Y_Ini = 0;
            int Y_Step = 280;
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
                    Y = Y_Ini + 30;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = 0;
                            double y = 0;

                            s = "C:\\ApProject\\Img\\" + (string)tabEti.Rows[i]["eti_art"] + ".jpg";

                            if (File.Exists(s))
                            {
                                x = X + 10;
                                y = Y + (iCol * Y_Step) + 2;

                                XImage img = XImage.FromFile(s);
                                gfx.DrawImage(img, y, x + 20, 240, 240);
                            }

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

                            x = X + 270;
                            y = Y + (iCol * Y_Step) + 150;

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            XRect rect = new XRect(y, x, 50, 350);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);


                            x = X + 255;
                            y = Y + (iCol * Y_Step) + 0;


                            XPen pen2 = new XPen(XColors.White, 2.5);
                            //gfx.DrawEllipse(pen2, 10, 0, 100, 60);
                            //gfx.DrawEllipse(XBrushes.Goldenrod, 130, 0, 100, 60);
                            gfx.DrawEllipse(pen2, XBrushes.Crimson, y, x, 130, 90);
                            //gfx.DrawEllipse(pen2, XBrushes.Goldenrod, 150, 80, 60, 60);






                            /******/

                            if (false) // && (string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ)
                            {
                                s = "OFFERTA";

                                //XTextFormatter tf = new XTextFormatter(gfx);
                                //tf.Alignment = XParagraphAlignment.Center;

                                //rect = new XRect(y, x - 50, 500, 400);
                                //gfx.DrawRectangle(XBrushes.Transparent, rect);
                                //tf.DrawString(s, font3, XBrushes.Black, rect);

                                x = X + 150;
                                y = Y + (iCol * Y_Step) - 10; // +(iCol * Y_Step) + 190;
                                tf = new XTextFormatter(gfx);
                                rect = new XRect(y, x, 400, 40);

                                gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                                //tf.Alignment = XParagraphAlignment.Default; //tf.Alignment = ParagraphAlignment.Left; 
                                //tf.DrawString(s, font7, XBrushes.Black, rect, XStringFormats.TopLeft);

                                if ((decimal)tabEti.Rows[i]["eti_pve"] > 0 && (decimal)tabEti.Rows[i]["eti_prv"] > 0)
                                    d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                if (d != 0)
                                {
                                    s = ((decimal)tabEti.Rows[i]["eti_pve"]).ToString("#,##0.00");
                                    gfx.DrawString(s, font8, XBrushes.Black, y + 20, x + 30, XStringFormats.Default);

                                    XPen pen = new XPen(XColors.Black, 3);
                                    gfx.DrawLine(pen, y + 17, x + 30, y + 100, x + 15);
                                }

                                s = "OFFERTA";
                                gfx.DrawString(s, font9, XBrushes.White, y + 150, x + 35, XStringFormats.Default);

                                if (d != 0)
                                {
                                    s = "Sc." + d.ToString("#,##0.00") + " %";
                                    gfx.DrawString(s, font7, XBrushes.Black, y + 320, x + 30, XStringFormats.Default);
                                }
                            }


                            /******/

                            x = X + 310;
                            y = Y + (iCol * Y_Step) + 5;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.White, y, x - 5, XStringFormats.Default);

                            x = X + 320;
                            y = Y + (iCol * Y_Step) + 120;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.White, y, x, frmDX);

                            x = X + 330;
                            y = Y + (iCol * Y_Step) + 150;
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

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 180;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 220;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            x = X + X_Pie - 35;
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
                                        gfx.DrawImage(img, y, x, 95, 25);
                                    }
                                    else
                                    {
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

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 5;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 40;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 305;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 300;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            
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

        public string PrnPdfEti018(DataTable tabEti, string strMsg)
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

            //page.Height = 842.0;
            //page.Width = 595;
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
            XFont font1 = new XFont("Courier new", 6.5, XFontStyle.Regular);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            //XFont font3 = new XFont("Impact", 12, XFontStyle.Bold);
            XFont font3 = new XFont("Courier new", 25, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 50, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 17, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 13, XFontStyle.Regular);

            XFont font7 = new XFont("Arial", 18, XFontStyle.Bold);
            XFont font8 = new XFont("Franklin Gothic Demi", 28, XFontStyle.Bold);
            XFont font9 = new XFont("Segoe Print", 30, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 2;
            int i = 0;
            int X_Ini = 25;
            int X_Step = 300;
            int Y_Ini = 0;
            int Y_Step = 421;
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
                    Y = Y_Ini + 30;
                    X = X_Ini + (iRow * X_Step);
                    for (int iCol = 0; iCol <= iCols - 1; iCol++)
                    {
                        if (i < tabEti.Rows.Count)
                        {
                            double x = 0;
                            double y = 0;

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

                                XImage img = XImage.FromFile(s);
                                x = X + 20;
                                y = Y + (iCol * Y_Step) + 180;
                                XRect rect1 = new XRect(y, x, 210, 210);
                                gfx.DrawImage(img, rect1);
                                //gfx.DrawImage(img, y, x + 20, 210, 210);

                                ////gfx.DrawImage(img, y, x + 20, 240, 200);
                                //Console.WriteLine("aaaaaaaa");
                                ////double x1 = (250 - img.PixelWidth * 72 / img.HorizontalResolution) / 2;
                                ////gfx.DrawImage(img, y, x + 20, dH, dW);

                                //x = X + 15;
                                //y = Y + (iCol * Y_Step) + 180;

                                //XImage img = XImage.FromFile(s);
                                //double dH = img.Size.Height;
                                //double dW = img.Size.Width;

                                //gfx.DrawImage(img, y, x + 20, 210, 210);

                            }

                            gfx.DrawLine(XPens.Black, 0, page.Height / 2, page.Width, page.Height / 2);
                            gfx.DrawLine(XPens.Black, page.Width / 2, 0, page.Width / 2, page.Height);

                            x = X - 30;
                            y = Y + (iCol * Y_Step)-30;

                            XRect rect = new XRect(y, x, page.Width / 2, 40);
                            gfx.DrawRectangle(XBrushes.Red, rect);

                            //s = "Migross conviene";
                            s = strMsg;
                            gfx.DrawString(s, font9, XBrushes.White, y + 2, x + 34, XStringFormats.Default);

                            s = "M3G";
                            gfx.DrawString(s, font8, XBrushes.White, y + 350, x + 33, XStringFormats.Default);


                            int iLen = 55;
                            s = (string)tabEti.Rows[i]["eti_ard"];
                            //if (s.Length > 45)
                            //    s = s.Substring(0, 45);

                            string s2 = "";

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

                            x = X + 25;
                            y = Y + (iCol * Y_Step) + -25;

                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            rect = new XRect(y, x, 180, 180);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            x = X + 140;
                            y = Y + (iCol * Y_Step) -20;
 
                            //XPen pen = new XPen(XColors.DarkBlue, 2.5);
                            //gfx.DrawRoundedRectangle(pen, 10, 0, 100, 60, 30, 20);
                            //gfx.DrawRoundedRectangle(XBrushes.Yellow, 130, 0, 100, 60, 30, 20);
                            //gfx.DrawRoundedRectangle(pen, XBrushes.Orange, 10, 80, 100, 60, 30, 20);
                            //gfx.DrawRoundedRectangle(pen, XBrushes.Orange, 150, 80, 60, 60, 20, 20); 
                            //#F8C000 
                            //248 192 0

                            XBrush myBrush = new XSolidBrush(new XColor { R = 248, G = 192, B = 0 });


                            XPen pen2 = new XPen(XColors.Crimson, 2.5);
                            //gfx.DrawEllipse(pen2, 10, 0, 100, 60);
                            //gfx.DrawEllipse(XBrushes.Goldenrod, 130, 0, 100, 60);
                            //gfx.DrawEllipse(pen2, XBrushes.Crimson, y, x, 130, 90);
                            gfx.DrawRoundedRectangle(XBrushes.Gold, y, x, 150, 90, 30, 20);
                            //gfx.DrawRoundedRectangle(pen2, XBrushes.Crimson, y, x, 130, 90);

                            x = X + 200;
                            y = Y + (iCol * Y_Step) - 15;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Red, y, x - 5, XStringFormats.Default);

                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 120;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Red, y, x, frmDX);

                            x = X + 224;
                            y = Y + (iCol * Y_Step) + 30;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0 && d <= 500m)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }

                            x = X + 237;
                            y = Y + (iCol * Y_Step) - 30;

                            s = "";

                            if ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ)
                                s = "Dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yyyy") + " al " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");

                            rect = new XRect(y, x, page.Width / 2, 33);
                            gfx.DrawRectangle(XBrushes.Green, rect);
                            //tf.DrawString(s, font7, XBrushes.Black, rect);

                            gfx.DrawString(s, font7, XBrushes.White, y+10, x+22, XStringFormats.Default);

                            /**/

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


                            x = X + 241;
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
                                        gfx.DrawImage(img, y, x, 95, 25);
                                    }
                                    else
                                    {
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
                            //y = Y + (iCol * Y_Step) + 40;
                            //s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            ////x = X + X_Pie;
                            ////y = Y + (iCol * Y_Step) + 305;
                            ////s = (string)tabEti.Rows[i]["eti_ean"];
                            ////gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 300;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            
                            /**/

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

        public string PrnPdfEti019(DataTable tabEti, string strMsg)
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
            XFont font3 = new XFont("Impact", 30, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 180, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 80, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 20, XFontStyle.Regular);

            //XFont font7 = new XFont("Arial", 30, XFontStyle.Bold);
            XFont font7 = new XFont("Franklin Gothic Demi", 45, XFontStyle.Bold);
            XFont font8 = new XFont("Franklin Gothic Demi", 50, XFontStyle.Bold);
            XFont font9 = new XFont("Franklin Gothic Demi", 25, XFontStyle.Bold);

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
                            double x = X;
                            double y = Y + (iCol * Y_Step);
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


                            x = X - 100;
                            y = Y + (iCol * Y_Step) - 55;

                            XRect rect = new XRect(y, x, page.Width, 80);
                            gfx.DrawRectangle(XBrushes.Red, rect);

                            //s = "Migross conviene";
                            s = strMsg;
                            gfx.DrawString(s, font7, XBrushes.White, y + 7, x + 60, XStringFormats.Default);

                            s = "M3G";
                            gfx.DrawString(s, font7, XBrushes.White, y + 500, x + 60, XStringFormats.Default);

                            s = "C:\\ApProject\\Img\\" + (string)tabEti.Rows[i]["eti_art"] + ".jpg";
                            s = (string)tabEti.Rows[i]["eti_art"] + ".jpg";

                            if (File.Exists("C:\\ApProject\\Img\\" + s))
                            {
                                string[] aa = Directory.GetFiles("C:\\ApProject\\Img\\", s);
                                if (File.Exists("C:\\ApProject\\Img\\" + s))
                                {
                                    //if (File.Exists(s))
                                    if (aa.Length > 0)
                                    {
                                        s = aa[0];

                                        if ((string)tabEti.Rows[i]["eti_art"] == "0005933")
                                            Console.WriteLine("aaaa");
                                        if (i == 0)
                                            Console.WriteLine("aaaa");

                                        XImage img = XImage.FromFile(s);
                                        x = X + 45;
                                        y = Y + (iCol * Y_Step) + 30;
                                        XRect rect1 = new XRect(y, x, 420, 420);
                                        gfx.DrawImage(img, rect1);
                                    }
                                }
                            }

                            x = X - 20;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            rect = new XRect(y, x, 500, 400);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(sDes, font3, XBrushes.Black, rect);

                            x = X + 410;
                            y = Y + (iCol * Y_Step) - 35;

                            if ((string)tabEti.Rows[i]["eti_bcl"] != "")
                            {
                                s = "Calibro: " + (string)tabEti.Rows[i]["eti_bcl"];
                                gfx.DrawString(s, font9, XBrushes.Black, y, x, XStringFormats.Default);
                            }
                            if ((string)tabEti.Rows[i]["eti_bct"] != "")
                            {
                                s = "Categoria: " + (string)tabEti.Rows[i]["eti_bct"];
                                gfx.DrawString(s, font9, XBrushes.Black, y, x+25, XStringFormats.Default);
                            }
                            if ((string)tabEti.Rows[i]["eti_bor"] != "")
                            {
                                s = "Origine: " + (string)tabEti.Rows[i]["eti_bor"];
                                gfx.DrawString(s, font9, XBrushes.Black, y, x+50, XStringFormats.Default);
                            }

                            x = X + 470;
                            y = Y + (iCol * Y_Step) - 35;
                            gfx.DrawRoundedRectangle(XBrushes.Gold, y, x, 550, 240, 30, 20);

                            x = X + 680;
                            y = Y + (iCol * Y_Step) + 500;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            x = X + 630;
                            y = Y + (iCol * Y_Step) - 20;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 675;
                            y = Y + (iCol * Y_Step) + 360;
                            d = 0;
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG" &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 180;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 220;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 680;
                            y = Y + (iCol * Y_Step) - 55;

                            rect = new XRect(y, x, page.Width, 65);
                            gfx.DrawRectangle(XBrushes.Green, rect);
                            //tf.DrawString(s, font7, XBrushes.Black, rect);

                            x = X + X_Pie -45;
                            y = Y + (iCol * Y_Step) + 440;
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
                                        gfx.DrawImage(img, y, x, 90, 35);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 5;

                            s = "";
                            if ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ)
                                s = "Dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yyyy") + " al " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");
                            gfx.DrawString(s, font9, XBrushes.White, y-50, x-15, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 5;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 58;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            //x = X + X_Pie;
                            //y = Y + (iCol * Y_Step) + 305;
                            //s = (string)tabEti.Rows[i]["eti_ean"];
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 470;
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

        public string PrnPdfEti020(DataTable tabEti)
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
            XFont font2 = new XFont("Arial", 8, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Verdana", 10, XFontStyle.Bold);
            XFont font4 = new XFont("Impact", 30, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 12, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 9, XFontStyle.Regular);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 3;
            int iRows = 8;
            int i = 0;
            int X_Ini = 0;
            //int X_Step = 108;
            int X_Step = 108;       //100
            int Y_Ini = 15;
            int Y_Step = 200;

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

                            XRect rect = new XRect(y, x, 120, 40);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            //tf.Alignment = ParagraphAlignment.Left; 
                            tf.DrawString(s, font3, XBrushes.Black, rect, XStringFormats.TopLeft);

                            x = X + 56;
                            y = Y + (iCol * Y_Step) + 95;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 56;
                            y = Y + (iCol * Y_Step) + 135;
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
                            y = Y + (iCol * Y_Step) + 105;
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

                            x = X + 83;
                            y = Y + (iCol * Y_Step) + 3;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 83;
                            y = Y + (iCol * Y_Step) + 26;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 83;
                            y = Y + (iCol * Y_Step) + 102;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

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

        public string PrnPdfEti021(DataTable tabEti)
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
            XFont font1 = new XFont("Courier new", 6.5, XFontStyle.Regular);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Impact", 35, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 100, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 30, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 10, XFontStyle.Regular);

            XFont font7 = new XFont("Arial", 12, XFontStyle.Bold);
            XFont font8 = new XFont("Arial", 25, XFontStyle.Bold);
            XFont font9 = new XFont("Impact", 20, XFontStyle.Regular);
            XFont font10 = new XFont("Arial", 80, XFontStyle.Bold);
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

                            double x = X;
                            double y = Y + (iCol * Y_Step) - 10;

                            XRect rect = new XRect(y, x, 400, 170);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            x = X + 110;
                            y = Y + (iCol * Y_Step) + 20;

                            if ((string)tabEti.Rows[i]["eti_bor"] != "")
                            {
                                s = "Origine: " + (string)tabEti.Rows[i]["eti_bor"];
                                gfx.DrawString(s, font9, XBrushes.Black, y, x + 0, XStringFormats.Default);
                            }
                            if ((string)tabEti.Rows[i]["eti_bct"] != "")
                            {
                                s = "Categoria: " + (string)tabEti.Rows[i]["eti_bct"];
                                gfx.DrawString(s, font9, XBrushes.Black, y, x + 25, XStringFormats.Default);
                            }
                            if ((string)tabEti.Rows[i]["eti_bcl"] != "")
                            {
                                s = "Calibro: " + (string)tabEti.Rows[i]["eti_bcl"];
                                gfx.DrawString(s, font9, XBrushes.Black, y, x + 50, XStringFormats.Default);
                            }

                            /**** PLU INIZIO ****/
                            if (((string)tabEti.Rows[i]["eti_plu"]).Trim() != "")
                            {

                                x = X + 82;
                                y = Y + (iCol * Y_Step) + 280;
                                s = "TASTO BILANCIA";
                                gfx.DrawString(s, font11, XBrushes.Black, y, x - 5, XStringFormats.Default);

                                x = X + 80;
                                y = Y + (iCol * Y_Step) + 270;
                                tf = new XTextFormatter(gfx);
                                rect = new XRect(y, x, 120, 90);

                                XPen pen = new XPen(XColors.Black, 3);
                                gfx.DrawLine(pen, y, x, y + 120, x + 90);


                                s = ((string)tabEti.Rows[i]["eti_plu"]).Substring(1);
                                if (_clsFun.Numerico(s))
                                {
                                    s = Convert.ToInt32(s).ToString();
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

                            x = X + 250;
                            y = Y + (iCol * Y_Step) + 10;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 270;
                            y = Y + (iCol * Y_Step) + 270;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y, x, frmDX);

                            x = X + 250;
                            y = Y + (iCol * Y_Step) + 290;
                            if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
                                s = "AL KG";
                            else
                            s = "AL NR";
                            gfx.DrawString(s, font12, XBrushes.Black, y, x - 5, XStringFormats.Default);


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
                            y = Y + (iCol * Y_Step) + 320;
                            s = DateTime.Today.ToString("dd.MM.yy");
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

        public string PrnPdfEti022(DataTable tabEti)
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
            XFont font3 = new XFont("Verdana", 18, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 70, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 25, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 4;
            int i = 0;
            int X_Ini = -10;
            int X_Step = 210;
            int Y_Ini = 15;
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

                            x = X + 42;
                            y = Y + (iCol * Y_Step) + 0;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            XRect rect = new XRect(y, x, 260, 100);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            x = X + 157;
                            y = Y + (iCol * Y_Step) + 6;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 180;
                            y = Y + (iCol * Y_Step) + 220;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);


                            x = X + 178;
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
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 50;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);



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
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }



                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 200;
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

        public string PrnPdfEti023(DataTable tabEti)
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
            XFont font1 = new XFont("Courier new", 8, XFontStyle.Regular);
            XFont font2 = new XFont("Courier new", 9.5, XFontStyle.Bold);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            XFont font3 = new XFont("Verdana", 18, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 60, XFontStyle.BoldItalic);
            XFont font5 = new XFont("Arial", 18, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 11, XFontStyle.Bold);
            XFont font7 = new XFont("Verdana", 28, XFontStyle.Bold);

            XFont font8 = new XFont("Arial", 20, XFontStyle.Bold);
            XFont font9 = new XFont("Impact", 25, XFontStyle.Bold);
            XFont font10 = new XFont("Arial", 12, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 4;
            int i = 0;
            int X_Ini = 0;
            int X_Step = 218;
            int Y_Ini = 10;
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
                            double x = X;
                            double y = Y + (iCol * Y_Step);
                            int iLen = 15;
                            //x = X + 10;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = DateTime.Today.ToString("dd.MM.yy");
                            //gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);


                            // NO PLU

                            //x = X + 48;
                            //y = Y + (iCol * Y_Step) + 195;
                            XTextFormatter tf = new XTextFormatter(gfx);
                            XRect rect = new XRect(y, x, 85, 40);
                            //if (((string)tabEti.Rows[i]["eti_plu"]).Trim() != "")
                            //{
                            //    s = (string)tabEti.Rows[i]["eti_plu"];
                            //    if (_clsFun.Numerico(s))
                            //    {
                            //        s = Convert.ToInt16(s).ToString();
                            //    }

                            //    //gfx.DrawString(s, font5, XBrushes.Black, y, x, XStringFormats.Default);
                            //    //XPen pen = new XPen(XColors.RoyalBlue, Math.PI);
                            //    gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                            //    tf.Alignment = XParagraphAlignment.Center; //tf.Alignment = ParagraphAlignment.Left; 
                            //    tf.DrawString(s, font7, XBrushes.Black, rect, XStringFormats.TopLeft);
                            //}

                            //x = X + 20;
                            //y = Y + (iCol * Y_Step) + 5;
                            //s = (string)tabEti.Rows[i]["eti_art"];
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

                            x = X + 20;
                            y = Y + (iCol * Y_Step) + 15;
                            tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;
                            //XRect rect = new XRect(y, x, 300, 100);
                            rect = new XRect(y, x, 260, 100);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            if (i == 1)
                                Console.WriteLine("aaaa");

                            decimal dPrv = 0;
                            if (!DBNull.Value.Equals(tabEti.Rows[i]["eti_prv"]))
                                dPrv = (decimal)tabEti.Rows[i]["eti_prv"];

                            if ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ)
                            {
                                //s = "OFFERTA";

                                x = X + 80;
                                y = Y + (iCol * Y_Step) - 10; // +(iCol * Y_Step) + 190;
                                tf = new XTextFormatter(gfx);
                                rect = new XRect(y, x, 290, 30);

                                //gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                                gfx.DrawRectangle(new SolidBrush(Color.White), rect);

                                if ((decimal)tabEti.Rows[i]["eti_pve"] > 0 && (decimal)tabEti.Rows[i]["eti_prv"] > 0)
                                    d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                if (d != 0)
                                {
                                    s = ((decimal)tabEti.Rows[i]["eti_pve"]).ToString("#,##0.00");
                                    gfx.DrawString(s, font8, XBrushes.Black, y + 20, x + 25, XStringFormats.Default);

                                    XPen pen = new XPen(XColors.Black, 3);
                                    gfx.DrawLine(pen, y + 17, x + 25, y + 60, x + 15);
                                }

                                s = "OFFERTA";
                                gfx.DrawString(s, font9, XBrushes.Black, y + 95, x + 25, XStringFormats.Default);

                                if (d != 0)
                                {
                                    s = "Sc." + d.ToString("#,##0.00") + " %";
                                    gfx.DrawString(s, font10, XBrushes.Black, y + 220, x + 25, XStringFormats.Default);
                                }

                                x = X + 177;
                                y = Y + (iCol * Y_Step) + 5;
                                s = "Validità dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yy") + " al " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yy");
                                gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);
                            }
                            else if ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFASCO)
                            {



                                x = X + 80;
                                y = Y + (iCol * Y_Step) - 10; // +(iCol * Y_Step) + 190;
                                tf = new XTextFormatter(gfx);
                                rect = new XRect(y, x, 290, 30);

                                //gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                                gfx.DrawRectangle(new SolidBrush(Color.White), rect);

                                //if ((decimal)tabEti.Rows[i]["eti_pve"] > 0 && (decimal)tabEti.Rows[i]["eti_prv"] > 0)
                                //    d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                //Prezzo scontato
                                if ((decimal)tabEti.Rows[i]["eti_pve"] != 0)
                                {
                                    d = (decimal)tabEti.Rows[i]["eti_prv"];
                                    //d = _clsFun.MenoPer(d, (decimal)tabEti.Rows[i]["eti_pve"]);

                                    s = d.ToString("#,##0.00");
                                    gfx.DrawString(s, font8, XBrushes.Black, y + 20, x + 25, XStringFormats.Default);

                                    XPen pen = new XPen(XColors.Black, 3);
                                    gfx.DrawLine(pen, y + 17, x + 25, y + 60, x + 15);
                                }

                                s = "OFFERTA";
                                gfx.DrawString(s, font9, XBrushes.White, y + 95, x + 25, XStringFormats.Default);

                                if ((decimal)tabEti.Rows[i]["eti_pve"] != 0)
                                {
                                    d = (decimal)tabEti.Rows[i]["eti_pve"];

                                    s = "Sc." + d.ToString("#,##0.00") + " %";
                                    gfx.DrawString(s, font10, XBrushes.Black, y + 220, x + 25, XStringFormats.Default);
                                }

                                x = X + 177;
                                y = Y + (iCol * Y_Step) + 5;
                                s = "Validità dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yy") + " al " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yy");
                                gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);


                                dPrv = (decimal)tabEti.Rows[i]["eti_prv"];
                                dPrv = _clsFun.MenoPer(dPrv, (decimal)tabEti.Rows[i]["eti_pve"]);


                            }

                            x = X + 160;
                            y = Y + (iCol * Y_Step) + 15;
                            s = "€";
                            gfx.DrawString(s, font7, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            x = X + 170;
                            y = Y + (iCol * Y_Step) + 220;
                            s = dPrv.ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            x = X + 185;
                            y = Y + (iCol * Y_Step) + 5;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 185;
                            y = Y + (iCol * Y_Step) + 65;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            x = X + 185;
                            d = 0;
                            y = Y + (iCol * Y_Step) + 120;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                            {
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                                //s = "€ al kg/L " + d.ToString("#####0.00");

                                if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) != d)
                                {
                                    if (d > 0 && d <= 500m)
                                    {
                                        s = "al Kg/Lt € " + d.ToString("#####0.00");
                                        gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                                    }
                                }
                            }
                            x = X + 168;
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

        public string OldPrnPdfEti024(DataTable tabEti)
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
            int iRows = 9;
            int i = 0;
            int X_Ini = -5;
            //int X_Step = 108;
            int X_Step = 92;       //100
            int Y_Ini = 30;
            int Y_Step = 196;

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
                            double x = X+3;
                            double y = Y + (iCol * Y_Step -10);
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

                            /*** Prezzo in chiaro ***/
                            x = X + 30;
                            y = Y + (iCol * Y_Step) + 0; ;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)

                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0 && d <= 500m)
                            {
                                s = d.ToString("#####0.00") + " € al kg/L";
                                gfx.DrawString(s.Trim(), font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }

                            /*** codice articolo ***/
                            x = X + 39;
                            y = Y + (iCol * Y_Step) - 10;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** Simbolo euro ***/
                            x = X + 51;
                            y = Y + (iCol * Y_Step) + 50;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            /*** Prezzo ***/
                            x = X + 58;
                            y = Y + (iCol * Y_Step) + 155;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            /*** Contenuto ***/
                            x = X + 48;
                            y = Y + (iCol * Y_Step) - 10;
                            s = "Cont. " + (string)tabEti.Rows[i]["eti_tgr"] + " " + ((decimal)tabEti.Rows[i]["eti_pne"]).ToString("#####0.00");
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** Grammatura ***/
                            x = X + 54;
                            y = Y + (iCol * Y_Step) + 50;
                            s = "al ";
                            if((string)tabEti.Rows[i]["eti_umi"] == "NR")
                                s += "PZ";
                            else
                                s += (string)tabEti.Rows[i]["eti_umi"];

                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** Data ***/
                            x = X + 56;
                            y = Y + (iCol * Y_Step) - 10;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** PxC ***/
                            x = X + 77;
                            y = Y + (iCol * Y_Step) - 20;
                            s = "PxC " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** Fornitore ***/
                            x = X + 77;
                            y = Y + (iCol * Y_Step) + 16;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** Barcode su arf di DADO ***/
                            x = X + 57;
                            y = Y + (iCol * Y_Step) -20;
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

                            /*** Barcode ***/
                            x = X + 57;
                            y = Y + (iCol * Y_Step) + 90;
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
                                        gfx.DrawImage(img, y, x, 70, 14);
                                    }
                                    else
                                    {
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 14);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            /*** Barcode in cifre ***/
                            x = X + 77;
                            y = Y + (iCol * Y_Step) + 94;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

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

        public string PrnPdfEti024(DataTable tabEti)
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
            int X_Ini = +10;
            int X_Step = 104; // 108;       //92
            int Y_Ini = 30;
            int Y_Step = 196;       //196

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

                            /*** Prezzo in chiaro ***/
                            x = X + 30;
                            y = Y + (iCol * Y_Step) + 0; ;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)

                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0 && d <= 500m)
                            {
                                s = d.ToString("#####0.00") + " € al kg/L";
                                gfx.DrawString(s.Trim(), font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }

                            /*** codice articolo ***/
                            x = X + 39;
                            y = Y + (iCol * Y_Step) - 10;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** Simbolo euro ***/
                            x = X + 51;
                            y = Y + (iCol * Y_Step) + 50;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);

                            /*** Prezzo ***/
                            x = X + 58;
                            y = Y + (iCol * Y_Step) + 155;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y + 10, x, frmDX);

                            /*** Contenuto ***/
                            x = X + 48;
                            y = Y + (iCol * Y_Step) - 10;
                            s = "Cont. " + (string)tabEti.Rows[i]["eti_tgr"] + " " + ((decimal)tabEti.Rows[i]["eti_pne"]).ToString("#####0.00");
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** Grammatura ***/
                            x = X + 54;
                            y = Y + (iCol * Y_Step) + 50;
                            s = "al ";
                            if ((string)tabEti.Rows[i]["eti_umi"] == "NR")
                                s += "PZ";
                            else
                                s += (string)tabEti.Rows[i]["eti_umi"];

                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** Data ***/
                            x = X + 56;
                            y = Y + (iCol * Y_Step) - 10;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** PxC ***/
                            x = X + 77;
                            y = Y + (iCol * Y_Step) - 20;
                            s = "PxC " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** Fornitore ***/
                            x = X + 77;
                            y = Y + (iCol * Y_Step) + 16;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);

                            /*** Barcode su arf di DADO ***/
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

                            /*** Barcode ***/
                            x = X + 57;
                            y = Y + (iCol * Y_Step) + 90;
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
                                        gfx.DrawImage(img, y, x, 70, 14);
                                    }
                                    else
                                    {
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 70, 14);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }

                            /*** Barcode in cifre ***/
                            x = X + 77;
                            y = Y + (iCol * Y_Step) + 94;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font2, XBrushes.Black, y, x, XStringFormats.Default);

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

        public string PrnPdfEti025(DataTable tabEti)
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
                                    s = Convert.ToInt32(s).ToString();
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
                            y = Y + (iCol * Y_Step-10);
                            s = "PxC " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            /*** PxC FINE ***/

                            /**** ARTICOLO INIZIO 2 ****/
                            x = X + 120;
                            y = Y + (iCol * Y_Step) + 300;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
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
                                s += "Etto";    // s += (string)tabEti.Rows[i]["eti_umi"];
                            gfx.DrawString(s, font7, XBrushes.Black, y, x, frmDX);
                            /*** UMI FINE 1 ***/

                            /**** PREZZO INIZIO 1 ****/
                            x = X + 180;
                            y = Y + (iCol * Y_Step) + 250;
                            d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]);
                            if ((string)tabEti.Rows[i]["eti_umi"] != "NR")
                                d = d / 10;
                            s = d.ToString("######0.00");
                            gfx.DrawString(s, font12, XBrushes.Black, y + 10, x, frmDX);
                            /**** PREZZO FINE 1 ****/

                            /*** PLU INIZIO 2 ***/
                            x = X + 150;
                            y = Y + (iCol * Y_Step) + 340;
                            s = "PLU " + (string)tabEti.Rows[i]["eti_plu"];
                            gfx.DrawString(s, font7, XBrushes.Black, y + 10, x, frmDX);
                            /*** PLU FINE 2 ***/

                            /**** EURO INIZIO 2 ****/
                            x = X + 180;
                            y = Y + (iCol * Y_Step) + 370;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);
                            /**** EURO FINE 2 ****/

                            /*** UMI INIZIO 2 ***/
                            x = X + 190;
                            y = Y + (iCol * Y_Step) + 387;
                            s = "al ";
                            if ((string)tabEti.Rows[i]["eti_umi"] == "NR")
                                s += "PZ";
                            else
                                s += "Etto";        //s += (string)tabEti.Rows[i]["eti_umi"];
                            gfx.DrawString(s, font7, XBrushes.Black, y + 10, x, frmDX);
                            /*** UMI FINE 2 ***/

                            /**** PREZZO INIZIO 2 ****/
                            x = X + 185;
                            y = Y + (iCol * Y_Step) + 550;

                            d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]);

                            if ((string)tabEti.Rows[i]["eti_umi"] != "NR")
                                d = d / 10;

                            s = d.ToString("######0.00");
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

                            /*** Barcode su arf di DADO ***/
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





                            /**** DATA INIZIO 1 ****/
                            x = X + 205;
                            y = Y + (iCol * Y_Step);
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /**** DATA FINE 1 ****/

                            /**** DATA INIZIO 2 ****/
                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 300;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
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
                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 100;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /**** EAN IN CIFRE FINE ****/

                            /**** EAN IN CIFRE INIZIO ****/
                            x = X + 205;
                            y = Y + (iCol * Y_Step) + 420;
                            s = ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
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

        public string PrnPdfEti026(DataTable tabEti)
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
            XFont font1 = new XFont("Courier new", 6.5, XFontStyle.Regular);
            XFont font2 = new XFont("Courier new", 10, XFontStyle.Regular);
            XFont font3 = new XFont("Calibri", 25, XFontStyle.Bold);
            //XFont font3 = new XFont("Calibri", 22, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 75, XFontStyle.Bold);
            XFont font5 = new XFont("Arial", 50, XFontStyle.Bold);
            XFont font6 = new XFont("Arial", 10, XFontStyle.Regular);

            XFont font7 = new XFont("Arial", 18, XFontStyle.Bold);
            XFont font8 = new XFont("Arial", 25, XFontStyle.Bold);
            XFont font9 = new XFont("Impact", 35, XFontStyle.Bold);

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            int iCols = 2;
            int iRows = 2;
            int i = 0;
            int X_Ini = 25;
            int X_Step = 305;
            int Y_Ini = 25;
            int Y_Step = 420;
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

                            /*** DESCRIZIONE INIZIO ***/
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


                            XTextFormatter tf = new XTextFormatter(gfx);
                            tf.Alignment = XParagraphAlignment.Center;

                            double x = X + 85;
                            double y = Y + (iCol * Y_Step) - 20;

                            XRect rect = new XRect(y, x, 450, 170);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);
                            /*** DESCRIZIONE FINE ***/

                            /*** OFFERTA INIZIO ***/

                            if ((string)tabEti.Rows[i]["eti_off"] != "" && (string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAPRZ)
                            {
                                x = X + 120;
                                y = Y + (iCol * Y_Step) - 10; // +(iCol * Y_Step) + 190;
                                tf = new XTextFormatter(gfx);
                                rect = new XRect(y, x, 400, 30);

                                gfx.DrawRectangle(new SolidBrush(Color.LightGray), rect);
                                //tf.Alignment = XParagraphAlignment.Default; //tf.Alignment = ParagraphAlignment.Left; 
                                //tf.DrawString(s, font7, XBrushes.Black, rect, XStringFormats.TopLeft);

                                if ((decimal)tabEti.Rows[i]["eti_pve"] > 0 && (decimal)tabEti.Rows[i]["eti_prv"] > 0)
                                    d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                if (d != 0)
                                {
                                    s = ((decimal)tabEti.Rows[i]["eti_pve"]).ToString("#,##0.00");
                                    gfx.DrawString(s, font8, XBrushes.Black, y + 20, x + 25, XStringFormats.Default);

                                    XPen pen = new XPen(XColors.Black, 3);
                                    gfx.DrawLine(pen, y + 10, x + 25, y + 80, x + 10);
                                }

                                s = "OFFERTA";
                                gfx.DrawString(s, font9, XBrushes.White, y + 150, x + 29, XStringFormats.Default);

                                if (d != 0)
                                {
                                    s = "Sc." + d.ToString("#,##0.00") + " %";
                                    gfx.DrawString(s, font7, XBrushes.Black, y + 300, x + 25, XStringFormats.Default);
                                }
                            }
                            /*** OFFERTA FINE ***/

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            /*** DICITURA FIDELITY INIZIO ***/
                            if (((string)tabEti.Rows[i]["eti_cam"]).Trim() != "")
                            {
                                x = X + 150;
                                y = Y + (iCol * Y_Step) - 280; // +(iCol * Y_Step) + 190;
                                s = "Prezzo Fidelity";
                                gfx.DrawString(s, font7, XBrushes.Black, y + 300, x + 28, XStringFormats.Default);
                            }

                            /*** DICITURA FIDELITY FINE ***/

                            /*** EURO INIZIO ***/
                            x = X + 230;
                            y = Y + (iCol * Y_Step) +100;
                            s = "€";
                            gfx.DrawString(s, font5, XBrushes.Black, y, x - 5, XStringFormats.Default);
                            /*** EURO FINE ***/

                            /*** PREZZO INIZIO ***/
                            x = X + 240;
                            y = Y + (iCol * Y_Step) + 350;
                            s = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y, x, frmDX);
                            /*** PREZZO FINE ***/

                            /*** P CHIARO INIZIO ***/
                            x = X + 240;
                            y = Y + (iCol * Y_Step) + 5;
                            if (Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) > 0 &&
                                Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]) > 0)
                                d = Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]) / (Convert.ToDecimal(tabEti.Rows[i]["eti_pne"]) / Convert.ToDecimal(tabEti.Rows[i]["eti_tgv"]));

                            if (d > 0 && d <= 500m)
                            {
                                s = "€ al kg/L " + d.ToString("#####0.00");
                                gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            }
                            /*** P CHIARO FINE ***/

                            /*** ARTICOLO INIZIO ***/
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 180;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /*** ARTICOLO FINE ***/

                            /*** ARTICOLO FOR INIZIO ***/
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 220;
                            s = (string)tabEti.Rows[i]["eti_fod"];
                            if (s.Length > 5)
                                s = s.Substring(0, 5);
                            if (s != "")
                                s += "/" + ((string)tabEti.Rows[i]["eti_arf"]).Trim();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /*** ARTICOLO FOR FINE ***/

                            /*** BARCODE INIZIO ***/
                            x = X + X_Pie - 15;
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
                                        gfx.DrawImage(img, y, x, 95, 20);
                                    }
                                    else
                                    {
                                        Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                        MemoryStream ms = new MemoryStream();
                                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        gfx.DrawImage(img, y, x, 95, 20);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _clsFun.ErrorLog(ex.Message, (string)tabEti.Rows[i]["eti_ean"]);
                                }
                            }
                            /*** BARCODE FINE ***/

                            /*** PEZZI INIZIO ***/
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 5;
                            s = "Pz. " + ((decimal)tabEti.Rows[i]["eti_pxc"]).ToString();
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /*** PEZZI FINE ***/

                            /*** DATA INIZIO ***/
                            x = X + X_Pie;
                            y = Y + (iCol * Y_Step) + 40;
                            s = DateTime.Today.ToString("dd.MM.yy");
                            gfx.DrawString(s, font1, XBrushes.Black, y, x, XStringFormats.Default);
                            /*** DATA FINE ***/

                            /*** EAN INIZIO ***/
                            x = X + X_Pie-20;
                            y = Y + (iCol * Y_Step) + 300;
                            s = (string)tabEti.Rows[i]["eti_ean"];
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            /*** EAN FINE ***/

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

        public string PrnPdfEti027(DataTable tabEti)
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
            XFont font3 = new XFont("Arial", 50, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 120, XFontStyle.Bold);
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

                            if (i == 0)
                                Console.WriteLine("aaaa");

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
                            x = X + 180;
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

                            XRect rect = new XRect(y, x, 500, 400);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            /*** DESCRIZIONE fine ***/

                            ///*** EURO inizio ***/
                            //x = X + 340;
                            //y = Y + (iCol * Y_Step) + 250;
                            //s = "€";
                            //gfx.DrawString(s, font4, XBrushes.Black, y, x - 5, XStringFormats.Default);
                            ///*** EURO fine ***/

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
                            gfx.DrawLine(pen, y+50, x, y + 160, x + -35);

                            d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;
                            s = "- " + d.ToString("#,##0.00") + " %";
                            gfx.DrawString(s, font3, XBrushes.Black, y + 210, x, XStringFormats.Default);

                            /*** OFFERTA fine ***/

                            /*** ARTICOLO inizio ***/
                            x = X + 470;
                            y = Y + (iCol * Y_Step) -10;
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
                            else if (Convert.ToString(tabEti.Rows[i]["eti_umi"]) == "KG")
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

        public string PrnPdfEti028(DataTable tabEti)
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
            //XFont font1 = new XFont("Courier new", 6.5, XFontStyle.Regular);
            XFont font2 = new XFont("Arial", 18, XFontStyle.Regular);
            //XFont font3 = new XFont("Impact", 11, XFontStyle.Bold);Verdana
            //XFont font3 = new XFont("Impact", 70, XFontStyle.Bold);
            XFont font3 = new XFont("Arial", 25, XFontStyle.Bold);
            XFont font4 = new XFont("Arial", 80, XFontStyle.Bold);
            //XFont font5 = new XFont("Arial", 80, XFontStyle.Regular);
            XFont font6 = new XFont("Arial", 18, XFontStyle.Regular);

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

                            string sTipOff = (string)tabEti.Rows[i]["eti_oft"];

                            /*** DATA OFFERTA inizio ***/
                            if ((string)tabEti.Rows[i]["eti_off"] != "" && (sTipOff == _clsDef.OFAPRZ || sTipOff == _clsDef.OFASCO || sTipOff == _clsDef.OFAMXN))
                            {
                                x = X + 170;
                                y = Y + (iCol * Y_Step) + 380;
                                //s = "Dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd/MM/yyyy") + " al " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd/MM/yyyy");
                                s = "Dal " + ((DateTime)tabEti.Rows[i]["eti_odi"]).ToString("dd");
                                s += " al " + ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("dd") + " ";
                                s += ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("MMMM") + " ";
                                s += ((DateTime)tabEti.Rows[i]["eti_odf"]).ToString("yyyy");
                                gfx.DrawString(s, font6, XBrushes.Black, y - 50, x - 15, XStringFormats.Default);
                            }
                            /*** DATA OFFERTA fine ***/

                            /*** DESCRIZIONE inizio ***/
                            x = X + 180;
                            y = Y + (iCol * Y_Step) - 20;
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

                            XRect rect = new XRect(y, x, 250, 400);
                            gfx.DrawRectangle(XBrushes.Transparent, rect);
                            tf.DrawString(s, font3, XBrushes.Black, rect);

                            /*** DESCRIZIONE fine ***/

                            ///*** EURO inizio ***/
                            //x = X + 340;
                            //y = Y + (iCol * Y_Step) + 250;
                            //s = "€";
                            //gfx.DrawString(s, font4, XBrushes.Black, y, x - 5, XStringFormats.Default);
                            ///*** EURO fine ***/

                            /*** PREZZO inizio ***/

                            d = (decimal)tabEti.Rows[i]["eti_prv"];

                            //Calcolare il prezzo scontato
                            if (sTipOff == _clsDef.OFASCO)
                                d = _clsFun.MenoPer(d, (decimal)tabEti.Rows[i]["eti_pve"]);

                            x = X + 280;
                            y = Y + (iCol * Y_Step) + 550;
                            //s = "€ " + Convert.ToDecimal(tabEti.Rows[i]["eti_prv"]).ToString("######0.00");
                            s = "€ " + d.ToString("######0.00");
                            gfx.DrawString(s, font4, XBrushes.Black, y, x, frmDX);
                            /*** PREZZO fine ***/

                            if (i == 0)
                                Console.WriteLine("aaaa");

                            if ((string)tabEti.Rows[i]["eti_arf"] == "0000801")
                                Console.WriteLine("aaaa");

                            /*** OFFERTA inizio ***/
                            if ((string)tabEti.Rows[i]["eti_oft"] == _clsDef.OFAMXN && (decimal)tabEti.Rows[i]["eti_pve"] > 0 && (decimal)tabEti.Rows[i]["eti_prv"] > 0)
                            {
                                x = X + 290;
                                y = Y + (iCol * Y_Step) - 20;

                                s = "OFFERTA MxN: Prendi " + ((decimal)tabEti.Rows[i]["eti_xem"]).ToString() + " Paghi " + ((decimal)tabEti.Rows[i]["eti_xen"]).ToString();
                                gfx.DrawString(s, font3, XBrushes.Black, y, x, XStringFormats.Default);

                            }
                            else if ((decimal)tabEti.Rows[i]["eti_pve"] > 0 && (decimal)tabEti.Rows[i]["eti_prv"] > 0 && (decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"] != 0)
                            {
                                x = X + 290;
                                y = Y + (iCol * Y_Step) - 20;

                                d = (decimal)tabEti.Rows[i]["eti_pve"];
                                if (sTipOff == _clsDef.OFASCO)
                                    d = (decimal)tabEti.Rows[i]["eti_prv"];

                                //s = "€  " + ((decimal)tabEti.Rows[i]["eti_pve"]).ToString("#,##0.00");
                                s = "€  " + d.ToString("#,##0.00");
                                gfx.DrawString(s, font3, XBrushes.Black, y, x, XStringFormats.Default);

                                XPen pen = new XPen(XColors.Black, 3);
                                gfx.DrawLine(pen, y + 20, x, y + 90, x + -15);

                                d = ((decimal)tabEti.Rows[i]["eti_pve"] - (decimal)tabEti.Rows[i]["eti_prv"]) / (decimal)tabEti.Rows[i]["eti_pve"] * 100;

                                if (sTipOff == _clsDef.OFASCO)
                                    d = (decimal)tabEti.Rows[i]["eti_pve"];

                                s = "- " + d.ToString("#,##0.00") + " %";
                                gfx.DrawString(s, font3, XBrushes.Black, y + 150, x, XStringFormats.Default);
                            }
                            /*** OFFERTA fine ***/

                            /*** ARTICOLO inizio ***/
                            x = X + 320;
                            y = Y + (iCol * Y_Step) - 10;
                            s = (string)tabEti.Rows[i]["eti_art"];
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            /*** ARTICOLO fine ***/

                            /*** ARTICOLO FORNITOE inizio ***/
                            x = X + 320;
                            y = Y + (iCol * Y_Step) + 80;
                            s = (string)tabEti.Rows[i]["eti_arf"];
                            gfx.DrawString(s, font6, XBrushes.Black, y, x, XStringFormats.Default);
                            /*** ARTICOLO FORNITORE fine ***/

                            /*** BARCODE inizio ***/
                            x = X + 300;
                            y = Y + (iCol * Y_Step) + 180;
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
                            /*** BARCODE fine ***/

                            /*** GRAMMATURA inizio ***/
                            x = X + 320;
                            y = Y + (iCol * Y_Step) + 480;
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

        public string PrnPdfEti029(DataTable tabEti)
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
            //int X_Step = 100;
            int X_Step = 108;
            int Y_Step = 310;

            double Y = Y_Ini;
            double X = X_Ini;

            int iRow = 0;

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
                            XRect rect = new XRect(y, x, 280 + 00, 60);
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
    
    }
}
