using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using BarcodeLib;

namespace APOffice
{
    class clsStampe
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        //public string _strPrn = "";
        //public DataTable _tabPrn = new DataTable("tabTmp");

        //private const string MDBPATH = "\\APproject\\dBase\\DataBase.mdb";

        //private string _strConMdb = "";

        public clsStampe()
        {
            //_strConMdb = _clsFun.ConMdb(MDBPATH);
        }

        public void PrintDivFornitore(DataTable tabDiv, string strFor, string strDes)
        {
            string s = "";
            int iRows = 37;
            int iRow = 0;
            decimal d = 0;

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Divulgazione fornitori";

            // Create an empty page
            PdfPage page = pd.AddPage();
            page.Height = 842.0;
            page.Width = 595;
            page.Orientation = PdfSharp.PageOrientation.Landscape;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 7, XFontStyle.Bold);
            XFont font3 = new XFont("Courier new", 6, XFontStyle.Bold);

            double Y = 15;
            double X = 20;
            double yPos = 0;
            double yLen = 0;
            double yAlf = 3;
            double yNum = 30;
            double x = 0;
            double xAlt = 0;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            XPen pen = new XPen(XColors.Black, 0.8);

            DataTable t = tabDiv.Copy();

            for (int i = 0; i <= t.Rows.Count - 1; i++)
            {
                if (iRow > iRows || iRow == 0)
                {
                    if (iRow > 0)
                    {
                        page = pd.AddPage();
                        page.Height = 842.0;
                        page.Width = 595;
                        page.Orientation = PdfSharp.PageOrientation.Landscape;

                        gfx = XGraphics.FromPdfPage(page);
                    }

                    X = 20;
                    iRow = 0;
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " DIVULGAZIONE FORNITORI " + strFor + " " + strDes;
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    x = X;
                    yPos = Y;
                    yLen = 55;
                    xAlt = 15;
                    s = "Codice";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 55;
                    s = "Cod.forn";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 80;
                    s = "Barcode";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 160;
                    s = "Descrizione";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 20;
                    s = "UM";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 35;
                    s = "Q.tà";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = "Cos An";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = "Cos dv";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 40;
                    s = "Diff.";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = "Prezzo anag.";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 30;
                    s = "Marg.%";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = "Prezzo Cons";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, x - 35);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 30;
                    s = "Marg.%";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, x - 35);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + 3, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 30;
                    s = "Stato";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, x - 35);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + 3, x, XStringFormats.Default);

                    X += 20;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    //s = (string)t.Rows[i]["div_art"];
                    //gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                    x = X;
                    yPos = Y;
                    yLen = 55;
                    s = (string)t.Rows[i]["div_art"];
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 55;
                    s = (string)t.Rows[i]["div_arf"];
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 80;
                    s = (string)t.Rows[i]["div_ean"];
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf+6, x - 5, XStringFormats.Default);

                    double xEan = x-4;
                    double yEan = yPos+5;

                    if (s.Trim().Length > 0)
                    {
                        try
                        {
                            if (s.Length <= 8)
                            {
                                Image img = _ean08.Encode(tpEan08, s, Color.Black, Color.White, 200, 200);
                                MemoryStream ms = new MemoryStream();
                                img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                gfx.DrawImage(img, yEan, xEan, 70, 7);
                            }
                            else
                            {
                                Image img = _ean13.Encode(tpEan13, s, Color.Black, Color.White, 200, 200);
                                MemoryStream ms = new MemoryStream();
                                img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                gfx.DrawImage(img, yEan, xEan, 70, 7);
                            }
                        }
                        catch (Exception ex)
                        {
                            _clsFun.ErrorLog(ex.Message, s);
                        }
                    }

                    yPos += yLen;
                    yLen = 160;
                    s = (string)t.Rows[i]["div_ard"];
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font3, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 20;
                    s = (string)t.Rows[i]["div_umi"];
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 35;
                    s = ((decimal)t.Rows[i]["div_qta"]).ToString("#0");
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yNum + 1, x, frmDX);

                    yPos += yLen;
                    yLen = 50;
                    s = ((decimal)t.Rows[i]["art_cos"]).ToString("#0.000");
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yNum + 10, x, frmDX);

                    yPos += yLen;
                    yLen = 50;
                    s = ((decimal)t.Rows[i]["div_cos"]).ToString("#0.000");
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yNum + 10, x, frmDX);

                    yPos += yLen;
                    yLen = 40;



                    s = ((decimal)t.Rows[i]["dif_cos"]).ToString("#0.000");
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);

                    XRect rect = new XRect(yPos, x-10, yLen, xAlt);
                    if ((decimal)t.Rows[i]["dif_cos"] > 0)
                        gfx.DrawRectangle(new SolidBrush(Color.OrangeRed), rect);
                    else if ((decimal)t.Rows[i]["dif_cos"] < 0)
                        gfx.DrawRectangle(new SolidBrush(Color.LightGreen), rect);

                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);

                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yNum + 1, x, frmDX);



                    yPos += yLen;
                    yLen = 50;
                    s = ((decimal)t.Rows[i]["art_prp"]).ToString("#0.00");
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yNum + 10, x, frmDX);

                    yPos += yLen;
                    yLen = 30;
                    s = ((decimal)t.Rows[i]["art_mav"]).ToString("#0.00");
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yNum - 1, x, frmDX);

                    yPos += yLen;
                    yLen = 50;
                    s = ((decimal)t.Rows[i]["div_prp"]).ToString("#0.00");
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yNum + 10, x, frmDX);

                    yPos += yLen;
                    yLen = 30;
                    s = ((decimal)t.Rows[i]["var_mav"]).ToString("#0.00");
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yNum - 1, x, frmDX);

                    yPos += yLen;
                    yLen = 30;
                    s = "";
                    if ((string)t.Rows[i]["div_off"] != "")
                        s = "O";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yNum - 20, x, XStringFormats.Default);
                }

                X += 15;
               // break;

            }

            X += 15;

            s = "Totale    " + d.ToString();
            gfx.DrawString(s, font1, XBrushes.Black, Y + 240, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\DivFor_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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

        public void PrintMovLotto(DataTable tabMov, string strPar)
        {
            string sDes = "";
            decimal dLotPes = 0;

            string[] a = strPar.Split('|');
            sDes = a[0];
            dLotPes = Convert.ToDecimal(a[1]);

            string s = "";
            int iRows = 37;
            int iRow = 0;
            decimal dPesTot = 0;
            string sEcr = "";

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Movimenti lotto";

            // Create an empty page
            PdfPage page = pd.AddPage();
            //page.Height = 842.0;
            //page.Width = 595;
            //page.Orientation = PdfSharp.PageOrientation.Landscape;

            page.Height = 845.0;
            page.Orientation = PdfSharp.PageOrientation.Portrait;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 7, XFontStyle.Bold);
            XFont font3 = new XFont("Courier new", 6, XFontStyle.Bold);

            double Y = 15;
            double X = 20;
            double yPos = 0;
            double yLen = 0;
            double yAlf = 3;
            double yNum = 30;
            double x = 0;
            double xAlt = 0;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            XPen pen = new XPen(XColors.Black, 0.8);

            //DataTable t = tabMov.Copy();
            DataTable t = new DataView(tabMov, "tmp_liv=3", "", DataViewRowState.CurrentRows).ToTable();

            for (int i = 0; i <= t.Rows.Count - 1; i++)
            {
                if (iRow > iRows || iRow == 0)
                {
                    if (iRow > 0)
                    {
                        page = pd.AddPage();
                        //page.Height = 842.0;
                        //page.Width = 595;
                        //page.Orientation = PdfSharp.PageOrientation.Landscape;

                        //PdfPage page = pd.AddPage();
                        page.Height = 845.0;
                        page.Orientation = PdfSharp.PageOrientation.Portrait;

                        gfx = XGraphics.FromPdfPage(page);
                    }

                    X = 20;
                    iRow = 0;
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " MOVIMENTI LOTTO " + sDes;
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    x = X;
                    yPos = Y;
                    yLen = 100;
                    xAlt = 15;
                    s = "Causale";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 28;
                    s = "Segno";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 80;
                    s = "Documento";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = "Data";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = "Articolo";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 135;
                    s = "Descrizione";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 30;
                    s = "Pezzi";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = "Peso";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    X += 20;
                }

                iRow++;

                s = (string)t.Rows[i]["tmp_l1c"] + (string)t.Rows[i]["tmp_l2c"] + (string)t.Rows[i]["tmp_l3c"];
                if (sEcr != s)
                {
                    Boolean b = true;
                    if (sEcr == "")
                        b = false;

                    sEcr = s;

                    x = X;
                    yPos = Y;
                    yLen = 100;
                    
                    if(b)
                        gfx.DrawLine(pen, yPos + 0, x - 5, yPos + 520, x - 5);

                    s = "    *** " + (string)t.Rows[i]["tmp_l2d"] + " - " + (string)t.Rows[i]["tmp_l3d"];
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x+8, XStringFormats.Default);

                    decimal dPes = 0;

                    DataRow[] j = t.Select("tmp_l1c+tmp_l2c+tmp_l3c='" + sEcr + "'");
                    if(j.Length > 0)
                    {
                        for (int ii = 0; ii < j.Length; ii++)
                        {
                            if ((string)j[ii]["tab_sgm"] == "-")
                                dPes -= (decimal)j[ii]["mov_qkg"];
                            else
                                dPes += (decimal)j[ii]["mov_qkg"];
                        }
                    }

                    yPos += 488;
                    yLen = 40;
                    s = dPes.ToString("#0.000");
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yNum + 1, x+8, frmDX);

                    X += 20;
                }

                if (t.Rows.Count - 1 >= i)
                {
                    //s = (string)t.Rows[i]["div_art"];
                    //gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                    x = X;
                    yPos = Y;
                    yLen = 100;
                    s = (string)t.Rows[i]["tab_des"];
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 28;
                    s = (string)t.Rows[i]["tab_sgm"];
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 80;
                    s = (string)t.Rows[i]["DocNdo"];
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf + 6, x - 5, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = ((DateTime)t.Rows[i]["DocDdo"]).ToString("dd/MM/yyyy");
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font3, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = (string)t.Rows[i]["mov_art"];
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 130;
                    s = (string)t.Rows[i]["mov_ard"];
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 40;
                    s = ((decimal)t.Rows[i]["mov_qta"]).ToString("#0");
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yNum + 1, x, frmDX);

                    yPos += yLen;
                    yLen = 65;
                    s = ((decimal)t.Rows[i]["mov_qkg"]).ToString("#0.000");
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yNum + 10, x, frmDX);

                    //dPesTot += (decimal)t.Rows[i]["mov_qkg"];

                    if ((string)t.Rows[i]["tab_sgm"] == "-")
                        dPesTot -= (decimal)t.Rows[i]["mov_qkg"];
                    else
                        dPesTot += (decimal)t.Rows[i]["mov_qkg"];
                }

                X += 15;
                // break;

            }

            X += 15;

            s = "Peso mezzena " + dLotPes.ToString("#0.000") + "    ";
            s += "Totale peso prodotti " + dPesTot.ToString("#0.000") + "   ";
            s += "Differenza " + (dLotPes - dPesTot).ToString("#0.000");
            gfx.DrawString(s, font1, XBrushes.Black, Y + 1, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = _clsDef.PATHPDF + "MovLot_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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

        public void PrintStaScontrini(DataTable tabMov, string strPar)
        {
            string sDes = "";
            decimal dTot = 0;

            string[] a = strPar.Split('|');
            sDes = a[0];

            string s = "";
            int iRows = 49;
            int iRow = 0;

            //Barcode _ean13 = new BarcodeLib.Barcode();
            //BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            //_ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            //Barcode _ean08 = new BarcodeLib.Barcode();
            //BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            //_ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Scontini";

            // Create an empty page
            PdfPage page = pd.AddPage();
            //page.Height = 842.0;
            //page.Width = 595;
            //page.Orientation = PdfSharp.PageOrientation.Landscape;

            page.Height = 845.0;
            page.Orientation = PdfSharp.PageOrientation.Portrait;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 7, XFontStyle.Bold);
            XFont font3 = new XFont("Courier new", 6, XFontStyle.Bold);

            double Y = 15;
            double X = 20;
            double yPos = 0;
            double yLen = 0;
            double yAlf = 3;
            double yNum = 30;
            double x = 0;
            double xAlt = 0;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            XPen pen = new XPen(XColors.Black, 0.8);

            //DataTable t = tabMov.Copy();
            DataTable t = new DataView(tabMov, "", "", DataViewRowState.CurrentRows).ToTable();

            for (int i = 0; i <= t.Rows.Count - 1; i++)
            {
                if (iRow > iRows || iRow == 0)
                {
                    if (iRow > 0)
                    {
                        page = pd.AddPage();
                        //page.Height = 842.0;
                        //page.Width = 595;
                        //page.Orientation = PdfSharp.PageOrientation.Landscape;

                        //PdfPage page = pd.AddPage();
                        page.Height = 845.0;
                        page.Orientation = PdfSharp.PageOrientation.Portrait;

                        gfx = XGraphics.FromPdfPage(page);
                    }

                    X = 20;
                    iRow = 0;
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " " + sDes;
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    x = X;
                    yPos = Y;
                    yLen = 50;
                    xAlt = 15;
                    s = "Negozio";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = "Causale";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = "Cassa";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = "Ora";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = "Scontrino";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 100;
                    s = "Importo";
                    gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    X += 20;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    //s = (string)t.Rows[i]["div_art"];
                    //gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                    x = X;
                    yPos = Y;
                    yLen = 50;
                    s = (string)t.Rows[i]["vet_neg"];
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = (string)t.Rows[i]["vet_cau"];
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = (string)t.Rows[i]["vet_pos"];
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf + 6, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 50;
                    s = (string)t.Rows[i]["vet_ora"];
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 100;
                    s = (string)t.Rows[i]["vet_sco"];
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yAlf, x, XStringFormats.Default);

                    yPos += yLen;
                    yLen = 150;
                    s = ((decimal)t.Rows[i]["vet_imp"]).ToString("#0.000");
                    //gfx.DrawRectangle(pen, yPos, x - 10, yLen, xAlt);
                    gfx.DrawString(s, font2, XBrushes.Black, yPos + yNum + 10, x, frmDX);

                    dTot += (decimal)t.Rows[i]["vet_imp"];
                }

                X += 15;
                // break;

            }

            X += 5;

            s = "Totale giorno Euro:                       " + (dTot).ToString("#0.00");
            gfx.DrawString(s, font1, XBrushes.Black, Y + 1, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = _clsDef.PATHPDF + "Scontrini_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
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
