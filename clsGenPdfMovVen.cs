using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using System.Diagnostics;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace APOffice
{
    class clsGenPdfMovVen
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public DataSet _dasGen = new DataSet();

        public string _strMovFat = "";
        public string _strCauTpd = "";
        public Boolean _bolMftPvi = true;           //Prezzo di vendita ivato
        public string _strConSql = "";
        public string _strTipDoc = "";  //
        public string _strCauTra = "";  //cmbCauTra.Text
        public string _strAspBen = "";  //cmbAspBen.Text
        public string _strNumCol = "";  //txtNumCol.Text;
        public string _strDocNot = "";  //txtDocNot.Text;
        public string _str022DocRaggruppArticoli = "";
        public string _strPrnPeso = "";
        public string _strDocTip = "";
        public string _strDes = "";
        public string _strCfo = "";
        public int _intPar024DecimalPesoPrezzo = 2;
        public Boolean _bolDocX2 = false;
        public Boolean _bolPrezzi = true;

        private string _strPar031Lotti2Pos = "";
             
        public clsGenPdfMovVen()
        {}

        public string PrnPdfMovVen()
        {
            string s = "";
            string sMsg = "";
            decimal d = 0;
            DataRow[] j;
            int iRows = 22;
            int iRow = 0;

            string sFrmDec = "00";

            _strPar031Lotti2Pos = _clsFun.ParGet(clsDefine.enuParametri.Par031Lotti2Pos, _strConSql);

            if(_intPar024DecimalPesoPrezzo != 2 && _intPar024DecimalPesoPrezzo > 0)
                sFrmDec = new string('0', 3);
            sFrmDec = "#,##0." + sFrmDec;

            s = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
            if (s == "")
                s = "001";
            DataTable tCnf = new clsQuery().ConfAzienda(s);

            DataTable tTes = _dasGen.Tables["DocTes"];
            DataTable tTmp = _dasGen.Tables["GesMovimenti"].Copy();
            DataTable tIva = _dasGen.Tables["TabIva"];
            DataTable tNot = _dasGen.Tables["TabNote"];

            DataTable t = tTmp.Clone();

            if (_str022DocRaggruppArticoli == "S")
            {
                foreach (DataRow y in tTmp.Rows)
                {
                    if (!(Boolean)y["mov_ann"])
                    {
                        j = t.Select("mov_art='" + (string)y["mov_art"] + "'");
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
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }

                    X = PdfMovVenTestata(tTes, iRow, X, Y, pd, gfx, page, font1, font5, font6);


                    iRow = 0;

                    /* TESTATA INIZIO 

                    X = 20;
                    gfx.DrawLine(XPens.Black, 0, X, page.Width, X);

                    XImage img = XImage.FromFile(_clsDef.PATHLOGHI + "LogoFattura.png");
                    if (_strDocTip == "" || !_strDocTip.Contains("P"))                      //Seck 20180227 no logo per preventivo
                        gfx.DrawImage(img, X, Y + 20, 430, 110);

                    iRow = 0;

                    X += 105;
                    X += 15;
                    gfx.DrawLine(XPens.Black, 0, X, page.Width, X);

                    X += 20;
                    s = "Spett.le ";
                    gfx.DrawString(s, font6, XBrushes.Black, Y + 300, X, XStringFormats.Default);

                    X += 5;

                    //s = ""; 
                    //if (_clsFun.Numerico(s))
                    //    s = Convert.ToInt64(s).ToString();

                    if (_strDocTip.Contains("P"))
                        s = _strTipDoc;
                    else if (_strMovFat == "M")
                        s = (_strTipDoc  + " numero           " + new string(' ', 21)).Substring(0,21);
                    else
                        s = "Documento numero     ";

                    if (!_strDocTip.Contains("P"))
                    {
                        if (_clsFun.Numerico((string)tTes.Rows[0]["tmp_ndo"], "1234567890"))
                            s += Convert.ToInt64(tTes.Rows[0]["tmp_ndo"]).ToString();
                        else
                            s += (string)tTes.Rows[0]["tmp_ndo"];

                        s +=  "/" + ((DateTime)tTes.Rows[0]["tmp_ddo"]).Year.ToString();
                    }
                    gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 15;
                    s = "Data                    " + ((DateTime)tTes.Rows[0]["tmp_ddo"]).ToString("dd/MM/yyyy");
                    gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 5;
                    s = (string)tTes.Rows[0]["tmp_rag"];
                    gfx.DrawString(s, font5, XBrushes.Black, Y + 300, X, XStringFormats.Default);

                    X += 15;
                    s = "Partita IVA       " + (string)tTes.Rows[0]["tmp_piv"];
                    gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 5;
                    s = ((string)tTes.Rows[0]["tmp_ind"]).Trim() + " " + tTes.Rows[0]["tmp_ncv"];
                    gfx.DrawString(s, font5, XBrushes.Black, Y + 300, X, XStringFormats.Default);

                    X += 15;
                    s = "Codice fiscale    " + (string)tTes.Rows[0]["tmp_cfi"];
                    gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 5;
                    s = (string)tTes.Rows[0]["tmp_loc"];
                    gfx.DrawString(s, font5, XBrushes.Black, Y + 300, X, XStringFormats.Default);

                    X += 15;
                    s = "Pagamento         " + (string)tTes.Rows[0]["tmp_tpg"];
                    gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

                    if (_strDes != "")
                    {
                        //X += 15;
                        s = "Destinazione: " + _strDes;
                        gfx.DrawString(s, font6, XBrushes.Black, Y + 200, X, XStringFormats.Default);
                    }

                    X += 5;
                    gfx.DrawLine(XPens.Black, 0, X, page.Width, X);

                    X += 20;
                    s = "Codice";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = "Descrizione";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 40, X, XStringFormats.Default);

                    if (_strPar031Lotti2Pos.Length > 0 && _strPar031Lotti2Pos.Substring(0, 1) == "S")
                    {
                        s = "Lotto";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 300, X, XStringFormats.Default);
                    }

                    s = "UM";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 350, X, XStringFormats.Default);

                    s = "Qta";
                    //if(_strPrnPeso == "S")
                    //    s = "CL/KG";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 390, X, XStringFormats.Default);

                    s = "Prezzo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 440, X, XStringFormats.Default);

                    s = "Importo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 505, X, XStringFormats.Default);

                    s = "IVA";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 560, X, XStringFormats.Default);

                    X += 15;

                    TESTATA FINE */
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    Boolean bCommento = false;
                    if (((string)t.Rows[i]["mov_iva"]).Trim() == "")
                        bCommento = true;

                    if (bCommento)
                    {
                        if (!DBNull.Value.Equals(t.Rows[i]["mov_ard"]))
                            s = (string)t.Rows[i]["mov_ard"];
                        gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);
                    }
                    else
                    {
                        s = (string)t.Rows[i]["mov_art"];
                        gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                        s = "";
                        if (!DBNull.Value.Equals(t.Rows[i]["mov_ard"]))
                            s = (string)t.Rows[i]["mov_ard"];
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 40, X, XStringFormats.Default);

                        s = "";
                        if (_strPar031Lotti2Pos.Length > 0 && _strPar031Lotti2Pos.Substring(0, 1) == "S")
                        {
                            if (!DBNull.Value.Equals(t.Rows[i]["mov_lot"]))
                            {
                                s = (string)t.Rows[i]["mov_lot"];
                                gfx.DrawString(s, font1, XBrushes.Black, Y + 300, X, XStringFormats.Default);
                            }
                        }

                        s = "";
                        if (!DBNull.Value.Equals(t.Rows[i]["mov_umi"]))
                        {
                            s = (string)t.Rows[i]["mov_umi"];
                            gfx.DrawString(s, font1, XBrushes.Black, Y + 350, X, XStringFormats.Default);
                        }

                        //if (_strPrnPeso == "S" && (string)t.Rows[i]["mov_umi"] == "KG")
                        //20190605 Seck se c'è il peso dovrebbe uscire il peso
                        if ((string)t.Rows[i]["mov_umi"] == "KG" && !DBNull.Value.Equals(t.Rows[i]["mov_qkg"]) && (decimal)t.Rows[i]["mov_qkg"] > 0)
                        {
                            //s = ((decimal)t.Rows[i]["mov_qkg"]).ToString("###,##0.00");
                            s = ((decimal)t.Rows[i]["mov_qkg"]).ToString(sFrmDec);
                            gfx.DrawString(s, font1, XBrushes.Black, Y + 410, X + 3, frmDX);
                        }
                        else
                        {
                            s = ((decimal)t.Rows[i]["mov_qta"]).ToString("###,##0");
                            gfx.DrawString(s, font1, XBrushes.Black, Y + 410, X + 3, frmDX);
                        }

                        if (_bolPrezzi)
                        {
                            d = (decimal)t.Rows[i]["mov_prv"];
                            if(_strCfo == "FOR")
                                d = (decimal)t.Rows[i]["mov_cos"];
                            j = tIva.Select("iva_cod='" + (string)t.Rows[i]["mov_iva"] + "'");
                            if (j.Length > 0 && !_bolMftPvi)
                                d = _clsFun.MenoIva((decimal)t.Rows[i]["mov_prv"], (decimal)j[0]["iva_ali"]);
                            //s = d.ToString("####,##0.00");
                            if (d > 0)
                            {
                                s = d.ToString(sFrmDec);
                                gfx.DrawString(s, font1, XBrushes.Black, Y + 478, X + 3, frmDX);
                            }

                            d = (decimal)t.Rows[i]["mov_imp"];
                            j = tIva.Select("iva_cod='" + (string)t.Rows[i]["mov_iva"] + "'");
                            if (j.Length > 0 && !_bolMftPvi)
                            {
                                d = _clsFun.MenoIva(d, (decimal)j[0]["iva_ali"]);
                            }
                            //s = d.ToString("####,##0.00");
                            if (d > 0)
                            {
                                s = d.ToString(sFrmDec);
                                gfx.DrawString(s, font1, XBrushes.Black, Y + 550, X + 3, frmDX);

                                s = (string)t.Rows[i]["mov_iva"];
                                gfx.DrawString(s, font1, XBrushes.Black, Y + 560, X, XStringFormats.Default);
                            }
                        }
                    }
                }

                X += 20;
            }

            Console.WriteLine("zzzzzzz");

            //X = 700;
            //gfx.DrawLine(XPens.Black, 0, X, page.Width, X);

            /* NOTE INIZIO */

            if (tNot.Rows.Count > 0)
            {
                int i = iRows - iRow - (tNot.Rows.Count / 2);

                if (i < 0)
                {
                    PdfMovVenPiede(X, Y, frmDX, gfx, page, font1, font2, font5, font6, font8, font7, tIva, iRows, iRow, pd);

                    page = pd.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    X = 20;

                    //page = pd.AddPage();
                    //gfx = XGraphics.FromPdfPage(page);
                    //X = 20;

                    X = PdfMovVenTestata(tTes, iRow, X, Y, pd, gfx, page, font1, font5, font6);

                    iRow = 0;
                    i = iRows - (tNot.Rows.Count / 2);
                }
                else
                    i = iRows - iRow - (tNot.Rows.Count / 2);

                for (int ii = 0; ii < i; ii++)
                {
                    X += 20;
                }

                X -= 20;

                int iNot = 0;

                foreach (DataRow y in tNot.Rows)
                {
                    s = (string)y["tab_txt"];
                    gfx.DrawString(s, font9, XBrushes.Black, Y + 10, X, XStringFormats.Default);
                    X += 10;
                    iNot++;

                    if(iNot % 2 == 0)
                        iRow++;

                    if(iNot == tNot.Rows.Count && iNot % 2 == 1)
                        iRow++;

                }
            }
            /* NOTE FINE */

            PdfMovVenPiede(X, Y, frmDX, gfx, page, font1, font2, font5, font6, font8, font7, tIva, iRows, iRow, pd);

            //PIEDE INIZIO

            //if (_bolPrezzi)
            //{
            //    X += 10;
            //    s = "TOTALE DOCUMENTO EURO ";
            //    gfx.DrawString(s, font6, XBrushes.Black, Y + 10, X + 5, XStringFormats.Default);

            //    s = ((decimal)tIva.Rows[tIva.Rows.Count - 1]["iva_tot"]).ToString("####,##0.00");
            //    gfx.DrawString(s, font7, XBrushes.Black, Y + 220, X + 9, frmDX);

            //    X += 10;
            //    gfx.DrawRectangle(XPens.Black, 0, X, 250, 122);

            //    X += 15;

            //    decimal dTot = 0;
            //    t = tIva.Copy();

            //    for (int i = 0; i <= t.Rows.Count - 1; i++)
            //    {
            //        if (iRow > iRows || iRow == 0 || i == 0)
            //        {
            //            if (iRow > iRows)
            //            {
            //                page = pd.AddPage();
            //                gfx = XGraphics.FromPdfPage(page);
            //                X = 20;
            //                iRow = 0;

            //            }

            //            s = "Descrizione";
            //            gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

            //            s = "Imponibile";
            //            gfx.DrawString(s, font1, XBrushes.Black, Y + 70, X, XStringFormats.Default);

            //            s = "IVA";
            //            gfx.DrawString(s, font1, XBrushes.Black, Y + 160, X, XStringFormats.Default);

            //            s = "Totale";
            //            gfx.DrawString(s, font1, XBrushes.Black, Y + 195, X, XStringFormats.Default);

            //            X += 10;
            //            gfx.DrawLine(XPens.Black, 0, X, 250, X);
            //            X += 8;
            //        }

            //        iRow++;
            //        X += 0;

            //        if (t.Rows.Count - 1 == i)
            //        {
            //            gfx.DrawLine(XPens.Black, 0, X, 250, X);
            //            X += 2;
            //        }
            //        X += 8;

            //        if (t.Rows.Count - 1 >= i)
            //        {
            //            //s = (string)t.Rows[i]["iva_cod"];
            //            //gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

            //            s = (string)t.Rows[i]["iva_des"];
            //            gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

            //            s = ((decimal)t.Rows[i]["iva_imp"]).ToString("####,##0.00");
            //            gfx.DrawString(s, font1, XBrushes.Black, Y + 130, X + 3, frmDX);

            //            s = ((decimal)t.Rows[i]["iva_iva"]).ToString("####,##0.00");
            //            gfx.DrawString(s, font1, XBrushes.Black, Y + 180, X + 3, frmDX);

            //            s = ((decimal)t.Rows[i]["iva_tot"]).ToString("####,##0.00");
            //            gfx.DrawString(s, font6, XBrushes.Black, Y + 240, X + 3, frmDX);
            //            dTot = (decimal)t.Rows[i]["iva_tot"];
            //            //break;
            //            //gfx.DrawRectangle(XPens.Black, 0, X, 250, X);
            //        }

            //        X += 8;
            //        //break;
            //    }
            //}



            //Console.WriteLine("aaaaaaaaa");

            //X = 720;
            //Y = 250;
            //gfx.DrawRectangle(XPens.Black, Y, X, 350, 122);

            //s = "CAUSALE DEL TRASPORTO                                   ASPETTO ESTERIONE DEI BENI                                                              N.COLLI";
            //gfx.DrawString(s, font8, XBrushes.Black, Y + 3, X + 7, XStringFormats.Default);

            //s = _strCauTra;
            //gfx.DrawString(s, font5, XBrushes.Black, Y + 3, X + 20, XStringFormats.Default);

            //s = _strAspBen;
            //gfx.DrawString(s, font5, XBrushes.Black, Y + 135, X + 20, XStringFormats.Default);

            //s = _strNumCol;
            //gfx.DrawString(s, font5, XBrushes.Black, Y + 330, X + 20, XStringFormats.Default);

            //X += 25;
            //gfx.DrawLine(XPens.Black, Y, X, 595, X);

            //s = "ORA RITIRO                                                                  DATA RITIRO";
            //gfx.DrawString(s, font8, XBrushes.Black, Y + 3, X + 7, XStringFormats.Default);

            //X += 25;
            //gfx.DrawLine(XPens.Black, Y, X, 595, X);

            //s = "VETTORE                                                                     FIRMA DEL CONDUCENTE                                    FIRMA DEL DESTINATARIO";
            //gfx.DrawString(s, font8, XBrushes.Black, Y + 3, X + 7, XStringFormats.Default);

            //X += 25;
            ////gfx.DrawLine(XPens.Black, Y, X, 595, X);
            //X += 25;
            //gfx.DrawLine(XPens.Black, Y, X, 595, X);

            //s = "NOTE";
            //gfx.DrawString(s, font8, XBrushes.Black, Y + 3, X + 7, XStringFormats.Default);

            //s = _strDocNot;
            //gfx.DrawString(s, font2, XBrushes.Black, Y + 3, X + 15, XStringFormats.Default);

            //PIEDE FINE

            ///****/

            //s = "Saldo    " + txtTotTot.Text;
            //gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = _clsDef.PATHPDF + "MovVen_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";

                if (_strTipDoc != "")
                {
                    string sNdo = ((string)tTes.Rows[0]["tmp_ndo"]);
                    if (_clsFun.Numerico(sNdo, "0123456789"))
                        sNdo = Convert.ToInt64(sNdo).ToString();
                    else
                    { if(sNdo.Length > 6)
                        sNdo = sNdo.Substring(5);
                    }

                    string sDes = ((string)tTes.Rows[0]["tmp_rag"]);
                    if (sDes.Length > 15)
                        sDes = sDes.Substring(0, 15);

                    s = _clsFun.CtrlCrtFile(_strTipDoc);

                    sFil = _clsDef.PATHPDF + s + "_" + ((DateTime)tTes.Rows[0]["tmp_ddo"]).ToString("yyyyMMdd") + "_" + sNdo + "_" + sDes + ".pdf";
                }

                pd.Save(sFil);
                // ...and start a viewer.

                if (_bolDocX2)
                    sFil = PdfDuplica(sFil);

                Process.Start(sFil);
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);

                sMsg = "File di stampa già aperto!";
            }

            return sMsg;
        }

        private double PdfMovVenTestata(DataTable tTes, int iRow, double X, double Y, PdfDocument pd, XGraphics gfx, PdfPage page, XFont font1, XFont font5, XFont font6)
        {
            string s = "";
            X = 20;
            gfx.DrawLine(XPens.Black, 0, X, page.Width, X);

            XImage img = XImage.FromFile(_clsDef.PATHLOGHI + "LogoFattura.png");
            if (_strDocTip == "" || !_strDocTip.Contains("P"))                      //Seck 20180227 no logo per preventivo
                gfx.DrawImage(img, X, Y + 20, 430, 110);

            iRow = 0;

            X += 105;
            X += 15;
            gfx.DrawLine(XPens.Black, 0, X, page.Width, X);

            X += 20;
            s = "Spett.le ";
            gfx.DrawString(s, font6, XBrushes.Black, Y + 300, X, XStringFormats.Default);

            X += 5;

            //s = ""; 
            //if (_clsFun.Numerico(s))
            //    s = Convert.ToInt64(s).ToString();

            if (_strDocTip.Contains("P"))
                s = _strTipDoc;
            else if (_strMovFat == "M")
                s = (_strTipDoc + " numero           " + new string(' ', 21)).Substring(0, 21);
            else
                s = "Documento numero     ";

            if (!_strDocTip.Contains("P"))
            {
                if (_clsFun.Numerico((string)tTes.Rows[0]["tmp_ndo"], "1234567890"))
                    s += Convert.ToInt64(tTes.Rows[0]["tmp_ndo"]).ToString();
                else
                    s += (string)tTes.Rows[0]["tmp_ndo"];

                s += "/" + ((DateTime)tTes.Rows[0]["tmp_ddo"]).Year.ToString();
            }
            gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

            X += 15;
            s = "Data                    " + ((DateTime)tTes.Rows[0]["tmp_ddo"]).ToString("dd/MM/yyyy");
            gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

            X += 5;
            s = (string)tTes.Rows[0]["tmp_rag"];
            gfx.DrawString(s, font5, XBrushes.Black, Y + 300, X, XStringFormats.Default);

            X += 15;
            s = "Partita IVA       " + (string)tTes.Rows[0]["tmp_piv"];
            gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

            X += 5;
            s = ((string)tTes.Rows[0]["tmp_ind"]).Trim() + " " + tTes.Rows[0]["tmp_ncv"];
            gfx.DrawString(s, font5, XBrushes.Black, Y + 300, X, XStringFormats.Default);

            X += 15;
            s = "Codice fiscale    " + (string)tTes.Rows[0]["tmp_cfi"];
            gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

            X += 5;
            s = (string)tTes.Rows[0]["tmp_loc"];
            gfx.DrawString(s, font5, XBrushes.Black, Y + 300, X, XStringFormats.Default);

            X += 15;
            s = "Pagamento         " + (string)tTes.Rows[0]["tmp_tpg"];
            gfx.DrawString(s, font6, XBrushes.Black, Y, X, XStringFormats.Default);

            if (_strDes != "")
            {
                //X += 15;
                s = "Destinazione: " + _strDes;
                gfx.DrawString(s, font6, XBrushes.Black, Y + 200, X, XStringFormats.Default);
            }

            X += 5;
            gfx.DrawLine(XPens.Black, 0, X, page.Width, X);

            X += 20;
            s = "Codice";
            gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

            s = "Descrizione";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 40, X, XStringFormats.Default);

            if (_strPar031Lotti2Pos.Length > 0 && _strPar031Lotti2Pos.Substring(0, 1) == "S")
            {
                s = "Lotto";
                gfx.DrawString(s, font1, XBrushes.Black, Y + 300, X, XStringFormats.Default);
            }

            s = "UM";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 350, X, XStringFormats.Default);

            s = "Qta";
            //if(_strPrnPeso == "S")
            //    s = "CL/KG";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 390, X, XStringFormats.Default);

            s = "Prezzo";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 440, X, XStringFormats.Default);

            s = "Importo";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 505, X, XStringFormats.Default);

            s = "IVA";
            gfx.DrawString(s, font1, XBrushes.Black, Y + 560, X, XStringFormats.Default);

            X += 15;

            return X;
        }

        private void PdfMovVenPiede(double X, double Y, XStringFormat frmDX, XGraphics gfx, PdfPage page, XFont font1, XFont font2, XFont font5, XFont font6, XFont font8, XFont font7, DataTable tIva, decimal iRows, int iRow, PdfDocument pd)
        {
            string s = "";

            if (iRow > iRows)
            {
                page = pd.AddPage();
                gfx = XGraphics.FromPdfPage(page);
                X = 20;
                iRow = 0;

            }
            X = 700;
            gfx.DrawLine(XPens.Black, 0, X, page.Width, X);

            if (_bolPrezzi)
            {
                X += 10;
                s = "TOTALE DOCUMENTO EURO ";
                gfx.DrawString(s, font6, XBrushes.Black, Y + 10, X + 5, XStringFormats.Default);

                s = ((decimal)tIva.Rows[tIva.Rows.Count - 1]["iva_tot"]).ToString("####,##0.00");
                gfx.DrawString(s, font7, XBrushes.Black, Y + 220, X + 9, frmDX);

                X += 10;
                gfx.DrawRectangle(XPens.Black, 0, X, 250, 122);

                X += 15;

                decimal dTot = 0;
                DataTable t = tIva.Copy();

                for (int i = 0; i <= t.Rows.Count - 1; i++)
                {
                    //if (iRow > iRows || iRow == 0 || i == 0)
                    if (i == 0)
                    {
                        //if (iRow > iRows)
                        //{
                        //    page = pd.AddPage();
                        //    gfx = XGraphics.FromPdfPage(page);
                        //    X = 20;
                        //    iRow = 0;

                        //}

                        s = "Descrizione";
                        gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                        s = "Imponibile";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 70, X, XStringFormats.Default);

                        s = "IVA";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 160, X, XStringFormats.Default);

                        s = "Totale";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 195, X, XStringFormats.Default);

                        X += 10;
                        gfx.DrawLine(XPens.Black, 0, X, 250, X);
                        X += 8;
                    }

                    iRow++;
                    X += 0;

                    if (t.Rows.Count - 1 == i)
                    {
                        gfx.DrawLine(XPens.Black, 0, X, 250, X);
                        X += 2;
                    }
                    X += 8;

                    if (t.Rows.Count - 1 >= i)
                    {
                        //s = (string)t.Rows[i]["iva_cod"];
                        //gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                        s = (string)t.Rows[i]["iva_des"];
                        gfx.DrawString(s, font2, XBrushes.Black, Y, X, XStringFormats.Default);

                        s = ((decimal)t.Rows[i]["iva_imp"]).ToString("####,##0.00");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 130, X + 3, frmDX);

                        s = ((decimal)t.Rows[i]["iva_iva"]).ToString("####,##0.00");
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 180, X + 3, frmDX);

                        s = ((decimal)t.Rows[i]["iva_tot"]).ToString("####,##0.00");
                        gfx.DrawString(s, font6, XBrushes.Black, Y + 240, X + 3, frmDX);
                        dTot = (decimal)t.Rows[i]["iva_tot"];
                        //break;
                        //gfx.DrawRectangle(XPens.Black, 0, X, 250, X);
                    }

                    X += 8;
                    //break;
                }
            }

            Console.WriteLine("aaaaaaaaa");

            X = 720;
            Y = 250;
            gfx.DrawRectangle(XPens.Black, Y, X, 350, 122);

            s = "CAUSALE DEL TRASPORTO                                   ASPETTO ESTERIONE DEI BENI                                                              N.COLLI";
            gfx.DrawString(s, font8, XBrushes.Black, Y + 3, X + 7, XStringFormats.Default);

            s = _strCauTra;
            gfx.DrawString(s, font5, XBrushes.Black, Y + 3, X + 20, XStringFormats.Default);

            s = _strAspBen;
            gfx.DrawString(s, font5, XBrushes.Black, Y + 135, X + 20, XStringFormats.Default);

            s = _strNumCol;
            gfx.DrawString(s, font5, XBrushes.Black, Y + 330, X + 20, XStringFormats.Default);

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
            gfx.DrawLine(XPens.Black, Y, X, 595, X);

            s = "NOTE";
            gfx.DrawString(s, font8, XBrushes.Black, Y + 3, X + 7, XStringFormats.Default);

            s = _strDocNot;
            gfx.DrawString(s, font2, XBrushes.Black, Y + 3, X + 15, XStringFormats.Default);

            /****/


        }

        public string PdfDuplica(string strFil)
        {
            string filename1 = strFil;
            string filename2 = strFil;

            // Open the input files
            PdfDocument inputDocument1 = PdfReader.Open(filename1, PdfDocumentOpenMode.Import);
            PdfDocument inputDocument2 = PdfReader.Open(filename2, PdfDocumentOpenMode.Import);

            // Create the output document
            PdfDocument outputDocument = new PdfDocument();

            // Show consecutive pages facing. Requires Acrobat 5 or higher.
            outputDocument.PageLayout = PdfPageLayout.TwoColumnLeft;

            XFont font = new XFont("Verdana", 10, XFontStyle.Bold);
            XStringFormat format = new XStringFormat();
            format.Alignment = XStringAlignment.Center;
            format.LineAlignment = XLineAlignment.Far;
            XGraphics gfx;
            XRect box;
            int count = Math.Max(inputDocument1.PageCount, inputDocument2.PageCount);
            for (int idx = 0; idx < count; idx++)
            {
                // Get page from 1st document
                PdfPage page1 = inputDocument1.PageCount > idx ?
                  inputDocument1.Pages[idx] : new PdfPage();

                // Get page from 2nd document
                PdfPage page2 = inputDocument2.PageCount > idx ?
                  inputDocument2.Pages[idx] : new PdfPage();

                // Add both pages to the output document
                page1 = outputDocument.AddPage(page1);

                page1.Height += 10;             //Seck 20180226 allargo il documento per fare in modo che si restringa sulla copia
                page1.Width += 10;

                page2 = outputDocument.AddPage(page2);


                page2.Height += 10;
                page2.Width += 10;


                //// Write document file name and page number on each page
                //gfx = XGraphics.FromPdfPage(page1);
                //box = page1.MediaBox.ToXRect();
                //box.Inflate(0, -10);
                //gfx.DrawString(String.Format("{0} • {1}", filename1, idx + 1), font, XBrushes.Red, box, format);

                //gfx = XGraphics.FromPdfPage(page2);
                //box = page2.MediaBox.ToXRect();
                //box.Inflate(0, -10);
                //gfx.DrawString(String.Format("{0} • {1}", filename2, idx + 1), font, XBrushes.Red, box, format);
            }

            // Save the document...

            string s = Path.GetDirectoryName(strFil) + Path.GetFileName(strFil) + "_2" + Path.GetExtension(strFil);

            Console.WriteLine(s);

            string filename = s;
            outputDocument.Save(filename);

            filename = PdfDuplica2in1(filename);

            return filename;

            //Process.Start(filename);
        }

        public string PdfDuplica2in1(string strFil)
        {

            //// Get a fresh copy of the sample PDF file
            //string filename = "Portable Document Format.pdf";
            //File.Copy(Path.Combine("../../../../../PDFs/", filename),
            //  Path.Combine(Directory.GetCurrentDirectory(), filename), true);

            string filename = strFil;

            // Create the output document
            PdfDocument outputDocument = new PdfDocument();

            // Show single pages
            // (Note: one page contains two pages from the source document)
            outputDocument.PageLayout = PdfPageLayout.SinglePage;

            XFont font = new XFont("Verdana", 8, XFontStyle.Bold);
            XStringFormat format = new XStringFormat();
            format.Alignment = XStringAlignment.Center;
            format.LineAlignment = XLineAlignment.Far;
            XGraphics gfx;
            XRect box;

            // Open the external document as XPdfForm object
            XPdfForm form = XPdfForm.FromFile(filename);

            for (int idx = 0; idx < form.PageCount; idx += 2)
            {
                // Add a new page to the output document
                PdfPage page = outputDocument.AddPage();
                page.Orientation = PdfSharp.PageOrientation.Landscape;
                double width = page.Width-2;
                //double height = page.Height;      
                double height = page.Height - 10;       //Seck 20180226

                int rotate = page.Elements.GetInteger("/Rotate");

                gfx = XGraphics.FromPdfPage(page);

                // Set page number (which is one-based)
                form.PageNumber = idx + 1;

                box = new XRect(0, 0, width / 2, height);
                // Draw the page identified by the page number like an image
                gfx.DrawImage(form, box);

                // Write document file name and page number on each page
                //box.Inflate(0, -10);
                //gfx.DrawString(String.Format("- {1} -", filename, idx + 1),
                //  font, XBrushes.Black, box, format);

                if (idx + 1 < form.PageCount)
                {
                    // Set page number (which is one-based)
                    form.PageNumber = idx + 2;

                    box = new XRect(width / 2, 0, width / 2, height);
                    // Draw the page identified by the page number like an image
                    gfx.DrawImage(form, box);

                    //// Write document file name and page number on each page
                    //box.Inflate(0, -10);
                    //gfx.DrawString(String.Format("- {1} -", filename, idx + 2),
                    //  font, XBrushes.Red, box, format);
                }
            }

            //// Save the document...
            //filename = "TwoPagesOnOne_tempfile.pdf";
            //outputDocument.Save(filename);
            //// ...and start a viewer.
            //Process.Start(filename);

            //string filename = s;
            outputDocument.Save(filename);

            return filename;

        }


    }
}
