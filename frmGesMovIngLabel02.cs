using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.IO;
using BarcodeLib;
using System.Data.SqlClient;

namespace APOffice
{
    public partial class frmGesMovIngLabel02 : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        //public DataTable _tabTab = new DataTable();
        public decimal _decEtiNum = 0;

        private string _str001_Print = "";
        private string _str002_Dim = "";

        private string _str003_Posz_NegRag = "";        
        private string _str003_Font_NegRag = "";
        private string _str003_Posz_NegInd = "";
        private string _str003_Font_NegInd = "";
        private string _str003_Posz_CliDes = "";
        private string _str003_Font_CliDes = "";

        private string _str004_Posz_ArtDes = "";
        private string _str004_Font_ArtDes = "";

        private string _str005_Posz_RicDes = "";
        private string _str005_Font_RicDes = "";

        private string _str006_Posz_PriLbl = "";
        private string _str006_Font_PriLbl = "";
        private string _str006_Posz_PriDta = "";
        private string _str006_Font_PriDta = "";

        private string _str007_Posz_ColLbl = "";
        private string _str007_Font_ColLbl = "";
        private string _str007_Posz_ColQta = "";
        private string _str007_Font_ColQta = "";

        private string _str008_Posz_PneLbl = "";
        private string _str008_Font_PneLbl = "";
        private string _str008_Posz_PneQta = "";
        private string _str008_Font_PneQta = "";

        private string _str009_Posz_PkgLbl = "";
        private string _str009_Font_PkgLbl = "";
        private string _str009_Posz_PkgVal = "";
        private string _str009_Font_PkgVal = "";

        private string _str010_Posz_ImpLbl = "";
        private string _str010_Font_ImpLbl = "";
        private string _str010_Posz_ImpVal = "";
        private string _str010_Font_ImpVal = "";

        private string _str011_PrnImmediata = "";

        public string _strConSql = "";
        public string _strCliDes = "";
        public string _strArtCod = "";
        public string _strArtDes = "";
        //public string _strLotCod = "";
        public string _strArtEan = "";
        public decimal _decMovQta = 0;
        public decimal _decMovQkg = 0;
        public decimal _decMovPrv = 0;                  //Prezzo al kilo
        public decimal _decMovImp = 0;                  //Importo
        public DateTime _dayArtSca = new DateTime();    //Data scadenza

        //public string _strEan = "";
        //public string _strPes = "";
        //public string _strTar = "";
        //public string _strPrv = "";
        //public string _strImp = "";
        //public string _strDsc = "";     //Data scadenza

        private int intStampati = 0;

        public frmGesMovIngLabel02()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmAnaArtIngLabel_Load(object sender, EventArgs e)
        {
            FillFont("R");

            if (_str001_Print == "")
                return;

            //if (_str001_Print != "")
            else
            {
                printDocument1.PrinterSettings.PrinterName = _str001_Print;

                printDocument1.DefaultPageSettings.Landscape = true;

                printDocument1.DefaultPageSettings.Landscape = true;

                int iH = 200;
                int iW = 200;

                if (_str002_Dim != "")
                {
                    string[] a = _str002_Dim.Split(',');

                    iH = Convert.ToInt16(a[0]);
                    iW = Convert.ToInt16(a[1]);
                }

                PaperSize pSize = new PaperSize();
                pSize.PaperName = "EtiNome";
                pSize.RawKind = (int)PaperKind.Custom;
                pSize.Width = iW;
                pSize.Height = iH;
                printDocument1.DefaultPageSettings.PaperSize = pSize;

                FillPrint();

                if (_str011_PrnImmediata == "S")
                {
                    printDocument1.Print();
                    Esci();
                }
            }
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmAnaArtIngLabel_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void printPreviewControl1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void FillPrint()
        {
            intStampati = 0;
            printPreviewControl1.InvalidatePreview();
        }

        private void OldprintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            string s = "";
            string sTxt = "";

            Boolean bBordi = true;

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            intStampati++;

            //e.PageSettings.Landscape = true;

            //printDocument1.DefaultPageSettings.Landscape = false;
            //printDocument1.PrinterSettings.DefaultPageSettings.Landscape = true;

            //System.Drawing.Printing.PageSettings ps = printDocument1..GetPageSettings();
            //ps.Landscape = true;
            //printDocument1.SetPageSettings(ps);


            //Font FonDe1 = _clsFun.Str2Font(new Font("Arial", 20f, FontStyle.Bold), Convert.ToString(tFon.Rows[0]["tab_fo3"]));
            //Brush BruDe3 = this._clsFun.Str2Color((string)tFon.Rows[0]["tab_fo5"]);

            s = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
            if (s == "")
                s = "001";
            DataTable tCnf = _clsQry.ConfAzienda(s);

            //Font Font2 = new System.Drawing.Font(sFonName, fFonDim, fSty);

            //Font FonDe1 = _clsFun.Str2Font(new Font("Arial", 20f, FontStyle.Bold), Convert.ToString(tFon.Rows[0]["tab_fo3"]));
            SolidBrush Bruh1 = new SolidBrush(Color.Black);

            int C = 10;
            int R = 100;
            int R2 = 15;

            //StringFormat stringFormat = new StringFormat();
            //stringFormat.FormatFlags = StringFormatFlags.DirectionVertical;

            //StringFormat stringFormat2 = new StringFormat();
            //stringFormat2.FormatFlags = StringFormatFlags.DirectionVertical;

            //String s = _clsDef.CRLF;

            int iX = 50;
            int iY = 100;
            int iXx = 50;
            int iYy = 100;
            int iIl = 20;   //Interlinea
            int iRow = 0;
            int iXProgr = 0;
            try
            {

                StringFormat format1 = new StringFormat(StringFormatFlags.NoClip);
                format1.LineAlignment = StringAlignment.Near;
                format1.Alignment = StringAlignment.Center;
                //format1.FormatFlags = StringFormatFlags.DirectionVertical;

                /*** RAGIONE SOCIALE ***/

                //s = _clsDef.PATHLOGHI + "logoPrint1.png";
                //e.Graphics.DrawImage(Image.FromFile(s), 10, -5, 50, 50);

                sTxt = (string)tCnf.Rows[0]["cnf_rag"];

                string[] aFon = _str003_Font_NegRag.Split(':');
                string sFonName = aFon[0];
                float fFonDim = (float)Convert.ToDouble(aFon[1]);
                FontStyle fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font Font = new System.Drawing.Font(sFonName, fFonDim, fSty);

                string[] a = _str003_Posz_NegRag.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                iX = iX + iIl + iXProgr;

                RectangleF rectF1 = new RectangleF(iY, iX, iYy, iXx);
                //e.Graphics.DrawString(sTxt, Font, Brushes.Black, rectF1, format1);
                //if (bBordi)
                //    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));

                iXProgr = iX;

                /*** DATI NEGOZIO ***/

                sTxt = (string)tCnf.Rows[0]["cnf_ind"] + "-";
                sTxt += (string)tCnf.Rows[0]["cnf_cap"] + " ";
                sTxt += (string)tCnf.Rows[0]["cnf_loc"] + _clsDef.CRLF;
                sTxt += (string)tCnf.Rows[0]["cnf_web"] + " ";

                a = _str003_Posz_NegInd.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                aFon = _str003_Font_NegInd.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                iX = iX + iIl + iXProgr;


                iY = 50;
                iX = 100;
                iYy = 100;
                iXx = 50;


                sTxt = "12345678901234567890";


                for (int n = 0; n < 5; n++)
                {
                    rectF1 = new RectangleF(iY, iX, iYy, iXx);
                    e.Graphics.DrawString(sTxt, Font, Brushes.Black, rectF1, format1);
                    if (bBordi)
                        e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                    iXProgr = iX;

                    sTxt = sTxt.Substring(n);
                    iX += 30;
                }

                ///*** DESCRIZIONE CLIENTE ***/

                //a = _str003_Posz_CliDes.Split(',');
                //iY = Convert.ToInt16(a[0]);
                //iX = Convert.ToInt16(a[1]);
                //iYy = Convert.ToInt16(a[2]);
                //iXx = Convert.ToInt16(a[3]);
                //iIl = Convert.ToInt16(a[4]);

                //aFon = _str003_Font_CliDes.Split(':');
                //sFonName = aFon[0];
                //fFonDim = (float)Convert.ToDouble(aFon[1]);
                //fSty = FontStyle.Bold;
                //if (aFon[2] == "0")
                //    fSty = FontStyle.Regular;
                //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                //iX = iX + iIl + iXProgr;

                //rectF1 = new RectangleF(iY, iX, iYy, iXx);
                //e.Graphics.DrawString(_strCliDes, Font, Brushes.Black, rectF1, format1);
                //if (bBordi)
                //    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                //iXProgr = iX;

                ///*** DESCRIZIONE ARTICOLO ***/

                //aFon = _str004_Font_ArtDes.Split(':');
                //sFonName = aFon[0];
                //fFonDim = (float)Convert.ToDouble(aFon[1]);
                //fSty = FontStyle.Bold;
                //if (aFon[2] == "0")
                //    fSty = FontStyle.Regular;
                //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);

                //a = _str004_Posz_ArtDes.Split(',');
                //iY = Convert.ToInt16(a[0]);
                //iX = Convert.ToInt16(a[1]);
                //iYy = Convert.ToInt16(a[2]);
                //iXx = Convert.ToInt16(a[3]);
                //iIl = Convert.ToInt16(a[4]);
                //iX = iX + iIl + iXProgr;

                //sTxt = _strArtDes;
                //rectF1 = new RectangleF(iY, iX, iYy, iXx);
                //e.Graphics.DrawString(sTxt, Font, Brushes.Black, rectF1, format1);
                //if (bBordi)
                //    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                //iXProgr += iX;


                ///*** Descrizione ***/

                //a = _str005_Posz_RicDes.Split(',');
                //iY = Convert.ToInt16(a[0]);
                //iX = Convert.ToInt16(a[1]);
                //iYy = Convert.ToInt16(a[2]);
                //iXx = Convert.ToInt16(a[3]);
                //iIl = Convert.ToInt16(a[4]);

                //aFon = _str005_Font_RicDes.Split(':');
                //sFonName = aFon[0];
                //fFonDim = (float)Convert.ToDouble(aFon[1]);
                //fSty = FontStyle.Bold;
                //if (aFon[2] == "0")
                //    fSty = FontStyle.Regular;
                //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                //iX = iX + iIl + iXProgr;

                //rectF1 = new RectangleF(iY, iX, iYy, iXx);
                //e.Graphics.DrawString("RIEPILOGO PRODOTTO", Font, Brushes.Black, rectF1, format1);
                //if (bBordi)
                //    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                //iXProgr = iX;

                //iX = 150;

                ///*** Descrizioni piede ***/

                //aFon = _str007_Font_FooDes.Split(':');
                //sFonName = aFon[0];
                //fFonDim = (float)Convert.ToDouble(aFon[1]);
                //fSty = FontStyle.Bold;
                //if (aFon[2] == "0")
                //    fSty = FontStyle.Regular;
                //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);

                //e.Graphics.DrawString("Preimballato il: ", Font, Brushes.Black, iY + 5, iX + 25);

                //e.Graphics.DrawString("Colli", Font, Brushes.Black, iY + 5, iX + 50);

                //e.Graphics.DrawString("Peso netto", Font, Brushes.Black, iY + 5, iX + 75);

                //e.Graphics.DrawString("€/Kg", Font, Brushes.Black, iY + 5, iX + 100);

                ////e.Graphics.DrawString("Da cons. entro il :", Font, Brushes.Black, iY + 5, iX + 100);

                //e.Graphics.DrawString("Importo", Font, Brushes.Black, iY + 5, iX + 125);

                ///*** Valori piede ***/

                //aFon = _str008_Font_FooVal.Split(':');
                //sFonName = aFon[0];
                //fFonDim = (float)Convert.ToDouble(aFon[1]);
                //fSty = FontStyle.Bold;
                //if (aFon[2] == "0")
                //    fSty = FontStyle.Regular;
                //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);

                //e.Graphics.DrawString(DateTime.Today.ToString("dd/MM/yyy"), Font, Brushes.Black, iY + 100, iX + 26);

                //e.Graphics.DrawString(_decMovQta.ToString("#0"), Font, Brushes.Black, iY + 100, iX + 50);

                //e.Graphics.DrawString(_decMovQkg.ToString("#0.000") + " Kg", Font, Brushes.Black, iY + 100, iX + 75);

                //e.Graphics.DrawString(_decMovPrv.ToString("#0.00") + " Kg", Font, Brushes.Black, iY + 100, iX + 100);

                //e.Graphics.DrawString(_decMovImp.ToString("#0.00") + " €", Font, Brushes.Black, iY + 100, iX + 125);


                ////aFon = _str009_Font_FooImp.Split(':');
                ////sFonName = aFon[0];
                ////fFonDim = (float)Convert.ToDouble(aFon[1]);
                ////fSty = FontStyle.Bold;
                ////if (aFon[2] == "0")
                ////    fSty = FontStyle.Regular;
                ////Font = new System.Drawing.Font(sFonName, fFonDim, fSty);


                ///*** Valori DATE ***/

                //aFon = _str010_Font_FooDay.Split(':');
                //sFonName = aFon[0];
                //fFonDim = (float)Convert.ToDouble(aFon[1]);
                //fSty = FontStyle.Bold;
                //if (aFon[2] == "0")
                //    fSty = FontStyle.Regular;
                //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);

                ////e.Graphics.DrawString(DateTime.Today.ToString("dd/MM/yyy"), Font, Brushes.Black, iY + 80, iX + 26);

                ////e.Graphics.DrawString(_dayArtSca.ToString("dd/MM/yyy"), Font, Brushes.Black, iY + 80, iX + 101);



            }
            catch (Exception ex)
            {
                e.Graphics.DrawString(s,
                    new Font("Arial", 40, FontStyle.Bold), Brushes.Black, 50, 125);

                string sMsg = ex.Message;
            }

            e.HasMorePages = true;
            if (intStampati >= _decEtiNum)
                e.HasMorePages = false;

        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            string s = "";
            string sTxt = "";

            Boolean bBordi = false;

            Barcode _ean13 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan13 = BarcodeLib.TYPE.EAN13;
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Barcode _ean08 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            _ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            intStampati++;

            //e.PageSettings.Landscape = true;

            //printDocument1.DefaultPageSettings.Landscape = false;
            //printDocument1.PrinterSettings.DefaultPageSettings.Landscape = true;

            //System.Drawing.Printing.PageSettings ps = printDocument1..GetPageSettings();
            //ps.Landscape = true;
            //printDocument1.SetPageSettings(ps);


            //Font FonDe1 = _clsFun.Str2Font(new Font("Arial", 20f, FontStyle.Bold), Convert.ToString(tFon.Rows[0]["tab_fo3"]));
            //Brush BruDe3 = this._clsFun.Str2Color((string)tFon.Rows[0]["tab_fo5"]);

            s = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
            if (s == "")
                s = "001";
            DataTable tCnf = _clsQry.ConfAzienda(s);

            //Font Font2 = new System.Drawing.Font(sFonName, fFonDim, fSty);

            //Font FonDe1 = _clsFun.Str2Font(new Font("Arial", 20f, FontStyle.Bold), Convert.ToString(tFon.Rows[0]["tab_fo3"]));
            SolidBrush Bruh1 = new SolidBrush(Color.Black);

            int C = 10;
            int R = 100;
            int R2 = 15;

            //StringFormat stringFormat = new StringFormat();
            //stringFormat.FormatFlags = StringFormatFlags.DirectionVertical;

            //StringFormat stringFormat2 = new StringFormat();
            //stringFormat2.FormatFlags = StringFormatFlags.DirectionVertical;

            //String s = _clsDef.CRLF;

            int iY = 5;
            int iX = -290;
            int iYy = 210;
            int iXx = 280;
            int iIl = 20;   //Interlinea
            int iRow = 0;
            int iXProgr = 0;
            try
            {

                //StringFormat format1 = new StringFormat(StringFormatFlags.NoClip);
                //format1.LineAlignment = StringAlignment.Near;
                //format1.Alignment = StringAlignment.Center;
                //format1.FormatFlags = StringFormatFlags.DirectionVertical;
                ////format1.FormatFlags = StringFormatFlags.DirectionRightToLeft;

                e.Graphics.RotateTransform(90);

                StringFormat format1 = new StringFormat(StringFormatFlags.NoClip);
                format1.LineAlignment = StringAlignment.Center;
                format1.Alignment = StringAlignment.Center;
                //format1.FormatFlags = StringFormatFlags.DirectionRightToLeft;
                //format1.FormatFlags = StringFormatFlags.DirectionVertical | StringFormatFlags.DirectionRightToLeft;
                //format1.FormatFlags = StringFormatFlags.DirectionVertical;

                //Pen blackPen = new Pen(Color.Black, 3);

                //iY = inizio riga
                //iX = inizio colonna
                //Yy = fine riga
                //Yx = fine colonna
                //rectF1 = new RectangleF(iY, iX, iYy, iXx);
                RectangleF rectF1 = new RectangleF(5, -290, 210, 280);
                //e.Graphics.DrawString(sTxt, new Font("Arial", 20f, FontStyle.Bold), Brushes.Black, rectF1, format1);
                if (bBordi)
                    e.Graphics.DrawRectangle(new Pen(Color.Black, 5), Rectangle.Round(rectF1));

                /*** RAGIONE SOCIALE ***/

                //s = _clsDef.PATHLOGHI + "logoPrint1.png";
                //e.Graphics.DrawImage(Image.FromFile(s), 10, -5, 50, 50);

                sTxt = (string)tCnf.Rows[0]["cnf_rag"];

                string[] aFon = _str003_Font_NegRag.Split(':');
                string sFonName = aFon[0];
                float fFonDim = (float)Convert.ToDouble(aFon[1]);
                FontStyle fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font Font = new System.Drawing.Font(sFonName, fFonDim, fSty);

                string[] a = _str003_Posz_NegRag.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                //format1.FormatFlags = StringFormatFlags.DirectionVertical | StringFormatFlags.DirectionRightToLeft;
                //format1.FormatFlags = StringFormatFlags.DirectionVertical;

                //iY = -5;
                //iX = -290;
                //iYy = 150;
                //iXx = 15;
                //iIl = Convert.ToInt16(a[4]);

                iX = iX + iIl + iXProgr;

                rectF1 = new RectangleF(iY, iX, iYy, iXx);
                e.Graphics.DrawString(sTxt, Font, Brushes.Black, rectF1, format1);
                if (bBordi)
                    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));

                //iXProgr = iX;

                /*** DATI NEGOZIO ***/

                sTxt = (string)tCnf.Rows[0]["cnf_ind"] + "-";
                sTxt += (string)tCnf.Rows[0]["cnf_cap"] + " ";
                sTxt += (string)tCnf.Rows[0]["cnf_loc"] + _clsDef.CRLF;
                sTxt += (string)tCnf.Rows[0]["cnf_web"] + " ";

                a = _str003_Posz_NegInd.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                aFon = _str003_Font_NegInd.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                iX = iX + iIl + iXProgr;

                rectF1 = new RectangleF(iY, iX, iYy, iXx);
                e.Graphics.DrawString(sTxt, Font, Brushes.Black, rectF1, format1);
                if (bBordi)
                    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                //iXProgr = iX;

                /*** DESCRIZIONE CLIENTE ***/

                a = _str003_Posz_CliDes.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                aFon = _str003_Font_CliDes.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                iX = iX + iIl + iXProgr;

                rectF1 = new RectangleF(iY, iX, iYy, iXx);
                e.Graphics.DrawString(_strCliDes, Font, Brushes.Black, rectF1, format1);
                if (bBordi)
                    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                //iXProgr = iX;

                /*** DESCRIZIONE ARTICOLO ***/

                aFon = _str004_Font_ArtDes.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);

                a = _str004_Posz_ArtDes.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);
                iX = iX + iIl + iXProgr;

                sTxt = _strArtDes;
                rectF1 = new RectangleF(iY, iX, iYy, iXx);
                e.Graphics.DrawString(sTxt, Font, Brushes.Black, rectF1, format1);
                if (bBordi)
                    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                //iXProgr += iX;


                /*** Descrizione ***/

                a = _str005_Posz_RicDes.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                aFon = _str005_Font_RicDes.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                iX = iX + iIl + iXProgr;

                rectF1 = new RectangleF(iY, iX, iYy, iXx);
                e.Graphics.DrawString("RIEPILOGO PRODOTTO", Font, Brushes.Black, rectF1, format1);
                if (bBordi)
                    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                //iXProgr = iX;

                iX = 150;

                /*** Preimballo label***/

                a = _str006_Posz_PriLbl.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                aFon = _str006_Font_PriLbl.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                iX = iX + iIl + iXProgr;

                //rectF1 = new RectangleF(iY, iX, iYy, iXx);
                //e.Graphics.DrawString("Preimballato il:", Font, Brushes.Black, rectF1, format1);
                e.Graphics.DrawString("Preimballato il:", Font, Brushes.Black, iY, iX);
                //if (bBordi)
                //    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));

                /*** Preimballo data ***/

                a = _str006_Posz_PriDta.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                aFon = _str006_Font_PriDta.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                iX = iX + iIl + iXProgr;

                e.Graphics.DrawString(DateTime.Today.ToString("dd/MM/yyy"), Font, Brushes.Black, iY, iX);



                /*** Colli label***/

                a = _str007_Posz_ColLbl.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                aFon = _str007_Font_ColLbl.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                iX = iX + iIl + iXProgr;

                //rectF1 = new RectangleF(iY, iX, iYy, iXx);
                //e.Graphics.DrawString("Preimballato il:", Font, Brushes.Black, rectF1, format1);
                e.Graphics.DrawString("Colli:", Font, Brushes.Black, iY, iX);
                //if (bBordi)
                //    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));

                /*** Colli qtà ***/

                a = _str007_Posz_ColQta.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                aFon = _str007_Font_ColQta.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                iX = iX + iIl + iXProgr;

                e.Graphics.DrawString(_decMovQta.ToString("#0"), Font, Brushes.Black, iY, iX);



                /*** Peso Netto label***/

                a = _str008_Posz_PneLbl.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                aFon = _str008_Font_PneLbl.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                iX = iX + iIl + iXProgr;

                //rectF1 = new RectangleF(iY, iX, iYy, iXx);
                //e.Graphics.DrawString("Preimballato il:", Font, Brushes.Black, rectF1, format1);
                e.Graphics.DrawString("Peso netto:", Font, Brushes.Black, iY, iX);
                //if (bBordi)
                //    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));

                /*** Peso Netto qtà ***/

                a = _str008_Posz_PneQta.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                aFon = _str008_Font_PneQta.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                iX = iX + iIl + iXProgr;

                e.Graphics.DrawString(_decMovQkg.ToString("#0.000") + " Kg", Font, Brushes.Black, iY, iX);


                /*** Prezzo al KG label***/

                a = _str009_Posz_PkgLbl.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                aFon = _str009_Font_PkgLbl.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                iX = iX + iIl + iXProgr;

                //rectF1 = new RectangleF(iY, iX, iYy, iXx);
                //e.Graphics.DrawString("Preimballato il:", Font, Brushes.Black, rectF1, format1);
                e.Graphics.DrawString("€/Kg:", Font, Brushes.Black, iY, iX);
                //if (bBordi)
                //    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));

                /*** Prezzo al KG prezzo ***/

                a = _str009_Posz_PkgVal.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                aFon = _str009_Font_PkgVal.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                iX = iX + iIl + iXProgr;

                e.Graphics.DrawString(_decMovPrv.ToString("#0.00") + " Kg", Font, Brushes.Black, iY, iX);



                /*** Prezzo al KG label***/

                a = _str010_Posz_ImpLbl.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                aFon = _str010_Font_ImpLbl.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                iX = iX + iIl + iXProgr;

                //rectF1 = new RectangleF(iY, iX, iYy, iXx);
                //e.Graphics.DrawString("Preimballato il:", Font, Brushes.Black, rectF1, format1);
                e.Graphics.DrawString("Importo:", Font, Brushes.Black, iY, iX);
                //if (bBordi)
                //    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));

                /*** Prezzo al KG prezzo ***/

                a = _str010_Posz_ImpVal.Split(',');
                iY = Convert.ToInt16(a[0]);
                iX = Convert.ToInt16(a[1]);
                iYy = Convert.ToInt16(a[2]);
                iXx = Convert.ToInt16(a[3]);
                iIl = Convert.ToInt16(a[4]);

                aFon = _str010_Font_ImpVal.Split(':');
                sFonName = aFon[0];
                fFonDim = (float)Convert.ToDouble(aFon[1]);
                fSty = FontStyle.Bold;
                if (aFon[2] == "0")
                    fSty = FontStyle.Regular;
                Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                iX = iX + iIl + iXProgr;

                e.Graphics.DrawString(_decMovImp.ToString("#0.00") + " €", Font, Brushes.Black, iY, iX);


















                //e.Graphics.DrawString("Preimballato il: ", Font, Brushes.Black, iY + 5, iX + 25);
                //e.Graphics.DrawString(DateTime.Today.ToString("dd/MM/yyy"), Font, Brushes.Black, iY + 100, iX + 26);

                //e.Graphics.DrawString("Colli", Font, Brushes.Black, iY + 5, iX + 50);
                //e.Graphics.DrawString(_decMovQta.ToString("#0"), Font, Brushes.Black, iY + 100, iX + 50);

                //e.Graphics.DrawString("Peso netto", Font, Brushes.Black, iY + 5, iX + 75);
                //e.Graphics.DrawString(_decMovQkg.ToString("#0.000") + " Kg", Font, Brushes.Black, iY + 100, iX + 75);

                //e.Graphics.DrawString("€/Kg", Font, Brushes.Black, iY + 5, iX + 100);
                //e.Graphics.DrawString(_decMovPrv.ToString("#0.00") + " Kg", Font, Brushes.Black, iY + 100, iX + 100);

                ////e.Graphics.DrawString("Da cons. entro il :", Font, Brushes.Black, iY + 5, iX + 100);

                //e.Graphics.DrawString("Importo", Font, Brushes.Black, iY + 5, iX + 125);
                //e.Graphics.DrawString(_decMovImp.ToString("#0.00") + " €", Font, Brushes.Black, iY + 100, iX + 125);

                ///*** Valori piede ***/

                //aFon = _str008_Font_FooVal.Split(':');
                //sFonName = aFon[0];
                //fFonDim = (float)Convert.ToDouble(aFon[1]);
                //fSty = FontStyle.Bold;
                //if (aFon[2] == "0")
                //    fSty = FontStyle.Regular;
                //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);






                ///*** Valori DATE ***/

                //aFon = _str010_Font_FooDay.Split(':');
                //sFonName = aFon[0];
                //fFonDim = (float)Convert.ToDouble(aFon[1]);
                //fSty = FontStyle.Bold;
                //if (aFon[2] == "0")
                //    fSty = FontStyle.Regular;
                //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);

            }
            catch (Exception ex)
            {
                e.Graphics.DrawString(s,
                    new Font("Arial", 40, FontStyle.Bold), Brushes.Black, 50, 125);

                string sMsg = ex.Message;
            }

            e.HasMorePages = true;
            if (intStampati >= _decEtiNum)
                e.HasMorePages = false;

        }

        private void stampantiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            intStampati = 0;

            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
                printDialog1.Document = printDocument1;
                printDialog1.ShowDialog();

                _str001_Print = printDocument1.PrinterSettings.PrinterName;

                FillFont("W");
            }
        }

        private void previewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            intStampati = 0;
            printPreviewDialog1.Document = this.printDocument1;
            printPreviewDialog1.ShowDialog();
        }

        private void stampaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            intStampati = 0;
            printDocument1.Print();
        }

        private void fontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //string s = "";
            //fontDialog1.ShowColor = true;
            ////fontDialog1.Font = _clsFun.Str2Font(new Font("Times New Roman", 7f), lbl.Text);
            ////fontDialog1.Color = lbl.ForeColor;
            //if (fontDialog1.ShowDialog() != DialogResult.Cancel)
            //{
            //    object[] name = new object[] { fontDialog1.Font.FontFamily.Name, ":", fontDialog1.Font.Size, ":", (int)fontDialog1.Font.Style, ":", null };
            //    Color color = fontDialog1.Color;
            //    name[6] = color.Name.ToString();
            //    s = string.Concat(name);

            //    _str003_Font_01 = s;

            //    FillFont("W");
            //}
        }

        private string FillFont(string strTip)
        {
            string s = "";

            string sFil = Path.GetDirectoryName(_clsDef.FILEINI) + "\\EtiNomeFont2.ini";

            if (strTip == "R")
            {
                if (File.Exists(sFil))
                {
                    try
                    {
                        using (StreamReader sr = new StreamReader(sFil))
                        {
                            string sRig = "";

                            while ((sRig = sr.ReadLine()) != null)
                            {
                                if (sRig.Length > 4)
                                {
                                    if (sRig.Substring(0, 3) == "001")
                                        _str001_Print = sRig.Substring(4);
                                    else if (sRig.Substring(0, 3) == "002")
                                        _str002_Dim = sRig.Substring(4);
                                    else if (sRig.Substring(0, 3) == "003")
                                    {
                                        string[] a = sRig.Substring(4).Split(';');
                                        _str003_Posz_NegRag = a[0];
                                        _str003_Font_NegRag = a[1];
                                        _str003_Posz_NegInd = a[2];
                                        _str003_Font_NegInd = a[3];
                                        _str003_Posz_CliDes = a[4];
                                        _str003_Font_CliDes = a[5];
                                    }
                                    else if (sRig.Substring(0, 3) == "004")
                                    {
                                        string[] a = sRig.Substring(4).Split(';');
                                        _str004_Posz_ArtDes = a[0];
                                        _str004_Font_ArtDes = a[1];
                                    }
                                    else if (sRig.Substring(0, 3) == "005")
                                    {
                                        string[] a = sRig.Substring(4).Split(';');
                                        _str005_Posz_RicDes = a[0];
                                        _str005_Font_RicDes = a[1];
                                    }
                                    else if (sRig.Substring(0, 3) == "006")
                                    {
                                        string[] a = sRig.Substring(4).Split(';');
                                        _str006_Posz_PriLbl = a[0];
                                        _str006_Font_PriLbl = a[1];
                                        _str006_Posz_PriDta = a[2];
                                        _str006_Font_PriDta = a[3];
                                    }
                                    else if (sRig.Substring(0, 3) == "007")
                                    {
                                        string[] a = sRig.Substring(4).Split(';');
                                        _str007_Posz_ColLbl = a[0];
                                        _str007_Font_ColLbl = a[1];
                                        _str007_Posz_ColQta = a[2];
                                        _str007_Font_ColQta = a[3];
                                    }
                                    else if (sRig.Substring(0, 3) == "008")
                                    {
                                        string[] a = sRig.Substring(4).Split(';');
                                        _str008_Posz_PneLbl = a[0];
                                        _str008_Font_PneLbl = a[1];
                                        _str008_Posz_PneQta = a[2];
                                        _str008_Font_PneQta = a[3];
                                    }
                                    else if (sRig.Substring(0, 3) == "009")
                                    {
                                        string[] a = sRig.Substring(4).Split(';');
                                        _str009_Posz_PkgLbl = a[0];
                                        _str009_Font_PkgLbl = a[1];
                                        _str009_Posz_PkgVal = a[2];
                                        _str009_Font_PkgVal = a[3];
                                    }
                                    else if (sRig.Substring(0, 3) == "010")
                                    {
                                        string[] a = sRig.Substring(4).Split(';');
                                        _str010_Posz_ImpLbl = a[0];
                                        _str010_Font_ImpLbl = a[1];
                                        _str010_Posz_ImpVal = a[2];
                                        _str010_Font_ImpVal = a[3];
                                    }
                                    else if (sRig.Substring(0, 3) == "011")
                                        _str011_PrnImmediata = sRig.Substring(4);

                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("The file could not be read:");
                        Console.WriteLine(e.Message);
                    }
                }

                //if (_str003_Font_01 == "")
                //{
                //    _str003_Font_01 = "Arial:12:1:ControlText";
                //    _str003_Font_02 = "Arial:12:1:ControlText";
                //}
            }
            //else
            //{
            //    s = "";

            //    StreamWriter sw = new StreamWriter(sFil, false);
            //    s = "001 " + _str01_Font;
            //    sw.Write(s + "\r\n");
            //    s = "002 " + _str02_Print;
            //    sw.Write(s + "\r\n");
            //    s = "003 " + _str03_Dim;
            //    sw.Write(s + "\r\n");
            //    s = "004 " + _str04_XY_1;
            //    sw.Write(s + "\r\n");
            //    s = "005 " + _str05_XY_2;
            //    sw.Write(s + "\r\n");
            //    ((TextWriter)sw).Flush();
            //    sw.Close();
            //    sw.Dispose();
            //}

            //_strFont = s;

            return s;
        }



    }
}
