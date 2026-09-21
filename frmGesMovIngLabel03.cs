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
    public partial class frmGesMovIngLabel03 : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        //public DataTable _tabTab = new DataTable();
        public decimal _decEtiNum = 0;

        private string _str001_Print = "";
        private string _str002_Dim = "";

        //private string _str003_Posz_NegRag = "";
        //private string _str003_Font_NegRag = "";
        //private string _str003_Posz_NegInd = "";
        //private string _str003_Font_NegInd = "";
        //private string _str003_Posz_RegUls = "";
        //private string _str003_Font_RegUls = "";

        private string _str004_Posz_ArtDes = "";
        private string _str004_Font_ArtDes = "";

        private string _str005_Posz_DaySca = "";
        private string _str005_Font_DaySca = "";

        private string _str006_Posz_ArtPxc = "";
        private string _str006_Font_ArtPxc = "";

        private string _str007_Posz_CodLot = "";
        private string _str007_Font_CodLot = "";
        //private string _str007_Posz_DesLot = "";
        //private string _str007_Font_DesLot = "";

        //private string _str007_Posz_PriLbl = "";
        //private string _str007_Font_PriLbl = "";
        //private string _str007_Posz_PriDta = "";
        //private string _str007_Font_PriDta = "";

        //private string _str008_Posz_PneLbl = "";
        //private string _str008_Font_PneLbl = "";
        //private string _str008_Posz_PneQta = "";
        //private string _str008_Font_PneQta = "";

        //private string _str009_Posz_ScaLbl = "";
        //private string _str009_Font_ScaLbl = "";
        //private string _str009_Posz_ScaDay = "";
        //private string _str009_Font_ScaDay = "";

        //private string _str010_Posz_PrkLbl = "";
        //private string _str010_Font_PrkLbl = "";
        //private string _str010_Posz_PrkKgr = "";
        //private string _str010_Font_PrkKgr = "";

        //private string _str011_Posz_ImpLbl = "";
        //private string _str011_Font_ImpLbl = "";
        //private string _str011_Posz_ImpVal = "";
        //private string _str011_Font_ImpVal = "";

        private string _str012_Posz_LogUno = "";
        private string _str012_Posz_LogDue = "";

        private string _str013_Posz_EanCod = "";
        private string _str013_Font_EanCod = "";

        //private string _str010_Posz_FooDay = "";
        //private string _str010_Font_FooDay = "";

        private string _str099_PrnImmediata = "";

        public string _strConSql = "";
        public string _strEtiTipo = "";                 //Se "" -> Rossetto
        public string _strCliCod = "";
        public string _strCliDes = "";
        public string _strCliInd = "";
        public string _strCliPar = "";
        public string _strArtCod = "";
        public string _strArtDes = "";
        public string _strLotCod = "";
        public string _strArtEan = "";
        public string _strArtUmi = "";
        public decimal _decMovQkg = 0;
        public decimal _decMovTar = 0;
        public decimal _decMovPrv = 0;                  //Prezzo al kilo
        public decimal _decMovImp = 0;                  //Importo
        public DateTime _dayDaySca = DateTime.Today;    //Data scadenza
        public string _strArtPxc = "";                  //Pezzi per confezione

        private string _strForArt = "";                 //Codice articolo fornitore

        //public string _strLotNas = "";
        //public string _strLotAll = "";
        //public string _strLotMac = "";
        //public string _strLotSez = "";

        //public string _strEan = "";
        //public string _strPes = "";
        //public string _strTar = "";
        //public string _strPrv = "";
        //public string _strImp = "";
        //public string _strDsc = "";     //Data scadenza

        private int intStampati = 0;

        private DataTable _tabLot = new DataTable();

        public frmGesMovIngLabel03()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmAnaArtIngLabel_Load(object sender, EventArgs e)
        {
            //Barcode _ean128 = new BarcodeLib.Barcode();
            //BarcodeLib.TYPE tpEan128 = BarcodeLib.TYPE.CODE128;
            //_ean128.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            Boolean bOk = true;

            string s = "";
            if (_strLotCod != "" && _strLotCod.Length > 2)
            {
                string[] a = _strLotCod.Split('-');

                if (a.Length >= 3)
                {
                    //s = _strLotCod.Substring(2).Trim();

                    string sTip = a[0];
                    string sCod = a[1];
                    string sYea = a[2];

                    s = "SELECT * FROM GesDocLotti WHERE lot_yea='" + sYea + "' AND lot_cod='" + sCod + "' AND lot_tip='" + sTip + "' ORDER BY lot_ddo DESC";
                    _tabLot = _clsFun.FillTabSql("GesDocLotti", s, true, _strConSql);

                    if(_tabLot.Rows.Count > 0)
                    {
                        string sFor = (string)_tabLot.Rows[0]["lot_for"];

                        s = "SELECT lia_arf FROM GesLisAcquisto WHERE lia_art='" + _strArtCod + "' AND lia_for='" + sFor + "' AND lia_ann=0";
                        DataTable t = _clsFun.FillTabSql("", s,true, _strConSql);
                        if(t.Rows.Count > 0)
                            _strForArt = ((string)t.Rows[0]["lia_arf"]).Trim();
                    }

                    Console.WriteLine("2222");
                }
            }

            if (_strForArt == "")
            {
                MessageBox.Show("Codice articoli del fornitore mancante (verificare il lotto)!", "CONTROLLO ETICHETTA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bOk = false;
            }
            FillFont("R");

            if (_str001_Print == "")
            {
                MessageBox.Show("Stampante non definita!", "CONTROLLO ETICHETTA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bOk = false;
            }

            if(!bOk)
                Esci();
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

                if (_str099_PrnImmediata == "S")
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

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            string s = "";
            string sTxt = "";

            Boolean bBordi = false;

            Barcode _ean128 = new BarcodeLib.Barcode();
            BarcodeLib.TYPE tpEan128 = BarcodeLib.TYPE.CODE128;
            _ean128.Alignment = BarcodeLib.AlignmentPositions.CENTER;

            //Barcode _ean08 = new BarcodeLib.Barcode();
            //BarcodeLib.TYPE tpEan08 = BarcodeLib.TYPE.EAN8;
            //_ean08.Alignment = BarcodeLib.AlignmentPositions.CENTER;

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
                Boolean bCliNom = false;
                Boolean bCliInd = false;
                Boolean bCliPrv = false;
                Boolean bCliLot = false;

                if (_strCliPar.Length > 0 && _strCliPar.Substring(0, 1) == "S")
                    bCliNom = true;
                if (_strCliPar.Length > 1 && _strCliPar.Substring(1, 1) == "S")
                    bCliInd = true;
                if (_strCliPar.Length > 2 && _strCliPar.Substring(2, 1) == "S")
                    bCliPrv = true;
                if (_strCliPar.Length > 3 && _strCliPar.Substring(3, 1) == "S")
                    bCliLot = true;

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

                if (_strEtiTipo == "02") 
                {
                    RectangleF rectF1 = new RectangleF(5, -290, 210, 280);
                    //e.Graphics.DrawString(sTxt, new Font("Arial", 20f, FontStyle.Bold), Brushes.Black, rectF1, format1);
                    if (bBordi)
                        e.Graphics.DrawRectangle(new Pen(Color.Black, 5), Rectangle.Round(rectF1));

                    string[] aFon; // = _str012_Posz_LogUno.Split(':');
                    string sFonName = ""; // aFon[0];
                    float fFonDim = (float)Convert.ToDouble(0);
                    FontStyle fSty = FontStyle.Bold;
                    //if (aFon[2] == "0")
                    //    fSty = FontStyle.Regular;
                    Font Font; // = new System.Drawing.Font(sFonName, fFonDim, fSty);

                    /** LOGO **/

                    string[] a = _str012_Posz_LogUno.Split(',');
                    iY = Convert.ToInt16(a[0]);
                    iX = Convert.ToInt16(a[1]);
                    iYy = Convert.ToInt16(a[2]);
                    iXx = Convert.ToInt16(a[3]);

                    string sLog = SeekLogo("LOG");
                    string sBol = SeekLogo("BOL");

                    if(sLog != "")
                    {
                        e.Graphics.DrawImage(Image.FromFile(sLog), iY, iX, iYy, iXx);

                        if(File.Exists(sBol))
                        {
                            a = _str012_Posz_LogDue.Split(',');
                            iY = Convert.ToInt16(a[0]);
                            iX = Convert.ToInt16(a[1]);
                            iYy = Convert.ToInt16(a[2]);
                            iXx = Convert.ToInt16(a[3]);

                            e.Graphics.DrawImage(Image.FromFile(sBol), iY, iX, iYy, iXx);
                        }
                    }

                    if (bBordi)
                        e.Graphics.DrawRectangle(new Pen(Color.Black, 5), Rectangle.Round(rectF1));

                    //string[] aFon; // = _str003_Font_NegRag.Split(':');
                    //string sFonName = "";
                    //float fFonDim = 8; // (float)Convert.ToDouble(aFon[1]);
                    //FontStyle fSty = FontStyle.Bold;
                    ////if (aFon[2] == "0")
                    ////    fSty = FontStyle.Regular;
                    //Font Font = new System.Drawing.Font("Courier new", fFonDim);

                    /*** 
                     * RAGIONE SOCIALE 
                     * ***/

                    //sTxt = "";

                    //if (_str003_Posz_NegRag.Length > 0 && _str003_Posz_NegRag.Substring(0, 1) != "N")
                    //{

                    //    if (bCliNom)
                    //        sTxt = _strCliDes;
                    //    else
                    //        sTxt = (string)tCnf.Rows[0]["cnf_rag"];

                    //    aFon = _str003_Font_NegRag.Split(':');
                    //    sFonName = aFon[0];
                    //    fFonDim = (float)Convert.ToDouble(aFon[1]);
                    //    fSty = FontStyle.Bold;
                    //    if (aFon[2] == "0")
                    //        fSty = FontStyle.Regular;
                    //    Font = new System.Drawing.Font(sFonName, fFonDim, fSty);

                    //    a = _str003_Posz_NegRag.Split(',');
                    //    iY = Convert.ToInt16(a[0]);
                    //    iX = Convert.ToInt16(a[1]);
                    //    iYy = Convert.ToInt16(a[2]);
                    //    iXx = Convert.ToInt16(a[3]);
                    //    iIl = Convert.ToInt16(a[4]);

                    //    rectF1 = new RectangleF(iY, iX, iYy, iXx);
                    //    e.Graphics.DrawString(sTxt, Font, Brushes.Black, rectF1, format1);
                    //    if (bBordi)
                    //        e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                    //}

                    /*** 
                     * DATI NEGOZIO 
                     * ***/

                    //if (_str003_Posz_NegInd.Length > 0 && _str003_Posz_NegInd.Substring(0, 1) != "N")
                    //{
                    //    sTxt = "";
                    //    if (bCliInd)
                    //        sTxt = _strCliInd;
                    //    else
                    //    {
                    //        sTxt = (string)tCnf.Rows[0]["cnf_ind"] + "-";
                    //        sTxt += (string)tCnf.Rows[0]["cnf_cap"] + " ";
                    //        sTxt += (string)tCnf.Rows[0]["cnf_loc"] + _clsDef.CRLF;
                    //    }

                    //    a = _str003_Posz_NegInd.Split(',');
                    //    iY = Convert.ToInt16(a[0]);
                    //    iX = Convert.ToInt16(a[1]);
                    //    iYy = Convert.ToInt16(a[2]);
                    //    iXx = Convert.ToInt16(a[3]);
                    //    iIl = Convert.ToInt16(a[4]);

                    //    aFon = _str003_Font_NegInd.Split(':');
                    //    sFonName = aFon[0];
                    //    fFonDim = (float)Convert.ToDouble(aFon[1]);
                    //    fSty = FontStyle.Bold;
                    //    if (aFon[2] == "0")
                    //        fSty = FontStyle.Regular;
                    //    Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                    //    //iX = iX + iIl + iXProgr;

                    //    rectF1 = new RectangleF(iY, iX, iYy, iXx);
                    //    e.Graphics.DrawString(sTxt, Font, Brushes.Black, rectF1, format1);
                    //    if (bBordi)
                    //        e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                    //    iXProgr = iX;
                    //}

                    /*** 
                     * Registrazione ULSS 
                     * ***/
                    //if (_str003_Posz_RegUls.Length > 0 && _str003_Posz_RegUls.Substring(0, 1) != "N")
                    //{

                    //    sTxt = (string)tCnf.Rows[0]["cnf_v01"] + " ";

                    //    a = _str003_Posz_RegUls.Split(',');
                    //    iY = Convert.ToInt16(a[0]);
                    //    iX = Convert.ToInt16(a[1]);
                    //    iYy = Convert.ToInt16(a[2]);
                    //    iXx = Convert.ToInt16(a[3]);
                    //    iIl = Convert.ToInt16(a[4]);

                    //    aFon = _str003_Font_RegUls.Split(':');
                    //    sFonName = aFon[0];
                    //    fFonDim = (float)Convert.ToDouble(aFon[1]);
                    //    fSty = FontStyle.Bold;
                    //    if (aFon[2] == "0")
                    //        fSty = FontStyle.Regular;
                    //    Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                    //    //iX = iX + iIl + iXProgr;

                    //    rectF1 = new RectangleF(iY, iX, iYy, iXx);
                    //    e.Graphics.DrawString(sTxt, Font, Brushes.Black, rectF1, format1);
                    //    if (bBordi)
                    //        e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                    //    iXProgr = iX;
                    //}

                    /*** 
                     * DESCRIZIONE ARTICOLO 
                     * ***/

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
                    //iX = iX + iIl + iXProgr;

                    sTxt = _strArtDes;
                    rectF1 = new RectangleF(iY, iX, iYy, iXx);
                    e.Graphics.DrawString(sTxt, Font, Brushes.Black, rectF1, format1);
                    if (bBordi)
                        e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                    iXProgr += iXx;

                    /*** 
                     * SCADENZA 
                     * ***/

                    aFon = _str005_Font_DaySca.Split(':');
                    sFonName = aFon[0];
                    fFonDim = (float)Convert.ToDouble(aFon[1]);
                    fSty = FontStyle.Bold;
                    if (aFon[2] == "0")
                        fSty = FontStyle.Regular;
                    Font = new System.Drawing.Font(sFonName, fFonDim, fSty);

                    a = _str005_Posz_DaySca.Split(',');
                    iY = Convert.ToInt16(a[0]);
                    iX = Convert.ToInt16(a[1]);
                    iYy = Convert.ToInt16(a[2]);
                    iXx = Convert.ToInt16(a[3]);
                    iIl = Convert.ToInt16(a[4]);
                    //iX = iX + iIl + iXProgr;

                    sTxt = "Scadenza " + _dayDaySca.ToString("dd/MM/yyyy");
                    rectF1 = new RectangleF(iY, iX, iYy, iXx);
                    e.Graphics.DrawString(sTxt, Font, Brushes.Black, rectF1, format1);
                    if (bBordi)
                        e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                    iXProgr += iXx;


                    /*** 
                     * PEZZI X COLLO 
                     * ***/

                    aFon = _str006_Font_ArtPxc.Split(':');
                    sFonName = aFon[0];
                    fFonDim = (float)Convert.ToDouble(aFon[1]);
                    fSty = FontStyle.Bold;
                    if (aFon[2] == "0")
                        fSty = FontStyle.Regular;
                    Font = new System.Drawing.Font(sFonName, fFonDim, fSty);

                    a = _str006_Posz_ArtPxc.Split(',');
                    iY = Convert.ToInt16(a[0]);
                    iX = Convert.ToInt16(a[1]);
                    iYy = Convert.ToInt16(a[2]);
                    iXx = Convert.ToInt16(a[3]);
                    iIl = Convert.ToInt16(a[4]);
                    //iX = iX + iIl + iXProgr;

                    sTxt = "PEZZI " + _strArtPxc;
                    rectF1 = new RectangleF(iY, iX, iYy, iXx);
                    e.Graphics.DrawString(sTxt, Font, Brushes.Black, rectF1, format1);
                    if (bBordi)
                        e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                    iXProgr += iXx;




                    /*** 
                     * RICETTA 
                     * ***/

                    //string sFil = _clsDef.PATHINGREDIENTI + "et01_" + _strArtCod + ".rtf";
                    //RichTextBox rtxBox = new RichTextBox();

                    //if (File.Exists(sFil))
                    //    rtxBox.LoadFile(sFil);

                    //if (_tabLot.Rows.Count > 0)
                    //{
                    //    //rtxBox.Text += _clsDef.CRLF;

                    //    string sAur = ((string)_tabLot.Rows[0]["lot_lot"]).Trim();          //Auricolare
                    //    string sNat = ((string)_tabLot.Rows[0]["lot_nat"]).Trim();          //Nato nazione
                    //    string sAll = ((string)_tabLot.Rows[0]["lot_all"]).Trim();
                    //    string sMna = ((string)_tabLot.Rows[0]["lot_mna"]).Trim();          //Macellazione nazione
                    //    string sMac = ((string)_tabLot.Rows[0]["lot_mac"]).Trim();          //Bollo macellazione     
                    //    string sSez = ((string)_tabLot.Rows[0]["lot_sez"]).Trim();          //Sezionato nazione
                    //    string sSbo = ((string)_tabLot.Rows[0]["lot_sbo"]).Trim();          //Sezionato bollo

                    //    string sMsg = "";

                    //    if (sAur != "")
                    //        sMsg += "AURICOLARE: " + sAur + _clsDef.CRLF;
                    //    //if (sNat != "")
                    //    //    sMsg += "NAZIONALITA'/INGRASSO: " + sNat + _clsDef.CRLF;
                    //    if (sNat != "")
                    //        sMsg += "NATO: " + sNat + _clsDef.CRLF;
                    //    if (sAll != "")
                    //        sMsg += "ALLEVATO: " + sAll + _clsDef.CRLF;
                    //    if (sMna != "")
                    //        sMsg += "MACELLAZIONE: " + sMna + " " + sMac + _clsDef.CRLF;
                    //    if (sSez != "")
                    //        sMsg += "SEZIONATO: " + sSez + " " + sSbo + _clsDef.CRLF;

                    //    rtxBox.Text = sMsg + _clsDef.CRLF + rtxBox.Text;
                    //}

                    //if(rtxBox.Text.Length > 0)
                    //{
                    //    aFon = _str005_Font_RicDes.Split(':');
                    //    sFonName = aFon[0];
                    //    fFonDim = (float)Convert.ToDouble(aFon[1]);
                    //    fSty = FontStyle.Bold;
                    //    if (aFon[2] == "0")
                    //        fSty = FontStyle.Regular;
                    //    Font = new System.Drawing.Font(sFonName, fFonDim, fSty);


                    //    a = _str005_Posz_RicDes.Split(',');
                    //    iY = Convert.ToInt16(a[0]);
                    //    iX = Convert.ToInt16(a[1]);
                    //    iYy = Convert.ToInt16(a[2]);
                    //    iXx = Convert.ToInt16(a[3]);
                    //    iIl = Convert.ToInt16(a[4]);
                    //    //iX = iX + iIl + iXProgr;

                    //    StringFormat format2 = new StringFormat(StringFormatFlags.NoClip);
                    //    format2.LineAlignment = StringAlignment.Center;
                    //    //format2.Alignment = StringAlignment.Center;
                    //    ////format1.FormatFlags = StringFormatFlags.DirectionRightToLeft;
                    //    ////format1.FormatFlags = StringFormatFlags.DirectionVertical | StringFormatFlags.DirectionRightToLeft;
                    //    ////format1.FormatFlags = StringFormatFlags.DirectionVertical;

                    //    rectF1 = new RectangleF(iY, iX, iYy, iXx);
                    //    e.Graphics.DrawString(rtxBox.Text, Font, Brushes.Black, rectF1, format2);
                    //    if (bBordi)
                    //        e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
                    //    //iXProgr += iX;
                    //}

                    /*** 
                     * LOTTO 
                     * ***/

                    aFon = _str007_Font_CodLot.Split(':');
                    sFonName = aFon[0];
                    fFonDim = (float)Convert.ToDouble(aFon[1]);
                    fSty = FontStyle.Bold;
                    if (aFon[2] == "0")
                        fSty = FontStyle.Regular;
                    Font = new System.Drawing.Font(sFonName, fFonDim, fSty);

                    a = _str007_Posz_CodLot.Split(',');
                    iY = Convert.ToInt16(a[0]);
                    iX = Convert.ToInt16(a[1]);
                    iYy = Convert.ToInt16(a[2]);
                    iXx = Convert.ToInt16(a[3]);
                    iIl = Convert.ToInt16(a[4]);

                    //s = "";
                    //a = _strLotCod.Split('-');
                    //if (a.Length > 2)
                    //{
                    //    s = a[0] + " ";
                    //    if(_clsFun.Numerico(a[1], "0123456789"))
                    //        s += Convert.ToInt16(a[1]).ToString();
                    //    else
                    //        s += a[1];
                    //}
                    //else if(a.Length > 1)
                    //{
                    //    s = a[0];
                    //}

                    s = "";
                    a = _strLotCod.Split('-');

                    if (bCliLot && _tabLot.Rows.Count > 0)
                        s = ((string)_tabLot.Rows[0]["lot_lot"]).Trim();          //Auricolare
                    else
                    {
                        if (a.Length > 1)
                        {
                            s = "00" + a[1];
                            //if (_clsFun.Numerico(a[1], "0123456789"))
                            //    s += Convert.ToInt16(a[1]).ToString();
                            //else
                            //    s += a[1];
                        }
                        else if (a.Length > 0)
                            s = a[0];
                        else
                            s = a[0];
                    }
                    string sLot = s;

                    e.Graphics.DrawString("LOTTO N. " + sLot, Font, Brushes.Black, iY, iX);

                    a = _str002_Dim.Split(',');
                    e.Graphics.DrawLine(new Pen(Brushes.Black, 2), iY - 5, iX + 15, Convert.ToInt16(a[1]) - 5, iX + 15);


                    /** Descrizione LOTTO **/
                    //if (_tabLot != null && _tabLot.Rows.Count > 0)
                    //{
                    //    aFon = _str006_Font_DesLot.Split(':');
                    //    sFonName = aFon[0];
                    //    fFonDim = (float)Convert.ToDouble(aFon[1]);
                    //    fSty = FontStyle.Bold;
                    //    if (aFon[2] == "0")
                    //        fSty = FontStyle.Regular;
                    //    Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                    //    a = _str006_Posz_DesLot.Split(',');
                    //    iY = Convert.ToInt16(a[0]);
                    //    iX = Convert.ToInt16(a[1]);
                    //    iYy = Convert.ToInt16(a[2]);
                    //    iXx = Convert.ToInt16(a[3]);
                    //    iIl = Convert.ToInt16(a[4]);
                    //    e.Graphics.DrawString("Auricolare", Font, Brushes.Black, iY, iX);
                    //    iY += iIl;
                    //    e.Graphics.DrawString("Nazionalità/ingrasso", Font, Brushes.Black, iY, iX);
                    //    iY += iIl;
                    //    e.Graphics.DrawString("Mecellazione", Font, Brushes.Black, iY, iX);
                    //    iY += iIl;
                    //    e.Graphics.DrawString("Sezionato", Font, Brushes.Black, iY, iX);
                    //}

                    /*** Preimballo label***/

                    //a = _str007_Posz_PriLbl.Split(',');
                    //iY = Convert.ToInt16(a[0]);
                    //iX = Convert.ToInt16(a[1]);
                    //iYy = Convert.ToInt16(a[2]);
                    //iXx = Convert.ToInt16(a[3]);
                    //iIl = Convert.ToInt16(a[4]);

                    //aFon = _str007_Font_PriLbl.Split(':');
                    //sFonName = aFon[0];
                    //fFonDim = (float)Convert.ToDouble(aFon[1]);
                    //fSty = FontStyle.Bold;
                    //if (aFon[2] == "0")
                    //    fSty = FontStyle.Regular;
                    //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                    ////iX = iX + iIl + iXProgr;

                    //rectF1 = new RectangleF(iY, iX, iYy, iXx);
                    //e.Graphics.DrawString("Preimballato il:", Font, Brushes.Black, rectF1, format1);
                    //e.Graphics.DrawString("Preimballato il:", Font, Brushes.Black, iY, iX); // Seck 20180323 tolto
                    //if (bBordi)
                    //    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));

                    /*** Preimballo data ***/

                    //a = _str007_Posz_PriDta.Split(',');
                    //iY = Convert.ToInt16(a[0]);
                    //iX = Convert.ToInt16(a[1]);
                    //iYy = Convert.ToInt16(a[2]);
                    //iXx = Convert.ToInt16(a[3]);
                    //iIl = Convert.ToInt16(a[4]);

                    //aFon = _str007_Font_PriDta.Split(':');
                    //sFonName = aFon[0];
                    //fFonDim = (float)Convert.ToDouble(aFon[1]);
                    //fSty = FontStyle.Bold;
                    //if (aFon[2] == "0")
                    //    fSty = FontStyle.Regular;
                    //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                    ////iX = iX + iIl + iXProgr;

                    //e.Graphics.DrawString(DateTime.Today.ToString("dd/MM/yyy"), Font, Brushes.Black, iY, iX); // Seck 20180323 tolto




                    /*** Peso Netto label***/

                    //a = _str008_Posz_PneLbl.Split(',');
                    //iY = Convert.ToInt16(a[0]);
                    //iX = Convert.ToInt16(a[1]);
                    //iYy = Convert.ToInt16(a[2]);
                    //iXx = Convert.ToInt16(a[3]);
                    //iIl = Convert.ToInt16(a[4]);

                    //aFon = _str008_Font_PneLbl.Split(':');
                    //sFonName = aFon[0];
                    //fFonDim = (float)Convert.ToDouble(aFon[1]);
                    //fSty = FontStyle.Bold;
                    //if (aFon[2] == "0")
                    //    fSty = FontStyle.Regular;
                    //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                    ////iX = iX + iIl + iXProgr;

                    ////rectF1 = new RectangleF(iY, iX, iYy, iXx);
                    ////e.Graphics.DrawString("Preimballato il:", Font, Brushes.Black, rectF1, format1);

                    //s = "Peso netto:";
                    //if (_strArtUmi == "NR")
                    //    s = "Peso:";
                    //e.Graphics.DrawString(s, Font, Brushes.Black, iY, iX);

                    //if (bBordi)
                    //    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));

                    /*** Peso Netto qtà ***/

                    //a = _str008_Posz_PneQta.Split(',');
                    //iY = Convert.ToInt16(a[0]);
                    //iX = Convert.ToInt16(a[1]);
                    //iYy = Convert.ToInt16(a[2]);
                    //iXx = Convert.ToInt16(a[3]);
                    //iIl = Convert.ToInt16(a[4]);

                    //aFon = _str008_Font_PneQta.Split(':');
                    //sFonName = aFon[0];
                    //fFonDim = (float)Convert.ToDouble(aFon[1]);
                    //fSty = FontStyle.Bold;
                    //if (aFon[2] == "0")
                    //    fSty = FontStyle.Regular;
                    //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                    ////iX = iX + iIl + iXProgr;

                    //s = " Kg";
                    //if (_strArtUmi == "NR")
                    //    s = " Kg circa";

                    //e.Graphics.DrawString(_decMovQkg.ToString("#0.000") + s, Font, Brushes.Black, iY, iX);

                    /*** Da consumarsi entro label ***/

                    //a = _str009_Posz_ScaLbl.Split(',');
                    //iY = Convert.ToInt16(a[0]);
                    //iX = Convert.ToInt16(a[1]);
                    //iYy = Convert.ToInt16(a[2]);
                    //iXx = Convert.ToInt16(a[3]);
                    //iIl = Convert.ToInt16(a[4]);

                    //aFon = _str009_Font_ScaLbl.Split(':');
                    //sFonName = aFon[0];
                    //fFonDim = (float)Convert.ToDouble(aFon[1]);
                    //fSty = FontStyle.Bold;
                    //if (aFon[2] == "0")
                    //    fSty = FontStyle.Regular;
                    //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                    //iX = iX + iIl + iXProgr;

                    //rectF1 = new RectangleF(iY, iX, iYy, iXx);
                    //e.Graphics.DrawString("Preimballato il:", Font, Brushes.Black, rectF1, format1);
                    //e.Graphics.DrawString("Da cons. entro il:", Font, Brushes.Black, iY, iX);
                    //if (bBordi)
                    //    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));

                    /*** Da consumarsi entro data ***/

                    //a = _str009_Posz_ScaDay.Split(',');
                    //iY = Convert.ToInt16(a[0]);
                    //iX = Convert.ToInt16(a[1]);
                    //iYy = Convert.ToInt16(a[2]);
                    //iXx = Convert.ToInt16(a[3]);
                    //iIl = Convert.ToInt16(a[4]);

                    //aFon = _str009_Font_ScaDay.Split(':');
                    //sFonName = aFon[0];
                    //fFonDim = (float)Convert.ToDouble(aFon[1]);
                    //fSty = FontStyle.Bold;
                    //if (aFon[2] == "0")
                    //    fSty = FontStyle.Regular;
                    //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                    //////iX = iX + iIl + iXProgr;

                    ////e.Graphics.DrawString(_dayArtSca.ToString("dd/MM/yyy"), Font, Brushes.Black, iY, iX);
                    //e.Graphics.DrawString(_strDaySca, Font, Brushes.Black, iY, iX);

                    //if (bCliPrv)
                    //{
                        /*** Prezzo al Kg label ***/

                        //a = _str010_Posz_PrkLbl.Split(',');
                        //iY = Convert.ToInt16(a[0]);
                        //iX = Convert.ToInt16(a[1]);
                        //iYy = Convert.ToInt16(a[2]);
                        //iXx = Convert.ToInt16(a[3]);
                        //iIl = Convert.ToInt16(a[4]);

                        //aFon = _str010_Font_PrkLbl.Split(':');
                        //sFonName = aFon[0];
                        //fFonDim = (float)Convert.ToDouble(aFon[1]);
                        //fSty = FontStyle.Bold;
                        //if (aFon[2] == "0")
                        //    fSty = FontStyle.Regular;
                        //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                        ////iX = iX + iIl + iXProgr;

                        //e.Graphics.DrawString("€/Kg: ", Font, Brushes.Black, iY, iX);


                        ///*** Prezzo al Kg valore ***/

                        //a = _str010_Posz_PrkKgr.Split(',');
                        //iY = Convert.ToInt16(a[0]);
                        //iX = Convert.ToInt16(a[1]);
                        //iYy = Convert.ToInt16(a[2]);
                        //iXx = Convert.ToInt16(a[3]);
                        //iIl = Convert.ToInt16(a[4]);

                        //aFon = _str010_Font_PrkKgr.Split(':');
                        //sFonName = aFon[0];
                        //fFonDim = (float)Convert.ToDouble(aFon[1]);
                        //fSty = FontStyle.Bold;
                        //if (aFon[2] == "0")
                        //    fSty = FontStyle.Regular;
                        //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                        ////iX = iX + iIl + iXProgr;

                        //e.Graphics.DrawString(_decMovPrv.ToString("#0.00"), Font, Brushes.Black, iY, iX);


                        ///*** Prezzo al Kg label ***/

                        //a = _str011_Posz_ImpLbl.Split(',');
                        //iY = Convert.ToInt16(a[0]);
                        //iX = Convert.ToInt16(a[1]);
                        //iYy = Convert.ToInt16(a[2]);
                        //iXx = Convert.ToInt16(a[3]);
                        //iIl = Convert.ToInt16(a[4]);

                        //aFon = _str011_Font_ImpLbl.Split(':');
                        //sFonName = aFon[0];
                        //fFonDim = (float)Convert.ToDouble(aFon[1]);
                        //fSty = FontStyle.Bold;
                        //if (aFon[2] == "0")
                        //    fSty = FontStyle.Regular;
                        //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                        ////iX = iX + iIl + iXProgr;

                        //e.Graphics.DrawString("Prezzo", Font, Brushes.Black, iY, iX);


                        /*** Prezzo al Kg valore ***/

                        //a = _str011_Posz_ImpVal.Split(',');
                        //iY = Convert.ToInt16(a[0]);
                        //iX = Convert.ToInt16(a[1]);
                        //iYy = Convert.ToInt16(a[2]);
                        //iXx = Convert.ToInt16(a[3]);
                        //iIl = Convert.ToInt16(a[4]);

                        //aFon = _str011_Font_ImpVal.Split(':');
                        //sFonName = aFon[0];
                        //fFonDim = (float)Convert.ToDouble(aFon[1]);
                        //fSty = FontStyle.Bold;
                        //if (aFon[2] == "0")
                        //    fSty = FontStyle.Regular;
                        //Font = new System.Drawing.Font(sFonName, fFonDim, fSty);
                        ////iX = iX + iIl + iXProgr;

                        //e.Graphics.DrawString(_decMovImp.ToString("#0.00"), Font, Brushes.Black, iY, iX);

                        /*** 
                         * Barcode 
                         * ***/

                        //s = "SELECT * FROM AnaBarcode WHERE ean_art='" + _strArtCod +"' AND ean_bil=1";
                        //DataTable t = _clsFun.FillTabSql("AnaBarcode", s, true, _strConSql);
                        //if(t.Rows.Count > 0)
                        //{
                        Console.WriteLine("xxx");

                        //string sEan1 = "01" + _strForArt + "17" + _dayDaySca.ToString("yyMMdd") + "30" + _strArtPxc + "10" + sLot;
                        //string sEan1 = "01" + _strForArt + "\\F17" + _dayDaySca.ToString("yyMMdd") + "\\F30" + _strArtPxc + "\\F10" + sLot;
                        //string sEan1 = "01" + _strForArt + (char)29 + "17" + _dayDaySca.ToString("yyMMdd") + (char)29 + "30" + _strArtPxc + (char)29 + "10" + sLot;
                        //string sEan1 = "01" + _strForArt + (char)63 + "17" + _dayDaySca.ToString("yyMMdd") + (char)63 + "30" + _strArtPxc + (char)63 + "10" + sLot;
                        string sEan1 = "01" + _strForArt + "17" + _dayDaySca.ToString("yyMMdd") + "30" + _strArtPxc + (char)63 + "10" + sLot;
                        //string sEan1 = "01" + _strForArt + (char)102 + "17" + _dayDaySca.ToString("yyMMdd") + (char)102 + "30" + _strArtPxc + (char)102 + "10" + sLot;
                        string sEan = "(01)" + _strForArt + "(17)" + _dayDaySca.ToString("yyMMdd") + "(30)" + _strArtPxc + "(10)" + sLot;

                        //if ((Boolean)t.Rows[0]["ean_ecp"])
                        //    s = _decMovQkg.ToString("00000");
                        //else
                        //    s = (_decMovImp * 100).ToString("00000");

                        //s = sEan.Substring(0, 7) + s;
                        //s = (_decMovImp * 100).ToString("00000");

                        //s = sEan.Substring(0, 7) + s;
                        //sEan = s + new clsCtrlCodici().FindMod10Digit(s);

                        //sEan.Substring(0, 7) + s;

                        /*
                        * Barcode
                        */
                        a = _str013_Posz_EanCod.Split(',');
                        iY = Convert.ToInt16(a[0]);
                        iX = Convert.ToInt16(a[1]);
                        iYy = Convert.ToInt16(a[2]);
                        iXx = Convert.ToInt16(a[3]);
                        iIl = Convert.ToInt16(a[4]);

                        aFon = _str013_Font_EanCod.Split(':');
                        sFonName = aFon[0];
                        fFonDim = (float)Convert.ToDouble(aFon[1]);
                        fSty = FontStyle.Bold;
                        if (aFon[2] == "0")
                            fSty = FontStyle.Regular;
                        Font = new System.Drawing.Font(sFonName, fFonDim, fSty);


                        if (sEan.Trim().Length > 0)
                        {
                            try
                            {
                                //if (s.Length <= 8)
                                //{
                                //    Image img = _ean08.Encode(tpEan08, sEan, Color.Black, Color.White, 200, 200);
                                //    MemoryStream ms = new MemoryStream();
                                //    img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                //    //gfx.DrawImage(img, iY, iX, iYy, iXx);
                                //    e.Graphics.DrawImage(img, iY, iX, iYy, iXx);
                                //}
                                //else
                                //{
                                //    sEan = sEan.PadLeft(13, Convert.ToChar('0'));
                                //    Image img = _ean13.Encode(tpEan13, sEan, Color.Black, Color.White, 200, 200);
                                //    MemoryStream ms = new MemoryStream();
                                //    img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                //    e.Graphics.DrawImage(img, iY, iX, iYy, iXx);
                                //}

                                Image img = _ean128.Encode(tpEan128, sEan1, Color.Black, Color.White, 700, 700);
                                MemoryStream ms = new MemoryStream();
                                //img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                //gfx.DrawImage(img, iY, iX, iYy, iXx);
                                e.Graphics.DrawImage(img, iY, iX, iYy, iXx);

                                //e.Graphics.DrawString(sEan1, Font, Brushes.Black, iY + 35, iX + 40);
                                //e.Graphics.DrawString(sEan, Font, Brushes.Black, iY + 35, iX + 80);

                                e.Graphics.DrawString(sEan, Font, Brushes.Black, iY + 30, iX + 40);
                                //e.Graphics.DrawString(sEan1, Font, Brushes.Black, iY + 30, iX + 50);

                            }
                            catch (Exception ex)
                            {
                                _clsFun.ErrorLog(ex.Message, _strArtCod + " " + sEan);
                            }
                        }
                        //}

                    //}
                }

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

            string sFil = Path.GetDirectoryName(_clsDef.FILEINI) + "\\EtiNomeFont3.ini";

            if (strTip == "R")
            {
                if (File.Exists(sFil))
                {
                    int i = 0;
                    try
                    {
                        using (StreamReader sr = new StreamReader(sFil))
                        {
                            string sRig = "";

                            while ((sRig = sr.ReadLine()) != null)
                            {
                                i++;
                                if (i == 21)
                                    s = "";

                                if (sRig.Length > 4)
                                {
                                    if (sRig.Substring(0, 3) == "001")
                                        _str001_Print = sRig.Substring(4);
                                    else if (sRig.Substring(0, 3) == "002")
                                        _str002_Dim = sRig.Substring(4);
                                    else if (sRig.Substring(0, 3) == "003")
                                    {
                                        string[] a = sRig.Substring(4).Split(';');
                                        //_str003_Posz_NegRag = a[0];
                                        //_str003_Font_NegRag = a[1];
                                        //_str003_Posz_NegInd = a[2];
                                        //_str003_Font_NegInd = a[3];
                                        //_str003_Posz_RegUls = a[4];
                                        //_str003_Font_RegUls = a[5];
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
                                        _str005_Posz_DaySca = a[0];
                                        _str005_Font_DaySca = a[1];
                                    }
                                    else if (sRig.Substring(0, 3) == "006")
                                    {
                                        string[] a = sRig.Substring(4).Split(';');
                                        _str006_Posz_ArtPxc = a[0];
                                        _str006_Font_ArtPxc = a[1];
                                    }
                                    else if (sRig.Substring(0, 3) == "007")
                                    {
                                        string[] a = sRig.Substring(4).Split(';');
                                        _str007_Posz_CodLot = a[0];
                                        _str007_Font_CodLot = a[1];
                                        //if (a.Length > 3)
                                        //{
                                        //    _str006_Posz_DesLot = a[2];
                                        //    _str006_Font_DesLot = a[3];
                                        //}
                                    }
                                    //else if (sRig.Substring(0, 3) == "007")
                                    //{
                                    //    string[] a = sRig.Substring(4).Split(';');
                                    //    _str007_Posz_PriLbl = a[0];
                                    //    _str007_Font_PriLbl = a[1];
                                    //    _str007_Posz_PriDta = a[2];
                                    //    _str007_Font_PriDta = a[3];
                                    //}
                                    else if (sRig.Substring(0, 3) == "008")
                                    {
                                        //string[] a = sRig.Substring(4).Split(';');
                                        //_str008_Posz_PneLbl = a[0];
                                        //_str008_Font_PneLbl = a[1];
                                        //_str008_Posz_PneQta = a[2];
                                        //_str008_Font_PneQta = a[3];
                                    }
                                    else if (sRig.Substring(0, 3) == "009")
                                    {
                                        //string[] a = sRig.Substring(4).Split(';');
                                        //_str009_Posz_ScaLbl = a[0];
                                        //_str009_Font_ScaLbl = a[1];
                                        //_str009_Posz_ScaDay = a[2];
                                        //_str009_Font_ScaDay = a[3];
                                    }
                                    else if (sRig.Substring(0, 3) == "010")
                                    {
                                        //string[] a = sRig.Substring(4).Split(';');
                                        //_str010_Posz_PrkLbl = a[0];
                                        //_str010_Font_PrkLbl = a[1];
                                        //_str010_Posz_PrkKgr = a[2];
                                        //_str010_Font_PrkKgr = a[3];
                                    }
                                    else if (sRig.Substring(0, 3) == "011")
                                    {
                                        //string[] a = sRig.Substring(4).Split(';');
                                        //_str011_Posz_ImpLbl = a[0];
                                        //_str011_Font_ImpLbl = a[1];
                                        //_str011_Posz_ImpVal = a[2];
                                        //_str011_Font_ImpVal = a[3];
                                    }
                                    else if (sRig.Substring(0, 3) == "012")
                                    {
                                        string[] a = sRig.Substring(4).Split(';');
                                        _str012_Posz_LogUno = a[0];
                                        _str012_Posz_LogDue = a[1];
                                    }
                                    else if (sRig.Substring(0, 3) == "013")
                                    {
                                        string[] a = sRig.Substring(4).Split(';');
                                        _str013_Posz_EanCod = a[0];
                                        _str013_Font_EanCod = a[1];
                                    }
                                    else if (sRig.Substring(0, 3) == "099")
                                        _str099_PrnImmediata = sRig.Substring(4);
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

        private string SeekLogo(string strTip)
        {
            string sPth = "C:\\ApProject\\Temp\\Img\\Clienti\\";
            string sLog = "";
            string s = "";

            if (strTip == "LOG")
            {
                string[] sFils = Directory.GetFiles(sPth);
                foreach (string sFil in sFils)
                {
                    s = Path.GetFileNameWithoutExtension(sFil);
                    if (s.Contains(_strCliCod))
                    {
                        sLog = sFil;
                        break;
                    }
                }
                if (sLog == "")
                {
                    sLog = sPth + "LOGO.png";

                    if (!File.Exists(sLog))
                        sLog = "";
                }
            }
            else if (strTip == "BOL")
            {
                sLog = sPth + "BOLLINO.png";
            }

            return sLog;
        }

    }
}
