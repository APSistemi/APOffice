using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Data;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace APOffice
{
    class clsGesMovIngLabel03
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public string _strPrn = "";
        //public string _strTxt = "";
        //public DataTable _tabSco = new DataTable();
        //public DataTable _tabPag = new DataTable();
        //public string _strPagDes = "";
        //public string _strConMdb = "";
        //public decimal _decPagNet = 0;
        //public decimal _decResto = 0;
        //public string _strScoNum = "";

        public string _strConSql = "";
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
        public DateTime _dayArtSca = new DateTime();    //Data scadenza

        public string _strLotNas = "";
        public string _strLotAll = "";
        public string _strLotMac = "";
        public string _strLotSez = "";

        private DataTable _tabLot = new DataTable();

        /* Dati fattura */
        //public string _strFatNum = "";
        //public string _strFatInf = "";

        string _strMsg = "OK";

        private DataTable _tabIva = new DataTable();

        public clsGesMovIngLabel03()
        {
            //string s = "SELECT * FROM TabIva";
            //_tabIva = _clsFun.FillTabMdb("TabIva", s, false, _strConMdb);
            //DataColumn[] keys = new DataColumn[5];
            //keys[0] = _tabIva.Columns["tab_cod"];
            //_tabIva.PrimaryKey = keys;

            string s = "";
            if (_strLotCod != "" && _strLotCod.Length > 2)
            {
                s = _strLotCod.Substring(2).Trim();

                s = "SELECT * FROM GesDocLotti WHERE lot_cod='" + s + "' ORDER BY lot_ddo DESC";
                _tabLot = _clsFun.FillTabSql("GesDocLotti", s, true, _strConSql);
            }
        }

        public string Stampa()
        {
            PrintDocument pd = new PrintDocument();

            try
            {
                try
                {
                    //DA ATTIVARE 
                    pd.PrinterSettings.PrinterName = _strPrn;

                    pd.PrintPage += new PrintPageEventHandler(this.pd_PrintPage);
                    pd.Print();
                }
                finally
                {
                    pd.Dispose();                    
                }
            }
            catch (Exception ex)
            {
                _strMsg = ex.Message;
            }

            return _strMsg;
        }

        // The PrintPage event is raised for each page to be printed.
        private void prova_pd_PrintPage(object sender, PrintPageEventArgs ev)
        {
            Font fo = new Font("Courier New", 10, FontStyle.Bold); ;

            float fX = 10;
            float fY = 10;

            //string[] a = ""; // _strTxt.Split(';');

            //foreach (string sRig in a)
            //{
            //    fX += 16;
            //    ev.Graphics.DrawString(sRig + _clsDef.CRLF, fo, Brushes.Black, fY, fX, new StringFormat());
            //}

            ev.Graphics.DrawString("." + _clsDef.CRLF, fo, Brushes.Black, fY, fX, new StringFormat());
            ev.Graphics.DrawString("." + _clsDef.CRLF, fo, Brushes.Black, fY, fX, new StringFormat());
            ev.Graphics.DrawString("." + _clsDef.CRLF, fo, Brushes.Black, fY, fX, new StringFormat());
            ev.Graphics.DrawString("." + _clsDef.CRLF, fo, Brushes.Black, fY, fX, new StringFormat());
            ev.Graphics.DrawString("." + _clsDef.CRLF, fo, Brushes.Black, fY, fX, new StringFormat());

            ev.HasMorePages = false;
        }

        private void pd_PrintPage(object sender, PrintPageEventArgs e)
        {
            string s = "";
            Boolean b = true;
            DataRow[] j;
            DataRow x;
            decimal d = 0;

            float fX = 10;
            float fStep = 16;
            float fY = 10;

            float iY = 0;
            float iX = 0;

            Boolean bCliNom = false;
            Boolean bCliInd = false;
            Boolean bCliPrv = false;

            if (_strCliPar.Length > 0 && _strCliPar.Substring(0, 1) == "S")
                bCliNom = true;
            if (_strCliPar.Length > 1 && _strCliPar.Substring(1, 1) == "S")
                bCliInd = true;
            if (_strCliPar.Length > 2 && _strCliPar.Substring(2, 1) == "S")
                bCliPrv = true;

            s = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
            if (s == "")
                s = "001";
            DataTable tCnf = _clsQry.ConfAzienda(s);

            //Font fo = new Font("Arial Narrow", 11, FontStyle.Regular); ;
            Font fo = new Font("Courier New", 11, FontStyle.Bold); ;

            //DataTable tIva = new clsGenTabTmp().TabTmpIva("TabIva");

            StringFormat format1 = new StringFormat(StringFormatFlags.NoClip);
            format1.LineAlignment = StringAlignment.Center;
            format1.Alignment = StringAlignment.Near;

            
            //int iTime = 90;

            for (int n = 1; n < 6; n++)
                fX += fStep;

            d = 0;

            s = "C:\\ApProject\\Temp\\Img\\ApShop\\scoTesta.png";
            if (File.Exists(s))
            {
                e.Graphics.DrawImage(Image.FromFile(s), 40, -10, 190, 100);
            }

            if (_strCliDes != "")
            {
                s = _strCliDes;
                fX += fStep;
                e.Graphics.DrawString(s, fo, Brushes.Black, fY, fX, new StringFormat());
                fX += fStep;
            }

            if (_strArtDes != "")
            {
                s = _strArtDes;
                fX += fStep;
                e.Graphics.DrawString(s, fo, Brushes.Black, fY, fX, new StringFormat());
                fX += fStep;
            }

            string sFil = _clsDef.PATHINGREDIENTI + "et01_" + _strArtCod + ".rtf";

            if (File.Exists(sFil))
            {
                fX += fStep;

                RichTextBox rtxBox = new RichTextBox();
                rtxBox.LoadFile(sFil);

                if (_tabLot.Rows.Count > 0)
                {
                    rtxBox.Text += _clsDef.CRLF;

                    string sNat = ((string)_tabLot.Rows[0]["lot_nat"]).Trim();
                    string sAll = ((string)_tabLot.Rows[0]["lot_all"]).Trim();
                    string sMac = ((string)_tabLot.Rows[0]["lot_mac"]).Trim();
                    string sSez = ((string)_tabLot.Rows[0]["lot_sez"]).Trim();

                    if (sNat != "")
                    {
                        s = "NATO: " + sNat;
                        rtxBox.Text += s + _clsDef.CRLF;
                    }
                    if (sAll != "")
                    {
                        s = "ALLEVATO: " + sAll;
                        rtxBox.Text += s + _clsDef.CRLF;
                    }
                    if (sMac != "")
                    {
                        s = "MACELLATO: " + sMac;
                        rtxBox.Text += s + _clsDef.CRLF;

                        if (sSez == "")
                        {
                            sSez = ((string)tCnf.Rows[0]["cnf_v01"]).Trim();
                            int pos = sSez.IndexOf(' ');
                            sSez = sSez.Substring(pos);
                        }

                        s = "SEZIONATO: " + sSez;
                        rtxBox.Text += s + _clsDef.CRLF;
                    }
                }

                //Stampa ingredienti da RichBox

                RectangleF rectF1 = new RectangleF(fY, fX, fY + 320, fX + 80); 
                
                e.Graphics.DrawString(rtxBox.Text, fo, Brushes.Black, rectF1, format1);
                //if (bBordi)
                //    e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1));
            }

            fX += fStep;
            iX = fX + 250;

            if (_strLotCod != "")
            {
                fX += fStep;
                e.Graphics.DrawString("Lotto N." + _strLotCod, fo, Brushes.Black, iY, iX);
            }

            iX += fStep;
            //e.Graphics.DrawString(s, fo, Brushes.Black, iY, iX);

            if (_dayArtSca != null)
            {
                iX += fStep;
                e.Graphics.DrawString("Da consumarsi entro il " + _dayArtSca.ToString("dd/MM/yyy"), fo, Brushes.Black, iY, iX);
            }

            if (_decMovQkg > 0)
            {
                iX += fStep;

                s = "Peso netto:";
                if (_strArtUmi == "NR")
                    s = "Peso: ";
                s += _decMovQkg.ToString("#0.000");

                string ss = " Kg";
                if (_strArtUmi == "NR")
                    ss = " Kg circa";
                s += ss;

                e.Graphics.DrawString(s, fo, Brushes.Black, iY, iX);
            }

            if (bCliPrv)
            {
                iX += fStep;

                e.Graphics.DrawString("€/Kg:    " + _decMovPrv.ToString("#0.00"), fo, Brushes.Black, iY, iX);

                iX += fStep;
                e.Graphics.DrawString("Importo: " + _decMovImp.ToString("#0.00"), fo, Brushes.Black, iY, iX);
            }

            for (int n = 1; n < 6;n++ )
                fX += fStep;

            e.Graphics.DrawString("." + _clsDef.CRLF, fo, Brushes.Black, fY, fX, new StringFormat());
            e.HasMorePages = false;

        }

    }
}
