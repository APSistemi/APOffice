using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;

namespace APOffice
{
    class clsBilOmega
    {
        public string _strOmegaBarcodeLun = "";
        //public string _strIngredienti = "";
        public string _strParBilance = "";

        public clsBilOmega()
        { }

        //static char asciiSymbol(byte val)
        //{
        //    if (val < 32) return '.';  // Non-printable ASCII
        //    if (val < 127) return (char)val;   // Normal ASCII
        //    // Workaround the hole in Latin-1 code page
        //    if (val == 127) return '.';
        //    if (val < 0x90) return "€.‚ƒ„…†‡ˆ‰Š‹Œ.Ž."[val & 0xF];
        //    if (val < 0xA0) return ".‘’“”•–—˜™š›œ.žŸ"[val & 0xF];
        //    if (val == 0xAD) return '.';   // Soft hyphen: this symbol is zero-width even in monospace fonts
        //    return (char)val;   // Normal Latin-1
        //}

        //private string ASCII8ToString(byte[] ASCIIData)
        //{
        //    var e = Encoding.GetEncoding("437");
        //    return e.GetString(ASCIIData);
        //}

        public string BilArtRow(DataRow rowBil, string strIng)
        {
            string s = "";
            string sRig = "";

            //string BilBan = "";     //    'Banco
            sRig += (string)rowBil["bil_reb"];          //Banco

            /*
            s = ((string)rowBil["bil_plu"]).Substring(1, 3);        //PLU

            if (_strOmegaBarcodeLun == "4")
            {
                s = ((string)rowBil["bil_plu"]).Trim();        //PLU

                Int32 i = Convert.ToInt32(s.Substring(0, 2));

                if (i > 9)
                {
                    i += 55;
                    //char c = asciiSymbol(Convert.ToByte(i));
                    //string ss = c.ToString();
                    //s = ss + s.Substring(2);

                    //var e = Encoding.GetEncoding(1252);
                    //var e = Encoding.GetEncoding("iso-8859-15");
                    //var e = Encoding.GetEncoding("utf-32");
                    var e = Encoding.GetEncoding(1250);
                    String ss = e.GetString(new byte[] { Convert.ToByte(i) });
                    s = ss + s.Substring(2);
                }
            }
            */

            //string BilCod = "";     //'Codice articolo bilancia
            string sPlu = ((string)rowBil["bil_plu"]).Substring(1, 3);        //PLU

            if (_strOmegaBarcodeLun == "4" || _strOmegaBarcodeLun == "8")
            {
                s = ((string)rowBil["bil_plu"]).Trim();        //PLU

                //s = Convert.ToInt32(s).ToString();

                Int32 i = Convert.ToInt32(s.Substring(0,2));

                if(i > 9 && s.Length > 3)
                {
                    i += 55;
                    //char c = asciiSymbol(Convert.ToByte(i));
                    //string ss = c.ToString();
                    //s = ss + s.Substring(2);

                    //var e = Encoding.GetEncoding(1252);
                    //var e = Encoding.GetEncoding("iso-8859-15");
                    //var e = Encoding.GetEncoding("utf-32");
                    var e = Encoding.GetEncoding(1250);
                    String ss = e.GetString(new byte[] { Convert.ToByte(i) });
                    sPlu = ss + s.Substring(2);
                }
            }
            //else
            //    s = ((string)rowBil["bil_plu"]).Substring(1, 3);        //PLU

            sRig += sPlu;

            /*
            s = ((string)rowBil["bil_plu"]).Trim();        //PLU

            if (s.Length <= 3)
                sRig += s;        //PLU
            else
            {
                char c = Convert.ToChar(65);
                string ss = c.ToString();
                sRig += s;        //PLU
            }
            */ 

            //string BilSco = "";     //'Sconto
            sRig += "0000";     //Sconto

            //string BilIva = "";
            sRig += ((string)rowBil["pos_iva"]).Substring(1,2);     //IVA

            //string BilF01 = "";    //'Filler
            sRig += "00";

            //string BilTip = "";    //'Tipo
            //1 peso
            //2 a corpo
            //3 peso modificabile
            //4 a corpo  modificabile
            s = "1";
            if(_strParBilance.Length > 0 && _strParBilance.Substring(0,1) == "S")
                s = "3";

            if ((Boolean)rowBil["bil_bpz"])
            {
                s = "2";
                if (_strParBilance.Length > 0 && _strParBilance.Substring(0, 1) == "S")
                    s = "4";
            }
            sRig += s;

            //string BilF02 = "";
            sRig += "2";

            //string BilPrz = "";    //'Prezzo
            s = Convert.ToInt32((decimal)rowBil["pos_prv"] * 100).ToString().Replace(",","").PadLeft(6, Convert.ToChar(" "));
            sRig += s;

            //string BilTar = "";    //'Tara

            s = "   0";
            if ((decimal)rowBil["bil_tar"] > 0)
            {
                s = Convert.ToInt16(rowBil["bil_tar"]).ToString().PadLeft(4, ' ');
                 
            }

            //sRig += "   0";
            sRig += s;

            //string Bil126 = "";    //'Chr(126)
            sRig += "~";

            //string BilF03 = "";
            sRig += "0";

            //string BilSog = "";    //'Soglia
            sRig += "    0";

            //string BilGla = "";    //'Glassa
            sRig += "  0";

            //string BilRic = "";    //'Ricetta
            sRig += "000";

            //string BilEst = "";    //'Codice esteso
            sRig += "      ";

            //string BilNRi = "";    //'N° riga
            sRig += "1";

            if (_strOmegaBarcodeLun == "3")
            {
                //string BilLay = "";    //'LayOut
                sRig += "01";
                //string BilEan = "";    //'Barcode
                s = ((string)rowBil["pos_ean"]).Substring(1, 5);
                sRig += s;
            }
            else if (_strOmegaBarcodeLun == "4")
            {
                //string BilLay = "";    //'LayOut
                sRig += "01";
                //string BilEan = "";    //'Barcode
                s = ((string)rowBil["pos_ean"]).Substring(1, 5);
                sRig += s;
            }
            else
            {
                string BilLay = "";    //'LayOut
                sRig += "01";
                string BilEan = "";    //'Barcode
                s = ((string)rowBil["pos_ean"]).Substring(2, 5);
                sRig += s;
            }

            //string BilUni = "";    //'Uniscr
            sRig += "01";

            //string BilLin = "";    //'Ling
            sRig += "00";

            //string BilLsc = "";    //'Lscad
            sRig += "00";

            //string BilLnt = "";    //'Lint
            sRig += "00";

            //string BilDa1 = "";    //'Data 1
            sRig += new string(' ', 6);

            //string BilDa2 = "";    //'Data conservazione dati
            //s = ((string)rowBil["bil_gsc"]).Trim();
            //if (s == "")
            //    s = new string(' ', 6);
            //else
            //{
            //    s = (new string(' ', 6) + s.Substring(s.Length));
            //    s = s.Substring(s.Length - 6, 6);
            //}
            s = "   000";
            sRig += s;

            //string BilDa3 = "";    //'Data 3
            sRig += new string(' ', 6);

            //sRig += (((string)rowBil["pos_ard"]) + new string(' ', 15)).Substring(0, 15);
            sRig += (((string)rowBil["pos_ard"]) + new string(' ', 30)).Substring(0, 30);

            if (strIng == "")
            {
                for (int i = 1; i <= 10; i++)
                    sRig += (char)29;
                sRig += new string(' ', 470);
            }
            else
            {
                sRig += (char)29;

                Console.WriteLine("aaaaaaaaaa");

                int i = 0;

                string sRi1 = "";

                string[] a = strIng.Split('|');
                foreach (string ss in a)
                {
                    sRi1 += ss.Trim() + (char)29;
                    i++;
                    if (i > 10)
                        break;
                }

                if(i<10)
                {

                    i = 10 - i-1;

                    for (int ii = 1; ii <= i; ii++)
                        sRi1 += (char)29;

                    sRi1 = (sRi1 + new string(' ', 479)).Substring(0,479);

                }
                //sRig += (char)29;
                sRig += sRi1;

            }
            //sRig += new string(' ', 485);

            //string Bil124 = "";    //'chr(124)
            //sRig += "     ";
            sRig += "|";
            sRig += "066500";
            if (_strOmegaBarcodeLun == "6" || _strOmegaBarcodeLun == "8") 
                sRig += "00000";
            else
                sRig += "000000";

            //BilPlu
            s =  ((string)rowBil["pos_ean"]).Substring(2, 5);
            if (_strOmegaBarcodeLun == "3")
                s = ((string)rowBil["pos_ean"]).Substring(1, 5);
            else if (_strOmegaBarcodeLun == "4")
                s = ((string)rowBil["pos_ean"]).Substring(1, 5);
            else if (_strOmegaBarcodeLun == "6")
                s = ((string)rowBil["pos_ean"]).Substring(1, 6);
            else if (_strOmegaBarcodeLun == "8")
                s = ((string)rowBil["pos_ean"]).Substring(1, 6);
            sRig += s;

            sRig += new string(' ', 49);
            sRig += " ";
            sRig += new string(' ', 6);

            //string BilBar = "";    //'Barcode SI/NO
            sRig += "00";

            //Repl(Bil029, 9) + _
            //Space(460) + _
            //Bil124 + _
            //"066500" + _
            //"000000" + BilPlu + Space(49) + " " + Space(6) + BilBar

            return sRig;
        }
    }
}
