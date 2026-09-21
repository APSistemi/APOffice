using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsBilBizerba
    {
        clsFuncs _clsFun = new clsFuncs();

        public clsBilBizerba()
        { }

        public string BilArtRow(DataRow rowBil)
        {
            string s = "";
            string sRig = "";

            //PLU (chiave primaria)   6   N  000001 - 999999 
            s = ((string)rowBil["bil_plu"]).PadLeft(6,Convert.ToChar('0'));        //PLU
            sRig += s;

            //gruppo merceologico   4   N   0001 – 0999 
            //sRig += "0001";
            sRig += ((string)rowBil["bil_reb"]).PadLeft(4, Convert.ToChar("0"));        //PLU


            //prezzo in €cent   6   N   000001 - 999999 
            s = Convert.ToInt32((decimal)rowBil["pos_prv"] * 100).ToString().Replace(",", "").PadLeft(6, Convert.ToChar("0"));
            sRig += s;

            //codice EAN    13   N 
            s = ((string)rowBil["pos_ean"]).Substring(0, 11) + "11";
            sRig += s;

            //Dimensioni etichetta  2  
            sRig += "55";

            //fisso 66 testo PLU    43   A
            s = (string)rowBil["pos_arb"];
            if (s.Length > 41)
                s = s.Substring(0, 41);
            sRig += (s + new string(' ', 41)).Substring(0, 41);

            //giorni di conservazione   4    N   0000 - 0999 
            sRig += "0000";

            //tara in grammi    6    N   000000 - ?????? (il valore max dipende dalla portata della bilancia) 
            //sRig += "000000";
            if (!DBNull.Value.Equals(rowBil["bil_tar"]) && _clsFun.Numerico(rowBil["bil_tar"]))
                sRig += Convert.ToInt16(rowBil["bil_tar"]).ToString("000000");
            else
                sRig += "000000";

            //Filler 3 
            sRig += "000";

            //Richiesta inserimento lotto tracciabilità
            sRig += "0";

            //Stile stampa tracciabilità
            sRig += "00";

            //codice IVA    1    N   0 - 9 
            sRig += "0";

            //variazione prezzo di vendita  1    N  0 - 1 
            sRig += "1";

            //flag prezzo scontato   1    N   0 - 1 
            sRig += "0";

            //tipo articolo    1    A    P, p, @ --> pesato (prodotto venduto a peso) M, m, H --> manuale (prodotto venduto a pezzo) E, e --> pesato con prezzo all'etto F, f --> peso fisso -, J --> in sottrazione 
            //sRig += "@";
            sRig += "P";

            //data applicazione   6    N   GGMMAA 
            sRig += DateTime.Now.ToString("ddMMyy");

            //flag delete/aggiungi   1    A  D --> delete A --> aggiungi o modifica 
            sRig += "A";


            return sRig;
        }

        public string BilArtIng(DataRow rowBil, DataTable tabIng)
        {
            string s = "";
            string sRig = "";

            //PLU (chiave primaria)   6   N  000001 - 999999 
            s = ((string)rowBil["bil_plu"]);        //PLU
            sRig += s;

            //Formato stringa 1
            sRig += "5";

            //Descrizione
            s = (string)rowBil["pos_arb"];
            if (s.Length > 50)
                s = s.Substring(0, 50);
            sRig += (s + new string(' ', 50)).Substring(0, 50);

            for (int i = 0; i < 9; i++)
            {

                //Formato stringa 1
                sRig += "5";
                //Testo riga 1

                if (tabIng.Rows.Count > i)
                {
                    sRig += (((string)tabIng.Rows[i]["ing_txt"]) + new string(' ', 50)).Substring(0,50);
                }
                else
                    sRig += new string(' ', 50);

            }

            //Formato stringa 2
            //Testo riga 

            //Formato stringa 3
            //Testo riga 

            //Formato stringa 4
            //Testo riga 

            //Formato stringa 5
            //Testo riga 

            //Formato stringa 6
            //Testo riga 

            //Formato stringa 7
            //Testo riga 

            //Formato stringa 8
            //Testo riga 

            //Formato stringa 9
            //Testo riga 

            //Formato stringa 10
            //Testo riga 

            sRig += DateTime.Now.ToString("ddMMyy");

            return sRig;
        }

        public string BilArtRow2(DataRow rowBil)
        {
            string s = "";
            string sRig = "";

            //PLU (chiave primaria)   6   N  000001 - 999999 
            s = ((string)rowBil["bil_plu"]).PadLeft(6, Convert.ToChar('0'));        //PLU
            sRig += s;

            //gruppo merceologico   4   N   0001 – 0999 
            //sRig += "0001";
            sRig += ((string)rowBil["bil_reb"]).PadLeft(4, Convert.ToChar("0"));        //PLU


            //prezzo in €cent   6   N   000001 - 999999 
            s = Convert.ToInt32((decimal)rowBil["pos_prv"] * 100).ToString().Replace(",", "").PadLeft(6, Convert.ToChar("0"));
            sRig += s;

            //codice EAN    13   N 
            s = ((string)rowBil["pos_ean"]).Substring(0, 11) + "11";
            sRig += s;

            //Dimensioni etichetta  2  
            sRig += "55";

            //fisso 66 testo PLU    43   A
            s = (string)rowBil["pos_arb"];
            if (s.Length > 41)
                s = s.Substring(0, 41);
            sRig += (s + new string(' ', 41)).Substring(0, 41);

            //giorni di conservazione   4    N   0000 - 0999 
            sRig += "0000";

            //tara in grammi    6    N   000000 - ?????? (il valore max dipende dalla portata della bilancia) 
            //sRig += "000000";
            if (!DBNull.Value.Equals(rowBil["bil_tar"]) && _clsFun.Numerico(rowBil["bil_tar"]))
                sRig += Convert.ToInt16(rowBil["bil_tar"]).ToString("000000");
            else
                sRig += "000000";

            //Filler 3 
            sRig += "000";

            //Richiesta inserimento lotto tracciabilità
            sRig += "0";

            //Stile stampa tracciabilità
            sRig += "00";

            //codice IVA    1    N   0 - 9 
            sRig += "0";

            //variazione prezzo di vendita  1    N  0 - 1 
            sRig += "1";

            //flag prezzo scontato   1    N   0 - 1 
            sRig += "0";

            //tipo articolo    1    A    P, p, @ --> pesato (prodotto venduto a peso) M, m, H --> manuale (prodotto venduto a pezzo) E, e --> pesato con prezzo all'etto F, f --> peso fisso -, J --> in sottrazione 
            sRig += "@";
            //sRig += "P";

            //data applicazione   6    N   GGMMAA 
            sRig += DateTime.Now.ToString("ddMMyy");

            //flag delete/aggiungi   1    A  D --> delete A --> aggiungi o modifica 
            sRig += "A";
            sRig += "000000000000000000000000000000000000";

            if (((string)rowBil["bil_img"]).Trim() != "")
                sRig += (((string)rowBil["bil_img"]).Trim() + ".jpg" + new string(' ', 24)).Substring(0,24) + "000";
            else
                sRig += "                        000";

            return sRig;
        }

    }

}
