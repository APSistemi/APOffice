using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;

namespace APOffice
{
    class clsBilBizWinvarp
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public Boolean _bolIng2Varp = false;
        public string _strIngTip = "";
        public string _strRecRebPlu = "";
        public string _strPar031Lotti2Pos = "";

        public clsBilBizWinvarp()
        { }

        public string BilArtRow(DataRow rowBil)
        {
            string s = "";
            string sRig = "";

            if(_strRecRebPlu != "")
            {
                if (_strRecRebPlu == "bbbppp")
                {
                    sRig = ((string)rowBil["bil_reb"]).PadLeft(3, Convert.ToChar("0"));

                    s = (string)rowBil["bil_plu"];
                    if (s.Length > 3)
                        s = s.Substring(s.Length-3, 3);

                    sRig += s.PadLeft(3, Convert.ToChar("0"));        //PLU
                }
            }
            else if (!DBNull.Value.Equals(rowBil["bil_gru"]) && ((string)rowBil["bil_gru"]).Trim() != "")
            {
                //Reparto + PLU N 1 6 1 – 999999
                s = ((string)rowBil["bil_gru"]).PadLeft(2, Convert.ToChar("0"));        //PLU
                sRig += s;

                //PLU + PLU N 1 6 1 – 999999
                s = ((string)rowBil["bil_plu"]).PadLeft(4, Convert.ToChar("0"));        //PLU
                sRig += s;
            }
            else
            {
                sRig += "0";
                //Reparto + PLU N 1 6 1 – 999999
                s = ((string)rowBil["bil_reb"]).PadLeft(1, Convert.ToChar("0"));        //PLU
                sRig += s;

                //PLU + PLU N 1 6 1 – 999999
                s = ((string)rowBil["bil_plu"]).PadLeft(4, Convert.ToChar("0"));        //PLU
                sRig += s;
            }

            //Numero offerta N 7 1 0 – 9
            sRig += "0";

            //Prezzo N 8 6 0 – 999999
            s = Convert.ToInt32((decimal)rowBil["pos_prv"] * 100).ToString().Replace(",", "").PadLeft(6, Convert.ToChar("0"));
            sRig += s;

            //Prezzo2 N 14 6 0 – 999999 non usato
            sRig += "000000";

            //Prezzo3 N 20 6 0 – 999999 non usato
            sRig += "000000";

            //Prezzo4 N 26 6 0 – 999999 non usato
            sRig += "000000";
            
            //Prezzo straordinario N 32 6 0 – 999999 non usato
            sRig += "000000";

            //Gruppo Merceologico N 38 4 1 – 999
            if (!DBNull.Value.Equals(rowBil["bil_rbi"]) && ((string)rowBil["bil_rbi"]).Trim() != "")
                //sRig += "000" + (string)rowBil["bil_rbi"];    // Seck 20190520
                sRig += ((string)rowBil["bil_rbi"]).PadLeft(4, Convert.ToChar("0"));
           else
                sRig += ((string)rowBil["bil_reb"]).PadLeft(4, Convert.ToChar("0"));

            //Tara N 42 6 0 – 9998 In grammi
            s = "000000";
            if ((decimal)rowBil["bil_tar"] > 0)
                s = Convert.ToInt16(rowBil["bil_tar"]).ToString("000000");
            sRig += s;

            //Giorni di conservazione 1 N 48 3 0 – 899 Giorni che verranno sommati alla DATA1 (data del giorno) per il calcolo della DATA2 (data di scadenza).
            s = "000";
            if (_clsFun.Numerico(rowBil["bil_gsc"], "0123456789"))
                s = Convert.ToInt16(rowBil["bil_gsc"]).ToString("000");
            sRig += s;

            //Giorni di conservazione 2 N 51 3 0 – 899 Giorni che verranno sommati alla DATA2 per il calcolo della DATA3
            sRig += "000";

            //Codice EAN N 54 13
            s = ((string)rowBil["pos_ean"]).Substring(0, 11) + "11";
            sRig += s;

            //Numero testo aggiuntivo 1 N 67 6 0 – 999999
            sRig += "000000";

            //Numero testo aggiuntivo 2 N 73 6 0 – 999999
            sRig += "000000";

            //Numero testo aggiuntivo 3 N 79 6 0 – 999999
            sRig += "000000";

            //Numero testo aggiuntivo 4 N 85 6 0 – 999999
            sRig += "000000";

            //Pezzi per scatola N 91 2 0 – 99
            sRig += "00";

            //Peso fisso N 93 6 0 – 999998 In grammi
            sRig += "000000";

            //Numero offerta N 99 1 0 – 9 Non usato
            sRig += "0";

            //Articolo in pubblicità N 100 1 0 – 1 0 = no (default) 1= si
            sRig += "0";

            //Sovrascrittura prezzo N 101 1 0 – 1 0 = no (default) 1= si
            sRig += "1";

            //Stile stampa tracciabilità N 102 1 0 – 4
            sRig += " ";

            //Richiesta inserimento lotto tracciabilità N 103 1 0 – 2 0 = no (default) 1 = manuale 2 = da configurazione
            //sRig += "0";

            s = "0";
            if ((string)rowBil["bil_tra"] == "S" && _strPar031Lotti2Pos == "")  //20190720 se gestione Lotti su parametro 31 non alzo il flag
                s = "1";
            sRig += s;


            //Tipo articolo N 104 2 0 – 10 00 = pesato 01 = manuale (a pezzo) 02 = sottrazione (reso) 03 = by count 04 = peso fisso 05 = non usato 06 = non usato 07 = non usato 08 = non usato 09 = non usato 10 = non usato Per approfondimenti vedi manuale programmazione della bilancia
            s = "00";
            if ((Boolean)rowBil["bil_bpz"])
                s = "01";
                sRig += s;

            //Base prezzo N 106 1 0 – 4 0 = prezzo al kilo 1 = prezzo all’etto 2 = non usato 3 = non usato 4 = non usato Per approfondimenti vedi manuale programmazione della bilancia
            sRig += "0";

            //Tipo prezzo N 107 1 0 – 3 0 = normale 1 = scalare 2 = speciale mix 3 = gratis (3x2)
            sRig += "0";

            //Codice IVA N 108 1 1 – 9
            sRig += "0"; //VERIFICARE TABELLA IVA

            ////Periodo offerta N 109 1 0 – 9 Non usato
            //sRig += "0";

            //Filler N 110 42
            sRig += new string(' ', 42);

            //Filler N 152 10
            sRig += new string(' ', 10);

            sRig += " ";

            //Testo articolo A 162 500

            //if (!_bolIng2Varp && _strPar031Lotti2Pos == "")
            //    sRig += "^5;" + (((string)rowBil["pos_ard"]) + new string(' ', 500)).Substring(0, 497);
            //else
            //{
                //int iLen = 0;
            string sRigg = "";

            Boolean bIng = false;

            if (_strIngTip == "2")
            {
                s = ((string)rowBil["pos_ard"]) + new string(' ', 54);
                s = s.Substring(0, 51).Trim() ;

                if(s.Length > 31)
                {
                    sRigg += "^5;" + s.Substring(0, 30) + _clsDef.LF;
                    s = s.Substring(30).Trim();
                    sRigg += "^5;" + s + _clsDef.LF;
                    bIng = true;
                }
                else
                {
                    sRigg += "^5;" + s + _clsDef.LF;
                    sRigg += _clsDef.LF;
                }

                //int iLoopText = 8;

                string sIng = _clsQry.GetFileIngredienti((string)rowBil["pos_art"]);
                if (sIng.Length > 0)
                {
                    sIng = _clsFun.CtrlCrt(sIng, "1234567890qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM/'@&.,+-*:%_\\=?!| ");
                    sIng = _clsFun.StrSplit2(sIng, 51);
                }
                string sLot = "";
                if ((string)rowBil["bil_tra"] == "S" && _strPar031Lotti2Pos != "")
                    sLot = _clsQry.GetDatiLotto(_strPar031Lotti2Pos, (string)rowBil["pos_ean"]);

                if (sLot != "")
                {
                    sLot = _clsFun.CtrlCrt(sLot, "1234567890qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM/'@&.,+-*:%_\\=?!| ");
                    sIng = sLot + sIng;
                }

                if (sIng.Length > 0)
                {
                    //sIng = _clsFun.CtrlCrt(sIng, "1234567890qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM/'@&.,+-*:%_\\=?!| ");
                    
                    bIng = true;

                    string[] aa = sIng.Split('|');

                    for (int i = 0; i < 8; i++)
                    {
                        if (aa.Length > i)
                        {
                            sRigg += "^2;" + aa[i].Trim();
                            if(i < 7)
                                sRigg += _clsDef.LF;
                        }
                        else if (i < 7)
                            sRigg += "^2;" + _clsDef.LF;
                        else
                            sRigg += "^2;";
                    }
                }
                else
                {
                    for(int i = 1; i <= 7; i++)
                        sRigg += "^2;" + _clsDef.LF;
                    sRigg += "^2;";
                }

                sRigg = (sRigg + new string(' ', 500)).Substring(0, 500);

                //sRig += sRigg;
            }
            if(!bIng)
               sRig += "^5;" + (((string)rowBil["pos_ard"]) + new string(' ', 500)).Substring(0, 497);
            else
                sRig += sRigg;

            //WALO N 662 1 0 – 1 tipo record 0 = aggiungi/modifica 1 = cancella
            sRig += "0";

            ////Testo generico 1 N 663 6 0 – 999999
            //for (int i = 1; i <= 10; i++)
            //    sRig += "000000";

            sRig += new string(' ', 72);


            ////Numero logo 1 N 723 4 0 - 9999
            //for (int i = 1; i <= 3; i++)
            //    sRig += "0000";

            //Data variazione N 735 8 GGMMAAAA 00000000
            sRig += DateTime.Today.ToString("ddMMyyyy");

            //sRig += _clsDef.CRLF;

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
                    sRig += (((string)tabIng.Rows[i]["ing_txt"]) + new string(' ', 50)).Substring(0, 50);
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

        public DataTable BizScontrino(string strPar, string strEan)
        {
            string[] arr = strPar.Split(',');

            Boolean b = true;
            string s = "";
            string sTip = arr[1];
            string sPth = arr[2];
            string sPrf = arr[4];
            string sSuf = arr[5];

            string sRig = "";

            //sPrf = arr[4];
            //sSuf = arr[5];

            DataTable tBil = new clsGenTabTmp().TabTmpBil("TabBil");

            //sPth = "C:\\WinSwGx-NET\\scontrini\\";

            string sFil = sPth + sPrf + strEan.Substring(2, 4) + sSuf;

            if (File.Exists(sFil))
            {
                using (StreamReader sr = new StreamReader(sFil, System.Text.Encoding.Default))
                {
                    Boolean bCan = false;

                    while ((sRig = sr.ReadLine()) != null)
                    {
                        Boolean bPez = false;

                        string sBan = sRig.Substring(19, 1);
                        string sPlu = sRig.Substring(20, 4);
                        string sPrv = sRig.Substring(40, 5);
                        string sQta = sRig.Substring(48, 5);
                        string sImp = sRig.Substring(57, 7);
                        string sSgn = sRig.Substring(53, 1);

                        if(sQta == "00000")                     //Qta a pezzi
                        {
                            bPez = true;
                            sQta = sRig.Substring(32, 5);
                        }

                        if (sSgn == "-")
                        {
                            s = "bil_ban='" + sBan + "' AND bil_plu='" + sPlu + "' AND bil_ims='" + sImp + "'";
                            DataRow[] j = tBil.Select(s);

                            if (j.Length > 0)
                                j[0]["bil_imp"] = 0;
                            bCan = true;
                        }
                        else
                        {
                            DataRow y = tBil.NewRow();
                            y["bil_ban"] = sBan;
                            y["bil_plu"] = sPlu;
                            y["bil_prv"] = Convert.ToDecimal(sPrv) / 100;

                            //y["bil_qta"] = Convert.ToDecimal(sQta) / 1000;
                            y["bil_qta"] = Convert.ToDecimal(sQta);
                            if(!bPez)
                                y["bil_qta"] = Convert.ToDecimal(sQta) / 1000;

                            y["bil_imp"] = Convert.ToDecimal(sImp) / 1000;
                            y["bil_ims"] = sImp;
                            tBil.Rows.Add(y);
                        }
                    }

                    if (bCan)
                    {
                        DataTable t = tBil.Clone();

                        foreach (DataRow y in tBil.Rows)
                        {
                            if ((decimal)y["bil_imp"] > 0)
                                t.ImportRow(y);
                        }

                        tBil = t.Copy();
                    }

                }
            }

            return tBil;
        }

    }
}
