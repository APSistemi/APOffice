using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace APOffice
{
    class clsDivEuroSpesa
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public clsDivEuroSpesa(){}

        public DataRow DivArt(string strRig, DataRow rowArt)
        {
            string s = "";
            //CCD0 Codice cliente                                   A    6    B     1    6
            //NUD0 Numero fattura                 (solo per fattura)S    6  0 B     7   12
            //DT10 Data fatt. Giorno                                S    2  0 B    13   14
            //DT20 Data fatt. Mese                                  S    2  0 B    15   16
            //DT30 Data fatt. Anno                                  S    4  0 B    17   20

            s = strRig.Substring(245, 8);

            rowArt["div_dva"] = _clsFun.Str2Day(s);

            //CAR0 Codice articolo                                  A    6    B    21   26
            rowArt["div_arf"] = strRig.Substring(5, 15).Trim(); //.PadLeft(9,Convert.ToChar("0"));

            //DES0 Descriz. articolo                                A   30    B    27   56
            rowArt["div_ard"] = strRig.Substring(42, 35).Trim();

            //FIL0 Filler                                           A    3    B    57   59
            //UMG0 U.M. x grammatura                                A    2    B    60   61
            rowArt["for_tgr"] = strRig.Substring(112, 2);
            //if ((string)rowArt["for_tgr"] == "CA")
            //    rowArt["for_tgr"] = "NR";

            s = strRig.Substring(21, 1);
            if (s == "T")
                rowArt["div_sta"] = _clsDef.STAATT;

            if ((string)rowArt["div_arf"] == "1111")
                Console.WriteLine("aaa");

            //LV10 Livello 1 Reparto                                A    2    B    62   63
            //LV20 Livello 2 Gruppo                                 A    2    B    64   65
            //LV30 Livello 3 Settore                                A    2    B    66   67
            //LV40 Livello 4 Famiglia                               A    2    B    68   69
            rowArt["for_rep"] = strRig.Substring(39, 3).Trim().PadLeft(3,Convert.ToChar("0"));
            //rowArt["div_ec1"] = strRig.Substring(63, 2);
            //rowArt["div_ec2"] = strRig.Substring(65, 2);
            //rowArt["div_ec3"] = strRig.Substring(67, 2);
            //rowArt["for_ecr"] = strRig.Substring(31, 8);
            rowArt["for_ecr"] = "";

            //IVV0 Cod. I.v.a.                                      A    2    B    70   71
            rowArt["for_iva"] = (strRig.Substring(84, 2)).Trim().PadLeft(3,Convert.ToChar("0"));

            //UNM0 Unità di misura                                  A    2    B    72   73
            rowArt["for_umi"] = strRig.Substring(114, 2);

            //PZC0 Pezzatura                                        S    4  0 B    74   77
            //rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(73, 4));
            rowArt["div_pxc"] = 1;

            //PEN0 Grammatura                                       S    6  3 B    78   83
            rowArt["div_pne"] = _clsFun.Txt2Dec(strRig.Substring(105, 7));

            //TPD0 N=Nota accredito                                 A    1    B    84   84
            rowArt["div_tpd"] = "";         // strRig.Substring(83, 1);

            //QTA0 Quantità fattura               (solo per fattura)S    7  2 B    85   91
            //s = strRig.Substring(84, 7);
            //if (_clsFun.Numerico(s))
            //    rowArt["div_qta"] = _clsFun.Txt2Dec(s) / 100;
            //else
            //    rowArt["div_qta"] = 0;
            rowArt["div_qta"] = 0;

            //PRZ0 Prezzo acq.  Netto   (arrotondato a due decimali)S    9  2 B    92  100
            rowArt["div_cof"] = _clsFun.Txt2Dec(strRig.Substring(77, 7)) / 1000;

            //IMN0 Importo riga                   (solo per fattura)S   11  2 B   101  111
            rowArt["div_imp"] = 0;      // _clsFun.Txt2Dec(strRig.Substring(100, 11)) / 100;

            //BAR0 Bar-codes    (ean12 senza check digit, non usare)A   16    B   112  127
            //TPR0 T.riga          (F=fatt  A=aggiornamento archivi)A    1    B   128  128
            rowArt["div_tva"] = "";         // _clsFun.Txt2Dec(strRig.Substring(127, 1));

            //PRN0 importo unitario                 (a sei decimali)S   13  6 B   129  141
            rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(77, 7)) / 1000;

            //BAR1 EAN13                                            A   16    B   142  157
            //s = strRig.Substring(141, 13);
            //if (_clsFun.Numerico(s))
            //    s = Convert.ToInt64(s).ToString();
            rowArt["div_ean"] = "";

            //PRZV Prezzo vendita consigliato                       A   11  2 B   158  168
            rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(88, 7)) / 100;

            //if (((string)rowArt["for_tgr"]).Trim() == "CA")  //Acquisto a  cartone
            //{
            //    rowArt["for_tgr"] = rowArt["for_umi"];
            //    rowArt["for_umi"] = "NR";
            //}

            return rowArt;
        }

        public DataRow DivArf(string strRig, DataRow rowArt)
        {
            string s = "";
            decimal d = 0;
            //CCD0 Codice cliente                                   A    6    B     1    6
            //NUD0 Numero fattura                 (solo per fattura)S    6  0 B     7   12
            //DT10 Data fatt. Giorno                                S    2  0 B    13   14
            //DT20 Data fatt. Mese                                  S    2  0 B    15   16
            //DT30 Data fatt. Anno                                  S    4  0 B    17   20

            //s = strRig.Substring(245, 8);
            //rowArt["div_dva"] = _clsFun.Str2Day(s);

            ////CAR0 Codice articolo                                  A    6    B    21   26
            //rowArt["div_arf"] = strRig.Substring(5, 15).Trim(); //.PadLeft(9,Convert.ToChar("0"));

            s = strRig.Substring(20, 2).Trim(); //.PadLeft(9,Convert.ToChar("0"));
            rowArt["div_stf"] = "";
            if(s=="F")
                rowArt["div_stf"] = "N";

            rowArt["div_fo2"] = strRig.Substring(38, 8).Trim();
            rowArt["div_af2"] = strRig.Substring(23, 15).Trim();

            if ((string)rowArt["div_arf"] == "1111")
                Console.WriteLine("aaaaaaaaaa");

            if (((string)rowArt["div_arf"]).Equals((string)rowArt["div_af2"]))
            {
                rowArt["div_fo2"] = "";
                rowArt["div_af2"] = "";
            }
            else
                Console.WriteLine("aaaaaaa");

            d = _clsFun.Txt2Dec(strRig.Substring(46, 7)) / 1000;
            //rowArt["div_cos"] = 0;
            if(d > 0)
                rowArt["div_cos"] = d;

            d = _clsFun.Txt2Dec(strRig.Substring(69, 7)) / 100;
            //rowArt["div_prp"] = 0;
            if (d > 0)
                rowArt["div_prp"] = d;

            rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(202, 4));

            ////FIL0 Filler                                           A    3    B    57   59
            ////UMG0 U.M. x grammatura                                A    2    B    60   61
            //rowArt["for_tgr"] = strRig.Substring(112, 2);
            ////if ((string)rowArt["for_tgr"] == "CA")
            ////    rowArt["for_tgr"] = "NR";

            //if ((string)rowArt["div_arf"] == "861083")
            //    Console.WriteLine("aaa");

            ////LV10 Livello 1 Reparto                                A    2    B    62   63
            ////LV20 Livello 2 Gruppo                                 A    2    B    64   65
            ////LV30 Livello 3 Settore                                A    2    B    66   67
            ////LV40 Livello 4 Famiglia                               A    2    B    68   69
            //rowArt["for_rep"] = strRig.Substring(39, 3);
            ////rowArt["div_ec1"] = strRig.Substring(63, 2);
            ////rowArt["div_ec2"] = strRig.Substring(65, 2);
            ////rowArt["div_ec3"] = strRig.Substring(67, 2);
            //rowArt["for_ecr"] = strRig.Substring(31, 8);

            ////IVV0 Cod. I.v.a.                                      A    2    B    70   71
            //rowArt["for_iva"] = (strRig.Substring(84, 2)).Trim().PadLeft(3, Convert.ToChar("0"));

            ////UNM0 Unità di misura                                  A    2    B    72   73
            //rowArt["for_umi"] = strRig.Substring(114, 2);

            ////PZC0 Pezzatura                                        S    4  0 B    74   77
            ////rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(73, 4));
            //rowArt["div_pxc"] = 1;

            ////PEN0 Grammatura                                       S    6  3 B    78   83
            //rowArt["div_pne"] = _clsFun.Txt2Dec(strRig.Substring(105, 7));

            ////TPD0 N=Nota accredito                                 A    1    B    84   84
            //rowArt["div_tpd"] = "";         // strRig.Substring(83, 1);

            ////QTA0 Quantità fattura               (solo per fattura)S    7  2 B    85   91
            ////s = strRig.Substring(84, 7);
            ////if (_clsFun.Numerico(s))
            ////    rowArt["div_qta"] = _clsFun.Txt2Dec(s) / 100;
            ////else
            ////    rowArt["div_qta"] = 0;
            //rowArt["div_qta"] = 0;

            ////PRZ0 Prezzo acq.  Netto   (arrotondato a due decimali)S    9  2 B    92  100
            //rowArt["div_cof"] = _clsFun.Txt2Dec(strRig.Substring(77, 7)) / 1000;

            ////IMN0 Importo riga                   (solo per fattura)S   11  2 B   101  111
            //rowArt["div_imp"] = 0;      // _clsFun.Txt2Dec(strRig.Substring(100, 11)) / 100;

            ////BAR0 Bar-codes    (ean12 senza check digit, non usare)A   16    B   112  127
            ////TPR0 T.riga          (F=fatt  A=aggiornamento archivi)A    1    B   128  128
            //rowArt["div_tva"] = "";         // _clsFun.Txt2Dec(strRig.Substring(127, 1));


            ////BAR1 EAN13                                            A   16    B   142  157
            ////s = strRig.Substring(141, 13);
            ////if (_clsFun.Numerico(s))
            ////    s = Convert.ToInt64(s).ToString();
            //rowArt["div_ean"] = "";


            //if (((string)rowArt["for_tgr"]).Trim() == "CA")  //Acquisto a  cartone
            //{
            //    rowArt["for_tgr"] = rowArt["for_umi"];
            //    rowArt["for_umi"] = "NR";
            //}

            return rowArt;
        }

        public void DivBil(string strRig, DataRow rowArt)
        {
            string sArf = strRig.Substring(5, 15).Trim();
            string sBan = strRig.Substring(22, 2).Trim();
            string sPlu = strRig.Substring(24, 4).Trim();
            string sEan = strRig.Substring(39, 13).Trim();

            if (sEan.Length == 7 && sEan.Substring(0, 2) == "20")
                sEan = sEan + "000000";

            rowArt["div_plu"] = sPlu;
            rowArt["div_reb"] = sBan;

        }

        public DataRow DivOff(string strRig, DataRow rowArt)
        {
            string s = "";
            //NUMERO OFFERTA	1	5
            //SEQUENZA	6	7 - ATTENZIONE PIU’ RIGHECON STESSO NUMERO  SE PIU EAN TRASMESSI
            //NON USATO	13	11
            //CODICE ARTICOLO	24	6
            rowArt["div_arf"] = strRig.Substring(82,15).Trim(); //.PadLeft(9,Convert.ToChar("0"));
            //UNITA’ DI VENDITA	30	2
            //rowArt["for_tgr"] = strRig.Substring(29, 2);
            ////if ((string)rowArt["for_tgr"] == "CA")
            ////    rowArt["for_tgr"] = "NR";

            ////PEZZATURA	32	4
            //rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(31, 4));
            ////NON USATO	36	3
            ////CESSIONE	39	11 - DI CUI 3 DECIMALI
            rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(112, 7)) / 100;
            ////PRZ VOLANTINO AL PUBBLICO	50	11 - DI CUI 3 DECIMALI
            rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(99, 7)) / 100;
            //MARGINALITA’	61	5 – DI CUI 2 DECIMALI
            //rowArt["div_mav"] = _clsFun.Txt2Dec(strRig.Substring(60, 5)) / 100;
            //DATA INIZIO SELL IN	66	8 - FORMATO AAAAMMGG
            //DATA FINE SELL IN	74	8 - FORMATO AAAAMMGG
            //DATA INIZIO SELL OUT	82	8 - FORMATO AAAAMMGG
            //DATA FINE SELL OUT	90	8 - FORMATO AAAAMMGG
            //NON USATO	98	8 - FORMATO AAAAMMGG
            //NON USATO	106	4
            //EAN13	110	12 - ATTENZIONE PIU’ RIGHE SE PIU EAN TRASMESSI
            //if (strRig.Length > 108)
            //    rowArt["div_ean"] = strRig.Substring(109, 13);
            ////DESCRIZIONE ARTICOLO	123	30
            //if (strRig.Length > 122)
            rowArt["div_ard"] = strRig.Substring(10, 35).Trim();

            s = strRig.Substring(61, 8);
            s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
            rowArt["div_dva"] = _clsFun.Str2Day(s);

            rowArt["off_tip"] = strRig.Substring(77, 1).Trim();
            rowArt["off_mxn"] = strRig.Substring(78, 2).Trim();
            ////UNITA’ MISURA X GRAMMATURA	153	2
            //if (strRig.Length > 152)
            //    rowArt["for_umi"] = strRig.Substring(152, 2);
            ////GRAMMATURA	155	6 DICUI 3 DECIMALI
            //if (strRig.Length > 155)
            //    rowArt["div_pne"] = _clsFun.Txt2Dec(strRig.Substring(154, 6)) / 1000;

            //if ((string)rowArt["div_arf"] == "240220" || (string)rowArt["div_arf"] == "792075")
            //    Console.WriteLine("aaaaaaaa");
            ////COD IVA	161	2 
            //if (strRig.Length > 160)
            //    rowArt["for_iva"] = strRig.Substring(160, 2);
            ////rowArt["div_iva"] = "0" + strRig.Substring(160, 2);
            ////REPARTO	163	2
            //if (strRig.Length > 161)
            //    rowArt["for_rep"] = strRig.Substring(162, 2);
            ////rowArt["div_rep"] = "0" + strRig.Substring(162, 2);
            ////GRUPPO	165	2
            //if (strRig.Length > 167)
            //    rowArt["for_ecr"] = strRig.Substring(164, 6);
            ////    rowArt["div_ec1"] = "0" + strRig.Substring(164, 2);
            //////SETTORE	167	2
            ////if (strRig.Length > 165)
            ////    rowArt["div_ec2"] = "0" + strRig.Substring(166, 2);
            //////FAMIGLIA	169	2
            ////if (strRig.Length > 167)
            ////    rowArt["div_ec3"] = "0" + strRig.Substring(168, 2);

            //if (((string)rowArt["for_tgr"]).Trim() == "CA")  //Acquisto a  cartone
            //{
            //    rowArt["for_tgr"] = rowArt["for_umi"];
            //    rowArt["for_umi"] = "NR";
            //}

            return rowArt;
        }

        public DataRow DivFat(string strRig, DataRow rowArt)
        {
            string s = "";

            //CAR0 Codice articolo                                
            rowArt["div_arf"] = strRig.Substring(43, 15).Trim();    //.PadLeft(9, Convert.ToChar("0"));

            ////BAR1 EAN13                                            A   16    B   142  157
            //s = strRig.Substring(23, 13).Trim();
            //if (_clsFun.Numerico(s))
            //    s = Convert.ToInt64(s).ToString();
            //rowArt["div_ean"] = s;

            //QTA0 Quantità fattura               (solo per fattura)S    7  2 B    85   91
            //s = strRig.Substring(36, 11);  colli
            s = strRig.Substring(114, 8);           //Pezzi
            if (_clsFun.Numerico(s))
                rowArt["div_qta"] = _clsFun.Txt2Dec(s);
            else
                rowArt["div_qta"] = 0;


            //PRN0 importo unitario                 (a sei decimali)S   13  6 B   129  141
            rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(122, 7)) / 100;

            //IVV0 Cod. I.v.a.                                      A    2    B    70   71
            rowArt["for_iva"] = "0" + strRig.Substring(136, 2);


            //PRZV Prezzo vendita consigliato                       A   11  2 B   158  168
            //Seck 20170417 tolto perchè riporta anche se il prezzo di offerta
            //rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(129, 7)) / 100;

            //DES0 Descriz. articolo                                A   30    B    27   56
            rowArt["div_ard"] = strRig.Substring(73, 35).Trim();

            rowArt["div_tva"] = _clsFun.Txt2Dec(strRig.Substring(146, 1));

            //FIL0 Filler                                           A    3    B    57   59
            //UMG0 U.M. x grammatura                                A    2    B    60   61
            rowArt["for_tgr"] = "";
            //if ((string)rowArt["for_tgr"] == "CA")
            //    rowArt["for_tgr"] = "NR";

            if ((string)rowArt["div_arf"] == "861083")
                Console.WriteLine("aaa");

            rowArt["div_imp"] = (decimal)rowArt["div_qta"] * (decimal)rowArt["div_cos"];

            //LV10 Livello 1 Reparto                                A    2    B    62   63
            //LV20 Livello 2 Gruppo                                 A    2    B    64   65
            //LV30 Livello 3 Settore                                A    2    B    66   67
            //LV40 Livello 4 Famiglia                               A    2    B    68   69
            //rowArt["for_rep"] = strRig.Substring(61, 2);
            //rowArt["div_ec1"] = strRig.Substring(63, 2);
            //rowArt["div_ec2"] = strRig.Substring(65, 2);
            //rowArt["div_ec3"] = strRig.Substring(67, 2);
            //rowArt["for_ecr"] = strRig.Substring(63, 6);

            //UNM0 Unità di misura                                  A    2    B    72   73
            rowArt["for_umi"] = strRig.Substring(108, 2);

            //PZC0 Pezzatura                                        S    4  0 B    74   77
            //rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(73, 4));

            //PEN0 Grammatura                                       S    6  3 B    78   83
            //rowArt["div_pne"] = _clsFun.Txt2Dec(strRig.Substring(77, 6)) / 1000;

            //TPD0 N=Nota accredito                                 A    1    B    84   84
            //rowArt["div_tpd"] = strRig.Substring(83, 1);


            //PRZ0 Prezzo acq.  Netto   (arrotondato a due decimali)S    9  2 B    92  100
            //rowArt["div_cof"] = _clsFun.Txt2Dec(strRig.Substring(91, 9)) / 100;

            //IMN0 Importo riga                   (solo per fattura)S   11  2 B   101  111
            //rowArt["div_imp"] = _clsFun.Txt2Dec(strRig.Substring(100, 11)) / 100;

            //BAR0 Bar-codes    (ean12 senza check digit, non usare)A   16    B   112  127
            //TPR0 T.riga          (F=fatt  A=aggiornamento archivi)A    1    B   128  128


            //if (((string)rowArt["for_tgr"]).Trim() == "CA")  //Acquisto a  cartone
            //{
            //    rowArt["for_tgr"] = rowArt["for_umi"];
            //    rowArt["for_umi"] = "NR";
            //}

            return rowArt;
        }

    }


}
