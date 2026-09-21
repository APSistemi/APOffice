using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsDivBrendolan
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public clsDivBrendolan()
        { }

        public DataRow DivArt(string strRig, DataRow rowArt)
        {
            string s = "";
            rowArt["div_dva"] = DateTime.Today;

            s = strRig.Substring(209, 8);
            if (_clsFun.Numerico(s))
                rowArt["div_dva"] = _clsFun.Str2Day(s);

            rowArt["div_arf"] = strRig.Substring(18, 13).Trim();

            s = strRig.Substring(31, 13).Trim();
            if (_clsFun.Numerico(s))
                s = Convert.ToInt64(s).ToString();
            rowArt["div_ean"] = s;

            //s = strRig.Substring(36, 11);
            //if (_clsFun.Numerico(s))
            //    rowArt["div_qta"] = _clsFun.Txt2Dec(s) / 100;
            //else
            //    rowArt["div_qta"] = 0;

            rowArt["div_ard"] = strRig.Substring(44, 50).Trim();
            rowArt["for_iva"] = "0" + strRig.Substring(109, 2);

            rowArt["for_rep"] = strRig.Substring(126, 3);
            rowArt["for_umi"] = strRig.Substring(129, 2);
            rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(131, 2));
            rowArt["div_pne"] = _clsFun.Txt2Dec(strRig.Substring(141, 6));
            rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(225, 9)) / 100;
            rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(234, 9)) / 100;

            rowArt["for_tgr"] = "GR"; // strRig.Substring(59, 2);

            //rowArt["div_tva"] = _clsFun.Txt2Dec(strRig.Substring(91, 1));

            if ((string)rowArt["div_arf"] == "861083")
                Console.WriteLine("aaa");

            return rowArt;
        }

        public DataRow DivFat(string strRig, DataRow rowArt)
        {
            string s = "";
            //CCD0 Codice cliente                                   A    6    B     1    6
            //NUD0 Numero fattura                 (solo per fattura)S    6  0 B     7   12
            //DT10 Data fatt. Giorno                                S    2  0 B    13   14
            //DT20 Data fatt. Mese                                  S    2  0 B    15   16
            //DT30 Data fatt. Anno                                  S    4  0 B    17   20
            //rowArt["div_dva"] = DateTime.Today; // _clsFun.Str2Day("20" + strRig.Substring(16, 4) + strRig.Substring(14, 2) + strRig.Substring(12, 2));

            //CAR0 Codice articolo                                  A    6    B    21   26
            rowArt["div_arf"] = strRig.Substring(8, 15).Trim();

            //BAR1 EAN13                                            A   16    B   142  157
            s = strRig.Substring(23, 13).Trim();
            if (_clsFun.Numerico(s))
                s = Convert.ToInt64(s).ToString();
            rowArt["div_ean"] = s;

            //QTA0 Quantità fattura               (solo per fattura)S    7  2 B    85   91
            //s = strRig.Substring(36, 11);  colli
            s = strRig.Substring(54, 11);           //Pezzi
            if (_clsFun.Numerico(s))
                rowArt["div_qta"] = _clsFun.Txt2Dec(s) / 100;
            else
                rowArt["div_qta"] = 0;


            //PRN0 importo unitario                 (a sei decimali)S   13  6 B   129  141
            rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(65, 11)) / 100;

            //IVV0 Cod. I.v.a.                                      A    2    B    70   71
            rowArt["for_iva"] = "0" + strRig.Substring(76, 2);


            //PRZV Prezzo vendita consigliato                       A   11  2 B   158  168
            rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(80, 11)) / 100;

            //DES0 Descriz. articolo                                A   30    B    27   56
            rowArt["div_ard"] = strRig.Substring(92, 25).Trim();

            rowArt["div_tva"] = _clsFun.Txt2Dec(strRig.Substring(91, 1));

            //FIL0 Filler                                           A    3    B    57   59
            //UMG0 U.M. x grammatura                                A    2    B    60   61
            rowArt["for_tgr"] = "GR"; // strRig.Substring(59, 2);
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
            rowArt["for_umi"] = ""; // strRig.Substring(71, 2);

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
