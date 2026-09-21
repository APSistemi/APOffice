using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace APOffice
{
    class clsDivApDiv
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public clsDivApDiv()
        { }
    
        public DataRow DivArt(string[] aryRig, DataRow rowArt)
        {
            string sPlu = aryRig[0];
            string sArd = aryRig[1];
            string sFod = aryRig[2];
            string sPrv = aryRig[3];
            string sIva = aryRig[4];
            string sCos = aryRig[5];
            string sRep = aryRig[6];
            string sFor = aryRig[7];

            rowArt["div_for"] = sFor;

            if (sPlu.Length > 2 && sPlu.Substring(0, 2).ToLower() == "sp")
                sPlu = sPlu.Substring(2);

            decimal dPrv = 0;
            if (sPrv.Length > 1 && sPrv.Substring(0, 2) == "€ ")
                sPrv = sPrv.Substring(2);
            sPrv = sPrv.Replace(".",",");
            if (_clsFun.Numerico(sPrv))
                dPrv = Convert.ToDecimal(sPrv);

            decimal dCos = 0;
            if (sCos.Length > 1 && sCos.Substring(0, 2) == "€ ")
                sCos = sCos.Substring(2);
            sCos = sCos.Replace(".", ",");
            if (_clsFun.Numerico(sCos))
                dCos = Convert.ToDecimal(sCos);

            string sEan = "";

            if (sPlu.Length < 5)
            {
                rowArt["div_plu"] = sPlu.PadLeft(4,Convert.ToChar("0"));
                rowArt["div_reb"] = "1";
            }

            if (sIva.Length > 1 && sIva.Substring(0, 2) == "€ ")
                sIva = sIva.Substring(2);
            sIva = sIva.Replace(".", ",");
            if (sIva.Length > 0)
            {
                decimal d = Convert.ToDecimal(sIva);

                sIva = Convert.ToInt16(d).ToString();
            }
            string s = "";
            //CCD0 Codice cliente                                   A    6    B     1    6
            //NUD0 Numero fattura                 (solo per fattura)S    6  0 B     7   12
            //DT10 Data fatt. Giorno                                S    2  0 B    13   14
            //DT20 Data fatt. Mese                                  S    2  0 B    15   16
            //DT30 Data fatt. Anno                                  S    4  0 B    17   20
            //rowArt["div_dva"] = _clsFun.Str2Day(strRig.Substring(16, 4) + strRig.Substring(14, 2) + strRig.Substring(12, 2));
            rowArt["div_dva"] = DateTime.Today;


            //CAR0 Codice articolo                                  A    6    B    21   26
            //rowArt["div_arf"] = strRig.Substring(20, 6);
            rowArt["div_arf"] = sPlu;

            //DES0 Descriz. articolo                                A   30    B    27   56
            //rowArt["div_ard"] = strRig.Substring(26, 30).Trim();
            rowArt["div_ard"] = sArd.Trim();

            //FIL0 Filler                                           A    3    B    57   59
            //UMG0 U.M. x grammatura                                A    2    B    60   61
            //rowArt["for_tgr"] = strRig.Substring(59, 2);
            rowArt["for_tgr"] = "KG";
            //UNM0 Unità di misura                                  A    2    B    72   73
            rowArt["for_umi"] = "KG";       // strRig.Substring(71, 2);
            if (sRep == "11" || sRep == "14" || sRep == "15" || sRep == "18" || sRep == "21")
            {
                rowArt["for_tgr"] = "NR";
                rowArt["for_umi"] = "NR";
            }
            else
            {
                sEan = "211" + sPlu.PadLeft(4, Convert.ToChar('0')) + "000000";
                //sEan = sEan + new clsCtrlCodici().FindMod10Digit(sEan);
            }

            if ((string)rowArt["div_arf"] == "861083")
                Console.WriteLine("aaa");

            //LV10 Livello 1 Reparto                                A    2    B    62   63
            //LV20 Livello 2 Gruppo                                 A    2    B    64   65
            //LV30 Livello 3 Settore                                A    2    B    66   67
            //LV40 Livello 4 Famiglia                               A    2    B    68   69
            //rowArt["for_rep"] = strRig.Substring(61, 2);
            rowArt["for_rep"] = sRep;
            //rowArt["div_ec1"] = strRig.Substring(63, 2);
            //rowArt["div_ec2"] = strRig.Substring(65, 2);
            //rowArt["div_ec3"] = strRig.Substring(67, 2);


            //***********************rowArt["for_ecr"] = strRig.Substring(63, 6);

            //IVV0 Cod. I.v.a.                                      A    2    B    70   71
            //rowArt["for_iva"] = strRig.Substring(69, 2);
            rowArt["for_iva"] = sIva;

            //PZC0 Pezzatura                                        S    4  0 B    74   77
            //rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(73, 4));
            rowArt["div_pxc"] = 1;

            //PEN0 Grammatura                                       S    6  3 B    78   83
            rowArt["div_pne"] = 1;          // _clsFun.Txt2Dec(strRig.Substring(77, 6)) / 1000;

            //TPD0 N=Nota accredito                                 A    1    B    84   84
            rowArt["div_tpd"] = ""; // strRig.Substring(83, 1);

            //QTA0 Quantità fattura               (solo per fattura)S    7  2 B    85   91
            //////s = strRig.Substring(84, 7);
            //////if (_clsFun.Numerico(s))
            //////    rowArt["div_qta"] = _clsFun.Txt2Dec(s) / 100;
            //////else
            //////    rowArt["div_qta"] = 0;

            //PRZ0 Prezzo acq.  Netto   (arrotondato a due decimali)S    9  2 B    92  100
            //rowArt["div_cof"] = dCos; // _clsFun.Txt2Dec(strRig.Substring(91, 9)) / 100;

            //IMN0 Importo riga                   (solo per fattura)S   11  2 B   101  111
            rowArt["div_imp"] = 0; // _clsFun.Txt2Dec(strRig.Substring(100, 11)) / 100;

            //BAR0 Bar-codes    (ean12 senza check digit, non usare)A   16    B   112  127
            //TPR0 T.riga          (F=fatt  A=aggiornamento archivi)A    1    B   128  128
            rowArt["div_tva"] = "";     // _clsFun.Txt2Dec(strRig.Substring(127, 1));

            //PRN0 importo unitario                 (a sei decimali)S   13  6 B   129  141
            rowArt["div_cos"] = dCos;       // _clsFun.Txt2Dec(strRig.Substring(128, 13)) / 1000000;

            //BAR1 EAN13                                            A   16    B   142  157
            //s = strRig.Substring(141, 13);
            //if(_clsFun.Numerico(s))
            //    s = Convert.ToInt64(s).ToString();
            rowArt["div_ean"] = sEan;                     

            //PRZV Prezzo vendita consigliato                       A   11  2 B   158  168
            rowArt["div_prp"] = dPrv; // _clsFun.Txt2Dec(strRig.Substring(157, 11)) / 100;

            //if (((string)rowArt["for_tgr"]).Trim() == "CA")  //Acquisto a  cartone
            //{
            //    rowArt["for_tgr"] = rowArt["for_umi"];
            //    rowArt["for_umi"] = "NR";
            //}

            return rowArt;
        }

    }
}
