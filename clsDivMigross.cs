using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsDivMigross
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public string _strPar017EcrDivFornitori = "";

        public clsDivMigross()
        {}

        public DataRow DivArt(string strRig, DataRow rowArt)
        {
            string sTre = strRig.Substring(0, 2);

            string s = "";

            if (sTre == "73")       //Offerta
            {
                rowArt["div_dva"] = _clsFun.Str2Day(strRig.Substring(184, 8));  //OK

                //Codice articolo 
                rowArt["div_arf"] = strRig.Substring(175, 9);          //OK

                //Descriz. articolo
                rowArt["div_ard"] = strRig.Substring(13, 30).Trim();    //

                rowArt["for_tgr"] = strRig.Substring(59, 2);

                rowArt["for_rep"] = strRig.Substring(77, 3);    //OK
                rowArt["for_ecr"] = ""; // strRig.Substring(63, 6);

                rowArt["for_iva"] = "";

                rowArt["for_umi"] = "NR";
                if (strRig.Substring(231, 1) == "2")
                    rowArt["for_umi"] = "KG";

                //rowArt["for_umi"] = "NR"; // strRig.Substring(71, 2);

                rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(72, 4));

                rowArt["div_pne"] = _clsFun.Txt2Dec(strRig.Substring(65, 4)) / 100;

                rowArt["div_tpd"] = ""; // strRig.Substring(83, 1);

                //s = strRig.Substring(84, 7);
                //if (_clsFun.Numerico(s))
                //    rowArt["div_qta"] = _clsFun.Txt2Dec(s) / 100;
                //else
                rowArt["div_qta"] = 0;

                rowArt["div_cof"] = 0;  // _clsFun.Txt2Dec(strRig.Substring(91, 9)) / 100;

                rowArt["div_imp"] = 0;  // _clsFun.Txt2Dec(strRig.Substring(100, 11)) / 100;

                rowArt["div_tva"] = ""; // _clsFun.Txt2Dec(strRig.Substring(127, 1));

                rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(116, 11)) / 10000;

                //EAN13
                //s = strRig.Substring(149, 13);
                //if (_clsFun.Numerico(s))
                //    s = Convert.ToInt64(s).ToString();
                //rowArt["div_ean"] = s;

                //PRZV Prezzo vendita consigliato
                rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(96, 10)) / 1000;

                //rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(116, 11)) / 10000;

                //if (((string)rowArt["for_tgr"]).Trim() == "CA")  //Acquisto a  cartone
                //{
                //    rowArt["for_tgr"] = rowArt["for_umi"];
                //    rowArt["for_umi"] = "NR";
                //}

            }
            else
            {
                rowArt["div_dva"] = _clsFun.Str2Day(strRig.Substring(184, 8));  //OK

                //Codice articolo 
                rowArt["div_arf"] = strRig.Substring(245, 9);          //OK

                //Descriz. articolo
                rowArt["div_ard"] = strRig.Substring(13, 30).Trim();    //

                rowArt["for_tgr"] = strRig.Substring(55, 2);

                if ((string)rowArt["div_arf"] == "002366601")
                    Console.WriteLine("aaa");

                rowArt["for_rep"] = strRig.Substring(72, 3);    //OK

                rowArt["for_ecr"] = "";
                if (_strPar017EcrDivFornitori == "S")
                {
                    //s = strRig.Substring(78, 8).Trim();

                    //s = "0" + s;

                    //string sLv1 = "0" + s.Substring(0, 2);
                    //string sLv2 = "0" + s.Substring(2, 2);
                    //string sLv3 = s.Substring(4, 3);

                    //Console.WriteLine("zzz");

                    //s = sLv1 + sLv2 + sLv3;

                    s = strRig.Substring(78, 8).Trim();

                    Console.WriteLine("zzz");

                    if(s != "")
                    {
                        s = "0" + s;

                        string sLv1 = "0" + s.Substring(0, 2);
                        string sLv2 = "0" + s.Substring(2, 2);
                        string sLv3 = s.Substring(4, 3);

                        Console.WriteLine("zzz");

                        s = sLv1 + sLv2 + sLv3;
                    }

                    rowArt["for_ecr"] = s;
                }

                rowArt["for_iva"] = strRig.Substring(194, 2);

                rowArt["for_umi"] = "NR";
                if (strRig.Substring(136, 1) == "2")
                    rowArt["for_umi"] = "KG";

                //rowArt["for_umi"] = "NR"; // strRig.Substring(71, 2);

                rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(68, 4));

                rowArt["div_pne"] = _clsFun.Txt2Dec(strRig.Substring(57, 6)) / 100;

                rowArt["div_tpd"] = ""; // strRig.Substring(83, 1);

                //s = strRig.Substring(84, 7);
                //if (_clsFun.Numerico(s))
                //    rowArt["div_qta"] = _clsFun.Txt2Dec(s) / 100;
                //else
                rowArt["div_qta"] = 0;

                rowArt["div_cof"] = 0;  // _clsFun.Txt2Dec(strRig.Substring(91, 9)) / 100;

                rowArt["div_imp"] = 0;  // _clsFun.Txt2Dec(strRig.Substring(100, 11)) / 100;

                rowArt["div_tva"] = ""; // _clsFun.Txt2Dec(strRig.Substring(127, 1));

                rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(231, 11)) / 10000;

                //EAN13
                s = strRig.Substring(149, 13);
                if (_clsFun.Numerico(s))
                    s = Convert.ToInt64(s).ToString();
                rowArt["div_ean"] = s;

                //PRZV Prezzo vendita consigliato
                 rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(87, 10)) / 1000;

                //if (((string)rowArt["for_tgr"]).Trim() == "CA")  //Acquisto a  cartone
                //{
                //    rowArt["for_tgr"] = rowArt["for_umi"];
                //    rowArt["for_umi"] = "NR";
                //}
            }

            return rowArt;
        }

        public DataRow DivFat(string strRig, DataRow rowArt)
        {
            string s = "";

            //rowArt["div_arf"] = strRig.Substring(28, 7).Trim();
            rowArt["div_arf"] = strRig.Substring(28, 9).Trim();

            //EAN13
            s = strRig.Substring(116).Trim();
            if (_clsFun.Numerico(s))
                s = Convert.ToInt64(s).ToString();
            rowArt["div_ean"] = s;

            //Quantità fattura               
            s = strRig.Substring(80, 9);
            if (_clsFun.Numerico(s))
                rowArt["div_qta"] = _clsFun.Txt2Dec(s) / 100;
            else
                rowArt["div_qta"] = 0;

            //Importo unitario                 
            rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(89, 11)) / 100;

            //Cod. IVA
            rowArt["for_iva"] = "0" + strRig.Substring(72, 2);

            //Prezzo vendita consigliato
            rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(100, 11)) / 100;
            //rowArt["div_prp"] = 0;

            //Descriz. articolo
            rowArt["div_ard"] = strRig.Substring(37, 35).Trim();

            //rowArt["div_tva"] = _clsFun.Txt2Dec(strRig.Substring(91, 1));

            //FIL0 Filler                                           A    3    B    57   59
            //UMG0 U.M. x grammatura                                A    2    B    60   61
            rowArt["for_tgr"] = "GR"; // strRig.Substring(59, 2);
            //if ((string)rowArt["for_tgr"] == "CA")
            //    rowArt["for_tgr"] = "NR";

            if ((string)rowArt["div_arf"] == "861083")
                Console.WriteLine("aaa");

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
            rowArt["div_tpd"] = strRig.Substring(6, 1);


            //PRZ0 Prezzo acq.  Netto   (arrotondato a due decimali)S    9  2 B    92  100
            //rowArt["div_cof"] = _clsFun.Txt2Dec(strRig.Substring(91, 9)) / 100;

            //IMN0 Importo riga                   (solo per fattura)S   11  2 B   101  111
            rowArt["div_imp"] = (decimal)rowArt["div_qta"] * (decimal)rowArt["div_cos"];

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
