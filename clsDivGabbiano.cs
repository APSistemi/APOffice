using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsDivGabbiano
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public clsDivGabbiano() { }

        public DataRow DivLis(DataRow rowLis, DataRow rowArt)
        {
            string s = "";

            string sRep = "";
            if (!DBNull.Value.Equals(rowLis[0]))
                sRep = ((string)rowLis[0]).Trim();

            string sEan = "";
            if (!DBNull.Value.Equals(rowLis[2]))
            {
                double dd = Convert.ToDouble(rowLis[2]);
                sEan =  dd.ToString();
            }

            string sArf = Convert.ToString(rowLis[3]);

            string sArd = "";
            if (!DBNull.Value.Equals(rowLis[7]))
            {
                sArd = (string)rowLis[7];
                if (sArd.Length > 50)
                    sArd = sArd.Substring(0, 50);
            }

            decimal dPxc = 0;
            if (!DBNull.Value.Equals(rowLis[5]))
                dPxc = Convert.ToDecimal(rowLis[5]);

            decimal dIva = 22;
            if (!DBNull.Value.Equals(rowLis[10]))
                dIva = Convert.ToDecimal(rowLis[10]);

            decimal dCos = 0;
            if (!DBNull.Value.Equals(rowLis[13]))
                dCos = Math.Round( Convert.ToDecimal(rowLis[13]),3,MidpointRounding.AwayFromZero);

            decimal dPrp = 0;
            if (!DBNull.Value.Equals(rowLis[12]) && Convert.ToDecimal(rowLis[12]) > 0)
                dPrp = Convert.ToDecimal(rowLis[12]);

            //Reparto

            rowArt["for_rep"] = "";
            if(sRep.Length >= 3)
                sRep.Substring(0, 3);

            //Codice articolo
            rowArt["div_arf"] = sArf.Trim();

            //EAN13
            rowArt["div_ean"] = sEan.Trim();

            //Descriz. articolo
            rowArt["div_ard"] = sArd.Trim();

            //Cod. I.v.a.                                     
            rowArt["for_iva"] = dIva.ToString("000");

            //U.M. x grammatura
            rowArt["for_tgr"] = "NR";

            //Pezzatura
            rowArt["div_pxc"] = dPxc;

            //Costo
            rowArt["div_cos"] = dCos;

            //Prezzo vendita consigliato
            rowArt["div_prp"] = dPrp;

            ////Grammatura
            //rowArt["div_pne"] = _clsFun.Txt2Dec(strRig.Substring(148, 4));

            ////ECR
            //rowArt["for_ecr"] = _clsFun.Txt2Dec(strRig.Substring(134, 3));
            //rowArt["for_rep"] = _clsFun.Txt2Dec(strRig.Substring(134, 3));

            rowArt["div_dva"] = DateTime.Now.ToShortDateString();

            ////Reparto
            //rowArt["for_rep"] = strRig.Substring(39, 3).Trim().PadLeft(3, Convert.ToChar("0"));

            ////UMI
            //rowArt["for_umi"] = strRig.Substring(114, 2);

            ////TPD0 N=Nota accredito                                 A    1    B    84   84
            //rowArt["div_tpd"] = "";         // strRig.Substring(83, 1);

            //rowArt["div_qta"] = 0;

            ////PRZ0 Prezzo acq.  Netto   (arrotondato a due decimali)S    9  2 B    92  100
            //rowArt["div_cof"] = _clsFun.Txt2Dec(strRig.Substring(77, 7)) / 1000;

            ////IMN0 Importo riga                   (solo per fattura)S   11  2 B   101  111
            //rowArt["div_imp"] = 0;      // _clsFun.Txt2Dec(strRig.Substring(100, 11)) / 100;

            ////BAR0 Bar-codes    (ean12 senza check digit, non usare)A   16    B   112  127
            ////TPR0 T.riga          (F=fatt  A=aggiornamento archivi)A    1    B   128  128
            //rowArt["div_tva"] = "";         // _clsFun.Txt2Dec(strRig.Substring(127, 1));

            //s = strRig.Substring(245, 8);

            //rowArt["div_dva"] = _clsFun.Str2Day(s);

            //s = strRig.Substring(21, 1);
            //if (s == "T")
            //    rowArt["div_sta"] = _clsDef.STAATT;

            //if ((string)rowArt["div_arf"] == "1111")
            //    Console.WriteLine("aaa");

            ////LV10 Livello 1 Reparto                                A    2    B    62   63
            ////LV20 Livello 2 Gruppo                                 A    2    B    64   65
            ////LV30 Livello 3 Settore                                A    2    B    66   67
            ////LV40 Livello 4 Famiglia                               A    2    B    68   69

            //UNM0 Unità di misura                                  A    2    B    72   73

            //if (((string)rowArt["for_tgr"]).Trim() == "CA")  //Acquisto a  cartone
            //{
            //    rowArt["for_tgr"] = rowArt["for_umi"];
            //    rowArt["for_umi"] = "NR";
            //}

            return rowArt;
        }


    }
}
