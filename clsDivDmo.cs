using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsDivDmo
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public clsDivDmo() { }

        public DataRow DivLis(DataRow rowLis, DataRow rowArt)
        {
            string s = "";

            string sArf = (string)rowLis[0];

            string sArd = "";
            if (!DBNull.Value.Equals(rowLis[1]))
                sArd = (string)rowLis[1];

            decimal dCos = 0;
            if (!DBNull.Value.Equals(rowLis[2]))
                dCos = Convert.ToDecimal(rowLis[2]);

            decimal dPxc = 0;
            if (!DBNull.Value.Equals(rowLis[4]))
                dPxc = Convert.ToDecimal(rowLis[4]);

            string sEan = "";
            if (!DBNull.Value.Equals(rowLis[5]))
                sEan = ((string)rowLis[5]).Trim();
            if (sEan.Length > 13)
                sEan = "";

            if (!_clsFun.Numerico(sEan))
                sEan = "";

            //decimal dIva = 22;
            //if (!DBNull.Value.Equals(rowLis[4]))
            //    dIva = Convert.ToDecimal(rowLis[4]);

            //Codice articolo
            rowArt["div_arf"] = "000" + sArf.Trim();

            //EAN13
            rowArt["div_ean"] = sEan.Trim();

            //Descriz. articolo
            rowArt["div_ard"] = sArd.Trim();

            ////Cod. I.v.a.                                     
            //rowArt["for_iva"] = dIva.ToString("000");

            //U.M. x grammatura
            rowArt["div_tgr"] = "PZ";
            rowArt["div_umi"] = "NR";

            //Pezzatura
            rowArt["div_pxc"] = dPxc;

            //Costo
            rowArt["div_cos"] = dCos;

            ////Prezzo vendita consigliato
            //rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(123, 11)) / 1000;

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

        public DataRow DivFat(string strRig, DataRow rowArt)
        {
            string s = "";

            //CPCO	S	6	0	PUNTO VENDITA	1
            //CART	S	9	0	CODICE ARTICOLO	7
            rowArt["div_arf"] = strRig.Substring(6, 9).Trim();

            //CEAN	A	20	 	CODICE A BARRE PIÙ RECENTE	16
            rowArt["div_ean"] = strRig.Substring(15, 13).Trim();

            //XART	A	40	 	DESCRIZIONE PRIMARIA ARTICOLO	36
            rowArt["div_ard"] = strRig.Substring(35, 40).Trim();

            //QTMO	S	6	0	QUANTITA SPEDITA DA MAGAZZINO IN PEZZI	76
            rowArt["div_qta"] = Convert.ToDecimal(strRig.Substring(75, 6).Trim());

            //NDOC	A	20	 	NUMERO DEL DOCUMENTO	82
            //TDOC	S	8	0	DATA DEL DOCUMENTO	102
            //SNAV	A	1	 	STANDARD OFFERTA	110

            //CFOR	S	6	0	CODICE FORNITORE	111
            //PCES	S	9	3	PREZZO DI CESSIONE	117
            rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(116, 9)) / 1000;

            //PVEN	S	9	3	PREZZO AL PUBBLICO	126
            rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(125, 9)) / 1000;

            rowArt["div_imp"] = (decimal)rowArt["div_cos"] * (decimal)rowArt["div_qta"];


            return rowArt;
        }

    }
}
