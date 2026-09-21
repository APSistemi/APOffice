using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsDivOrtoFrutticola
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public clsDivOrtoFrutticola() { }

        public DataRow DivLis(DataRow rowLis, DataRow rowArt)
        {
            string s = "";

            string sArf = (string)rowLis[0];

            string sEan = "";
            if(!DBNull.Value.Equals(rowLis[1]))
                sEan = ((string)rowLis[1]).Trim();
            if (sEan.Length > 13)
                sEan = "";

            string sArd = "";
            if (!DBNull.Value.Equals(rowLis[2]))
                sArd = (string)rowLis[2];

            decimal dCos = 0;
            if (!DBNull.Value.Equals(rowLis[3]))
                dCos = Convert.ToDecimal(rowLis[3]);

            decimal dIva = 0;
            if (!DBNull.Value.Equals(rowLis[4]))
                dIva = Convert.ToDecimal(rowLis[4]);

            decimal dPxc = 0;
            if (!DBNull.Value.Equals(rowLis[5]))
                dPxc = Convert.ToDecimal(rowLis[5]);

            if (!_clsFun.Numerico(sEan))
                sEan = "";

            //Codice articolo
            rowArt["div_arf"] = sArf.Trim();

            //EAN13
            rowArt["div_ean"] = sEan.Trim();

            //Descriz. articolo
            rowArt["div_ard"] = sArd.Trim();

            //Cod. I.v.a.                                     
            rowArt["for_iva"] = dIva.ToString("000");

            //U.M. x grammatura
            //rowArt["div_tgr"] = "PZ";
            //rowArt["div_umi"] = "NR";

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

        public DataRow DivOff(DataRow rowLis, DataRow rowArt)
        {
            string s = "";

            string sArf = (string)rowLis[0];

            string sEan = "";
            if (!DBNull.Value.Equals(rowLis[1]))
                sEan = ((string)rowLis[1]).Trim();
            if (sEan.Length > 13)
                sEan = "";
            if (!_clsFun.Numerico(sEan))
                sEan = "";

            string sArd = "";
            if (!DBNull.Value.Equals(rowLis[2]))
                sArd = (string)rowLis[2];

            decimal dCos = 0;
            if (!DBNull.Value.Equals(rowLis[3]))
                dCos = Convert.ToDecimal(rowLis[3]);

            decimal dPxc = 0;
            if (!DBNull.Value.Equals(rowLis[4]))
                dPxc = Convert.ToDecimal(rowLis[4]);

            decimal dIva = 0;
            if (!DBNull.Value.Equals(rowLis[5]))
                dIva = Convert.ToDecimal(rowLis[5]);

            //Data fine offerta di acquisto

            s = Convert.ToString(rowLis[6]);
            DateTime dSif = _clsDef.DAYOUT;
            string[] a = s.Split('/');
            if (a.Length > 2)
                dSif = new DateTime(Convert.ToInt32(((string)a[2]).Substring(0, 4)), Convert.ToInt32(a[1]), Convert.ToInt32(a[0]));

            rowArt["off_tip"] = _clsDef.OFAPRZ;

            ////Data fine offerta di acquisto 
            //rowArt["div_sif"] = _clsDef.OFAPRZ;

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

            rowArt["div_prp"] = dCos;

            ////Prezzo vendita consigliato
            //rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(123, 11)) / 1000;

            ////Grammatura
            //rowArt["div_pne"] = _clsFun.Txt2Dec(strRig.Substring(148, 4));

            ////ECR
            //rowArt["for_ecr"] = _clsFun.Txt2Dec(strRig.Substring(134, 3));
            //rowArt["for_rep"] = _clsFun.Txt2Dec(strRig.Substring(134, 3));

            rowArt["div_dva"] = DateTime.Now.ToShortDateString();
            rowArt["div_sif"] = dSif;

            //Bollini 108,6

            //Tipo modifica 115,1

            //Listino 116,4

            //Numero decimali 120,1

            //s = strRig.Substring(6, 8);
            ////s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
            //rowArt["dit_soi"] = _clsFun.Str2Day(s);

            //s = strRig.Substring(14, 8);
            ////s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
            //rowArt["dit_sof"] = _clsFun.Str2Day(s);

            return rowArt;
        }

        public DataRow DivFat(string strRig, DataRow rowArt)
        {
            string s = "";

            //1	1	4	4	N	Versione tracciato	0103
            //2	5	5	1	X	Tipo Documento	D/F/B/N
            //3	6	13	8	N	Numero documento	
            //4	14	21	8	N	Data documento	aaaammgg
            //5	22	22	1	X	Tipo Cessione	V/O
            //6	23	37	15	X	Codice articolo (interno del fornitore)	
            rowArt["div_arf"] = strRig.Substring(22, 15).Trim();

            //7	38	45	8	N	Quantità	nnnnn,nnn
            rowArt["div_qta"] = Convert.ToDecimal(strRig.Substring(37, 8).Trim()) /1000;

            //8	46	54	9	N	Costo unitario	nnnnnnn,nn
            rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(45, 9)) / 100;

            //9	55	63	9	N	Importo complessivo	nnnnnnn,nn
            rowArt["div_imp"] = _clsFun.Txt2Dec(strRig.Substring(54, 9)) / 100;

            //10	64	65	2	N	aliquota iva	
            rowArt["for_iva"] = "0" + strRig.Substring(63, 2).Trim();

            //11	66	78	13	X	codice a barre	
            rowArt["div_ean"] = strRig.Substring(65,13).Trim();

            //12	79	128	50	X	descrizione	
            rowArt["div_ard"] = strRig.Substring(78, 50).Trim();

            //13	129	130	2	X	riservato	
            //14	131	138	8	N	riservato	
            //15	139	142	4	X	riservato	
            //16	143	146	4	N	riservato	
            //17	147	155	9	N	riservato	
            //18	156	163	8	N	riservato	
            //19	164	171	8	N	riservato	
            //20	172	179	8	X	codice destinazione merce secondo Ortofrutticola	

            //rowArt["div_dva"] = dayDdo;

            //rowArt["for_umi"] = strRig.Substring(52, 2).Trim();
            //rowArt["div_pxc"] = Convert.ToDecimal(strRig.Substring(78, 5).Trim());

            return rowArt;
        }

    }
}
