using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsDivApOffice
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public string _strForArf = "";

        public clsDivApOffice() { }

        public string DivVar(string strTes, string strRig, DataRow rowArt)
        {
            string s = "";
            string sEan = "";

            string[] aTes = strTes.Split('|');

            Console.WriteLine("aaaaaaaaaaaa");

            string[] aRig = strRig.Split('|');

            //int i = aTes
            //int i = aTes.in.Equals("tmp_art");

            rowArt["div_dva"] = DateTime.Today;

            //Codice articolo
            int i = Array.IndexOf(aTes, "tmp_art");
            rowArt["div_arf"] = aRig[i];
            rowArt["div_art"] = aRig[i];

            if (_strForArf == "FORARF")
            {
                //Codice articolo fornitore
                i = Array.IndexOf(aTes, "tmp_arf");
                rowArt["div_arf"] = aRig[i];
            }

            //Stato articolo
            i = Array.IndexOf(aTes, "tmp_sta");
            rowArt["div_sta"] = aRig[i];

            //Descriz. articolo
            i = Array.IndexOf(aTes, "tmp_ard");
            rowArt["div_ard"] = aRig[i];

            //Descriz. articolo
            i = Array.IndexOf(aTes, "tmp_arb");
            rowArt["div_adb"] = aRig[i];

            //Cod. I.v.a.        
            i = Array.IndexOf(aTes, "tmp_iva");
            rowArt["div_iva"] = aRig[i];
            //rowArt["for_iva"] = aRig[i];

            //UM
            i = Array.IndexOf(aTes, "tmp_umi");
            rowArt["div_umi"] = aRig[i];

            //U.M. x grammatura
            i = Array.IndexOf(aTes, "tmp_tgr");
            rowArt["div_tgr"] = aRig[i];

            //Peso netto
            i = Array.IndexOf(aTes, "tmp_pne");
            rowArt["div_pne"] = Convert.ToDecimal(aRig[i]);

            //Sfrido
            i = Array.IndexOf(aTes, "tmp_sfr");
            rowArt["art_sfr"] = Convert.ToDecimal(aRig[i]);

            //Tara
            i = Array.IndexOf(aTes, "tmp_tar");
            rowArt["art_tar"] = Convert.ToDecimal(aRig[i]);

            //Pezzatura
            i = Array.IndexOf(aTes, "tmp_pxc");
            rowArt["div_pxc"] = Convert.ToDecimal(aRig[i]);

            //Reparto
            i = Array.IndexOf(aTes, "tmp_rep");
            rowArt["div_rep"] = aRig[i];

            //Costo
            i = Array.IndexOf(aTes, "tmp_cos");
            rowArt["div_cos"] = Convert.ToDecimal(aRig[i]);

            //Prezzo vendita consigliato
            i = Array.IndexOf(aTes, "tmp_prv");
            rowArt["div_prp"] = Convert.ToDecimal(aRig[i]);

            //Listino vendita annullato
            i = Array.IndexOf(aTes, "tmp_lva");
            rowArt["div_lva"] = aRig[i];

            //ECR
            i = Array.IndexOf(aTes, "tmp_ecr");
            s = aRig[i];

            if (s.Length == 9)
            {
                rowArt["for_ecr"] = s;
                rowArt["div_ec1"] = s.Substring(0, 3);
                rowArt["div_ec2"] = s.Substring(3, 3);
                rowArt["div_ec3"] = s.Substring(6, 3);
            }

            //Fornitore consegnatario
            i = Array.IndexOf(aTes, "tmp_for");
            rowArt["div_fo2"] = aRig[i];

            //Fornitore descrizione
            i = Array.IndexOf(aTes, "tmp_fod");
            rowArt["div_fod"] = aRig[i];

            ////EAN13
            //i = Array.IndexOf(aTes, "tmp_ean");
            //rowArt["div_ean"] = aRig[i];

            //A pz in bilancia
            rowArt["div_bil"] = "";
            i = Array.IndexOf(aTes, "tmp_bil");
            if ((string)aRig[i] == "1")
                rowArt["div_bil"] = "S";

            //PLU
            i = Array.IndexOf(aTes, "tmp_plu");
            rowArt["div_plu"] = aRig[i];

            //Reparto bilancia
            i = Array.IndexOf(aTes, "tmp_reb");
            rowArt["div_reb"] = aRig[i];

            //Origine
            i = Array.IndexOf(aTes, "tmp_ori");
            rowArt["div_ori"] = aRig[i];

            //Calibro
            i = Array.IndexOf(aTes, "tmp_cal");
            rowArt["div_cal"] = aRig[i];

            //Categoria
            i = Array.IndexOf(aTes, "tmp_cat");
            rowArt["div_cat"] = aRig[i];

            //A pz in bilancia
            i = Array.IndexOf(aTes, "tmp_bpz");
            if ((string)aRig[i] == "1")
                rowArt["div_bpz"] = "S";
            else
                rowArt["div_bpz"] = "";


           // x["ean_bil"] = (Boolean)j2[i]["die_bil"];

            //rowArt["die_bil"] = true;


            i = Array.IndexOf(aTes, "ean_ean");
            sEan = aRig[i];

            //return rowArt;
            return sEan;
        }

        public string DivVarOff(string strTes, string strRig, DataRow rowArt)
        {
            string s = "";
            string sEan = "";

            string[] aTes = strTes.Split('|');

            Console.WriteLine("aaaaaaaaaaaa");

            string[] aRig = strRig.Split('|');

            //int i = aTes
            //int i = aTes.in.Equals("tmp_art");

            rowArt["div_dva"] = DateTime.Today;

            //Codice articolo
            int i = Array.IndexOf(aTes, "tmp_art");
            rowArt["div_arf"] = aRig[i];
            rowArt["div_art"] = aRig[i];

            //Stato articolo
            i = Array.IndexOf(aTes, "tmp_sta");
            rowArt["div_sta"] = aRig[i];

            //Descriz. articolo
            i = Array.IndexOf(aTes, "tmp_ard");
            rowArt["div_ard"] = aRig[i];

            //Descriz. articolo
            i = Array.IndexOf(aTes, "tmp_arb");
            rowArt["div_adb"] = aRig[i];

            //Cod. I.v.a.        
            i = Array.IndexOf(aTes, "tmp_iva");
            rowArt["div_iva"] = aRig[i];
            //rowArt["for_iva"] = aRig[i];

            //UM
            i = Array.IndexOf(aTes, "tmp_umi");
            rowArt["div_umi"] = aRig[i];

            //U.M. x grammatura
            i = Array.IndexOf(aTes, "tmp_tgr");
            rowArt["div_tgr"] = aRig[i];

            //Peso netto
            i = Array.IndexOf(aTes, "tmp_pne");
            rowArt["div_pne"] = Convert.ToDecimal(aRig[i]);

            //Sfrido
            i = Array.IndexOf(aTes, "tmp_sfr");
            rowArt["art_sfr"] = Convert.ToDecimal(aRig[i]);

            //Tara
            i = Array.IndexOf(aTes, "tmp_tar");
            rowArt["art_tar"] = Convert.ToDecimal(aRig[i]);

            //Pezzatura
            i = Array.IndexOf(aTes, "tmp_pxc");
            rowArt["div_pxc"] = Convert.ToDecimal(aRig[i]);

            //Reparto
            i = Array.IndexOf(aTes, "tmp_rep");
            rowArt["div_rep"] = aRig[i];

            //Costo
            i = Array.IndexOf(aTes, "tmp_cos");
            rowArt["div_cos"] = Convert.ToDecimal(aRig[i]);

            //Prezzo vendita consigliato
            i = Array.IndexOf(aTes, "tmp_prv");
            rowArt["div_prp"] = Convert.ToDecimal(aRig[i]);

            //ECR
            i = Array.IndexOf(aTes, "tmp_ecr");
            s = aRig[i];

            if (s.Length == 9)
            {
                rowArt["for_ecr"] = s;
                rowArt["div_ec1"] = s.Substring(0, 3);
                rowArt["div_ec2"] = s.Substring(3, 3);
                rowArt["div_ec3"] = s.Substring(6, 3);
            }

            //Fornitore consegnatario
            i = Array.IndexOf(aTes, "tmp_for");
            rowArt["div_fo2"] = aRig[i];

            ////EAN13
            //i = Array.IndexOf(aTes, "tmp_ean");
            //rowArt["div_ean"] = aRig[i];

            //A pz in bilancia
            rowArt["div_bil"] = "";
            i = Array.IndexOf(aTes, "tmp_bil");
            if ((string)aRig[i] == "1")
                rowArt["div_bil"] = "S";

            //PLU
            i = Array.IndexOf(aTes, "tmp_plu");
            rowArt["div_plu"] = aRig[i];

            //Reparto bilancia
            i = Array.IndexOf(aTes, "tmp_reb");
            rowArt["div_reb"] = aRig[i];

            //Origine
            i = Array.IndexOf(aTes, "tmp_ori");
            rowArt["div_ori"] = aRig[i];

            //Calibro
            i = Array.IndexOf(aTes, "tmp_cal");
            rowArt["div_cal"] = aRig[i];

            //Categoria
            i = Array.IndexOf(aTes, "tmp_cat");
            rowArt["div_cat"] = aRig[i];

            //A pz in bilancia
            i = Array.IndexOf(aTes, "tmp_bpz");
            if ((string)aRig[i] == "1")
                rowArt["div_bpz"] = "S";
            else
                rowArt["div_bpz"] = "";


            //Dati offerta

            //Tipo
            i = Array.IndexOf(aTes, "ofa_tip");
            rowArt["off_tip"] = aRig[i];

            //Valore
            i = Array.IndexOf(aTes, "ofa_val");
            rowArt["div_prp"] = aRig[i];

            //Annullato
            i = Array.IndexOf(aTes, "ofa_ann");
            

            //rowArt["off_ann"] = aRig[i];

            rowArt["off_ann"] = false;
            s = aRig[i];
            if (s == "1")
                rowArt["off_ann"] = true;

            // x["ean_bil"] = (Boolean)j2[i]["die_bil"];

            //rowArt["die_bil"] = true;


            i = Array.IndexOf(aTes, "ean_ean");
            sEan = aRig[i];

            //return rowArt;
            return sEan;
        }

        public DataRow DivArt(string strRig, DataRow rowArt)
        {
            string s = "";


            //MERCATO NUMERICO 6  codice del negozio

            //CODVARIAZ ALFANUM. 2
            /*
            10 Cancellazione di un articolo
            15 Cancellazione di un codice di vendita
            19 Inserimento/Modifica dati articolo
            20 Inserimento di un nuovo articolo
            21 Modifica dati di un articolo
            22 Modifica costo di un articolo
            23 Modifica prezzo di un articolo
            25 Inserimento di un nuovo codice di vendita
            30 Inserimento di una nuova relazione articolo-fornitore
            31 Modifica dati di una relazione articolo-fornitore
            32 Cancellazione di una relazione articolo-fornitore
            */

            s = strRig.Substring(6, 2).Trim(); 
            rowArt["div_stf"] = "";
            if (s == "10" || s == "15" || s == "32")
                rowArt["div_stf"] = "N";

            //DATAVALID DATA 8
            s = strRig.Substring(8, 8).Trim();
            rowArt["div_dva"] = _clsFun.Str2Day(s);

            //COD_ARTI ALFANUM. 14
            rowArt["div_arf"] = strRig.Substring(16, 7).Trim();

            if ((string)rowArt["div_arf"] == "0012872")
                Console.WriteLine("aaaa");

            //CODVEND NUMERICO 13
            rowArt["div_ean"] = strRig.Substring(30, 13).Trim();

            //DESCRIZIO ALFANUM. 30
            rowArt["div_ard"] = strRig.Substring(43, 30).Trim();

            //SGOCCIOLAT NUMERICO 6
            rowArt["div_pne"] = _clsFun.Txt2Dec(strRig.Substring(79, 6));

            if ((decimal)rowArt["div_pne"] == 0)
            {
                //PEZZATURA NUMERICO 6  ( sta peso unitario)
                s = strRig.Substring(73, 6);
                rowArt["div_pne"] = Convert.ToDecimal(s);
            }


            //STATO ALFANUM. 2
            /*
            0: ATTIVO
            1: IN ATTESA
            2: SOSPESO
            8: CANCELLATO
            9: FUORI ASSORTIMENTO
            */
            s = strRig.Substring(85, 2).Trim(); 
            if(s == "2" || s == "8" || s == "9")
                rowArt["div_stf"] = "N";

            //UNITA_MIS ALFANUM. 2
            rowArt["for_tgr"] = strRig.Substring(87, 2);

            //COD_IVA ALFANUM. 2
            //Cod. I.v.a.                                     
            rowArt["for_iva"] = (strRig.Substring(89, 2)).Trim().PadLeft(3, Convert.ToChar("0"));

            //FLAG_PESO CARATTERE 1
            rowArt["for_umi"] = "NR";
            s = strRig.Substring(91, 1);
            if (s == "S")
                rowArt["for_umi"] = "KG";

            //CAUZIONE NUMERICO 14

            //TESTOCASSA ALFANUM. 20
            s = strRig.Substring(106, 20);
            rowArt["div_adb"] = s;

            //TESTO_ETI ALFANUM. 30

            //STATO_PRE NUMERICO 1

            //EXT-U-MERC ALFANUM. 12
            rowArt["for_ecr"] = strRig.Substring(157, 12).Trim();
            rowArt["for_rep"] = (string)rowArt["for_ecr"];

            //TIPO_ETI ALFANUM. 1

            //PREZZO NUMERICO 8
            rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(170, 8)) / 100;

            //COSTO NUMERICO 8
            rowArt["div_cos"] = 0;  // _clsFun.Txt2Dec(strRig.Substring(178, 8)) / 100;

            //COD_FORNIT NUMERICO 6

            //CODARTFORN ALFANUM. 14

            //CONFEZIONE NUMERICO 5

            //FORN_PRI NUMERICO 1

            //MULT NUMERICO 3

            //MIX NUMERICO 3
            //REPARTO NUMERICO 4
            //CODVENDPOS NUMERICO 13
            //COD_PERMA ALFANUM. 2
            //TCALPREMED ALFANUM. 2
            //COD_POLIT ALFANUM. 2
            //GIA_MIN NUMERICO 8
            //GIA_MAX NUMERICO 8
            //LIV_RIORDI NUMERICO 5
            //GIACENZA CARATTERE 1
            //TEMPO_APPR NUMERICO 3
            //ORDINE_MIN NUMERICO 8
            //ORD_FORN NUMERICO 3
            //NUM_ETI ALFANUM. 2
            //BOLLINI NUMERICO 6
            //FLAG_DISP ALFANUM. 10
            //LISTINO_P ALFANUM. 4
            //LISTINO_C ALFANUM. 4
            //DECLIST_P NUMERICO 1
            //DECLIST_C NUMERICO 1
            //DESCRIZIOL ALFANUM. 40
            //TESTO_ETIL ALFANUM. 40
            //PRZZERO ALFANUM. 1
            //LINKBILA ALFANUM. 1
            //LINKTARA NUMERICO 6


            return rowArt;
        }

        public DataRow DivOff(string strRig, DataRow rowArt)
        {
            string s = "";

            /* TIPO PROMO
              
              Promozioni normali
              - 1 MxN 
              - 2 Promozione sconto %
              - 3 Promozione taglio prezzo
              - 4 Promozione bollini
             
              Promozioni fidelity
              - 1 Mxn
              - 2 Premio
              - 3 SET
             
            */

            rowArt["off_tip"] = _clsDef.OFAPRZ;

            s = strRig.Substring(27, 2).Trim();
            if (s == "02")
            {
                rowArt["off_tip"] = _clsDef.OFASCO;

                //Valore
                rowArt["div_imp"] = _clsFun.Txt2Dec(strRig.Substring(90, 6));
            }
            //Descrizione 29,30
            //Sconto 59,7 (2 dec)

            //Codice articolo
            rowArt["div_arf"] = strRig.Substring(66, 14).Trim();


            if ((string)rowArt["div_arf"] == "0416370")
                Console.WriteLine("aaaaaaaa");

            //Codice Mix  80,3

            //Sconto MIx 83,8

            //Multiplo 91,3

            //Percentuale sconto 94,7 (2 dec)
            /*
             1 -> 5
             2 -> 10
             3 -> 15
             4 -> 20
             5 -> 25
             6 -> 30
             7 -> 40
             8 -> 50
             9 -> 33
             */



            //Prezzo  
            rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(101, 8)) / 100;           


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

            //Data
            s = strRig.Substring(10, 8);
            rowArt["div_dva"] = _clsFun.Str2Day(s);

            //Codice articolo                                
            rowArt["div_arf"] = strRig.Substring(63, 7).Trim();   

            //Descriz. articolo              
            rowArt["div_ard"] = strRig.Substring(70, 35).Trim();

            //Quantità fattura               
            s = strRig.Substring(105, 9);          
            if (_clsFun.Numerico(s))
                rowArt["div_qta"] = Convert.ToDecimal(s)/100;
            else
                rowArt["div_qta"] = 0;

            //Importo
            decimal d = Convert.ToDecimal( strRig.Substring(115, 11))/1000; 
            rowArt["div_imp"] = d;

            //IVA
            rowArt["for_iva"] = "0" + strRig.Substring(127, 2);

            //Tipo documento
            rowArt["div_tva"] = strRig.Substring(129, 2);

            //Importo unitario         
            rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(131, 11)) / 1000;

            //Pezzatura NON C'E'
            //rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(73, 4));

            ////PRZV Prezzo vendita consigliato                       A   11  2 B   158  168
            //rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(129, 7)) / 100;

            ////FIL0 Filler                                           A    3    B    57   59
            ////UMG0 U.M. x grammatura                                A    2    B    60   61
            //rowArt["for_tgr"] = "";
            ////if ((string)rowArt["for_tgr"] == "CA")
            ////    rowArt["for_tgr"] = "NR";

            //if ((string)rowArt["div_arf"] == "861083")
            //    Console.WriteLine("aaa");

            ////LV10 Livello 1 Reparto                                A    2    B    62   63
            ////LV20 Livello 2 Gruppo                                 A    2    B    64   65
            ////LV30 Livello 3 Settore                                A    2    B    66   67
            ////LV40 Livello 4 Famiglia                               A    2    B    68   69
            ////rowArt["for_rep"] = strRig.Substring(61, 2);
            ////rowArt["div_ec1"] = strRig.Substring(63, 2);
            ////rowArt["div_ec2"] = strRig.Substring(65, 2);
            ////rowArt["div_ec3"] = strRig.Substring(67, 2);
            ////rowArt["for_ecr"] = strRig.Substring(63, 6);

            ////UNM0 Unità di misura                                  A    2    B    72   73
            //rowArt["for_umi"] = strRig.Substring(108, 2);

            ////PZC0 Pezzatura                                        S    4  0 B    74   77
            ////rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(73, 4));

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
