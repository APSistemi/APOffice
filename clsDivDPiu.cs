using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    class clsDivDPiu
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public clsDivDPiu() { }

        public DataRow DivLis(string strRig, DataRow rowArt)
        {
            string s = "";

            //Codice articolo
            rowArt["div_arf"] = strRig.Substring(6, 7).Trim(); 

            //EAN13
            rowArt["div_ean"] = strRig.Substring(13, 13).Trim();

            //Descriz. articolo
            rowArt["div_ard"] = strRig.Substring(26, 35).Trim();

            //Cod. I.v.a.                                     
            rowArt["for_iva"] = (strRig.Substring(96, 2)).Trim().PadLeft(3, Convert.ToChar("0"));

            //U.M. x grammatura
            rowArt["for_tgr"] = strRig.Substring(98, 2);
            //if ((string)rowArt["for_tgr"] == "CA")
            //    rowArt["for_tgr"] = "NR";

            //Pezzatura
            s = strRig.Substring(106, 6);
            rowArt["div_pxc"] = Convert.ToDecimal(s);

            //Costo
            rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(112, 11)) / 1000;

            //Prezzo vendita consigliato
            rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(123, 11)) / 1000;

            //Grammatura
            rowArt["div_pne"] = _clsFun.Txt2Dec(strRig.Substring(148, 4));

            //ECR
            rowArt["for_ecr"] = _clsFun.Txt2Dec(strRig.Substring(134, 3));
            rowArt["for_rep"] = _clsFun.Txt2Dec(strRig.Substring(134, 3));
            
            s = strRig.Substring(138, 8);
            rowArt["div_dva"] = _clsFun.Str2Day(s.Substring(4, 4) + s.Substring(2, 2) + s.Substring(0, 2));

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

            if ((string)rowArt["div_arf"] == "0428330")
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

        //public DataRow DivArt_2(string strRig, DataRow rowArt)
        //{
        //    string s = "";

        //    //CDA Codice articolo 
        //    //A 7 1 7

        //    rowArt["div_arf"] = strRig.Substring(46, 7).Trim();

        //    //DES Descrizione 
        //    //A 30 8 37
        //    rowArt["div_ard"] = strRig.Substring(53, 30).Trim();


        //    //MUL Multiplo.Numero di singole confezioni che compongono l’unita’ di vendita 
        //    //S 4,00 38 41
        //    // ???? rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(170, 8)) / 100;

        //    //CON Confezionamento Cfr. tab. 0005 
        //    //A 2 42 43

        //    //TPG Tipo grammatura (o unita’ di misura) Cr. Tab. 0006 
        //    //A 2 44 45
        //    rowArt["for_tgr"] = strRig.Substring(89, 2);

        //    //GRA Grammatura 
        //    //S 4,00 46 49
        //    rowArt["div_pne"] = _clsFun.Txt2Dec(strRig.Substring(91, 4));

        //    //SGO Sgocciolato 
        //    //S 4,00 50 53

        //    //MER Codice merceologia Crf. Tab. 0127 livello 01
        //    //A 12 54 65
        //    rowArt["for_ecr"] = strRig.Substring(99, 12).Trim();
        //    //rowArt["for_rep"] = (string)rowArt["for_ecr"];

        //    //TPA Tipo articolo in acquisto ‘1’ a numero - ‘2’ a peso
        //    //A 1 66 66
        //    rowArt["for_umi"] = "NR";
        //    s = strRig.Substring(111, 1);
        //    if (s == "2")
        //        rowArt["for_umi"] = "KG";

        //    //TRA Tipo merce Cfr. tab. 0014 
        //    //A 1 67 67

        //    //FOR Fornitore abituale 
        //    //A 6 68 73

        //    //UMS Tipo scadenza ‘GG’ giorni
        //    //‘SS’ settimane
        //    //‘MM’ mesi
        //    //‘AA’ anni
        //    //A 2 74 75

        //    //DUR Durata.
        //    //Precisa l’ampiezza del periodo definito dal
        //    //tipo di scadenza dal momento del
        //    //ricevimento della merce presso l’azienda
        //    //distributrice.
        //    //S 3,00 76 78

        //    //CAL Calo peso Non usato 
        //    //S 4,02 79 82

        //    //MAR Marchio 
        //    //A 12 83 94

        //    //FMA Indicatore prodotto a marchio interno ‘S’ - Si ‘N’ – No
        //    //A 1 95 95

        //    //RA1 Codice articolo vuoto 1 Non usato 
        //    //A 7 96 102

        //    //CO1 Numero elementi in vuoto 1 Non usato 
        //    //S 4,00 103 106

        //    //RA2 Codice articolo vuoto 2 Non usato 
        //    //A 7 107 113

        //    //CO2 Numero elementi in vuoto 2 Non usato 
        //    //S 4,00 114 117

        //    //VAR Codice PLU bilance 
        //    //A 4 118 121

        //    //FEI Libero Non usato 
        //    //A 1 122 122

        //    //ELS Tipo etichetta scaffale.
        //    //Assume dei valori che nel tempo possono mutare. Non usato. Dato ad uso interno della Commerciale Brendolan.
        //    //A 1 123 123

        //    //ETS Spedizione bilance S/N ‘S’ se presente VAR, altrimenti ‘N’
        //    //A 1 124 124

        //    //DEB Banchi bilance 
        //    //A 2 125 126

        //    //SAA Anno stagionalita' 
        //    //A 4 127 130

        //    //STP Codice stagionalita' Cfr. tab. 0047 
        //    //A 3 131 133

        //    //FL3 Stato articolo ‘1’ valido ‘0’ annullato
        //    //A 1 134 134
        //    s = strRig.Substring(179, 1).Trim();
        //    if (s == "1")
        //        rowArt["div_stf"] = "N";

        //    //DC1 Tipologia vini DOC Non fornita la tabella di controllo e decodifica
        //    //A 4 135 138

        //    //DC2 Provenienza vini DOC Non fornita la tabella di controllo e decodifica
        //    //A 4 139 142
        //    rowArt["for_iva"] = (strRig.Substring(188, 2)).Trim().PadLeft(3, Convert.ToChar("0"));

        //    //CDI Codice iva Cfr. tab. 0007 
        //    //A 2 143 144

        //    //DVI Data validita' iva Forma 'aaaammgg' 
        //    //S 8,00 145 152

        //    //OCI Codice iva precedente 
        //    //A 2 153 154

        //    //PEN Aliquota encc 
        //    //S 5,02 155 159
            
        //    //CTR Valore contrassegno di stato e ifa 
        //    //S 2,03 160 171

        //    //GAN Gradi anidri 
        //    //S 4,00 172 175

        //    //VOU Codice UTIF Cfr. tab. 0009 
        //    //A 2 176 177

        //    //UMU UTIF Grammatura Cfr. tab. 0006 
        //    //A 2 178 179

        //    //GRU UTIF Valore grammatura 
        //    //S 5,00 180 184

        //    //SCL Descrizione scontrino 
        //    //A 20 185 204
        //    rowArt["div_adb"] = (strRig.Substring(230, 20)).Trim();

        //    //TPV Tipo articolo in vendita ‘1’ a numero ‘2’ a peso
        //    //A 1 205 205

        //    //FPP Flag primo prezzo Fisso ‘N’ 
        //    //A 1 206 206

        //    //VA2 Secondo PLU bilancia 
        //    //A 4 207 210

        //    //FPB Indicatore prezzo bloccato Fisso ‘N’ 
        //    //A 1 211 211

        //    //PEZ Pezzi per collo in vendita da deposito 
        //    //S 4,00 212 215

        //    //BOL Soggetto a stampa bolla 
        //    //A 1 216 216

        //    //PZC Pezzi per collo in acquisto da fornitore 
        //    //S 4,00 217 220
        //    rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(262, 4));

        //    //DRF Deposito rifornitore 
        //    //A 4 221 224

        //    //BLK Pallet: colli per piano 
        //    //S 5,00 225 229

        //    //TIR Pallet: numero di piani 
        //    //S 5,00 230 234

        //    //CLM Pallet: colli eccedenti a blk & tir 
        //    //S 5,00 235 239

        //    //RPE Rilevazione pesata Uso interno Commerciale Brendolan
        //    //A 1 240 240

        //    //DAE Descrizione aggiuntiva etichette Cfr. tab. BR67. Il campo va scomposto in singoli caratteri
        //    //A 4 241 244

        //    //s = strRig.Substring(6, 2).Trim();
        //    //rowArt["div_stf"] = "";
        //    //if (s == "10" || s == "15" || s == "32")
        //    //    rowArt["div_stf"] = "N";

        //    //DATAVALID DATA 8
        //    //s = strRig.Substring(8, 8).Trim();
        //    //rowArt["div_dva"] = _clsFun.Str2Day(s);

        //    //COD_ARTI ALFANUM. 14
        //    //rowArt["div_arf"] = strRig.Substring(16, 7).Trim();

        //    //if ((string)rowArt["div_arf"] == "0012872")
        //    //    Console.WriteLine("aaaa");

        //    ////CODVEND NUMERICO 13
        //    //rowArt["div_ean"] = strRig.Substring(30, 13).Trim();

        //    //////DESCRIZIO ALFANUM. 30
        //    ////rowArt["div_ard"] = strRig.Substring(43, 30).Trim();

        //    ////SGOCCIOLAT NUMERICO 6
        //    //rowArt["div_pne"] = _clsFun.Txt2Dec(strRig.Substring(79, 6));

        //    //if ((decimal)rowArt["div_pne"] == 0)
        //    //{
        //    //    //PEZZATURA NUMERICO 6  ( sta peso unitario)
        //    //    s = strRig.Substring(73, 6);
        //    //    rowArt["div_pne"] = Convert.ToDecimal(s);
        //    //}


        //    ////STATO ALFANUM. 2
        //    ///*
        //    //0: ATTIVO
        //    //1: IN ATTESA
        //    //2: SOSPESO
        //    //8: CANCELLATO
        //    //9: FUORI ASSORTIMENTO
        //    //*/
        //    ////s = strRig.Substring(85, 2).Trim();
        //    ////if (s == "2" || s == "8" || s == "9")
        //    ////    rowArt["div_stf"] = "N";

        //    ////UNITA_MIS ALFANUM. 2
        //    //rowArt["for_tgr"] = strRig.Substring(87, 2);

        //    ////COD_IVA ALFANUM. 2
        //    ////Cod. I.v.a.                                     
        //    //rowArt["for_iva"] = (strRig.Substring(89, 2)).Trim().PadLeft(3, Convert.ToChar("0"));

        //    ////FLAG_PESO CARATTERE 1
        //    //rowArt["for_umi"] = "NR";
        //    //s = strRig.Substring(91, 1);
        //    //if (s == "S")
        //    //    rowArt["for_umi"] = "KG";

        //    ////CAUZIONE NUMERICO 14

        //    ////TESTOCASSA ALFANUM. 20
        //    //s = strRig.Substring(106, 20);
        //    //rowArt["div_adb"] = s;

        //    ////TESTO_ETI ALFANUM. 30

        //    ////STATO_PRE NUMERICO 1

        //    ////EXT-U-MERC ALFANUM. 12
        //    //rowArt["for_ecr"] = strRig.Substring(157, 12).Trim();
        //    //rowArt["for_rep"] = (string)rowArt["for_ecr"];

        //    ////TIPO_ETI ALFANUM. 1

        //    ////PREZZO NUMERICO 8
        //    //rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(170, 8)) / 100;

        //    ////COSTO NUMERICO 8
        //    //rowArt["div_cos"] = 0;  // _clsFun.Txt2Dec(strRig.Substring(178, 8)) / 100;

        //    ////COD_FORNIT NUMERICO 6

        //    ////CODARTFORN ALFANUM. 14

        //    ////CONFEZIONE NUMERICO 5

        //    ////FORN_PRI NUMERICO 1

        //    ////MULT NUMERICO 3

        //    ////MIX NUMERICO 3
        //    ////REPARTO NUMERICO 4
        //    ////CODVENDPOS NUMERICO 13
        //    ////COD_PERMA ALFANUM. 2
        //    ////TCALPREMED ALFANUM. 2
        //    ////COD_POLIT ALFANUM. 2
        //    ////GIA_MIN NUMERICO 8
        //    ////GIA_MAX NUMERICO 8
        //    ////LIV_RIORDI NUMERICO 5
        //    ////GIACENZA CARATTERE 1
        //    ////TEMPO_APPR NUMERICO 3
        //    ////ORDINE_MIN NUMERICO 8
        //    ////ORD_FORN NUMERICO 3
        //    ////NUM_ETI ALFANUM. 2
        //    ////BOLLINI NUMERICO 6
        //    ////FLAG_DISP ALFANUM. 10
        //    ////LISTINO_P ALFANUM. 4
        //    ////LISTINO_C ALFANUM. 4
        //    ////DECLIST_P NUMERICO 1
        //    ////DECLIST_C NUMERICO 1
        //    ////DESCRIZIOL ALFANUM. 40
        //    ////TESTO_ETIL ALFANUM. 40
        //    ////PRZZERO ALFANUM. 1
        //    ////LINKBILA ALFANUM. 1
        //    ////LINKTARA NUMERICO 6


        //    return rowArt;
        //}
  
        public DataRow DivOff(string strRig, DataRow rowArt)
        {
            string s = "";
            decimal d = 0;

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

                d = _clsFun.Txt2Dec(strRig.Substring(94, 7));

                decimal dPerc = 0;

                if (d == 1)
                    dPerc = 5;
                else if (d == 2)
                    dPerc = 10;
                else if (d == 3)
                    dPerc = 15;
                else if (d == 4)
                    dPerc = 20;
                else if (d == 5)
                    dPerc = 25;
                else if (d == 6)
                    dPerc = 30;
                else if (d == 7)
                    dPerc = 40;
                else if (d == 8)
                    dPerc = 50;
                else if (d == 9)
                    dPerc = 33;

                //Valore
                rowArt["div_imp"] = dPerc;
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

        public DataTable Fil2Tab(string strFil, ProgressBar progressBar1)
        {
            //s = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + s + ";Extended Properties=\"Excel 8.0;HDR={1};IMEX=1\"";

            DataRow x;

            ArrayList aNac = new ArrayList();

            string strDayOff = "";

            string sCon = _clsFun.ConMdb(Path.GetDirectoryName(_clsDef.FILEINI) + "\\ApDPiu.mdb");
            string s = "SELECT * FROM AnaArt ORDER BY art_arf";
            DataTable tArt = _clsFun.FillTabMdb("AnaArt", s, false, sCon);
            
            DataTable tVar = tArt.Clone();
            tVar.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ean",
                Caption = "Ean",
                MaxLength = 200,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            tVar.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "fat_tdo",
                Caption = "Tipo documento",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)"F"
            });
            tVar.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "fat_ndo",
                Caption = "Numero documento",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tVar.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "fat_ddo",
                Caption = "Data documento",
                MaxLength = 8,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tVar.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "fat_tri",
                Caption = "Tipo riga",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tVar.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "fat_qta",
                Caption = "Pezzi spediti",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            tVar.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "fat_cos",
                Caption = "Costo",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            tVar.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "fat_prp",
                Caption = "Prezzo al pubblico",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            tVar.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "fat_imp",
                Caption = "Importo",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            tVar.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "off_tip",
                Caption = "Tipo offerta",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tVar.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "off_day",
                Caption = "Date offerta",
                MaxLength = 30,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            //Evito indice univoco su tVar
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_arf"];
            tArt.PrimaryKey = keys;

            s = "SELECT * FROM AnaEan ORDER BY ean_cdb";
            DataTable tEan = _clsFun.FillTabMdb("AnaEan", s, false, sCon);
            keys = new DataColumn[1];
            keys[0] = tEan.Columns["ean_cdb"];
            tEan.PrimaryKey = keys;

            s = "SELECT * FROM AnaEan ORDER BY tab_ger";
            DataTable tEcr = _clsFun.FillTabMdb("AnaEcr", s, false, sCon);
            keys = new DataColumn[1];
            keys[0] = tEcr.Columns["tab_ger"];
            tEcr.PrimaryKey = keys;

            if (File.Exists(strFil))
            {
                DataRow[] j;

                using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
                {
                    FileInfo fInfo = new FileInfo(strFil);

                    progressBar1.Value = 0;
                    progressBar1.Maximum = (int)fInfo.Length;
                    progressBar1.Minimum = 0;

                    string sRig = "";

                    ArrayList aWhe = new ArrayList();
                    aWhe.Add("art_arf");
                    ArrayList aExl = new ArrayList();

                    while ((sRig = sr.ReadLine()) != null)
                    {
                        progressBar1.Increment(sRig.Length);
                        System.Windows.Forms.Application.DoEvents();

                        if (false && sRig.Length > 15 && sRig.Substring(10, 3) == "TAB" && sRig.Substring(100, 4) == "0127")
                        {
                            //string sCod = sRig.Substring(50, 12);
                            //string sRep = sRig.Substring(107, 3);

                            //j = tEcr.Select("art_arf='" + x["art_arf"] + "'");
                            //if (j.Length > 0)
                            //{
                            //    s = _clsFun.SqlUpdRow("AnaArt", tEcr, j[0], x, aWhe, aExl);
                            //}
                            //else
                            //{
                            //    s = _clsFun.SqlInsertRow("AnaArt", tEcr, x);
                            //    tEcr.Rows.Add(x);
                            //}
                            //if (s != "")
                            //    _clsFun.MdbWrite(s, sCon);
                        }
                        else if(sRig.Length > 15 && sRig.Substring(10,3) == "MDA")
                        {
                            aWhe = new ArrayList();
                            aWhe.Add("art_arf");
                            aExl = new ArrayList();

                            x = tArt.NewRow();
                            //x["art_arf"] = (string)y[10];
                            //x["art_ard"] = (string)y[11];
                            //x["art_ean"] = (string)y[12];

                            //x["art_iva"] = "";
                            //s = (string)y[18];
                            //if (_clsFun.Numerico(s, "0123456789,."))
                            //    s = Convert.ToDecimal(s).ToString("000");
                            //x["art_iva"] = s;

                            //s = (string)y[19];

                            //Console.WriteLine("xxx");

                            //x["art_pne"] = 0;
                            //x["art_tgr"] = "KG";

                            //string[] a = s.Split('/');
                            //if (a.Length > 1)
                            //{
                            //    s = a[0].Trim();
                            //    Console.WriteLine("xxx");

                            //    if (_clsFun.Numerico(s, "0123456789.,"))
                            //        x["art_pne"] = Convert.ToDecimal(s); // .Replace(",",".");
                            //    x["art_tgr"] = a[1].Trim();
                            //}

                            //x["art_umi"] = "NR";

                            //x["art_pxc"] = 1;
                            //s = Convert.ToString(y[22]);
                            //if (_clsFun.Numerico(s, "0123456789.,"))
                            //    x["art_pxc"] = Convert.ToString(s);

                            //CDA Codice articolo 
                            //A 7 1 7

                            x["art_arf"] = sRig.Substring(46, 7).Trim();

                            //DES Descrizione 
                            //A 30 8 37
                            x["art_ard"] = sRig.Substring(53, 30).Trim();


                            //MUL Multiplo.Numero di singole confezioni che compongono l’unita’ di vendita 
                            //S 4,00 38 41
                            // ???? rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(170, 8)) / 100;

                            //CON Confezionamento Cfr. tab. 0005 
                            //A 2 42 43

                            //TPG Tipo grammatura (o unita’ di misura) Cr. Tab. 0006 
                            //A 2 44 45
                            x["art_tgr"] = sRig.Substring(89, 2);

                            //GRA Grammatura 
                            //S 4,00 46 49
                            x["art_pne"] = _clsFun.Txt2Dec(sRig.Substring(91, 4));

                            //SGO Sgocciolato 
                            //S 4,00 50 53

                            //MER Codice merceologia Crf. Tab. 0127 livello 01
                            //A 12 54 65
                            x["art_ecr"] = sRig.Substring(99, 12).Trim();
                            //rowArt["for_rep"] = (string)rowArt["for_ecr"];

                            //TPA Tipo articolo in acquisto ‘1’ a numero - ‘2’ a peso
                            //A 1 66 66
                            x["art_umi"] = "NR";
                            s = sRig.Substring(111, 1);
                            if (s == "2")
                                x["art_umi"] = "KG";

                            //TRA Tipo merce Cfr. tab. 0014 
                            //A 1 67 67

                            //FOR Fornitore abituale 
                            //A 6 68 73

                            //UMS Tipo scadenza ‘GG’ giorni
                            //‘SS’ settimane
                            //‘MM’ mesi
                            //‘AA’ anni
                            //A 2 74 75

                            //DUR Durata.
                            //Precisa l’ampiezza del periodo definito dal
                            //tipo di scadenza dal momento del
                            //ricevimento della merce presso l’azienda
                            //distributrice.
                            //S 3,00 76 78

                            //CAL Calo peso Non usato 
                            //S 4,02 79 82

                            //MAR Marchio 
                            //A 12 83 94

                            //FMA Indicatore prodotto a marchio interno ‘S’ - Si ‘N’ – No
                            //A 1 95 95

                            //RA1 Codice articolo vuoto 1 Non usato 
                            //A 7 96 102

                            //CO1 Numero elementi in vuoto 1 Non usato 
                            //S 4,00 103 106

                            //RA2 Codice articolo vuoto 2 Non usato 
                            //A 7 107 113

                            //CO2 Numero elementi in vuoto 2 Non usato 
                            //S 4,00 114 117

                            //VAR Codice PLU bilance 
                            //A 4 118 121

                            //FEI Libero Non usato 
                            //A 1 122 122

                            //ELS Tipo etichetta scaffale.
                            //Assume dei valori che nel tempo possono mutare. Non usato. Dato ad uso interno della Commerciale Brendolan.
                            //A 1 123 123

                            //ETS Spedizione bilance S/N ‘S’ se presente VAR, altrimenti ‘N’
                            //A 1 124 124

                            //DEB Banchi bilance 
                            //A 2 125 126

                            //SAA Anno stagionalita' 
                            //A 4 127 130

                            //STP Codice stagionalita' Cfr. tab. 0047 
                            //A 3 131 133

                            //FL3 Stato articolo ‘1’ valido ‘0’ annullato
                            //A 1 134 134
                            x["art_stf"] = "";
                            s = sRig.Substring(179, 1).Trim();
                            if (s == "1")
                                x["art_stf"] = "N";

                            //DC1 Tipologia vini DOC Non fornita la tabella di controllo e decodifica
                            //A 4 135 138

                            //DC2 Provenienza vini DOC Non fornita la tabella di controllo e decodifica
                            //A 4 139 142
                            x["art_iva"] = (sRig.Substring(188, 2)).Trim().PadLeft(3, Convert.ToChar("0"));

                            //CDI Codice iva Cfr. tab. 0007 
                            //A 2 143 144

                            //DVI Data validita' iva Forma 'aaaammgg' 
                            //S 8,00 145 152

                            //OCI Codice iva precedente 
                            //A 2 153 154

                            //PEN Aliquota encc 
                            //S 5,02 155 159

                            //CTR Valore contrassegno di stato e ifa 
                            //S 2,03 160 171

                            //GAN Gradi anidri 
                            //S 4,00 172 175

                            //VOU Codice UTIF Cfr. tab. 0009 
                            //A 2 176 177

                            //UMU UTIF Grammatura Cfr. tab. 0006 
                            //A 2 178 179

                            //GRU UTIF Valore grammatura 
                            //S 5,00 180 184

                            //SCL Descrizione scontrino 
                            //A 20 185 204
                            x["art_adb"] = (sRig.Substring(230, 20)).Trim();

                            //TPV Tipo articolo in vendita ‘1’ a numero ‘2’ a peso
                            //A 1 205 205

                            //FPP Flag primo prezzo Fisso ‘N’ 
                            //A 1 206 206

                            //VA2 Secondo PLU bilancia 
                            //A 4 207 210

                            //FPB Indicatore prezzo bloccato Fisso ‘N’ 
                            //A 1 211 211

                            //PEZ Pezzi per collo in vendita da deposito 
                            //S 4,00 212 215

                            //BOL Soggetto a stampa bolla 
                            //A 1 216 216

                            //PZC Pezzi per collo in acquisto da fornitore 
                            //S 4,00 217 220
                            x["art_pxc"] = _clsFun.Txt2Dec(sRig.Substring(262, 4));

                            x["art_rep"] = sRig.Substring(99, 3).Trim();

                            x["art_tva"] = "ART";                            //Anagrafica articoli

                            j = tArt.Select("art_arf='" + x["art_arf"] + "'");
                            if (j.Length > 0)
                            {
                                if (!DBNull.Value.Equals(j[0]["art_pve"]))
                                    x["art_pve"] = (decimal)j[0]["art_pve"];
                                s = _clsFun.SqlUpdRow("AnaArt", tArt, j[0], x, aWhe, aExl);
                            }
                            else
                            {
                                s = _clsFun.SqlInsertRow("AnaArt", tArt, x);
                                tArt.Rows.Add(x);
                            }
                            if (s != "")
                                _clsFun.MdbWrite(s, sCon);

                            DataRow xx = tVar.NewRow();
                            foreach (DataColumn c in tArt.Columns)
                                xx[c.ColumnName] = x[c.ColumnName];

                            if ((string)x["art_arf"] == "0428647")
                                Console.WriteLine("xxxx");

                            s = "";
                            j = tEan.Select("ean_cda='" + x["art_arf"] + "'", "tmp_day DESC");
                            if(j.Length > 0)
                            {
                                for(int i = 0; i < j.Length; i++)
                                {
                                    if (i > 13)
                                        break;
                                    s += (string)j[i]["ean_cdb"] + ";";
                                }
                            }

                            xx["tmp_ean"] = s;
                            tVar.Rows.Add(xx);                        
                        }
                        else if (sRig.Length > 15 && sRig.Substring(10, 3) == "MHE")    //Prezzi al pubblico
                        {
                            aWhe = new ArrayList();
                            aWhe.Add("art_arf");
                            aExl = new ArrayList();

                            //MHE – PREZZI AL PUBBLICO
                            //CDA Codice articolo 
                            //A 7 1 7
                            string sArt = sRig.Substring(46, 7);

                            //PRV Prezzo al pubblico 
                            //S 12,03 8 19
                            s = sRig.Substring(53,12);
                            decimal dPrv = Convert.ToDecimal(s) / 1000;

                            //AVA Anno inizio validita' 
                            //S 4,00 20 23

                            //MVA Mese inizio validita' 
                            //S 2,00 24 25

                            //GVA Giorno inizio validita' 
                            //S 2,00 26 27

                            //STA Stato validita' record Fisso ‘1’ 
                            //A 1 28 28

                            //VAL Codice valuta Fisso ‘EUR’ 
                            //A 3 29 31

                            //ING Non usato 
                            //A 1 32 32

                            //FPF Non usato 
                            //A 1 33 33

                            j = tArt.Select("art_arf='" + sArt + "'");
                            if (j.Length > 0)
                            {
                                //j[0]["art_pve"] = dPrv;

                                if (DBNull.Value.Equals(j[0]["art_pve"]))
                                    j[0]["art_pve"] = 0;


                                if (dPrv > 0 && (decimal)j[0]["art_pve"] != dPrv)
                                {
                                    s = "UPDATE AnaArt SET art_pve=" + dPrv.ToString().Replace(",", ".") + " WHERE art_arf='" + sArt + "'";

                                    //DataRow xx = j[0];
                                    //xx["art_pve"] = dPrv;

                                    //s = _clsFun.SqlUpdRowIdx("AnaArt", tArt, j[0], xx, aExl);
                                    //if (s != "")
                                    _clsFun.MdbWrite(s, sCon);
                                }
                            }
                        }
                        else if (false && sRig.Length > 15 && sRig.Substring(10, 3) == "MHN")    //Costi
                        {
                            //CDA Codice articolo 
                            //A 7 1 7
                            string sArt = sRig.Substring(46,7);


                            //CDF Codice fornitore 
                            //A 6 8 13

                            //AVA Anno inizio validita' S 4,00 14 17
                            //MVA Mese inizio validita' S 2,00 18 19
                            //GVA Giorno inizio validita' S 2,00 20 21

                            //CLR Costo Fisso 0,001€ 
                            //S 12,03 22 33

                            //NET Costo Fisso 0,001€ 
                            //S 12,03 34 45

                            //SNE Costo Fisso 0,001€ 
                            //S 12,03 46 57

                            //FAT Costo Fisso 0,001€ 
                            //S 12,03 58 69

                            //STA Stato validita' record ‘1’ valido ‘0’ annullato
                            //A 1 70 70

                            //VAL Codice valuta Fisso ‘EUR’ 
                            //A 3 71 73


                        }
                        else if (sRig.Length > 15 && sRig.Substring(10, 3) == "MHA")    //Barcode
                        {
                            aWhe = new ArrayList();
                            aWhe.Add("ean_cdb");
                            aExl = new ArrayList();
                            aExl.Add("tmp_day");

                            x = tEan.NewRow();

                            //MHA – CODICI A BARRE

                            //CDA Codice articolo 
                            //A 7 1 7
                            x["ean_cda"] = sRig.Substring(46, 7);

                            //CDF Codice fornitore 
                            //A 6 8 13

                            //CDB Codice a barre Allineato a sinistra 
                            //A 13 14 26
                            x["ean_cdb"] = sRig.Substring(59, 13);

                            //PEZ Pezzi confezione Fisso zero 
                            //S 4,00 27 30
                            x["ean_pez"] = Convert.ToInt16(sRig.Substring(72, 4));

                            //STA Stato validita' record ‘1’ valido ‘9’ annullato
                            //A 1 31 31
                            x["ean_sta"] = sRig.Substring(76, 1);

                            x["tmp_day"] = DateTime.Today.ToString("yyyyMMdd");

                            j = tEan.Select("ean_cdb='" + x["ean_cdb"] + "'");
                            if (j.Length > 0)
                            {
                                s = _clsFun.SqlUpdRow("AnaEan", tEan, j[0], x, aWhe, aExl);
                            }
                            else
                            {
                                s = _clsFun.SqlInsertRow("AnaEan", tEan, x);
                                tEan.Rows.Add(x);
                            }
                            if (s != "")
                                _clsFun.MdbWrite(s, sCon);
                        }
                        else if (sRig.Length > 15 && sRig.Substring(10, 3) == "BLE")    //Bolle elettroniche
                        {
                            string sTdo = sRig.Substring(46, 2);
                            if (sTdo == "01")
                            {
                                string sNdo = sRig.Substring(56, 7);
                                string sDdo = sRig.Substring(67, 8);
                                string sArf = sRig.Substring(75, 7);
                                string sArd = sRig.Substring(82, 30).Trim();
                                string sTgr = sRig.Substring(114, 2);
                                string sPne = sRig.Substring(116, 4);
                                string sPxc = sRig.Substring(127, 4);
                                string sUmi = sRig.Substring(131, 1);       //2 a peso
                                string sIva = sRig.Substring(132, 2);       //Codice IVA
                                string sAli = sRig.Substring(134, 4);       //Aliquota IVA
                                string sTme = sRig.Substring(138, 1);       //Tipo merce
                                string sEcr = sRig.Substring(139, 12);
                                string sEan = sRig.Substring(151, 13);

                                string sTri = sRig.Substring(164, 1);       //Tipo riga: ‘N’ riga “normale” ‘V’ vuoti in addebito ‘A’ vuoti in accredito ‘O’ omaggio ‘M’ omaggio ‘B’ riga “bollo” 
                                string sCsp = sRig.Substring(165, 5);       //Colli spediti
                                string sQta = sRig.Substring(170, 9);       //Pezzi spediti (2 decimali)
                                string sCos = sRig.Substring(179, 14);      //Costo (5 decimali)
                                string sPrv = sRig.Substring(193, 10);      //Prezzo pubblico (3 dec)
                                string sImp = sRig.Substring(203, 14);      //Importo riga (5 dec)

                                if (sArf.Trim() != "")
                                {
                                    aWhe = new ArrayList();
                                    aWhe.Add("art_arf");
                                    aExl = new ArrayList();

                                    x = tArt.NewRow();

                                    x["art_arf"] = sArf;
                                    x["art_ard"] = sArd;
                                    x["art_tgr"] = sTgr;
                                    x["art_pne"] = _clsFun.Txt2Dec(sPne);
                                    x["art_ecr"] = sEcr;

                                    x["art_umi"] = "NR";
                                    if (sUmi == "2")
                                        x["art_umi"] = "KG";

                                    x["art_stf"] = "";
                                    x["art_iva"] = sIva.Trim().PadLeft(3, Convert.ToChar("0"));
                                    x["art_pxc"] = _clsFun.Txt2Dec(sPxc);

                                    x["art_rep"] = sEcr.Substring(1, 3).Trim();
                                    x["art_tva"] = "FAT";                            //Bolla elettronica

                                    j = tArt.Select("art_arf='" + x["art_arf"] + "'");
                                    if (j.Length > 0)
                                    {
                                        if (DBNull.Value.Equals(j[0]["art_pve"]))
                                            j[0]["art_pve"] = 0;
                                        x["art_pve"] = (decimal)j[0]["art_pve"];
                                        s = _clsFun.SqlUpdRow("AnaArt", tArt, j[0], x, aWhe, aExl);
                                    }
                                    else
                                    {
                                        s = _clsFun.SqlInsertRow("AnaArt", tArt, x);
                                        tArt.Rows.Add(x);
                                    }
                                    if (s != "")
                                        _clsFun.MdbWrite(s, sCon);

                                    DataRow xx = tVar.NewRow();
                                    foreach (DataColumn c in tArt.Columns)
                                        xx[c.ColumnName] = x[c.ColumnName];

                                    if ((string)x["art_arf"] == "0428647")
                                        Console.WriteLine("xxxx");

                                    s = "";
                                    j = tEan.Select("ean_cda='" + x["art_arf"] + "'", "tmp_day DESC");
                                    if (j.Length > 0)
                                    {
                                        for (int i = 0; i < j.Length; i++)
                                        {
                                            if (i > 13)
                                                break;
                                            s += (string)j[i]["ean_cdb"] + ";";
                                        }
                                    }

                                    xx["tmp_ean"] = s;

                                    xx["fat_tdo"] = "F";
                                    xx["fat_ndo"] = sNdo;
                                    xx["fat_ddo"] = sDdo;
                                    xx["fat_tri"] = sTri;
                                    xx["fat_qta"] = _clsFun.Txt2Dec(sQta) / 100;
                                    xx["fat_cos"] = _clsFun.Txt2Dec(sCos) / 100000;
                                    xx["fat_prp"] = _clsFun.Txt2Dec(sPrv) / 1000;
                                    xx["fat_imp"] = _clsFun.Txt2Dec(sImp) / 100000;

                                    tVar.Rows.Add(xx);
                                }
                            }
                            else if (sTdo == "03")
                            {
                                tVar.Rows[tVar.Rows.Count - 1]["fat_tdo"] = "N";
                            }

                        }
                        else if (sRig.Length > 60 && sRig.Substring(10, 3) == "OFT")    //Offerte testata
                            strDayOff += sRig.Substring(46, 5) + "-" + sRig.Substring(85, 8) + "-" + sRig.Substring(97, 8) + "|" ;
                        else if (sRig.Length > 15 && sRig.Substring(10, 3) == "OFD")    //Offerte
                        {
                            if (sRig.Substring(46, 1) == "1")
                            {
                                string  sDayOff = "";

                                string sCod = sRig.Substring(47, 5);        //Codice offerta

                                string[] a = strDayOff.Split('|');
                                foreach(String ss in a)
                                {
                                    if(ss.Length > 10 && ss.Substring(0,5) == sCod)
                                    {
                                        sDayOff = ss;
                                        break;
                                    }
                                }

                                string sSta = sRig.Substring(46, 1);        //Stato offerta 0=annullato
                                string sTof = sRig.Substring(52, 2);        //11 = taglio prezzo, 10 sconto %
                                string sArf = sRig.Substring(54, 7);        //Codice articolo
                                string sVma = sRig.Substring(61, 12);       //Valore massimo
                                string sVmi = sRig.Substring(73, 12);       //Valore minimo
                                string sCas = sRig.Substring(85, 5);        //Casualità aplicazione

                                string sSos = sRig.Substring(90, 12);       //Soglia applicazione
                                string sSov = sRig.Substring(102, 12);      //Soglia applicazione
                                string sSoq = sRig.Substring(114, 12);      //Soglia applicazione
                                string sSoa = sRig.Substring(126, 12);      //Soglia applicazione

                                string sPrz = sRig.Substring(138, 12);      //Prezzo 3 dec

                                string sVao = sRig.Substring(150, 12);      //Valore offerta
                                string sLeg = sRig.Substring(162, 3);       //Legami
                                string sCva = sRig.Substring(165, 3);       //Codice valuta

                                aWhe = new ArrayList();
                                aWhe.Add("art_arf");
                                aExl = new ArrayList();

                                x = tVar.NewRow();

                                x["fat_ndo"] = sCod;
                                x["art_arf"] = sArf;
                                x["off_tip"] = sTof;
                                x["off_day"] = sDayOff;
                                x["art_pve"] = _clsFun.Txt2Dec(sPrz) / 1000; ;
                                //x["art_pne"] = _clsFun.Txt2Dec(sPne);
                                //x["art_ecr"] = sEcr;

                                //x["art_umi"] = "NR";
                                //if (sUmi == "2")
                                //    x["art_umi"] = "KG";

                                //x["art_stf"] = "";
                                //x["art_iva"] = sIva.Trim().PadLeft(3, Convert.ToChar("0"));
                                //x["art_pxc"] = _clsFun.Txt2Dec(sPxc);

                                //x["art_rep"] = sEcr.Substring(1, 3).Trim();
                                x["art_tva"] = "OFF";                            //Offerta

                                j = tArt.Select("art_arf='" + x["art_arf"] + "'");
                                if (j.Length > 0)
                                {
                                    DataRow xx = tVar.NewRow();
                                    foreach (DataColumn c in tArt.Columns)
                                        xx[c.ColumnName] = j[0][c.ColumnName];

                                    if ((string)x["art_arf"] == "0428647")
                                        Console.WriteLine("xxxx");

                                    s = "";
                                    j = tEan.Select("ean_cda='" + x["art_arf"] + "'", "tmp_day DESC");
                                    if (j.Length > 0)
                                    {
                                        for (int i = 0; i < j.Length; i++)
                                        {
                                            if (i > 13)
                                                break;
                                            s += (string)j[i]["ean_cdb"] + ";";
                                        }
                                    }

                                    x["tmp_ean"] = s;

                                    //xx["fat_ndo"] = sNdo;
                                    //xx["fat_ddo"] = sDdo;
                                    //xx["fat_tri"] = sTri;
                                    //xx["fat_qta"] = _clsFun.Txt2Dec(sQta) / 100;
                                    //xx["fat_cos"] = _clsFun.Txt2Dec(sCos) / 100000;
                                    //xx["fat_prp"] = _clsFun.Txt2Dec(sPrv) / 1000;
                                    //xx["fat_imp"] = _clsFun.Txt2Dec(sImp) / 100000;

                                    tVar.Rows.Add(x);
                                }


                            }
                        }
                    }

                    sr.Close();
                    sr.Dispose();

                    //s = strFil + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                    //File.Move(strFil, s);
                }
            }

            return tVar;
        }

    }
}
