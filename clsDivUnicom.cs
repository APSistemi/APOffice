using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsDivUnicom
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public string _strPar017EcrDivFornitori = "";

        public clsDivUnicom() { }

        public DataRow DivArt(string strRig, DataRow rowArt, DateTime dayDay)
        {
            string s = "";
            Double i = 0;

            rowArt["div_dva"] = dayDay;

            //BODCDE	4 (C)	1-4	codice negozio	
            //BODCDA	7 (C)	5-11	codice articolo	
            rowArt["div_arf"] = strRig.Substring(4, 7).Trim();

            //BODSET	3 (C)	12-14	settore merceologico (livello 1)	
            //BODFAM	3 (C)	15-17	Famiglia/reparto (livello 2)	
            //BODSUF	6 (C)	18-23	Sottofamiglia/categoria  (livello 3 ECR)
            rowArt["for_ecr"] = "";
            if (_strPar017EcrDivFornitori == "S")
                rowArt["for_ecr"] = strRig.Substring(11, 12);

            //BODEAN	13 (C)	24-36	codice a barre	
            s = strRig.Substring(23, 13).Trim();
            if (s.Length == 13 && s.Substring(0, 1) == "2" && s.Substring(7, 5) == "00000")
                s = s.Substring(0, 7) + "000000";
            rowArt["div_ean"] = s;

            //BODDSA	40 (C)	37-76	descrizione articolo	se BODTPR=D:
            rowArt["div_ard"] = strRig.Substring(36, 40).Trim();

            //da 1 a 10 = calibro
            //da 11 a 15 = categoria
            //da 16 a 30 = provenienza/categoria pesce
            //da 31 a 40 = filler

            //BODDS1	20 (C)	77-96	descrizione articolo scontrino per BODTPR diverso da T e D	se BODTPR=D: provenienza
            rowArt["div_adb"] = strRig.Substring(76, 20).Trim();

            if ((string)rowArt["div_arf"] == "2092611")
                Console.WriteLine("aaaaaaaa");

            //BODREU	4 (C)	97-100	reparto casse	
            s = strRig.Substring(97, 3);

            //if (s == "002")
            //{
            //    i = Convert.ToDouble(rowArt["for_ecr"]);
            //    if (i >= 7111111000)
            //        s = "002";      //Surgelati
            //}
            if (s == "004")
            {
                if (((string)rowArt["for_ecr"]).Trim() != "" && _clsFun.Numerico((string)rowArt["for_ecr"]))
                {
                    i = Convert.ToDouble(rowArt["for_ecr"]);
                    if (i >= 4007007300 && i < 5000000000)
                        s = "202";      //Surgelati
                }
            }
            else if (s == "005")
            {
                if (((string)rowArt["for_ecr"]).Trim() != "" && _clsFun.Numerico((string)rowArt["for_ecr"]))
                {
                    if (((string)rowArt["for_ecr"]).Substring(0, 3) == "006")
                        s = "006";      //Gastronomia
                }
            }
            rowArt["for_rep"] = s;

            //BODPRV	7 (C)	101-107	prezzo di vendita
            //(se negativo nel primo byte c’è il segno - )
            //(se a ZERO nel primo byte c’è il segno + )
            //(se prezzo richiesto al primo byte c’è il segno ? )	*se BODTPR=T: codice listino
            rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(101, 6)) / 100;

            //25/07/2011 nuova implementazione gestita solo in Mersy (no IND)
            //BODIVA	2 (C)	108-109	codice iva	
            rowArt["for_iva"] = (strRig.Substring(107, 2)).Trim().PadLeft(3, Convert.ToChar("0"));

            //BODCNF	2 (C)	110-111	confezione	

            //BODTAR	3 (C)	112-114	tara per prodotti a peso	
            //rowArt["for_tar"] = strRig.Substring(112, 3);              da gestire ??????????????

            //BODRIO	1 (C)	115-115	C = centrale
            //D = diretti
            //P = piattaforma (dismesso dal 2013)
            //V = ventilazione da SPG 
            //        (07/10/2014: nuova implementazione)	
            
            //BODSCO	4 (C)	116-119	sconto percentuale	2 interi + 2 decimali

            //BODTPG	2 (C)	120-121	tipo grammatura	KG, LT, MT, PZ
            rowArt["for_tgr"] = strRig.Substring(119, 2);

            //BODGRA	6 (C)	122-127	Grammatura	3 interi + 3 decimali
            //rowArt["div_pne"] = Convert.ToDecimal(strRig.Substring(121, 6))/1000;
            s = strRig.Substring(121, 6);
            if (_clsFun.Numerico(s))
                rowArt["div_pne"] = Convert.ToDecimal(s) / 1000;

            //BODSGO	6 (C)	128-133	Sgocciolato	3 interi + 3 decimali
            s = strRig.Substring(127, 6);
            if(_clsFun.Numerico(s))
                rowArt["div_pne"] = Convert.ToDecimal(s) / 1000;

            //BODUMV	3 (C)	134-136	unità misura di vendita	001 = pz
            //002 = kg
            //s = "KG";
            //if (strRig.Substring(119, 2) == "PZ")
            //    s = "PZ";
            //rowArt["for_umi"] = s;

            s = "KG";
            if (strRig.Substring(133, 3) == "001")
                s = "PZ";
            rowArt["for_umi"] = s;

            //BODMPS	1 (C)	137-137	molteplicità punti set	da 0 a 9
            //BODFCL	1 (C)	138-138	categoria cliente fedele
            //Saldo punti	A=campagna attuale
            //B=campagna precedente
            //BODFPR	1 (C)	139-139	flag per premi fidelity	S = premio omaggio
            //N = premio con contributo
            //BODFTM	1 (C)	140-140	flag offerte fidelity	M = modifiche C = cancellazioni

            //BODCDF	6 (C)	141-146	codice fornitore	
            rowArt["div_fo2"] = strRig.Substring(140, 6);

            //BODARF	7 (C)	147-153	codice articolo fornitore	
            rowArt["div_af2"] = strRig.Substring(141, 7);

            if ((string)rowArt["div_arf"] == "2034013")
                Console.WriteLine("aaaaaaa");

            //BODBO2	2 (C)	154-155	bollini (in migliaia)	
            //BODQCC	2 (C)	156-157	q.tà componenti cesti	
            //BODPMP	1 (C)	158-158	vendite a peso o a corpo	
            //•	P = vendita a peso per articoli con codice bilancia
            //•	M = vendita a corpo per articoli con codice bilancia
            s = strRig.Substring(157, 1).Trim();
            rowArt["div_bpz"] = "";
            if (s == "M")
                rowArt["div_bpz"] = "S";
            if (s != "")
            {
                rowArt["div_bil"] = "S";

                if ((string)rowArt["for_rep"] == "003") //Orto
                    rowArt["div_reb"] = "3";
                else if ((string)rowArt["for_rep"] == "002") //Macelleria
                    rowArt["div_reb"] = "1";
                else if ((string)rowArt["for_rep"] == "006") //Gatronomia
                    rowArt["div_reb"] = "5";
                else if ((string)rowArt["for_rep"] == "007") //Pescheria
                    rowArt["div_reb"] = "7";
            }

            //BODBOL	3 (C)	159-161	bollini fidelity	

            //BODCES	7 (C)	162-168	prezzo di cessione	*se BODTPR=T: codice listino
            rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(161, 7)) / 1000;

            //BODUMA	3 (C)	169-171	unità misura di acquisto	003 fisso

            //BODPEZ	4 (C)	172-175	pezzi/collo	
            rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(172, 4)) ;

            //BODCOL	3 (C)	176-178	colli per pallet	

            //BODPLU	6 (C)	179-184	codice bilancia (plu)	
            s = strRig.Substring(180, 4).Trim();

            if (s != "")
            {
                Console.WriteLine("aaa");
                s = Convert.ToInt32(s).ToString();
                //rowArt["div_plu"] = strRig.Substring(179, 5).Trim();
                rowArt["div_plu"] = s;
            }

            //BODDT1	8 (C)	185-192	data validità/inizio off.	formato AAAAMMGG
            //BODDT1	8 (C)	185-192	data validità/inizio off.	formato AAAAMMGG
            //BODDT2	8 (C)	193-200	data fine offerta	formato AAAAMMGG
            //BODDT3	8 (C)	201-208	data fine validità	formato AAAAMMGG
            //BODPAD	7 (C)	209-215	Sconto a valore
            //codice padre per cesti se BODTPR=’X’
            //codice contributo se BODTPR=’G’	
            //BODRAG	4 (C)	216-219	raggruppamento per MxN	Stesso valore se prezzo e merceologia uguali
            //BODMIX	2 (C)	220-221	mix (M nel MxN)	
            //BODMCH	2 (C)	222-223	match (N nel MxN)	
            //BODSTA	3 (C)	224-226	Da 224 a 225 Evidenziazione articolo
            //SE = articolo Selex
            //PB = il prezzo più basso
            //BI = biologico
            //SG = senza glutine
            //RB = ribassato e bloccato

            //Da 226 a 226 filler
            //Dal 04/03/2013: buono pasto: 
            //" "= lascia inalterato il valore presente
            //“0“ = articolo NON pabagile tramite BuonoPasto
            //“1“ = articolo pagabile tramite BuonoPasto
	
            //BODTPA	3 (C)	227-229	Tipo articolo	multiplo di vendita
            //BODARC	3 (C)	230-232	articolo collegato

            //03/06/2014
            //da 230 a 230:
            //flag articolo scontabile per assistiti celiaci:
            //‘0’ = disattiva flag articolo per celiaci
            //‘1’ = attiva flag articolo per celiaci
            //Blank = Lascia inalterato il valore precedente

            //Da 231 a 232: filler
            //    non gestito
            //BODSCV	7 (C)	233-239	sconto a valore fidelity (in centesimi)	*se BODTPR=T: codice listino
            //BODDET	2 (C)	240-241	% detrazione

            //Mix-Match per gestione promozioni:
            //- sconto
            //- cartolina

            //----------------------------------------------------------
            //25/07/2011: Nuova implementazione 

            //    non gestito

            //SCONTO
            //"  "= lascia inalterato il valore presente
            //"00"=articolo soggetto a sconto 
            //"01"=articolo esente da sconto: ricariche telefoniche, ricariche telefoniche e gratta e vinci
            //“02“ = articolo esente da sconto: alimenti per celiaci

            //--------------------------------------------------
            //CARTOLINE
            //"11"=1 articolo 1 cartolina
            //"12"=1 articolo 2 cartoline
            //"13"=1 articolo 3 cartoline
            //"14"=1 articolo 4 cartoline
            //"15"=1 articolo 5 cartoline

            //"21"=2 articoli 1 cartolina
            //"22"=2 articoli 2 cartoline
            //"23"=2 articoli 3 cartoline
            //"24"=2 articoli 4 cartoline
            //"25"=2 articoli 5 cartoline

            //…

            //"61"=6 articoli 1 cartolina
            //"62"=6 articoli 2 cartoline
            //"63"=6 articoli 3 cartoline
            //"64"=6 articoli 4 cartoline
            //"65"=6 articoli 5 cartoline

            //Liberi dal 66 al 99

            //BODFL1
            //    1 (C)	242-242	Tipo articolo 1
            //Erogazione di PIN tramite il software NCR UPB pin-dispatching
            //25/07/2011: Nuova implementazione 
            //    Aggiornamento del 25/07/2011
            //" "= lascia inalterato il valore presente
            //"0" = articolo che NON eroga PIN 
            //          (ricarica telefonica alla cassa)
            //"1" = articolo che eroga PIN
            //         (ricarica telefonica alla cassa)

            //Aggiornamento del 06/11/2015
            //"4" = articolo "bollettino postale"
            //"5" = articolo "gift-card“

            //BODTRA	1 (C)	243-243	flag tracciabilità plu per carne bovine	‘ ‘ = articolo non tracciabile
            //0 = articolo non tracciabile
            //1 = articolo tracciabile selex
            //3 = articolo tracciabile
            //4 = articolo rintracciabile macinato
            //BODTIE	1 (C)	244-244	tipo etichette	non gestito
            //BODSNE	1 (C)	245-245	stampa etichette
            //Tipo articolo 2	non gestito
            //C = articolo contributo (es. smaltimento pneumatici“)

            //BODTPR	1 (C)	246-246	tipo record	T = testata
            //A = articolo annullato food
            //B = cancellazione del solo codice a barre 
            //       (non usato)
            //C = cancellazione completa articolo 
            //       (non usato)
            //D = dettaglio anagrafica prodotti freschi
            //       (ortofrutta)
            //E = articolo eliminato nofood
            //F = offerte fidelity
            //L = descrizioni aggiuntive bilance
            //       (da attivare con Mersy)
            //M = modifiche e/o inserimenti
            //N = inserimento Ean
            //       (non usato)
            //O = offerte speciali
            //P = premi fidelity
            //V = solo modifiche
            //       (non usato)
            //X = articolo figlio per articoli composti
            //      BODPAD = codice articolo padre
            //G = articolo a cui associare il contributo (es. smaltimento pneumatici)
            //       BODPAD = codice articolo contributo 

            //Nel record di tipo “T” vengono valorizzati solo i seguenti campi:
            //BODCDE, BODDSA, BODPRV, BODCES, BODSCV, BODTPR
            //*codice listino : 	blank = valori espressi in lire
            //        EURO2 = valori espressi in €uro con 2 decimali

            //Nei records di tipo “D” vengono valorizzati solo i seguenti campi:
            //BODCDE, BODCDA, BODSET, BODFAM, BODSUF, BODDSA, BODDS1, BODTPR


            //Tipo Record L:
            //codice negozio	4 (C)
            //Codice articolo	14 (C) allineato a sinistra
            //Giorni di conservazione	2 (C)
            //Filler	20 (C)
            //Testi bilancia	205 (C)
            //Tipo record (L)	1 (C)

            string sTip = strRig.Substring(245, 1);

            if (sTip == "A" || sTip == "C" || sTip == "E")
                rowArt["div_stf"] = "N";                        //Stato articolo del fornitore

            //Campo “Testi bilancia” è formato da:
            //Dimensione carattere 	1(C)
            //1° testo bilancia 	50(C)
            //carattere separatore: TAB o | (PIPE) 1(C)
            //Dimensione carattere 	1 (C)
            //2° testo bilancia 	50(C)
            //carattere separatore: TAB o | (PIPE) 1(C)
            //ecc.

            //rowArt["div_stf"] = "";
            //rowArt["div_dva"] = _clsFun.Str2Day(s);
            //rowArt["div_stf"] = "N";

            return rowArt;
        }

        public DataRow DivBil(string strRig, DataRow rowArt)
        {
            string s = "";

            //BODCDE	4 (C)	1-4	codice negozio	
            //BODCDA	7 (C)	5-11	codice articolo	
            //rowArt["div_arf"] = strRig.Substring(4, 7).Trim();

            //BODSET	3 (C)	12-14	settore merceologico (livello 1)	
            //BODFAM	3 (C)	15-17	Famiglia/reparto (livello 2)	
            //BODSUF	6 (C)	18-23	Sottofamiglia/categoria  (livello 3 ECR)	

            //BODEAN	13 (C)	24-36	codice a barre	

            //BODDSA	40 (C)	37-76	descrizione articolo	se BODTPR=D:

            //da 1 a 10 = calibro
            //da 11 a 15 = categoria
            //da 16 a 30 = provenienza/categoria pesce
            //da 31 a 40 = filler

            s = strRig.Substring(36, 10).Trim();
            if (s.Length > 20)
                s = s.Substring(0, 20);
            rowArt["for_cal"] = s;

            s = strRig.Substring(46, 5).Trim();
            if (s.Length > 20)
                s = s.Substring(0, 20);
            rowArt["for_cat"] = s;

            s = strRig.Substring(76, 20).Trim();
            if (s.Length > 20)
                s = s.Substring(0, 20);
            rowArt["for_ori"] = _clsFun.FaiLApice(s);

            //BODDS1	20 (C)	77-96	descrizione articolo scontrino per BODTPR diverso da T e D	se BODTPR=D: provenienza

            //BODREU	4 (C)	97-100	reparto casse	
            //rowArt["for_rep"] = strRig.Substring(97, 3);

            //BODPRV	7 (C)	101-107	prezzo di vendita
            //(se negativo nel primo byte c’è il segno - )
            //(se a ZERO nel primo byte c’è il segno + )
            //(se prezzo richiesto al primo byte c’è il segno ? )	*se BODTPR=T: codice listino
            //rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(101, 6)) / 100;

            //25/07/2011 nuova implementazione gestita solo in Mersy (no IND)
            //BODIVA	2 (C)	108-109	codice iva	
            //rowArt["for_iva"] = (strRig.Substring(107, 2)).Trim().PadLeft(3, Convert.ToChar("0"));

            //BODCNF	2 (C)	110-111	confezione	

            //BODTAR	3 (C)	112-114	tara per prodotti a peso	
            //rowArt["for_tar"] = strRig.Substring(112, 3);              da gestire ??????????????

            //BODRIO	1 (C)	115-115	C = centrale
            //D = diretti
            //P = piattaforma (dismesso dal 2013)
            //V = ventilazione da SPG 
            //        (07/10/2014: nuova implementazione)	

            //BODSCO	4 (C)	116-119	sconto percentuale	2 interi + 2 decimali

            //BODTPG	2 (C)	120-121	tipo grammatura	KG, LT, MT, PZ
            //rowArt["for_tgr"] = strRig.Substring(119, 2);

            //BODGRA	6 (C)	122-127	Grammatura	3 interi + 3 decimali
            //rowArt["div_pne"] = Convert.ToDecimal(strRig.Substring(121, 6))/1000;
            //s = strRig.Substring(121, 6);
            //if (_clsFun.Numerico(s))
            //    rowArt["div_pne"] = Convert.ToDecimal(s) / 1000;

            //BODSGO	6 (C)	128-133	Sgocciolato	3 interi + 3 decimali
            //s = strRig.Substring(127, 6);
            //if (_clsFun.Numerico(s))
            //    rowArt["div_pne"] = Convert.ToDecimal(s) / 1000;

            //BODUMV	3 (C)	134-136	unità misura di vendita	001 = pz
            //002 = kg
            //s = "KG";
            //if (strRig.Substring(119, 2) == "001")
            //    s = "NR";
            //rowArt["for_umi"] = s;

            //BODMPS	1 (C)	137-137	molteplicità punti set	da 0 a 9
            //BODFCL	1 (C)	138-138	categoria cliente fedele
            //Saldo punti	A=campagna attuale
            //B=campagna precedente
            //BODFPR	1 (C)	139-139	flag per premi fidelity	S = premio omaggio
            //N = premio con contributo
            //BODFTM	1 (C)	140-140	flag offerte fidelity	M = modifiche C = cancellazioni

            //BODCDF	6 (C)	141-146	codice fornitore	
            //rowArt["div_fo2"] = strRig.Substring(140, 6);

            //BODARF	7 (C)	147-153	codice articolo fornitore	
            //rowArt["div_af2"] = strRig.Substring(141, 7);

            //BODBO2	2 (C)	154-155	bollini (in migliaia)	
            //BODQCC	2 (C)	156-157	q.tà componenti cesti	
            //BODPMP	1 (C)	158-158	vendite a peso o a corpo	
            //•	P = vendita a peso per articoli con codice bilancia
            //•	M = vendita a corpo per articoli con codice bilancia
            //rowArt["div_bpz"] = false;
            //if (strRig.Substring(157, 1) == "M")
            //    rowArt["div_bpz"] = true;

            //BODBOL	3 (C)	159-161	bollini fidelity	

            //BODCES	7 (C)	162-168	prezzo di cessione	*se BODTPR=T: codice listino
            //rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(161, 7)) / 1000;

            //BODUMA	3 (C)	169-171	unità misura di acquisto	003 fisso

            //BODPEZ	4 (C)	172-175	pezzi/collo	
            //rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(172, 4));

            //BODCOL	3 (C)	176-178	colli per pallet	

            //BODPLU	6 (C)	179-184	codice bilancia (plu)	
            //rowArt["div_plu"] = strRig.Substring(178, 6).Trim();

            //BODDT1	8 (C)	185-192	data validità/inizio off.	formato AAAAMMGG
            //BODDT1	8 (C)	185-192	data validità/inizio off.	formato AAAAMMGG
            //BODDT2	8 (C)	193-200	data fine offerta	formato AAAAMMGG
            //BODDT3	8 (C)	201-208	data fine validità	formato AAAAMMGG
            //BODPAD	7 (C)	209-215	Sconto a valore
            //codice padre per cesti se BODTPR=’X’
            //codice contributo se BODTPR=’G’	
            //BODRAG	4 (C)	216-219	raggruppamento per MxN	Stesso valore se prezzo e merceologia uguali
            //BODMIX	2 (C)	220-221	mix (M nel MxN)	
            //BODMCH	2 (C)	222-223	match (N nel MxN)	
            //BODSTA	3 (C)	224-226	Da 224 a 225 Evidenziazione articolo
            //SE = articolo Selex
            //PB = il prezzo più basso
            //BI = biologico
            //SG = senza glutine
            //RB = ribassato e bloccato

            //Da 226 a 226 filler
            //Dal 04/03/2013: buono pasto: 
            //" "= lascia inalterato il valore presente
            //“0“ = articolo NON pabagile tramite BuonoPasto
            //“1“ = articolo pagabile tramite BuonoPasto

            //BODTPA	3 (C)	227-229	Tipo articolo	multiplo di vendita
            //BODARC	3 (C)	230-232	articolo collegato

            //03/06/2014
            //da 230 a 230:
            //flag articolo scontabile per assistiti celiaci:
            //‘0’ = disattiva flag articolo per celiaci
            //‘1’ = attiva flag articolo per celiaci
            //Blank = Lascia inalterato il valore precedente

            //Da 231 a 232: filler
            //    non gestito
            //BODSCV	7 (C)	233-239	sconto a valore fidelity (in centesimi)	*se BODTPR=T: codice listino
            //BODDET	2 (C)	240-241	% detrazione

            //Mix-Match per gestione promozioni:
            //- sconto
            //- cartolina






            //----------------------------------------------------------
            //25/07/2011: Nuova implementazione 


            //    non gestito

            //SCONTO
            //"  "= lascia inalterato il valore presente
            //"00"=articolo soggetto a sconto 
            //"01"=articolo esente da sconto: ricariche telefoniche, ricariche telefoniche e gratta e vinci
            //“02“ = articolo esente da sconto: alimenti per celiaci

            //--------------------------------------------------
            //CARTOLINE
            //"11"=1 articolo 1 cartolina
            //"12"=1 articolo 2 cartoline
            //"13"=1 articolo 3 cartoline
            //"14"=1 articolo 4 cartoline
            //"15"=1 articolo 5 cartoline

            //"21"=2 articoli 1 cartolina
            //"22"=2 articoli 2 cartoline
            //"23"=2 articoli 3 cartoline
            //"24"=2 articoli 4 cartoline
            //"25"=2 articoli 5 cartoline

            //…

            //"61"=6 articoli 1 cartolina
            //"62"=6 articoli 2 cartoline
            //"63"=6 articoli 3 cartoline
            //"64"=6 articoli 4 cartoline
            //"65"=6 articoli 5 cartoline

            //Liberi dal 66 al 99

            //BODFL1
            //    1 (C)	242-242	Tipo articolo 1
            //Erogazione di PIN tramite il software NCR UPB pin-dispatching
            //25/07/2011: Nuova implementazione 
            //    Aggiornamento del 25/07/2011
            //" "= lascia inalterato il valore presente
            //"0" = articolo che NON eroga PIN 
            //          (ricarica telefonica alla cassa)
            //"1" = articolo che eroga PIN
            //         (ricarica telefonica alla cassa)

            //Aggiornamento del 06/11/2015
            //"4" = articolo "bollettino postale"
            //"5" = articolo "gift-card“

            //BODTRA	1 (C)	243-243	flag tracciabilità plu per carne bovine	‘ ‘ = articolo non tracciabile
            //0 = articolo non tracciabile
            //1 = articolo tracciabile selex
            //3 = articolo tracciabile
            //4 = articolo rintracciabile macinato
            //BODTIE	1 (C)	244-244	tipo etichette	non gestito
            //BODSNE	1 (C)	245-245	stampa etichette
            //Tipo articolo 2	non gestito
            //C = articolo contributo (es. smaltimento pneumatici“)

            //BODTPR	1 (C)	246-246	tipo record	T = testata
            //A = articolo annullato food
            //B = cancellazione del solo codice a barre 
            //       (non usato)
            //C = cancellazione completa articolo 
            //       (non usato)
            //D = dettaglio anagrafica prodotti freschi
            //       (ortofrutta)
            //E = articolo eliminato nofood
            //F = offerte fidelity
            //L = descrizioni aggiuntive bilance
            //       (da attivare con Mersy)
            //M = modifiche e/o inserimenti
            //N = inserimento Ean
            //       (non usato)
            //O = offerte speciali
            //P = premi fidelity
            //V = solo modifiche
            //       (non usato)
            //X = articolo figlio per articoli composti
            //      BODPAD = codice articolo padre
            //G = articolo a cui associare il contributo (es. smaltimento pneumatici)
            //       BODPAD = codice articolo contributo 

            //Nel record di tipo “T” vengono valorizzati solo i seguenti campi:
            //BODCDE, BODDSA, BODPRV, BODCES, BODSCV, BODTPR
            //*codice listino : 	blank = valori espressi in lire
            //        EURO2 = valori espressi in €uro con 2 decimali

            //Nei records di tipo “D” vengono valorizzati solo i seguenti campi:
            //BODCDE, BODCDA, BODSET, BODFAM, BODSUF, BODDSA, BODDS1, BODTPR


            //Tipo Record L:
            //codice negozio	4 (C)
            //Codice articolo	14 (C) allineato a sinistra
            //Giorni di conservazione	2 (C)
            //Filler	20 (C)
            //Testi bilancia	205 (C)
            //Tipo record (L)	1 (C)

            //string sTip = strRig.Substring(245, 1);

            //if (sTip == "A" || sTip == "C" || sTip == "E")
            //    rowArt["div_stf"] = "N";                        //Stato articolo del fornitore


            //Campo “Testi bilancia” è formato da:
            //Dimensione carattere 	1(C)
            //1° testo bilancia 	50(C)
            //carattere separatore: TAB o | (PIPE) 1(C)
            //Dimensione carattere 	1 (C)
            //2° testo bilancia 	50(C)
            //carattere separatore: TAB o | (PIPE) 1(C)
            //ecc.


            //rowArt["div_stf"] = "";
            //rowArt["div_dva"] = _clsFun.Str2Day(s);
            //rowArt["div_stf"] = "N";

            return rowArt;
        }

        public DataRow DivIng(string strRig, DataRow rowArt)
        {
            string s = "";

            //BODCDE	4 (C)	1-4	codice negozio	
            //BODCDA	7 (C)	5-11	codice articolo	
            //rowArt["div_arf"] = strRig.Substring(4, 7).Trim();

            //BODSET	3 (C)	12-14	settore merceologico (livello 1)	
            //BODFAM	3 (C)	15-17	Famiglia/reparto (livello 2)	
            //BODSUF	6 (C)	18-23	Sottofamiglia/categoria  (livello 3 ECR)	

            //BODEAN	13 (C)	24-36	codice a barre	

            //BODDSA	40 (C)	37-76	descrizione articolo	se BODTPR=D:

            //da 1 a 10 = calibro
            //da 11 a 15 = categoria
            //da 16 a 30 = provenienza/categoria pesce
            //da 31 a 40 = filler

            s = strRig.Substring(40, 200).Trim();
            //if (s.Length > 20)
            //    s = s.Substring(0, 20);
            rowArt["div_ing"] = s;

            //s = strRig.Substring(46, 5).Trim();
            //if (s.Length > 20)
            //    s = s.Substring(0, 20);
            //rowArt["for_cat"] = s;

            //s = strRig.Substring(76, 20).Trim();
            //if (s.Length > 20)
            //    s = s.Substring(0, 20);
            //rowArt["for_ori"] = _clsFun.FaiLApice(s);

            //BODDS1	20 (C)	77-96	descrizione articolo scontrino per BODTPR diverso da T e D	se BODTPR=D: provenienza

            //BODREU	4 (C)	97-100	reparto casse	
            //rowArt["for_rep"] = strRig.Substring(97, 3);

            //BODPRV	7 (C)	101-107	prezzo di vendita
            //(se negativo nel primo byte c’è il segno - )
            //(se a ZERO nel primo byte c’è il segno + )
            //(se prezzo richiesto al primo byte c’è il segno ? )	*se BODTPR=T: codice listino
            //rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(101, 6)) / 100;

            //25/07/2011 nuova implementazione gestita solo in Mersy (no IND)
            //BODIVA	2 (C)	108-109	codice iva	
            //rowArt["for_iva"] = (strRig.Substring(107, 2)).Trim().PadLeft(3, Convert.ToChar("0"));

            //BODCNF	2 (C)	110-111	confezione	

            //BODTAR	3 (C)	112-114	tara per prodotti a peso	
            //rowArt["for_tar"] = strRig.Substring(112, 3);              da gestire ??????????????

            //BODRIO	1 (C)	115-115	C = centrale
            //D = diretti
            //P = piattaforma (dismesso dal 2013)
            //V = ventilazione da SPG 
            //        (07/10/2014: nuova implementazione)	

            //BODSCO	4 (C)	116-119	sconto percentuale	2 interi + 2 decimali

            //BODTPG	2 (C)	120-121	tipo grammatura	KG, LT, MT, PZ
            //rowArt["for_tgr"] = strRig.Substring(119, 2);

            //BODGRA	6 (C)	122-127	Grammatura	3 interi + 3 decimali
            //rowArt["div_pne"] = Convert.ToDecimal(strRig.Substring(121, 6))/1000;
            //s = strRig.Substring(121, 6);
            //if (_clsFun.Numerico(s))
            //    rowArt["div_pne"] = Convert.ToDecimal(s) / 1000;

            //BODSGO	6 (C)	128-133	Sgocciolato	3 interi + 3 decimali
            //s = strRig.Substring(127, 6);
            //if (_clsFun.Numerico(s))
            //    rowArt["div_pne"] = Convert.ToDecimal(s) / 1000;

            //BODUMV	3 (C)	134-136	unità misura di vendita	001 = pz
            //002 = kg
            //s = "KG";
            //if (strRig.Substring(119, 2) == "001")
            //    s = "NR";
            //rowArt["for_umi"] = s;

            //BODMPS	1 (C)	137-137	molteplicità punti set	da 0 a 9
            //BODFCL	1 (C)	138-138	categoria cliente fedele
            //Saldo punti	A=campagna attuale
            //B=campagna precedente
            //BODFPR	1 (C)	139-139	flag per premi fidelity	S = premio omaggio
            //N = premio con contributo
            //BODFTM	1 (C)	140-140	flag offerte fidelity	M = modifiche C = cancellazioni

            //BODCDF	6 (C)	141-146	codice fornitore	
            //rowArt["div_fo2"] = strRig.Substring(140, 6);

            //BODARF	7 (C)	147-153	codice articolo fornitore	
            //rowArt["div_af2"] = strRig.Substring(141, 7);

            //BODBO2	2 (C)	154-155	bollini (in migliaia)	
            //BODQCC	2 (C)	156-157	q.tà componenti cesti	
            //BODPMP	1 (C)	158-158	vendite a peso o a corpo	
            //•	P = vendita a peso per articoli con codice bilancia
            //•	M = vendita a corpo per articoli con codice bilancia
            //rowArt["div_bpz"] = false;
            //if (strRig.Substring(157, 1) == "M")
            //    rowArt["div_bpz"] = true;

            //BODBOL	3 (C)	159-161	bollini fidelity	

            //BODCES	7 (C)	162-168	prezzo di cessione	*se BODTPR=T: codice listino
            //rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(161, 7)) / 1000;

            //BODUMA	3 (C)	169-171	unità misura di acquisto	003 fisso

            //BODPEZ	4 (C)	172-175	pezzi/collo	
            //rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(172, 4));

            //BODCOL	3 (C)	176-178	colli per pallet	

            //BODPLU	6 (C)	179-184	codice bilancia (plu)	
            //rowArt["div_plu"] = strRig.Substring(178, 6).Trim();

            //BODDT1	8 (C)	185-192	data validità/inizio off.	formato AAAAMMGG
            //BODDT1	8 (C)	185-192	data validità/inizio off.	formato AAAAMMGG
            //BODDT2	8 (C)	193-200	data fine offerta	formato AAAAMMGG
            //BODDT3	8 (C)	201-208	data fine validità	formato AAAAMMGG
            //BODPAD	7 (C)	209-215	Sconto a valore
            //codice padre per cesti se BODTPR=’X’
            //codice contributo se BODTPR=’G’	
            //BODRAG	4 (C)	216-219	raggruppamento per MxN	Stesso valore se prezzo e merceologia uguali
            //BODMIX	2 (C)	220-221	mix (M nel MxN)	
            //BODMCH	2 (C)	222-223	match (N nel MxN)	
            //BODSTA	3 (C)	224-226	Da 224 a 225 Evidenziazione articolo
            //SE = articolo Selex
            //PB = il prezzo più basso
            //BI = biologico
            //SG = senza glutine
            //RB = ribassato e bloccato

            //Da 226 a 226 filler
            //Dal 04/03/2013: buono pasto: 
            //" "= lascia inalterato il valore presente
            //“0“ = articolo NON pabagile tramite BuonoPasto
            //“1“ = articolo pagabile tramite BuonoPasto

            //BODTPA	3 (C)	227-229	Tipo articolo	multiplo di vendita
            //BODARC	3 (C)	230-232	articolo collegato

            //03/06/2014
            //da 230 a 230:
            //flag articolo scontabile per assistiti celiaci:
            //‘0’ = disattiva flag articolo per celiaci
            //‘1’ = attiva flag articolo per celiaci
            //Blank = Lascia inalterato il valore precedente

            //Da 231 a 232: filler
            //    non gestito
            //BODSCV	7 (C)	233-239	sconto a valore fidelity (in centesimi)	*se BODTPR=T: codice listino
            //BODDET	2 (C)	240-241	% detrazione

            //Mix-Match per gestione promozioni:
            //- sconto
            //- cartolina






            //----------------------------------------------------------
            //25/07/2011: Nuova implementazione 


            //    non gestito

            //SCONTO
            //"  "= lascia inalterato il valore presente
            //"00"=articolo soggetto a sconto 
            //"01"=articolo esente da sconto: ricariche telefoniche, ricariche telefoniche e gratta e vinci
            //“02“ = articolo esente da sconto: alimenti per celiaci

            //--------------------------------------------------
            //CARTOLINE
            //"11"=1 articolo 1 cartolina
            //"12"=1 articolo 2 cartoline
            //"13"=1 articolo 3 cartoline
            //"14"=1 articolo 4 cartoline
            //"15"=1 articolo 5 cartoline

            //"21"=2 articoli 1 cartolina
            //"22"=2 articoli 2 cartoline
            //"23"=2 articoli 3 cartoline
            //"24"=2 articoli 4 cartoline
            //"25"=2 articoli 5 cartoline

            //…

            //"61"=6 articoli 1 cartolina
            //"62"=6 articoli 2 cartoline
            //"63"=6 articoli 3 cartoline
            //"64"=6 articoli 4 cartoline
            //"65"=6 articoli 5 cartoline

            //Liberi dal 66 al 99

            //BODFL1
            //    1 (C)	242-242	Tipo articolo 1
            //Erogazione di PIN tramite il software NCR UPB pin-dispatching
            //25/07/2011: Nuova implementazione 
            //    Aggiornamento del 25/07/2011
            //" "= lascia inalterato il valore presente
            //"0" = articolo che NON eroga PIN 
            //          (ricarica telefonica alla cassa)
            //"1" = articolo che eroga PIN
            //         (ricarica telefonica alla cassa)

            //Aggiornamento del 06/11/2015
            //"4" = articolo "bollettino postale"
            //"5" = articolo "gift-card“

            //BODTRA	1 (C)	243-243	flag tracciabilità plu per carne bovine	‘ ‘ = articolo non tracciabile
            //0 = articolo non tracciabile
            //1 = articolo tracciabile selex
            //3 = articolo tracciabile
            //4 = articolo rintracciabile macinato
            //BODTIE	1 (C)	244-244	tipo etichette	non gestito
            //BODSNE	1 (C)	245-245	stampa etichette
            //Tipo articolo 2	non gestito
            //C = articolo contributo (es. smaltimento pneumatici“)

            //BODTPR	1 (C)	246-246	tipo record	T = testata
            //A = articolo annullato food
            //B = cancellazione del solo codice a barre 
            //       (non usato)
            //C = cancellazione completa articolo 
            //       (non usato)
            //D = dettaglio anagrafica prodotti freschi
            //       (ortofrutta)
            //E = articolo eliminato nofood
            //F = offerte fidelity
            //L = descrizioni aggiuntive bilance
            //       (da attivare con Mersy)
            //M = modifiche e/o inserimenti
            //N = inserimento Ean
            //       (non usato)
            //O = offerte speciali
            //P = premi fidelity
            //V = solo modifiche
            //       (non usato)
            //X = articolo figlio per articoli composti
            //      BODPAD = codice articolo padre
            //G = articolo a cui associare il contributo (es. smaltimento pneumatici)
            //       BODPAD = codice articolo contributo 

            //Nel record di tipo “T” vengono valorizzati solo i seguenti campi:
            //BODCDE, BODDSA, BODPRV, BODCES, BODSCV, BODTPR
            //*codice listino : 	blank = valori espressi in lire
            //        EURO2 = valori espressi in €uro con 2 decimali

            //Nei records di tipo “D” vengono valorizzati solo i seguenti campi:
            //BODCDE, BODCDA, BODSET, BODFAM, BODSUF, BODDSA, BODDS1, BODTPR


            //Tipo Record L:
            //codice negozio	4 (C)
            //Codice articolo	14 (C) allineato a sinistra
            //Giorni di conservazione	2 (C)
            //Filler	20 (C)
            //Testi bilancia	205 (C)
            //Tipo record (L)	1 (C)

            //string sTip = strRig.Substring(245, 1);

            //if (sTip == "A" || sTip == "C" || sTip == "E")
            //    rowArt["div_stf"] = "N";                        //Stato articolo del fornitore


            //Campo “Testi bilancia” è formato da:
            //Dimensione carattere 	1(C)
            //1° testo bilancia 	50(C)
            //carattere separatore: TAB o | (PIPE) 1(C)
            //Dimensione carattere 	1 (C)
            //2° testo bilancia 	50(C)
            //carattere separatore: TAB o | (PIPE) 1(C)
            //ecc.


            //rowArt["div_stf"] = "";
            //rowArt["div_dva"] = _clsFun.Str2Day(s);
            //rowArt["div_stf"] = "N";

            return rowArt;
        }

        //public DataRow DivItt(string strRig, DataRow rowArt)
        //{
        //    string s = "";
        //    s = strRig.Substring(40, 200).Trim();
        //    rowArt["itt_arf"] = s;
        //    return rowArt;
        //}

        public DataRow DivFat(string strRig, DataRow rowArt)
        {
            string s = "";

            //string[] a = strRig.Split('|');

            ////string sRi1 = a[0];
            ////string sRi2 = a[1];

            ////Data
            //s = strRig.Substring(11, 8);
            //rowArt["div_dva"] = _clsFun.Str2Day(s);

            ////Codice articolo                                
            //rowArt["div_arf"] = strRig.Substring(33, 7).Trim();

            ////Descriz. articolo              
            //rowArt["div_ard"] = strRig.Substring(40).Trim();

            ////Quantità fattura               
            //s = strRig.Substring(26, 9);
            //if (_clsFun.Numerico(s))
            //    rowArt["div_qta"] = Convert.ToDecimal(s) / 100;
            //else
            //    rowArt["div_qta"] = 0;

            ////Importo unitario         
            //rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(35, 11)) / 1000;


            s = strRig.Substring(11, 8);
            rowArt["div_dva"] = _clsFun.Str2Day(s);

            s = strRig.Substring(20, 13).Trim();
            if (s.Length == 13 && s.Substring(0, 1) == "2" && s.Substring(7, 5) == "00000")
                s = s.Substring(0, 7) + "000000";
            rowArt["div_ean"] = s.Trim();

            rowArt["div_arf"] = strRig.Substring(33, 7).Trim();

            if ((string)rowArt["div_arf"] == "2010065")
                Console.WriteLine("aaaaaaaaaa");

            rowArt["div_ard"] = strRig.Substring(40, 21).Trim();

            s = strRig.Substring(64, 2).Trim();
            if (s == "$$")
                s = "";
            //rowArt["for_tgr"] = strRig.Substring(64, 2).Trim();
            rowArt["for_tgr"] = s;

            //s = "KG";
            //if ((string)rowArt["for_tgr"] == "PZ")
            //    s = "PZ";
            //rowArt["for_umi"] = s;

            s = strRig.Substring(66, 7).Trim();
            if (_clsFun.Numerico(s))
                rowArt["div_pne"] = Convert.ToDecimal(s) / 100;

            s = strRig.Substring(80, 2).Trim();
            if (s.Length > 0)
                rowArt["for_iva"] = s.PadLeft(3, Convert.ToChar("0"));
            //rowArt["div_iva"] = s.PadLeft(3, Convert.ToChar("0"));

            //s = strRig.Substring(82, 7);
            //if (_clsFun.Numerico(s))
            //    rowArt["div_pxc"] = Convert.ToDecimal(s);

            s = strRig.Substring(82, 9);
            if (_clsFun.Numerico(s))
                rowArt["div_qta"] = Convert.ToDecimal(s) / 100;
            else
                rowArt["div_qta"] = 0;
            //rowArt["div_qta"] = 1;

            s = strRig.Substring(93, 9).Trim();
            if (_clsFun.Numerico(s))
                rowArt["div_cos"] = _clsFun.Txt2Dec(s) / 1000;

            s = strRig.Substring(103, 9).Trim();
            if (_clsFun.Numerico(s))
                rowArt["div_prp"] = _clsFun.Txt2Dec(s) / 100;

            return rowArt;
        }

        public DataRow DivBol(string strRig, DataRow rowArt)
        {
            string s = "";

            s = strRig.Substring(11, 8);
            rowArt["div_dva"] = _clsFun.Str2Day(s);

            //s = strRig.Substring(20, 13);
            s = strRig.Substring(20, 13).Trim();
            if (s.Length == 13 && s.Substring(0, 1) == "2" && s.Substring(7, 5) == "00000")
                s = s.Substring(0, 7) + "000000";
            rowArt["div_ean"] = s.Trim();

            rowArt["div_arf"] = strRig.Substring(33, 7).Trim();

            if ((string)rowArt["div_arf"] == "2010065")
                Console.WriteLine("aaaaaaaaaa");


            rowArt["div_ard"] = strRig.Substring(40, 21).Trim();

            rowArt["for_tgr"] = strRig.Substring(64, 2).Trim();

            //s = "KG";
            //if ((string)rowArt["for_tgr"] == "PZ")
            //    s = "PZ";
            //rowArt["for_umi"] = s;

            s = strRig.Substring(66, 7).Trim();
            if (_clsFun.Numerico(s))
                rowArt["div_pne"] = Convert.ToDecimal(s) / 100;

            s = strRig.Substring(80, 2).Trim();
            if(s.Length > 0)
                rowArt["for_iva"] = s.PadLeft(3, Convert.ToChar("0"));
            //rowArt["div_iva"] = s.PadLeft(3, Convert.ToChar("0"));

            //s = strRig.Substring(82, 7);
            //if (_clsFun.Numerico(s))
            //    rowArt["div_pxc"] = Convert.ToDecimal(s);

            s = strRig.Substring(82, 9);
            if (_clsFun.Numerico(s))
                rowArt["div_qta"] = Convert.ToDecimal(s) / 100;
            else
                rowArt["div_qta"] = 0;
            //rowArt["div_qta"] = 1;

            s = strRig.Substring(93, 9).Trim();
            if (_clsFun.Numerico(s))
                rowArt["div_cos"] = _clsFun.Txt2Dec(s) / 1000;

            //s = strRig.Substring(103, 9).Trim();
            //if (_clsFun.Numerico(s))
            //    rowArt["div_prp"] = _clsFun.Txt2Dec(s) / 100;

            s = strRig.Substring(103, 9).Trim();
            if (_clsFun.Numerico(s))
                rowArt["div_prp"] = _clsFun.Txt2Dec(s) / 100;

            return rowArt;
        }

    }
}
