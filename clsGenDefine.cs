using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace APOffice
{
    class clsDefine
    {
        public enum enuParametri
        {
            ParForDefault = 1,
            ParDivForDownLoad = 2,
            ParRitardoChiusura = 3,
            ParOrdForDividi = 4,
            ParGesGraph = 5,
            ParDivPrezziAZero = 6,
            ParPrefixEabBilancia = 7,
            ParOmegaBarcodeLun = 8,
            ParStatVisProf = 9,                     //Vendite proforma incluse
            ParEcrDefault = 10,
            Par011InsArtCodMod = 11,                //Codice articolo modificabile in inserimento
            Par012CodIVAxDefault = 12,
            Par013CodRepxDefault = 13,
            Par014AggPVenditaSoloFattura = 14,      //Aggiornamento prezzi fornitore principale solo fattura elettronica + tipo per confronto costo
            Par015AggDesArtFromFornitore = 15,
            Par016PathDivNegozi = 16,               //Path e codice divulgazione negozi
            Par017EcrDivFornitori = 17,             //Controllo ECR da divulgazione fornitori
            Par018NuovoArticoloNo = 18,             //Disattivazione inserimento articolo
            Par019VarForzata = 19,                  //Variazione forzata da div fornitori
            Par020EsclusioneEan2021davDiv = 20,     //Esclusione barcode pref 20 21 da div fornitori
            Par021GesIngredienti = 21,              //Tipo gestione ingredienti 1/2, su tracciato bilancia winvarp
            Par022DocRaggruppArticoli = 22,         //Stampa documenti con raggruppamento articoli
            Par023Clienti2Pos = 23,                 //Divulgazione clienti alle casse per fattura 
            Par024DecimalPesoPrezzo = 24,           //Decimali su peso e importo in stampa DDT
            Par025DocX2Pdf = 25,                    //Doppio documento su stessa pagina
            Par026FattFromCassa = 26,               //Parametri inserimento fattura in automatico da cassa
            Par027Fidelity = 27,                    //Tipo fidelity 2 = multitessera
            Par028ParBilance = 28,                  //Parametri bilance 1-Prezzo modificabile
            Par029ParPathWeb = 29,                  //Parametri percorso generazione file per il WEB su art_web
            Par030ParPathIng = 30,                  //Parametri percorso generazione file ingredienti per il libro
            Par031Lotti2Pos = 31,                   //Nuova gestione lotti con gestione da riga di carico divulgazione alle casse + prefisso barcode
            Par032ImportEanizzati = 32,             //Regola import articoli eanizzati
            Par033FattureElettroniche = 33,         //Gestione fatture elettroniche agenzia entrate
            Par034ParamFattureDiff = 34,            //Parametri fatture differite
            Par035BilScontrino = 35,                //Parametri lettura scontrino bilancia
            Par036Div2Sede = 36,                    //Divulgazione a sede: 1 path, 2 Documenti, 3 Inventario, 4 Import ApPhone
            Par037ParCeliaci = 37,                  //Parametri gestione celiaci
            Par038Password = 38,                    //Parametri password
            Par039SoloPrezziNeg2Pos = 39,           //Divulgazioni in cassa solo prezzi negozio
        }

        public enum enuNumeratori
        {
            NumAnaArticoli = 1,
            NumAnaClienti = 2,
            NumAnaFornitori = 3,
            NumGesOfferte = 4,
            NumGesMovimenti = 5,        //Movimenti
            NumGesMovFatture = 6,       //Movimenti fatture
            NumGesDocFatture = 7,       //Documenti fatture
            NumGesDocDdt = 8,           //Documenti DDT
            NumGesOrdFor = 9,
            NumGesTesGru = 10,          //Legami tessere tra nuova e vecchia (vecchio sistema)
            NumGesDocLotti = 11,
            NumAnaTessere = 12,         //Codice cliente tessera
            NumGesDocCeliaci = 13,      //Numeratore documenti celiaci
            NumGesFatElettroniche = 14, //Numeratore fatture elettroniche
            NumGesDocRiparazioni = 15,  //Numeratore Riparazioni
            NumGesMovInterni = 16,      //Numeratore Movimenti interni
        }

        public enum enuIni
        {
            Ini00APPath = 0,                    //Path ApOffice
            IniAPOfficeSql = 1,                 //Connessione SQL ApOffice
            IniAPOfficeMdb = 2,                 //Connessione MDB
            IniAPOfficeSqlStat = 3,             //Connessione SQL ApStat
            IniUserLast = 4,                    //Ultimo utente
            IniCrenziali = 5,                   //Accesso senza richiesta utente
            Ini06PathApShopCasse = 6,           //Non usato
            Ini07PathDivCasse = 7,              //Path di generazione file di divulgazione alle casse
            Ini08PathDivBilance = 8,            //Path di generazione file di divulgazione alle bilance
            Ini09CodiceAzienda = 9,             //Codice azienda per CnfConfig in caso di multizienda
            Ini10NumeroTerminale = 10,          //Numero terminale attivo per evitare di entrare in gestione variazioni da 2 postazioni
            Ini11VenditaTouch = 11,             //Impostazioni gestione documenti touch
            Ini12BilCheckOut = 12,              //Impostazioni bilancia check-out
            Ini13Ubicazione = 13,               //00 per Sede oppure = al codice negozio
            Ini14Password = 14,                 //Password varie
            Ini15AIAgent = 15,                   //Agente AI abilitato (1=Si, 0=No)
            Ini16DashBoard = 16,                 //Dashboard abilitata all'avvio (1=Si, 0=No)
            Ini17Modal = 17                      //Apertura maschere modale (1=Si, 0=No)
        }

        public enum enuPos
        {
            posDitron = 1,
            posBrainpos = 2,
            posNcr745x = 3,
            posUgaSid20 = 4,
            posApShop01 = 5
        }

        public enum enuBilance
        {
            bilOmega = 1,
            bilBizerba = 2,
            bilBizWinVarp = 3,
            bilBizerba2 = 4
        }

        public string ERRLOG = "C:\\APproject\\Log\\Error.log";
        public string FILLOG = "C:\\APproject\\Log\\Log.log";
        public string FILEDIR = "C:\\APproject\\DirApOffice.txt";
        public string FILLOGDIVFOR = "C:\\APproject\\Log\\LogDivFor.log";
        public string FILEINI = "C:\\APproject\\APOffice\\DataBase\\ApOffice.ini";
        public string PATHINGREDIENTI = "C:\\APproject\\Temp\\Ingredienti\\";
        public string PATHREPORT = "C:\\APproject\\APOffice\\Report\\";
        public string PATHLOGHI = "C:\\ApProject\\Temp\\Img\\Loghi\\";
        public string PATHPDF = "C:\\APproject\\PDF\\";
        public string BATCHBACKUP = "C:\\APproject\\Temp\\SQL\\BckApOffice.bat";
        public string PATHTABCSV = "C:\\APproject\\Temp\\SQL\\Tabelle.csv";
        public DateTime DAYOUT = new DateTime(2050, 1, 1, 0, 0, 0);
        public string MDBGEN = "DataBase.mdb";
        public string CODNEW = "NEW";
        public string COD03Z = "000";
        public string COD04Z = "0000";
        public string COD03X = "XXX";
        public string COD04X = "XXXX";
        public string COD06X = "XXXXXX";
        public string TIPCLI = "CLI";
        public string TIPFOR = "FOR";

        public string LISPOS = "001";
        public string LISFOR = "002";
        public string LISOFF = "003";
        public string LISPRO = "004";

        public string VARPOS = "POS";
        public string VARETI = "ETI";
        public string VARALL = "ALL";
        public string CRLF = "\r\n";
        public string LF = "\n";

        public string STAATT = "A";
        public string STAREP = "R";
        public string STANOA = "N";
        public string STADAA = "D";     //Da attivare
        public string STADAC = "Y";     //Da chiudere
        public string STACLO = "X";     //Chiusa
        public string STAPRE = "P";     //Premio
        public string STACAN = "C";     //Cancellata

        public string DIVDAD = "0";     //Da divulgare
        public string DIVDIV = "1";     //Divulgato
        public string DIVSOS = "3";     //Sospesa
        public string DIVANN = "4";     //Annullata

        public string OFAPRZ = "001";
        public string OFASCO = "002";
        public string OFAMXN = "003";
        public string OFASCV = "004";   //Sconto valore

        public string TMPPOSSTOP = "C:\\APproject\\Temp\\DivCasse\\Stop.smf";
        public string SMFVARPOS = "C:\\APproject\\APOffice\\DataBase\\GesVariaz";     //+ numero terminale + ".smf"
        public string TMPPOSVAR = "C:\\APproject\\Temp\\DivCasse\\PosVar.txt";
        public string TMPBILVAR = "C:\\APproject\\Temp\\DivBilance\\BilVar.txt";
        public string TMPBILING = "C:\\APproject\\Temp\\DivBilance\\BilIng.txt";
        public string TMPPOSOFF = "C:\\APproject\\Temp\\DivCasse\\PosOff.txt";
        public string TMPPOSFID = "C:\\APproject\\Temp\\DivCasse\\PosFid.txt";
        public string TMPPOSCLI = "C:\\APproject\\Temp\\DivCasse\\PosCli.txt";
        public string TMPPOSCLD = "C:\\APproject\\Temp\\DivCasse\\PosCld.txt";        //Destinazioni clienti
        public string TMPPOSOTR = "C:\\APproject\\Temp\\DivCasse\\PosOtr.txt";
        public string TMPPOSFIDMON = "C:\\APproject\\Temp\\DivCasse\\PosFidMon.txt";
        //public string TMPFORDIV = "\\";
        public string TMPPOSVEN = "C:\\APproject\\Temp\\Venduto\\";

        public string PATHORD = "C:\\APproject\\Temp\\OrdFornitori\\";
        public string DIVTERRON = "TERRON";
        public string DIVBRENDO = "BRENDOLAN";
        public string DIVEUROSPESA = "EUROSPESA";
        public string DIVDPIU = "DPIU";
        public string DIVAPMATCH = "APMATCH";
        public string DIVMIGROSS = "MIGROSS";
        public string DIVUNICOM = "UNICOM";

        public string PATHMEMOR = "C:\\APproject\\Temp\\DivMemor\\";
        public string TERMMEMOR = "001";
        public string TERMCSV1 = "002";
        public string TERMPDT1 = "003";
        public string TERMTDIV = "004";
        public string TERMOPH3 = "005";
        public string PUNTITESSERAX = "PUNTITESSERAX";

        public clsDefine()
        {
            SetPathParam();
        }

        private void SetPathParam()
        {
            string sCod = "000";
            string sPar = "";

            if (File.Exists("Ini.ini"))
            {
                using (StreamReader sr = new StreamReader("Ini.ini"))
                {
                    string sRig = "";

                    while ((sRig = sr.ReadLine()) != null)
                    {
                        if (sRig.Length > 3 && sRig.Substring(0, 3) == sCod)
                        {
                            if (sRig.Length > 4)
                                sPar = sRig.Substring(4).Trim();
                            break;
                        }
                    }
                    sr.Close();
                    sr.Dispose();
                }

                if (sPar != "")
                {
                    string sPathIni = sPar;

                    if (sPathIni != "")
                    {
                        ERRLOG = sPathIni + "\\Log\\Error.log";
                        FILLOG = sPathIni + "\\Log\\Log.log";
                        FILEDIR = sPathIni + "\\DirApOffice.txt";
                        FILLOGDIVFOR = sPathIni + "\\Log\\LogDivFor.log";
                        FILEINI = sPathIni + "\\APOffice\\DataBase\\ApOffice.ini";
                        PATHINGREDIENTI = sPathIni + "\\Temp\\Ingredienti\\";
                        PATHREPORT = sPathIni + "\\APOffice\\Report\\";
                        PATHLOGHI = sPathIni + "\\Temp\\Img\\Loghi\\";
                        PATHPDF = sPathIni + "\\PDF\\";
                        BATCHBACKUP = sPathIni + "\\Temp\\SQL\\BckApOffice.bat";
                        PATHTABCSV = sPathIni + "\\Temp\\SQL\\Tabelle.csv";

                        SMFVARPOS = sPathIni + "\\APOffice\\DataBase\\GesVariaz";     //+ numero terminale + ".smf"
                        TMPPOSVAR = sPathIni + "\\Temp\\DivCasse\\PosVar.txt";
                        TMPBILVAR = sPathIni + "\\Temp\\DivBilance\\BilVar.txt";
                        TMPBILING = sPathIni + "\\Temp\\DivBilance\\BilIng.txt";
                        TMPPOSOFF = sPathIni + "\\Temp\\DivCasse\\PosOff.txt";
                        TMPPOSFID = sPathIni + "\\Temp\\DivCasse\\PosFid.txt";
                        TMPPOSCLI = sPathIni + "\\Temp\\DivCasse\\PosCli.txt";
                        TMPPOSCLD = sPathIni + "\\Temp\\DivCasse\\PosCld.txt";        //Destinazioni clienti
                        TMPPOSOTR = sPathIni + "\\Temp\\DivCasse\\PosOtr.txt";
                        TMPPOSFIDMON = sPathIni + "\\Temp\\DivCasse\\PosFidMon.txt";
                        //TMPFORDIV = "\\";
                        TMPPOSVEN = sPathIni + "\\Temp\\Venduto\\";
                        PATHORD = "\\APproject\\Temp\\OrdFornitori\\";
                    }

                }
            }
        }

    }
}
