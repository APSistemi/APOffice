using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace APOffice
{
    class clsPosDitron
    {
        clsDefine _clsDef = new clsDefine();

        public clsPosDitron()
        { }

        public string PosArtRow(DataRow rowPos )
        {
            string s = "";
            string sRig = "";
            Boolean bPes = false;

            decimal dPrv = (decimal)rowPos["pos_prv"];
            if ((decimal)rowPos["pos_eap"] > 0)
                dPrv = (decimal)rowPos["pos_eap"];

            if (!DBNull.Value.Equals(rowPos["pos_eaq"]) && (decimal)rowPos["pos_eaq"] > 1)
                dPrv = dPrv * (decimal)rowPos["pos_eaq"];

            //dPrv = dPrv * 100;

            //s = Convert.ToInt32((decimal)rowPos["pos_prv"] * (decimal)rowPos["pos_eaq"] * 100).ToString().PadLeft(9, Convert.ToChar(" "));

            // D=cancellato I=inserimento X=modifica
            if ((string)rowPos["pos_sta"] == "C")
                sRig += "D";
            else
                sRig += "I";
            sRig += "!";
            //Barcode, se a peso primi 7 crt
            s = (string)rowPos["pos_ean"];
            if (s.Length == 13 && s.Substring(0, 1) == "2" && s.Substring(7, 5) == "00000")
            {
                s = "   " + new string(' ', 6) + s.Substring(0, 7);
                bPes = true;
            }
            else
                s = "   " + (s.Trim() + new string(' ', 13)).Substring(0, 13);
            //s = "   " + s.PadLeft(13, Convert.ToChar(" "));

            sRig += s + "!";
            // 1=attivo per la vendita, 2 x fidelity
            sRig += "1!";
            //Descrizione
            s = (string)rowPos["pos_arb"];
            if (s.Length > 20)
                s = s.Substring(0, 20);
            s = (s + "                   ").Substring(0, 20);
            sRig += s + "!";

            //Prezzo di vendita
            //s = Convert.ToInt32((decimal)rowPos["pos_prv"] * 100).ToString().PadLeft(9, Convert.ToChar(" "));
            //if (!DBNull.Value.Equals(rowPos["pos_eaq"]) && (decimal)rowPos["pos_eaq"] > 1)
            //    s = Convert.ToInt32((decimal)rowPos["pos_prv"] * (decimal)rowPos["pos_eaq"] * 100).ToString().PadLeft(9, Convert.ToChar(" "));

            s = Convert.ToInt32(dPrv * 100).ToString().PadLeft(9, Convert.ToChar(" "));
            sRig += s + "!";
            //Aliquota IVA
            s = ((string)rowPos["pos_iva"]).Substring(1, 2) + "." + "00";
            sRig += s + "!";
            //Prezzo massimo
            sRig += "999999999" + "!";
            //Prezzo minimo
            sRig += "        0" + "!";
            //Offerte dalla 1 alla 20
            s = Convert.ToInt32(rowPos["pos_mix"]).ToString("00000");
            sRig += s + "!";
            for (int i = 1; i <= 19; i++)
                sRig += "00000!";
            //Offerte in q.tà
            s = "00";
            if (Convert.ToInt32(rowPos["pos_mix"]) > 0)
                s = "01";
            sRig += s + "!";
            //Offerte in ammontare
            sRig += "00!";
            //Impostazionemanuale manuale prezzo
            s = "3";
            if (Convert.ToInt32(rowPos["pos_mix"]) > 0)
                s = "3";
            sRig += s + "!";
            //Vidima solo per SF3ST/FP300
            sRig += "2!";
            //Impostazione q.tà
            s = "3";
            if (((string)rowPos["pos_tva"]).Trim() != "")
                s = "3";
            sRig += s + "!";
            //Impostazione stringa
            sRig += "2!";
            //Impostazione reso
            sRig += "1!";
            //Scontabile al subtotale
            sRig += "1!";
            //Articolo manualmente scontabile
            sRig += "1!";
            //Flag ean 2
            if(bPes)
            {
                if((Boolean)rowPos["pos_ecp"])
                    sRig += "7!";
                else
                    sRig += "6!";
            }
            else
                sRig += "3!";
            //Utilizzo di prezzi alternativi
            sRig += "1!";
            //Finalizzazione automatica
            sRig += "2!";
            //Codice vuoto
            sRig += "                !";
            //Collegamento al vuoto
            sRig += "00!";
            //Prezzo alternativo 1
            //s = Convert.ToInt32((decimal)rowPos["pos_prv"] * 100).ToString().PadLeft(9, Convert.ToChar(" "));
            s = Convert.ToInt32(dPrv * 100).ToString().PadLeft(9, Convert.ToChar(" "));
            sRig += s + "!";
            //Prezzo alternativo 2
            sRig += "        0!";
            //Prezzo alternativo 3
            //s = Convert.ToInt32((decimal)rowPos["pos_prv"] * 100).ToString().PadLeft(9, Convert.ToChar(" "));
            s = Convert.ToInt32(dPrv * 100).ToString().PadLeft(9, Convert.ToChar(" "));
            sRig += s + "!";
            //Prezzo alternativo 4
            sRig += "        0!";
            //Prezzo alternativo 5
            sRig += "        0!";
            //Rendita bollini
            sRig += "00000!";
            //Codice promozione fidelity
            sRig += "000!";
            //Codice ultima promozione
            sRig += "000!";
            //Codice penultima promozione
            sRig += "000!";
            //Costo bollini per fidelity
            sRig += "00000!";
            //Costo bollini fidelity nell'ultima promozione 
            sRig += "00000!";
            //Costo bollini per fidelity
            sRig += "00000!";
            //Decimali significativi per moltiplicazione
            sRig += "0!";
            //Flag moltiplicatori
            sRig += "1!";
            //Unità di misura
            sRig += "PEZ.!";
            //Reparto
            s = ((string)rowPos["pos_rep"]).PadLeft(4, Convert.ToChar("0"));
            sRig += s + "!";
            //Codice interno
            sRig += "!";
            //Descrizione estesa
            s = ((string)rowPos["pos_ard"] + "                                                 ").Substring(0, 40);
            sRig += s + "!";
            //Codice del fornitore
            s = "                ";
            sRig += s + "!";
            //giorno mese anno ultima vendita
            sRig += "000000!";
            //Data dell'inserimnto dell'articolo
            sRig += "000000!";
            //Panieri dal 1 al 9
            sRig += "  0!  0!  0!  0!  0!  0!  0!  0!  0!";
            //Codice unione per paniere
            sRig += "             ";

            return sRig;
        }

        public string PosOffArtRowSco(DataRow rowPos)
        {
            string s = "";
            string sRig = "";

            // D=cancellato I=inserimento X=modifica
            if ((string)rowPos["pos_sta"] == "C")
                sRig += "D";
            else
                sRig += "I";
            sRig += "!";
            //Offerte dalla 1 alla 20
            s = Convert.ToInt32(rowPos["pos_mix"]).ToString("    0");
            sRig += s + "!";
            //Stato 1=attiva 2=disattivata
            s = "2";
            if ((string)rowPos["pos_ost"] == _clsDef.STAATT)
                s = "1";
            sRig += s + "!";
            //Descrizione
            s = "SCONTO " + (string)rowPos["pos_ofd"];
            s = (s + "                   ").Substring(0, 20);
            sRig += s + "!";
            //Minuti + ore inizio
            sRig += "00!00!";
            //Data inizio offerta
            s = ((DateTime)rowPos["pos_dti"]).ToString("dd!MM!yy!");
            sRig += s;
            //Minuti + ore fine
            sRig += "59!23!";
            //Data fine offerta
            s = ((DateTime)rowPos["pos_dtf"]).ToString("dd!MM!yy!");
            sRig += s;
            //Giorni settimanali validi
            sRig += "127!";
            //Limite massimo e minimo
            sRig += "999999999!        1!";
            //Casualità offerta
            sRig += "00000!";
            //Soglia subtotale
            sRig += "        0!        0!        0!        0!        0!";
            //Prima soglia articolo in q.tà
            //s = ((decimal)rowPos["pos_xem"]).ToString("0.000").Replace(",", ".");
            //s = "    " + s;
            s = "    1.000";
            sRig += s + "!";
            //2,3,4^ soglia
            sRig += "    0.000!    0.000!    0.000!";
            //Soglia di attivazione
            sRig += "        0!";
            //Tipo offerta
            sRig += "10!";
            //Codice di riferimento
            sRig += "                !";
            //Offerta cumulativa mix-match
            sRig += "1!";
            //Valore offerta 1
            s = ((decimal)rowPos["pos_ova"]).ToString("0.00").Replace(",", ".");
            s = "    " + s;
            sRig += s + "!";
            //Valore offerta 2,3,4
            sRig += "        0!        0!        0!";
            //Applicabilità
            sRig += "2!";
            //Livello offerta
            sRig += "00!";
            //???
            sRig += "0  !";

            return sRig;
        }

        public string PosOffArtRowMxN(DataRow rowPos)
        {
            string s = "";
            string sRig = "";

            // D=cancellato I=inserimento X=modifica
            if ((string)rowPos["pos_sta"] == "C")
                sRig += "D";
            else
                sRig += "I";
            sRig += "!";
            //Offerte dalla 1 alla 20
            s = Convert.ToInt32(rowPos["pos_mix"]).ToString("    0");
            sRig += s + "!";
            //Stato 1=attiva 2=disattivata
            s = "2";
            if ((string)rowPos["pos_ost"] == _clsDef.STADAA)
                s = "1";
            sRig += s + "!";
            //Descrizione
            s = "Offerta " + Convert.ToString(rowPos["pos_xem"]) + "x" + Convert.ToString(rowPos["pos_xen"]);
            s = (s + "                   ").Substring(0, 20);
            sRig += s + "!";
            //Minuti + ore inizio
            sRig += "00!00!";
            //Data inizio offerta
            s = ((DateTime)rowPos["pos_dti"]).ToString("dd!MM!yy!");
            sRig += s;
            //Minuti + ore fine
            sRig += "59!23!";
            //Data fine offerta
            s = ((DateTime)rowPos["pos_dtf"]).ToString("dd!MM!yy!");
            sRig += s;
            //Giorni settimanali validi
            sRig += "127!";
            //Limite massimo e minimo
            sRig += "999999999!        1!";
            //Casualità offerta
            sRig += "00000!";
            //Soglia subtotale
            sRig += "        0!        0!        0!        0!        0!";
            //Prima soglia articolo in q.tà
            s = ((decimal)rowPos["pos_xem"]).ToString("0.000").Replace(",", ".");
            s = "    " + s;
            sRig += s + "!";
            //2,3,4^ soglia
            sRig += "    0.000!    0.000!    0.000!";
            //Soglia di attivazione
            sRig += "        0!";
            //Tipo offerta
            sRig += " 5!";
            //Codice di riferimento
            sRig += "                !";
            //Offerta cumulativa mix-match
            sRig += "1!";
            //Valore offerta 1
            s = ((decimal)rowPos["pos_xen"]).ToString("0.000").Replace(",", ".");
            s = "    " + s;
            sRig += s + "!";
            //Valore offerta 2,3,4
            sRig += "        0!        0!        0!";
            //Applicabilità
            sRig += "2!";
            //Livello offerta
            sRig += "00!";
            //???
            sRig += "0  !";

            return sRig;
        }

        public string PosOffTrans(DataRow rowPos)
        {
            string s = "";
            string sRig = "";

            // D=cancellato I=inserimento X=modifica
            if ((string)rowPos["pos_sta"] == "C")
                sRig += "D";
            else
                sRig += "I";
            sRig += "!";
            //Offerte dalla 1 alla 20
            s = Convert.ToInt32(rowPos["pos_mix"]).ToString("    0");
            sRig += s + "!";
            //Stato 1=attiva 2=disattivata
            s = "2";
            if ((string)rowPos["pos_ost"] == _clsDef.STADAA)
                s = "1";
            sRig += s + "!";
            //Descrizione
            s = "Offerta " + Convert.ToString(rowPos["pos_xem"]) + "x" + Convert.ToString(rowPos["pos_xen"]);
            s = (s + "                   ").Substring(0, 20);
            sRig += s + "!";
            //Minuti + ore inizio
            sRig += "00!00!";
            //Data inizio offerta
            s = ((DateTime)rowPos["pos_dti"]).ToString("dd!MM!yy!");
            sRig += s;
            //Minuti + ore fine
            sRig += "59!23!";
            //Data fine offerta
            s = ((DateTime)rowPos["pos_dtf"]).ToString("dd!MM!yy!");
            sRig += s;
            //Giorni settimanali validi
            sRig += "127!";
            //Limite massimo e minimo
            sRig += "999999999!        1!";
            //Casualità offerta
            sRig += "00000!";
            //Soglia subtotale
            sRig += "        0!        0!        0!        0!        0!";
            //Prima soglia articolo in q.tà
            s = ((decimal)rowPos["pos_xem"]).ToString("0.000").Replace(",", ".");
            s = "    " + s;
            sRig += s + "!";
            //2,3,4^ soglia
            sRig += "    0.000!    0.000!    0.000!";
            //Soglia di attivazione
            sRig += "        0!";
            //Tipo offerta
            sRig += " 5!";
            //Codice di riferimento
            sRig += "                !";
            //Offerta cumulativa mix-match
            sRig += "1!";
            //Valore offerta 1
            s = ((decimal)rowPos["pos_xen"]).ToString("0.000").Replace(",", ".");
            s = "    " + s;
            sRig += s + "!";
            //Valore offerta 2,3,4
            sRig += "        0!        0!        0!";
            //Applicabilità
            sRig += "2!";
            //Livello offerta
            sRig += "00!";
            //???
            sRig += "0  !";

            return sRig;
        }

        public string PosFidCliRow(DataRow rowPos)
        {
            string s = "";
            string sRig = "";

            //Tipo record I inserimento D delete
            if ((Boolean)rowPos["tes_ann"])
                sRig += "D";
            else
                sRig += "I";
            sRig += "!";

            //Codice cliente
            s = ((string)rowPos["tes_cod"]).PadRight(16, Convert.ToChar(" "));
            sRig += s + "!";

            //Stato
            if ((Boolean)rowPos["tes_ann"])
                sRig += "2";
            else
                sRig += "1";
            sRig += "!";

            //Scadenza
            s = ((DateTime)_clsDef.DAYOUT).ToString("dd!MM!yy!");
            sRig += s;

            //Ragione sociale
            sRig += new string(' ', 40) + "!";

            //Indirizzo
            sRig += new string(' ', 40) + "!";

            //Numero civico
            sRig += new string(' ', 6) + "!";

            //CAP
            sRig += new string(' ', 5) + "!";

            //Località
            sRig += new string(' ', 30) + "!";

            //Provincia
            sRig += new string(' ', 2) + "!";

            //P.IVA
            sRig += new string(' ', 14) + "!";

            //C.fiscale
            sRig += new string(' ', 16) + "!";

            //Codice cliente correlato
            sRig += new string(' ', 16) + "!";

            //Tipo livello cliente
            //s = ((string)rowPos["tes_gru"]).Substring(1, 2) + "!";
            s = "21" + "!";
            sRig += s;

            sRig += "00!00!00!00!00!00!00!00!00!";

            return sRig;
        }

        public string PosFidMonCliRow(DataRow rowPos, string strCam)
        {
            string s = "";
            string sRig = "";

            //Tipo record
            sRig += "X!";

            //Codice cliente
            s = ((string)rowPos["tes_cod"]).PadRight(16, Convert.ToChar(" "));
            sRig += s + "!!!!!!!!!";

            //Campagna
            sRig += "00" + strCam + "!!!";

            //Montante
            s = Convert.ToString(Convert.ToInt32(rowPos["TesPun"])).PadLeft(5, Convert.ToChar("0"));
            sRig += s;

            sRig += "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!";

            return sRig;
        }

    }
}
