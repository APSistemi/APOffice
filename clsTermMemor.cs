using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace APOffice
{
    class clsTermMemor
    {
        clsDefine _clsDef = new clsDefine();

        public clsTermMemor()
        { }

        public string TermArtRow(DataRow rowPos, DataRow[] rowArf)
        {
            string s = "";
            string sRig = "";
            Boolean bPes = false;
            // 1=inserimento
            sRig += "1";
            //Barcode, se a peso primi 6 crt
            s = ((string)rowPos["pos_ean"]).PadLeft(13, Convert.ToChar("0"));
            sRig += s;
            //Descrizione
            s = ((string)rowPos["pos_ard"] + new string(' ', 30)).Substring(0, 30);
            sRig += s;
            //Descrizione breve
            s = ((string)rowPos["pos_arb"] + new string(' ', 15)).Substring(0, 15);
            sRig += s;
            //Costo
            s = new string('0', 7);
            if (rowArf.Length > 0)
                s = Convert.ToInt32((decimal)rowArf[0]["lia_cos"] * 100).ToString("0000000");
            sRig += s;
            //Prezzo di vendita
            s = Convert.ToInt32((decimal)rowPos["pos_prv"] * 100).ToString("0000000");
            sRig += s;
            //Pxc
            s = new string('0', 4);
            if (rowArf.Length > 0)
                s = Convert.ToInt32((decimal)rowArf[0]["lia_pxc"]).ToString("0000");
            sRig += s;
            //UMI
            sRig += ((string)rowPos["pos_umi"]).Substring(0, 2);
            //Aliquota IVA
            sRig += ((string)rowPos["pos_iva"]).Substring(1, 2) + "000";
            //Reparto
            sRig += ((string)rowPos["pos_rep"] + new string(' ', 3)).Substring(1, 2);
            //Bilancia
            sRig += " ";
            //Produttore
            sRig += new string(' ', 10);
            //Codice articolo
            s = new string(' ', 15);
            if (rowArf.Length > 0)
                s = ((string)rowArf[0]["lia_arf"] + new string(' ', 15)).Substring(0, 15);
            sRig += s;
            //Codice articolo del fornitore
            s = new string(' ', 15);
            if (rowArf.Length > 0)
                s = ((string)rowPos["pos_art"] + new string(' ', 15)).Substring(0, 15);
            sRig += s;
            //Tipo contenuto
            sRig += ((string)rowPos["pos_tgr"]).Substring(0, 1);
            //Codice associazione prezzo
            sRig += "  ";
            //Codice fornitore
            s = new string(' ', 6);
            if (rowArf.Length > 0)
                s = "0" + (string)rowArf[0]["lia_for"];
            sRig += s;
            //Contenuto netto
            sRig += Convert.ToInt32((decimal)rowPos["pos_pne"] * 1000).ToString("000000");
            //Ecr Famiglia (settore)
            sRig += ((string)rowPos["pos_ecr"]).Substring(0,3);
            //Ecr settore
            sRig += ((string)rowPos["pos_ecr"]).Substring(1, 2);
            //Ecr fam
            sRig += ((string)rowPos["pos_ecr"]).Substring(4, 2);
            //Ecr sotto fam
            sRig += ((string)rowPos["pos_ecr"]).Substring(7, 2);
            //Indice conversione prezzo di acquisto
            sRig += "000000";
            //Imballo di vendita al pubblico
            sRig += "0001";
            //Tipo etichetta da stampare
            sRig += " ";
            //Numero etichetta da stampare
            sRig += "01";
            //Facing
            sRig += "00";
            //Colli x pallet
            s = "0001";
            if (rowArf.Length > 0)
                s = Convert.ToInt32((decimal)rowArf[0]["lia_cxp"]).ToString("0000");
            sRig += s;
            //Colli x strato
            sRig += "0001";
            //EAN a Peso/Valore
            s = " ";
            if ((Boolean)rowPos["bil_bil"])
                s = "2";
            sRig += s;
            //Master
            sRig += "0000000";
            //Costo con 3 decimali
            s = new string('0', 7);
            if (rowArf.Length > 0)
                s = Convert.ToInt32((decimal)rowArf[0]["lia_cos"] * 1000).ToString("0000000");
            sRig += s;
            //Punti concorso premi
            sRig += "0000";
            //Livello ottimale scorta
            sRig += "0001";
            //Magazzino: '1'=magazzino/'0'=T.D.
            sRig += "1";
            //Ordinabile: ' '=riordinabile/'N'=non riordinabile
            sRig += " ";
            //Sconti clienti permesso S/N
            sRig += "S";

            return sRig;
        }

    }
}
