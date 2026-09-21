using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsTermCsv2
    {
        clsDefine _clsDef = new clsDefine();

        private string sSep = "|";

        public clsTermCsv2()
        { }

        public string TermArtTes()
        {
            string sRig = "T" + sSep;
            sRig += "art_art" + sSep;
            sRig += "art_ard" + sSep;
            sRig += "art_umi" + sSep;
            sRig += "art_tgr" + sSep;
            sRig += "art_red" + sSep;
            sRig += "art_pxc" + sSep;
            sRig += "art_pne" + sSep;
            sRig += "art_pve" + sSep;
            sRig += "art_iva" + sSep;
            return sRig;
        }

        public string TermEanTes()
        {
            string sRig = "T" + sSep;
            sRig += "ean_art" + sSep;
            sRig += "ean_ean" + sSep;
            return sRig;
        }

        public string TermLiaTes()
        {
            string sRig = "T" + sSep;
            sRig += "lia_tip" + sSep;
            sRig += "lia_arf" + sSep;
            sRig += "lia_art" + sSep;
            sRig += "lia_for" + sSep;
            sRig += "lia_fod" + sSep;
            sRig += "lia_pxc" + sSep;
            sRig += "lia_cos" + sSep;
            sRig += "lia_mrg" + sSep;
            sRig += "lia_dti" + sSep;
            return sRig;
        }

        public string TermArtRow(DataRow rowPos, DataTable tabRep)
        {
            string s = "";
            string sRig = "R" + sSep;
            DataRow[] j;

            //Codice articolo
            s = ((string)rowPos["pos_art"]).Trim();
            sRig += s + sSep;

            //Descrizione
            s = ((string)rowPos["pos_ard"]).Trim();
            sRig += s + sSep;

            //UMI
            s = "NR";
            if (((string)rowPos["pos_umi"]).Length > 1)
                s = ((string)rowPos["pos_umi"]).Substring(0, 2);
            sRig += s + sSep;

            //TGR
            s = "PZ";
            if (((string)rowPos["pos_tgr"]).Length > 1)
                s = ((string)rowPos["pos_tgr"]).Substring(0, 2);
            sRig += s + sSep;

            //Reparto descrizione
            s = "";

            if (tabRep == null)
            {
                s = "";
                if (!DBNull.Value.Equals((string)rowPos["RepDes"]) && ((string)rowPos["RepDes"]).Length == 2)
                s = ((string)rowPos["RepDes"]).Substring(0, 2);
            }
            else
            {
                j = tabRep.Select("tab_cod='" + (string)rowPos["pos_rep"] + "'");
                if (j.Length > 0)
                    s = (string)j[0]["tab_des"];
            }
            sRig += s + sSep;

            //Pxc
            //s = "1";
            s = Convert.ToInt32(rowPos["pos_pxc"]).ToString();
            sRig += s + sSep;

            //Pne 
            s = ((decimal)rowPos["pos_pne"]).ToString().Replace(",", ".");
            sRig += s + sSep;

            //Prezzo di vendita
            s = ((decimal)rowPos["pos_prv"]).ToString().Replace(",", ".");
            sRig += s + sSep;

            //IVA
            s = ((string)rowPos["pos_iva"]).ToString();
            sRig += s + sSep;

            ////Data di modifica
            //s = DateTime.Today.ToString("dd/MM/yyyy");
            //sRig += s + sSep;

            return sRig;
        }

        public string TermEanRow(DataRow rowPos)
        {
            string s = "";
            string sRig = "R" + sSep;
            DataRow[] j;
            Boolean b = true;

            //Codice articolo
            s = ((string)rowPos["pos_art"]).Trim();
            if (s.Length > 0)
                sRig += s + sSep;
            else
                b = false;

            //Barcode
            s = ((string)rowPos["pos_ean"]).Trim();
            if (s.Length > 0)
                sRig += s + sSep;
            else
                b = false;

            if (!b)
                sRig = "";

            return sRig;
        }

        public string TermLiaRow(DataRow rowPos, DataRow rowLia)
        {
            Boolean b = true;

            //string sArt = (string)rowPos["lia_art"];

            string s = "";
            string sRig = "R" + sSep;

            //Tipo listino
            s = ((string)rowLia["lia_tip"]).Trim();
            sRig += s + sSep;

            //Codice articolo fornitore
            s = ((string)rowLia["lia_arf"]).Trim();
            sRig += s + sSep;

            //Articolo
            s = ((string)rowLia["lia_art"]).Trim();
            if (s.Length == 0)
                b = false;
            else
                sRig += s + sSep;

            //Fornitore
            s = ((string)rowLia["lia_for"]).Trim();
            if (s.Length == 0)
                b = false;
            else
                sRig += s + sSep;

            //Fornitore
            s = (string)rowLia["LiaFod"];
            sRig += s + sSep;

            //Pxc
            //s = "1";
            s = Convert.ToInt32((decimal)rowLia["lia_pxc"]).ToString();
            sRig += s + sSep;

            //Costo
            s = ((decimal)rowLia["lia_cos"]).ToString().Replace(",", ".");
            sRig += s + sSep;

            //Data di modifica
            s = ((DateTime)rowLia["lia_dti"]).ToString("yyyy-MM-dd");
            sRig += s + sSep;

            if (!b)
                sRig = "";

            return sRig;
        }


    }
}
