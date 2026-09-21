using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsTermCsv1
    {
        clsDefine _clsDef = new clsDefine();

        private string sSep = ",";

        public clsTermCsv1()
        { }

        public string TermArtTes()
        {
            string sRig = "T" + sSep;
            sRig += "art_cod" + sSep;
            sRig += "art_ean" + sSep;
            sRig += "art_des" + sSep;
            sRig += "art_umi" + sSep;
            sRig += "art_pxc" + sSep;
            sRig += "art_cxp" + sSep;
            sRig += "art_pco" + sSep;
            sRig += "art_pve" + sSep;
            sRig += "art_dmo" + sSep;
            return sRig;
        }

        public string TermArtRow(DataRow rowPos, DataRow[] rowArf)
        {
            string s = "";
            string sRig = "R" + sSep;

            //Codice articolo
            s = new string(' ', 15);
            if (rowArf.Length > 0)
                s = ((string)rowArf[0]["lia_arf"]).Trim();
            sRig += s + sSep;

            //Barcode, se a peso primi 6 crt
            s = ((string)rowPos["pos_ean"]).Trim();
            sRig += s + sSep;

            //Descrizione
            s = ((string)rowPos["pos_ard"]).Trim();
            sRig += s + sSep;

            //UMI
            s = ((string)rowPos["pos_umi"]).Substring(0, 2);
            sRig += s + sSep;

            //Pxc
            s = "1";
            if (rowArf.Length > 0)
                s = Convert.ToInt32((decimal)rowArf[0]["lia_pxc"]).ToString();
            sRig += s + sSep;

            //Colli x strato
            s = "1";
            if (rowArf.Length > 0)
                s = Convert.ToInt32((decimal)rowArf[0]["lia_cxp"]).ToString();
            sRig += s + sSep;

            //Costo con 3 decimali
            s = "0.0";
            if (rowArf.Length > 0)
                s = ((decimal)rowArf[0]["lia_cos"]).ToString().Replace(",",".");
            sRig += s + sSep;

            //Prezzo di vendita
            s = ((decimal)rowPos["pos_prv"]).ToString().Replace(",", ".");
            sRig += s + sSep;

            //Data di modifica
            s = DateTime.Today.ToString("dd/MM/yyyy");
            sRig += s + sSep;

            return sRig;
        }

    }
}
