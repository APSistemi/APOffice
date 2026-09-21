using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsDivDado
    {

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public clsDivDado()
        {}

        public DataRow DivArt(string strRig, DataRow rowArt)
        {

            string[] a = strRig.Split(';');

            string sArt = a[4];
            string sEst = a[5];
            string sDes = a[6];
            string sUmi = a[8];
            string sCos = a[9];
            string sPve = a[10];
            string sIva = a[11];
            string sPxc = a[12];
            string sEan = a[15];

            sArt = sArt.PadLeft(5, Convert.ToChar('0'));
            sEst = sEst.PadLeft(2, Convert.ToChar('0'));

            sArt = sArt + sEst;

            if (sIva.Length > 2)
                sIva = "";

            if (sEan.Length > 13 || sEan.Length < 8)
                sEan = "";

            string s = "";
            rowArt["div_dva"] = DateTime.Today.ToShortDateString();  //OK

            //Codice articolo 
            rowArt["div_arf"] = sArt;          //OK

            if (sArt == "2534801")
                Console.WriteLine("AAAA");

            //Descriz. articolo
            rowArt["div_ard"] = sDes;    //

            rowArt["for_tgr"] = "GR";

            if ((string)rowArt["div_arf"] == "861083")
                Console.WriteLine("aaa");

            rowArt["for_rep"] = "001";    //OK
            rowArt["for_ecr"] = ""; // strRig.Substring(63, 6);

            rowArt["for_iva"] = sIva.PadLeft(3,Convert.ToChar('0'));

            rowArt["for_umi"] = sUmi; // strRig.Substring(71, 2);

            rowArt["div_pxc"] = _clsFun.Txt2Dec(sPxc);

            rowArt["div_pne"] = 1; // _clsFun.Txt2Dec(strRig.Substring(57, 6)) / 100;

            rowArt["div_tpd"] = ""; // strRig.Substring(83, 1);

            //s = strRig.Substring(84, 7);
            //if (_clsFun.Numerico(s))
            //    rowArt["div_qta"] = _clsFun.Txt2Dec(s) / 100;
            //else
            rowArt["div_qta"] = 0;

            rowArt["div_cof"] = 0; // _clsFun.Txt2Dec(sCos);

            rowArt["div_imp"] = 0;  // _clsFun.Txt2Dec(strRig.Substring(100, 11)) / 100;

            rowArt["div_tva"] = ""; // _clsFun.Txt2Dec(strRig.Substring(127, 1));

            rowArt["div_cos"] = _clsFun.Txt2Dec(sCos);

            //EAN13
            //s = strRig.Substring(149, 13);
            //if (_clsFun.Numerico(s))
            //    s = Convert.ToInt64(s).ToString();
            rowArt["div_ean"] = sEan;

            //PRZV Prezzo vendita consigliato
            rowArt["div_prp"] = _clsFun.Txt2Dec(sPve);

            //if (((string)rowArt["for_tgr"]).Trim() == "CA")  //Acquisto a  cartone
            //{
            //    rowArt["for_tgr"] = rowArt["for_umi"];
            //    rowArt["for_umi"] = "NR";
            //}

            return rowArt;
        }


    }
}
