using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsDivVefri
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public clsDivVefri() { }

        public DataRow DivFat(string strRig, DataRow rowArt, DateTime dayDdo)
        {
            string s = "";

            //TIPO RECORD	2	1	2		        Lasciare in bianco
            //COD CLIENTE	6	3	8		        Codice cliente

            //NUMERO BOLLA	6	9	14		
            //DATA BOLLA	6	15	20		        AAMMGG
            rowArt["div_dva"] = _clsFun.Str2Day(strRig.Substring(14,6));

            //TIPO MERCE	1	21	21		        0=vendita; 1 resi; 2=omaggio

            //COD.ARTICOLO	15	22	36		
            rowArt["div_arf"] = strRig.Substring(21, 15).Trim();

            //QUANTITA'	8	37	44                  2 decimali
            rowArt["div_qta"] = Convert.ToDecimal(strRig.Substring(36, 8).Trim())/100;

            //PREZZO NETTO	9	45	53		        4 decimali
            rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(44, 9)) / 1000;

            //UNITA' MISURA	2	54	55				PZ/KG
            rowArt["for_umi"] = strRig.Substring(53, 2).Trim();

            //DESCR.ARTICOLO	22	56	77		
            rowArt["div_ard"] = strRig.Substring(55, 22).Trim();

            //COD. ARTICOLO VEGA	7	78	84		

            //NUMERO FATTURA	15	85	99
            //DATA FATTURA	6	100	105				AAMMGG
            rowArt["div_imp"] = (decimal)rowArt["div_cos"] * (decimal)rowArt["div_qta"];

            //LIBERO 1	9	106	114		
            //VALUTA	1	115	115		
            //LIBERO 2	13	116	128		

            rowArt["div_pxc"] = Convert.ToDecimal(strRig.Substring(78, 5).Trim());
            rowArt["for_iva"] = "0" + strRig.Substring(84, 2).Trim();
            rowArt["div_ean"] = strRig.Substring(89).Trim();

            return rowArt;
        }

    }
}
