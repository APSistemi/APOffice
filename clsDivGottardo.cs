using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsDivGottardo
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public string _strPar017EcrDivFornitori = "";

        public clsDivGottardo() { }

        public DataRow DivArt(string strRig, DataRow rowArt)
        {
            string s = "";

            if (strRig.Substring(26, 6) == "817995")
                Console.WriteLine("aaaa");

            //codice punto vendita	N	1	3	

            //data variazione	N	4	8	SSAAMMGG
            rowArt["div_dva"] = _clsFun.Str2Day(strRig.Substring(3, 8));

            //tipo operazione	A	12	1	" "=inser "C"=canc "V"=variaz
            rowArt["div_stf"] = "";
            if(strRig.Substring(11,1)== "C")
                rowArt["div_stf"] = "N";

            //tipo barcode	N	13	1	0=ean/upc/interno/etc 4=prod.lettura prezzo 5=prod.lettura peso
            //barcode	N	14	13	
            rowArt["div_ean"] = "";
            s = strRig.Substring(13, 13).Trim();
            if (_clsFun.Numerico(s))
            {
                s = Convert.ToDouble(s).ToString();
                rowArt["div_ean"] = s;
            }

            //codice prodotto	N	27	6	
            rowArt["div_arf"] = strRig.Substring(26, 6).Trim();

            //descrizione	A	33	30	
            rowArt["div_ard"] = strRig.Substring(32, 30).Trim();

            //unità di misura	A	63	2	
            rowArt["for_umi"] = strRig.Substring(62, 2).Trim();

            //reparto cassa	N	65	2	
            rowArt["for_rep"] = strRig.Substring(64, 2).Trim();

            //I livello	N	67	2	
            //II livello	N	69	2	
            //III livello	N	71	2	
            //IV livello	N	73	2	
            //valuta	A	75	3	

            rowArt["for_ecr"] = "";
            if (_strPar017EcrDivFornitori == "S")
                rowArt["for_ecr"] = strRig.Substring(66, 8);

            rowArt["for_rep"] = strRig.Substring(64, 2).Trim();

            //n.decimali prezzi	N	78	1	
            //prezzo acquisto	N	79	8	(3 decimali)
            rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(78, 8)) / 1000;

            //prezzo vendita ivato	N	87	8	(2 decimali)
            rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(86, 8)) / 100;

            //prezzo unitario	N	95	8	=prezzo vendita se tipo=4, altrimenti =0 (3 decimali)
            //%sconto 1	N	103	4	(2 decimali)
            //%sconto 2	N	107	4	(2 decimali)
            //data inizio sconto 2	N	111	8	SSAAMMGG
            //data fine sconto2	N	119	8	SSAAMMGG
            //UM prezzo in chiaro	A	127	2	"  " o GR o ML
            rowArt["for_tgr"] = strRig.Substring(126, 2);

            //Qta prezzo in chiaro	N	129	4	
            s = strRig.Substring(128, 4);
            if (_clsFun.Numerico(s))
                rowArt["div_pne"] = Convert.ToDecimal(s);

            //flag peso-prezzo	A	133	1	L=Lettura a prezzo, P=Lettura a peso
            //codice iva	A	134	3	
            rowArt["for_iva"] = (strRig.Substring(133, 2)).Trim().PadLeft(3, Convert.ToChar("0"));

            //aliquota iva	N	137	4	(2 decimali)
            //segnalazione	A	141	2	
            //descrizione breve	A	143	15	
            rowArt["div_adb"] = strRig.Substring(142, 15).Trim();

            //codice fornitore	N	158	6	
            //calo peso	N	164	4	(2 decimali)

            //unità vendita Cedis	N	168	6	(2 decimali)

            //pezzi x cartone (fornitore)	N	174	6	
            rowArt["div_pxc"] = _clsFun.Txt2Dec(strRig.Substring(173, 6));

            //pezzi termo (fornitore)	N	180	6	
            //flag tipo offerta	A	186	1	" " no offerta, G=2x1, H=3x2, I=4x2, K=6x4, L=taglio prezzo
            //numero offerta	N	187	8	SSAAXXXX
            //data inizio offerta	N	195	8	SSAAMMGG
            //data fine offerta	N	203	8	SSAAMMGG
            //flag1	A	211	1	prodotto con etichetta S/N
            //flag2	A	212	1	Categoria offerta Fidelity (0=no fid/ 1...9)
            //flag3	A	213	1	
            //flag4	A	214	1	
            //flag5	A	215	1	magazzino " "=cedis  "*" diretto
            //etichetta	A	216	1	S=variato elemento x etichetta, N=no

            //prezzo consigliato	N	217	8	(2 decimali)
            //rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(126, 8)) / 1000;

            //ScAcqOs	N	225	4	(2 decimali) non gestito
            //ScVenOs	N	229	4	(2 decimali) sconto % offerta
            //Reparto OS	N	233	2	Punti fidelity
            //NxM-AA2	N	235	2	numeratore NxM
            //NxM-NUM	N	237	4	
            //NxM-MASTER	N	241	6	
            //        246	1	
            //Espositore/Singolo		247	1	E= Espositore / S= Prodotto venduto singolarmente

            return rowArt;
        }

        public DataRow DivFat(string strRig, DataRow rowArt)
        {
            string s = "";

            //FT	1	12		Num	Numero Fattura	Compilato a Zero (0)
            //DTFT	13	8		Num	Data Fattura	aaaammgg Compilato a Zero (0)
            //NDOC	21	11		Alfa	Numero Bolla	

            //DTDOC	32	8		Num	Data Bolla	aaaammgg
            s = strRig.Substring(31, 8);
            rowArt["div_dva"] = _clsFun.Str2Day(s);

            //TDOC	40	1		Alfa	Tipo Documento	F=fattura, Q=Rettifica a Qta, V=Rettifica a Valore
            rowArt["div_tip"] = strRig.Substring(39, 1);

            //CODFOR	41	6		Num	Codice Fornitore	Compilato a Zero (0)
            //CODCLI	47	6		Num	Codice Cliente	

            //TE-PROCOMP	53	1		Num	Test prodotto composto	0=Prodotto normale 1=Prodotto composto 2=Componente prodotto composto
            rowArt["div_tri"] = strRig.Substring(52, 1);

            //CODPRO	54	6		Num	Codice prodotto	
            rowArt["div_arf"] = strRig.Substring(53, 6).Trim();

            //DES	60	30		Alfa	Descrizione prodotto	
            rowArt["div_ard"] = strRig.Substring(59, 30).Trim();

            //REPARTO	90	2		Num	Reparto	
            rowArt["for_rep"] = strRig.Substring(89, 2).Trim();

            //UM	92	2		Alfa	Unità di misura	
            rowArt["for_umi"] = strRig.Substring(91, 2).Trim();

            //PUM-UM	94	2		Alfa	Unità di misura per prezzo al Kilo/Litro	
            rowArt["for_tgr"] = strRig.Substring(93, 2).Trim();

            //PUM-QTA	96	8	2	Num	Quantità per prezzo al Kilo/Litro in gr/ml	
            s = strRig.Substring(95, 8);
            if (_clsFun.Numerico(s))
                rowArt["div_pne"] = Convert.ToDecimal(s);

            //QTAV	104	8	2	Num	Quantità vendita Cedis	
            s = strRig.Substring(103, 8);
            if (_clsFun.Numerico(s))
                rowArt["div_pxc"] = Convert.ToDecimal(s) / 100;
            else
                rowArt["div_pxc"] = 0;

            //CODIVA	112	3		Alfa	Codice IVA	
            rowArt["for_iva"] = strRig.Substring(111, 2).Trim();

            //ACQ	115	8	3	Num	Prezzo di acquisto	
            rowArt["div_cos"] = _clsFun.Txt2Dec(strRig.Substring(114, 8)) / 1000;

            //VEN	123	8	2	Num	Prezzo di Vendita al pubblico	
            rowArt["div_prp"] = _clsFun.Txt2Dec(strRig.Substring(122, 8)) / 100;

            //SEGNO	131	1		Alfa	Segno	

            //QTASPE	132	8	2	Num	Quantità spedita	
            s = strRig.Substring(131, 8);
            if (_clsFun.Numerico(s))
                rowArt["div_qta"] = Convert.ToDecimal(s) / 100;
            else
                rowArt["div_qta"] = 0;

            rowArt["div_stf"] = "";
            if ((decimal)rowArt["div_qta"] == 0)
                rowArt["div_stf"] = "N";

            //BCODE	140	13		Num	Barcode	
            //BCODE	153	13		Num	Barcode	
            //BCODE	166	13		Num	Barcode	
            //BCODE	179	13		Num	Barcode	
            //BCODE	192	13		Num	Barcode	
            //BCODE	205	13		Num	Barcode	
            //BCODE	218	13		Num	Barcode	
            //BCODE	231	13		Num	Barcode	
            //BCODE	244	13		Num	Barcode	
            //BCODE	257	13		Num	Barcode	
            //BCODE	270	13		Num	Barcode	
            //BCODE	283	13		Num	Barcode	
            //BCODE	296	13		Num	Barcode	
            //BCODE	309	13		Num	Barcode	
            //BCODE	322	13		Num	Barcode	

            return rowArt;
        }



    }
}
