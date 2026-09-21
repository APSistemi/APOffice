using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsPosNcr745x
    {
        clsDefine _clsDef = new clsDefine();

        public Boolean _bolImpianto = false;

        public clsPosNcr745x()
        { }

        public string PosArtRow(DataRow rowPos )
        {
            string s = "";
            //Dim sRiga As String
            string sRig = "";
            //Dim sOfTyp As String          'Tipo offerta
            string sOfTyp = "";
            //Dim cOfVal As Currency        'Valore offerta
            decimal cOfVal = 0;
            //Dim bOk As Boolean
            Boolean bOk = false;
            //Dim PosFi1 As String          'Filler 3
            string sPosFi1 = "   ";
            //Dim PosPlu As String          'Ean
            string sPosPlu = "";
            //Dim PosRep As String          'Reparto
            string sPosRep = "";
            //Dim PosCo1 As String          'Code 1 = '0', se promozione = '1'
            string sPosCo1 = "";
            //Dim PosCo2 As String          'Code 2 fisso '0'
            string sPosCo2 = "";
            //Dim PosSco As String          'Discount code fisso '1'
            string sPosSco = "";
            //Dim PosIva As String          'Codici IVA ( 0=esente, 1=4%, 2=10%, 3=20% )
            string sPosIva = "";
            //Dim PosCo3 As String          'Tara fisso 00
            string sPosCo3 = "";
            //Dim PosUmi As String          'Pack Type -> unità di misura
            string sPosUmi = "";
            //Dim PosDec As String          'Pack unit -> decimali €: fisso '0010'
            string sPosDec = "";
            //Dim PosDep As String          'Deposit link fisso '0000'
            string sPosDep = "";
            //Dim PosDes As String          'Descrizione articolo
            string sPosDes = "";
            //Dim PosCo5 As String          'Esenzione punti transazionali '0'
            string sPosCo5 = "";
            //Dim PosCo6 As String          'Etichetta a prezzo variabile in €
            string sPosCo6 = "";
            //Dim PosFi2 As String          'Filler 11
            string sPosFi2 = "";
            //Dim PosCo4 As String          'Filler 1
            string sPosCo4 = "";
            //Dim PosPrz As String          'Prezzo
            string sPosPrz = "";
            //Dim PosBol As Currency        'Bollini
            decimal cPosBol = 0;
            //Dim PosMix As String          'Codice Mix
            string sPosMix = "";

            //   PosFi1 = Space(3)

            //   'Controlli su barcode
   
            //   If Len(Trim(Nul2Space(m_rsVarPos!pos_ean))) = 11 Then
            //      PosPlu = " 0" + Trim(m_rsVarPos!pos_ean)
            //'   ElseIf Len(Trim(Nul2Space(m_rsVarPos!pos_ean))) = 4 Then
            //'      PosPlu = "     0000" + Trim(m_rsVarPos!pos_ean)
            //   Else
            //      PosPlu = FillChar(Trim(m_rsVarPos!pos_ean), 13, " ", "D")
            //   End If

            Boolean bPes = false;

            s = ((string)rowPos["pos_ean"]).Trim();
            if (s.Length == 11)
                sPosPlu = " 0" + s;
            //else if (s.Length == 4)
            //    sPosPlu = "     0000" + s;
            else
                sPosPlu = s.PadLeft(13, Convert.ToChar(" "));

            if (sPosPlu.Substring(0, 1) == "2" && sPosPlu.Length == 13 && sPosPlu.Substring(7, 6) == "000000")
            {
                s = sPosPlu.Substring(0, 12);
                sPosPlu = s + new clsCtrlCodici().FindMod10Digit(s);
                bPes = true;
            }
            //   PosRep = FillChar(Nul2Zero(m_rsVarPos!pos_rev), 4, "0", "D")
            sPosRep = ((string)rowPos["pos_rep"]).Trim().PadLeft(4,Convert.ToChar("0")); 
   
            //   If Not IsVuoto(Nul2Space(m_rsVarPos!pos_ovt)) Or _
            //      Not IsVuoto(Nul2Space(m_rsVarPos!pos_fta)) Or _
            //      Not IsVuoto(Nul2Space(m_rsVarPos!pos_ftb)) Then
            //      PosCo1 = "1"
            //   Else
            //      PosCo1 = "0"
            //   End If
            sPosCo1 = "0";
            if ((string)rowPos["pos_ost"] == _clsDef.STAATT)
                sPosCo1 = "1";
  
            sPosCo2 = "0";
            sPosSco = "1";
   
            //   m_tbTabIva.Seek Array("0" + m_rsVarPos!pos_iva), adSeekFirstEQ
            //   If m_tbTabIva.EOF Then
            //      PosIva = "0"
            //   Else
            //     PosIva = Left(Str(Nul2Zero(m_tbTabIva!tab_ncr)), 1)
            //     If IsVuoto(PosIva) Then
            //        PosIva = FillChar(m_tbTabIva!tab_ncr, 1, "0", "D")
            //     End If
            //   End If
            sPosIva = Convert.ToString(rowPos["IvaNcr"]);
   
            sPosCo3 = "00";

            sPosUmi = "PZ";
            if(bPes)
                sPosUmi = "KG";

            sPosDec = "0010";
            sPosDep = "0000";
   
            //   'Definizione descrizione
            //   If IsVuoto(m_rsVarPos!pos_des) Then
            //      PosDes = Space(20)
            //   Else
            //      PosDes = Left(m_rsVarPos!pos_des + Space(20), 20)
            //      If Ncr2190CtrlDes(PosDes) Then  'Controllo sulla descrizione
            //         PosDes = DesModify(PosDes)
            //         Print #3, "Articolo " + m_rsVarPos!pos_art + " modificata la descrizione x caratteri non ammessi"
            //      End If
            //   End If
            sPosDes = new string(' ', 20);
            s = ((string)rowPos["pos_arb"]).Trim();
            if (s != "")
            {
                if (s.Length > 20)
                    s = s.Substring(0, 20);
                else if (s.Length < 20)
                    s = s.PadRight(20, Convert.ToChar(" "));
                sPosDes = s; // +new string(' ', 20).Substring(0, 20);

            }
            sPosCo5 = "0";
            sPosCo6 = "1";

            sPosFi2 = "0" +  new string(' ', 10);
            if(bPes)
                sPosFi2 = "1" + new string(' ', 10);

            sPosCo4 = " ";
   
            //   'Se taglio prezzo lo imposto da qui
            //   '
            //   If Nul2Space(m_rsVarPos!pos_ovt) = "P" And _
            //      Nul2Zero(m_rsVarPos!pos_ova) > 0 Then
            //      PosPrz = FillChar(Multi(Nul2Zero(m_rsVarPos!pos_ova), 100), 8, "0", "D")
            //   Else
            //      PosPrz = FillChar(Multi(Nul2Zero(m_rsVarPos!pos_prz), 100), 8, "0", "D")
            //   End If
            sPosPrz = Convert.ToInt32((decimal)rowPos["pos_prv"] * 100).ToString().PadLeft(8, Convert.ToChar("0"));

            //   bOk = True
   
            //   'Controllo PLU
            //   If Nul2Zero(PosPlu) > 999999 Then
            //      If PosPlu <> CtrlDigit(Left(PosPlu, 12)) Then
            //         Print #3, "Articolo " + m_rsVarPos!pos_art + " non inviato perchè BARCODE (" + PosPlu + ")errato"
            //         bOk = False
            //      End If
            //   End If

            //   If Not bOk Then
            //      'Non faccio nulla
            //   ElseIf Nul2Zero(PosPlu) = 0 Then
            //      Print #3, "Articolo " + m_rsVarPos!pos_art + " non inviato perchè BARCODE (" + PosPlu + ") non valido"
            //   ElseIf m_rsVarPos!pos_tip = "A" Then
            //      'Cancellato
            //      sRiga = PosFi1 + PosPlu + "-0" + Space(60)
            //      Print #1, sRiga
            //   ElseIf PosPrz <= 0 Then             'Controllo sul prezzo
            //      Print #3, "Articolo " + m_rsVarPos!pos_art + " non inviato perchè manca il PREZZO"
   
            //   ElseIf TipVar = "Z" Then            'Impianto
            //      'Record di aggiornamento prezzo
            //      sRiga = PosFi1 + PosPlu + PosRep + _
            //              PosCo1 + PosCo2 + PosSco + _
            //              PosIva + PosCo3 + PosUmi + _
            //              PosDec + PosDep + PosDes + _
            //              PosCo5 + PosCo6 + PosFi2 + _
            //              PosCo4 + PosPrz
            //      Print #1, sRiga
            sRig = "";
            if (_bolImpianto)
            {
                sRig += sPosFi1;
                sRig += sPosPlu;
                sRig += sPosRep;
                sRig += sPosCo1;
                sRig += sPosCo2;
                sRig += sPosSco;
                sRig += sPosIva;
                sRig += sPosCo3;
                sRig += sPosUmi;
                sRig += sPosDec;
                sRig += sPosDep;
                sRig += sPosDes;
                sRig += sPosCo5;
                sRig += sPosCo6;
                sRig += sPosFi2;
                sRig += sPosCo4;
                sRig += sPosPrz;
            }
            else
            //   Else

            //      'Per I cosa cancello
            //      sRiga = PosFi1 + PosPlu + "-0" + Space(60)
            //      Print #1, sRiga
            //      sRiga = "P  " + PosPlu + "-" + Space(61)
            //      Print #1, sRiga
            //      sRiga = "P01" + PosPlu + "-" + Space(61)
            //      Print #1, sRiga
            //      sRiga = "P02" + PosPlu + "-" + Space(61)
            //      Print #1, sRiga
            //      sRiga = "P08" + PosPlu + "-" + Space(61)           'Seck 000005 P08 per la circolarità
            //      Print #1, sRiga
            {
                sRig += sPosFi1 + sPosPlu + "-0" + new string(' ', 60);
                sRig += _clsDef.CRLF;

                //Print #1, sRiga
                sRig += "P  " + sPosPlu + "-" + new string(' ', 61);
                sRig += _clsDef.CRLF;

                //Print #1, sRiga
                sRig += "P01" + sPosPlu + "-" + new string(' ', 61);
                sRig += _clsDef.CRLF;
                //Print #1, sRiga
                sRig += "P02" + sPosPlu + "-" + new string(' ', 61);
                sRig += _clsDef.CRLF;
                //Print #1, sRiga
                sRig += "P08" + sPosPlu + "-" + new string(' ', 61);   //'Seck 000005 P08 per la circolarità
                sRig += _clsDef.CRLF;
                //Print #1, sRiga

                //      'Per II cosa scrivo la variazione
                //      sRiga = PosFi1 + PosPlu + PosRep + _
                //              PosCo1 + PosCo2 + PosSco + _
                //              PosIva + PosCo3 + PosUmi + _
                //              PosDec + PosDep + PosDes + _
                //              PosCo5 + PosCo6 + PosFi2 + _
                //              PosCo4 + PosPrz
                //      Print #1, sRiga

                sRig += sPosFi1;
                sRig += sPosPlu;
                sRig += sPosRep;
                sRig += sPosCo1;
                sRig += sPosCo2;
                sRig += sPosSco;
                sRig += sPosIva;
                sRig += sPosCo3;
                sRig += sPosUmi;
                sRig += sPosDec;
                sRig += sPosDep;
                sRig += sPosDes;
                sRig += sPosCo5;
                sRig += sPosCo6;
                sRig += sPosFi2;
                sRig += sPosCo4;
                sRig += sPosPrz;
                sRig += _clsDef.CRLF;

            }

            //      'Per III cosa vedo se articolo in offerta
         
            //      'PROMOZIONI - Offerta normale
            //      '
            //      sOfTyp = Nul2Space(m_rsVarPos!pos_ovt)
            sOfTyp = (string)rowPos["pos_tva"];
            //      cOfVal = Nul2Zero(m_rsVarPos!pos_ova)
      
            //      PosMix = "000"
            sPosMix = "000";
      
            //      If Not IsVuoto(sOfTyp) And Not IsVuoto(cOfVal) Then
      
            //         If sOfTyp = "X" Then
         
            //            'Offerta MIX-MATCH

            if (sPosPlu == "8718951001572")
                Console.WriteLine("aaaaaaaa");

            if ((string)rowPos["pos_tva"] == _clsDef.OFAMXN)
            {
                //            PosMix = Nul2Space(m_rsVarPos!pos_mix)
                sPosMix = Convert.ToString(rowPos["pos_mix"]);
                if (sPosMix.Length > 3)
                    sPosMix.Substring(sPosMix.Length - 3, 3);
                else if (sPosMix.Length < 3)
                    sPosMix = sPosMix.PadLeft(3, Convert.ToChar("0"));

                //            'Controllo che MxN sia valorizzato correttamente
                //            If cOfVal < 11 Or cOfVal > 99 Then
                //               Print #3, "Articolo " + m_rsVarPos!pos_art + " con offerta non inviato perchè MxN non valido: " + Str(cOfVal)
                //            ElseIf Nul2Zero(PosMix) = 0 Then
                //               Print #3, "Articolo " + m_rsVarPos!pos_art + " con offerta non inviato perchè Mix non valido: " + PosMix
                //            Else
                //               sRiga = Ncr745xMxN(FillChar(Nul2Zero(PosMix), 3, " ", "D"), PosRep, PosIva, Trim(Str(cOfVal)), PosPrz)

                sRig += Ncr745xMxN(sPosMix, sPosRep, sPosIva, Convert.ToString(rowPos["pos_xem"]),  Convert.ToString(rowPos["pos_xen"]), (decimal)rowPos["pos_prv"]);
                sRig += _clsDef.CRLF;

                //               Print #1, sRiga
                //            End If
                //         End If
            }
            //         If sOfTyp <> "P" Then
            else if ((string)rowPos["pos_tva"] == _clsDef.OFAPRZ)
            {
                //            PosBol = "000"
                cPosBol = 0;
                //            sRiga = Ncr745xFidy("P  ", _
                //                                PosPlu, _
                //                                sOfTyp, _
                //                                cOfVal, _
                //                                PosBol, _
                //                                PosPrz, _
                //                                PosMix)

                sRig += Ncr745xFidy("P  ", sPosPlu, sOfTyp, cOfVal, cPosBol, (decimal)rowPos["pos_prv"], sPosMix);
                sRig += _clsDef.CRLF;

                //            Print #1, sRiga
                //         End If
                //      End If
            }
            //      'PROMOZIONI - Fidelity A
            //      '
            //      PosPrz = FillChar(Multi(Nul2Zero(m_rsVarPos!pos_prz), 100), 8, "0", "D")

            sPosPrz = Convert.ToInt32( (decimal)rowPos["pos_prv"] * 100).ToString("00000000");
      
            //      sOfTyp = Nul2Space(m_rsVarPos!pos_fta)
            sOfTyp = (string)rowPos["pos_tva"];

            //      cOfVal = Nul2Zero(m_rsVarPos!pos_fva)
            cOfVal = (decimal)rowPos["pos_prv"];

            //      PosBol = Nul2Zero(m_rsVarPos!pos_fpa)
            cPosBol = 0;

            //      If (Not IsVuoto(sOfTyp) And Not IsVuoto(cOfVal)) Or Not IsVuoto(PosBol) Then
            if (cPosBol > 0)
            {
                //         sRiga = Ncr745xFidy("P01", _
                //                             PosPlu, _
                //                             sOfTyp, _
                //                             cOfVal, _
                //                             PosBol, _
                //                             PosPrz, _
                //                             PosMix)
                sRig = Ncr745xFidy("P  ", sPosPlu, sOfTyp, cOfVal, cPosBol, (decimal)rowPos["pos_prv"], sPosMix);
                sRig += _clsDef.CRLF;

                //         Print #1, sRiga

                //         Print #1, "P08" + Mid(sRiga, 4)                'Seck 000005 00001 P08 per la circolarità

                //      End If
            }
            //      'PROMOZIONI - Fidelity B
            //      '
            //      sOfTyp = Nul2Space(m_rsVarPos!pos_ftb)
            //      cOfVal = Nul2Zero(m_rsVarPos!pos_fvb)
            //      PosBol = Nul2Zero(m_rsVarPos!pos_fpb)
      
            //      If Not IsVuoto(sOfTyp) And (Not IsVuoto(cOfVal) Or Not IsVuoto(PosBol)) Then
         
            //         sRiga = Ncr745xFidy("P02", _
            //                             PosPlu, _
            //                             sOfTyp, _
            //                             cOfVal, _
            //                             PosBol, _
            //                             PosPrz, _
            //                             PosMix)
            //         Print #1, sRiga
            //      End If
            //   End If
        
            return sRig;
        }

        //Private Function Ncr745xMxN(sMix, sRep, sIva, sMxN, cPrz)
        private string Ncr745xMxN(string sMix, string sRep, string sIva, string sXem, string sXen, decimal cPrz)
        {
            //Dim sRiga As String
            string sRig = "";
            //Dim sKey As String
            string sKey = "";
            //Dim sTip As String
            string sTip = "";
            //Dim sSco As String
            string sSco = "";
            //Dim sFlg As String
            string sFlg = "";
            //Dim sDes As String
            string sDes = "";
            //Dim sQta As String
            string sQta = "";
            //Dim sQtx As String
            string sQtx = "";
            //Dim sVal As String
            string sVal = "";
            //Dim iMmm As Integer
            int iMmm = 0;
            //Dim iNnn As Integer
            int iNnn = 0;

            //iMmm = Nul2Zero(Mid(sMxN, 1, 1))
            //iNnn = Nul2Zero(Mid(sMxN, 2, 1))
            iMmm = Convert.ToInt16(sXem);
            iNnn = Convert.ToInt16(sXen);

            //   sKey = "MM" + Space(11)
            sKey = "MM" + new string(' ', 11);

            //   sTip = "0"
            sTip = "0";

            //   sSco = "0"
            sSco = "0";

            //   sFlg = "0"
            sFlg = "0";

            //   sDes = "PROMOZIONE " + FillChar(iMmm, 1, "0", "D") + "x" + FillChar(iNnn, 1, "0", "D") + Space(2)
            sDes = "PROMOZIONE " + sXem + "x" + sXen + new string(' ', 50);
            sDes = sDes.Substring(0, 16);

            //   sQta = FillChar(iMmm, 2, "0", "D")
            sQta = iMmm.ToString("00");

            //   sQtx = Repl("00", 14)
            sQtx = new string('0', 28);

            //   sVal = FillChar(Multi(Sottrai(iMmm, iNnn), cPrz), 8, "0", "D")
            sVal = ((iMmm - iNnn) * cPrz).ToString("00000000");

            //   sRiga = sKey + sMix + sRep + sTip + sSco + sFlg + sIva + sDes + _
            //           sQta + sQtx + sVal

            sRig = sKey + sMix + sRep + sTip + sSco + sFlg + sIva + sDes + sQta + sQtx + sVal;

            //   Ncr745xMxN = sRiga
            return sRig;
            //End Function
        }

        //Private Function Ncr745xFidy(sCat, sEan, sTip, cSco, cBol, cPrz, sMix)
        private string Ncr745xFidy(string sCat, string sEan, string sTip, decimal cSco, decimal cBol, decimal cPrz, string sMix)
        {

            String sRig = "";
            //Dim sCo1 As String      'Codice 1
            string sCo1 = "";
            //Dim sCo2 As String      'Tipo sconto
            string sCo2 = "";
            //Dim sSco As String      'Valore sconto
            string sSco = "";
            //Dim sBol As String      'Bollini
            string sBol = "0";
            //Dim sSet As String      'Numero molteplicità riferimento tab MxN
            string sSet = "";
            //Dim sInd As String      'Indice tabella bollini
            string sInd = "";
            //Dim sPre As String      'Prezzo premio
            string sPre = "";
            //Dim sDpt As String      'Riferimento DPT
            string sDpt = "";
            //Dim sTot As String      'Indice totalizzatore secondario
            string sTot = "";

            //   sCo1 = "0"
            sCo1 = "0";

            //   If sTip = "X" Then
            //      sCo2 = "0"
            //      sSco = "000000"

            if(sTip == _clsDef.OFAMXN)
            {
                sCo2 = "0";
                sSco = "000000";
            }

            //   ElseIf sTip = "P" Then
            //      If cBol > 0 Then
            //         sCo2 = "5"
            //      Else
            //         sCo2 = "1"
            //      End If
            //      cSco = Multi(cSco, 100)
            //      sSco = FillChar(Sottrai(cPrz, cSco), 6, "0", "D")
            else if (sTip == _clsDef.OFAPRZ)
            {
                sCo2 = "0";
                sSco = "000000";

                sCo2 = "1";
                if (cBol > 0)
                    sCo2 = "5";

                cSco = cSco * 100;
                sSco = ((cPrz * 100 ) - cSco).ToString("000000");
            }
            //   ElseIf sTip = "S" Then
            //      If cBol > 0 Then
            //         sCo2 = "5"
            //      Else
            //         sCo2 = "1"
            //      End If
            //      sSco = FillChar(Multi(cSco, 100), 6, "0", "D")
            else if(sTip == _clsDef.OFASCV)
            {
                sCo2 = "1";
                if (cBol > 0)
                    sCo2 = "5";
                sSco = (cPrz - cSco).ToString("000000");
            }
            //   ElseIf sTip = "%" Then
            //      If cBol > 0 Then
            //         sCo2 = "6"
            //      Else
            //         sCo2 = "2"
            //      End If
            //      sSco = FillChar(Multi(cSco, 10), 6, "0", "D")
            else if (sTip == _clsDef.OFASCO)
            {
                sCo2 = "2";
                if (cBol > 0)
                    sCo2 = "6";
                sSco = (cSco * 10).ToString("000000");
            }
            //   ElseIf cBol > 0 Then
            //      sCo2 = "4"
            //      sSco = FillChar(Multi(cSco, 10), 6, "0", "D")
            //   End If
            else
            {
                sCo2 = "4";
                sSco = (cSco * 10).ToString("000000");
            }

            //   sBol = FillChar(cBol, 3, "0", "D")
            sBol = cBol.ToString("000");

            //   If Nul2Zero(sMix) > 0 Then
            //      sSet = "01"
            //   Else
            //      sSet = "00"
            //   End If

            if (Convert.ToInt16(sMix) > 0)
                sSet = "01";
            else
                sSet = "00";

            //   sInd = "00"
            //   sPre = "0000000000"
            //   sDpt = "0"
            //   sTot = "0"

               sInd = "00";
               sPre = "0000000000";
               sDpt = "0";
               sTot = "0";

            //   sRiga = sCat + sEan + sCo1 + _
            //           sCo2 + sSco + sBol + _
            //           sMix + sSet + sInd + _
            //           sPre + sDpt + sTot + Space(32)

            sRig = sCat + sEan + sCo1 + sCo2 + sSco + sBol + sMix + sSet + sInd + sPre + sDpt + sTot;
            sRig = sRig + new string(' ', 32);
            
            //   Ncr745xFidy = sRiga

            //End Function
            return sRig;
        }

    }
}
