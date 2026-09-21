using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace APOffice
{
    class clsPosUgaSid20
    {
        clsDefine _clsDef = new clsDefine();

        public Boolean _bolImpianto = false;

        public clsPosUgaSid20()
        { }

        //'
        //'  casse UNIGALAXY ( SIMENS / NIXDORF )
        //'
        //'Tracciato record per MANEXT.SID
        //Private Sub VarUgaSid20(TipVar)
        public string PosArtRow(DataRow rowPos)
        {
            string s = "";
            //Dim sRiga As String
            string sRig = "";
            //Dim bOk As Boolean
            Boolean bOk = true;
            //Dim bFlgErr As Boolean
            Boolean bFlgErr = true;
            //Dim Err_______ As String
            string Err_______ = "";
            //Dim Cod_fil___ As String
            string Cod_fil___ = "";
            //Dim Data_var__ As String
            string Data_var__ = "";
            //Dim Ora_var___ As String
            string Ora_var___ = "";
            //Dim Tipo_rec__ As String
            string Tipo_rec__ = "";
            //Dim Tipo_var__ As String
            string Tipo_var__ = "";
            //Dim Cod_art___ As String
            string Cod_art___ = "";
            //Dim Cod_art_s_ As String
            string Cod_art_s_ = "";
            //Dim Cod_bar___ As String
            string Cod_bar___ = "";
            //Dim Cod_rio___ As String
            string Cod_rio___ = "";
            //Dim Tipo_bila_ As String
            string Tipo_bila_ = "";
            //Dim Descr_____ As String
            string Descr_____ = "";
            //Dim Descr_cas_ As String
            string Descr_cas_ = "";
            //Dim Cod_rep___ As String
            string Cod_rep___ = "";
            //Dim Cod_rep_ex As String
            string Cod_rep_ex = "";
            //Dim Cod_fam___ As String
            string Cod_fam___ = "";
            //Dim Cod_gam___ As String
            string Cod_gam___ = "";
            //Dim Cod_vuoto_ As String
            string Cod_vuoto_ = "";
            //Dim Cod_iva___ As String
            string Cod_iva___ = "";
            //Dim Cod_iva_ex As String
            string Cod_iva_ex = "";
            //Dim Prz_costo_ As String
            string Prz_costo_ = "";
            //Dim Prz_vend__ As String
            string Prz_vend__ = "";
            //Dim Un_misura_ As String
            string Un_misura_ = "";
            //Dim Un_vendita As String
            string Un_vendita = "  ";
            //Dim Peso______ As String
            string Peso______ = "";
            //Dim Tara______ As String
            string Tara______ = "";
            //Dim Corsia____ As String
            string Corsia____ = "";
            //Dim Scafale___ As String
            string Scafale___ = "";
            //Dim Fila______ As String
            string Fila______ = "";
            //Dim File_scaf_ As String
            string File_scaf_ = "";
            //Dim Sc_minima_ As String
            string Sc_minima_ = "";
            //Dim Qta_rio___ As String
            string Qta_rio___ = "";
            //Dim Qta_x_cart As String
            string Qta_x_cart = "";
            //Dim Fornitore_ As String
            string Fornitore_ = "";
            //Dim Prz_rich__ As String
            string Prz_rich__ = "";
            //Dim Art_n_rend As String
            string Art_n_rend = "";
            //Dim Art_n_sc__ As String
            string Art_n_sc__ = "";
            //Dim Qta_n_per_ As String
            string Qta_n_per_ = "";
            //Dim Over_n_per As String
            string Over_n_per = "";
            //Dim No_boll___ As String
            string No_boll___ = "";
            //Dim Sc_man_lim As String
            string Sc_man_lim = "";
            //Dim Per_max_sc As String
            string Per_max_sc = "";
            //Dim Art_ris___ As String
            string Art_ris___ = "";
            //Dim Tipo_off__ As String
            string Tipo_off__ = "";
            //Dim Codice_off As String
            string Codice_off = "";
            //Dim Paniere___ As String
            string Paniere___ = "";
            //Dim Codice_mix As String
            string Codice_mix = "";
            //Dim Sconto_per As String
            string Sconto_per = "";
            //Dim Sconto_val As String
            string Sconto_val = "";
            //Dim NxM_______ As String
            string NxM_______ = "";
            //Dim Qta_sc_1__ As String
            string Qta_sc_1__ = "";
            //Dim Per_sc_1__ As String
            string Per_sc_1__ = "";
            //Dim Val_sc_1__ As String
            string Val_sc_1__ = "";
            //Dim Ricalc_1__ As String
            string Ricalc_1__ = "";
            //Dim Qta_sc_2__ As String
            string Qta_sc_2__ = "";
            //Dim Per_sc_2__ As String
            string Per_sc_2__ = "";
            //Dim Val_sc_2__ As String
            string Val_sc_2__ = "";
            //Dim Ricalc_2__ As String
            string Ricalc_2__ = "";
            //Dim Qta_sc_3__ As String
            string Qta_sc_3__ = "";
            //Dim Max_qta_a_ As String
            string Max_qta_a_ = "";
            //Dim Num_boll__ As String
            string Num_boll__ = "";
            //Dim Num_art_b_ As String
            string Num_art_b_ = "";
            //Dim Off_boll__ As String
            string Off_boll__ = "";
            //Dim Pnt_prem1_ As String
            string Pnt_prem1_ = "";
            //Dim Imp_prem1_ As String
            string Imp_prem1_ = "";
            //Dim Pnt_prem2_ As String
            string Pnt_prem2_ = "";
            //Dim Imp_prem2_ As String
            string Imp_prem2_ = "";

            //   bOk = True
            bOk = true;

            //   Err_______ = Repl(" ", 3)
            Err_______ = new string(' ', 3);

            //   Cod_fil___ = " 0"                     'Repl(" ", 2)
            Cod_fil___ = " 0";

            //   Data_var__ = Format(m_rsVarPos!pos_day, "yyyymmdd")
            Data_var__ = ((DateTime)rowPos["pos_day"]).ToString("yyyymmdd");

            //   Ora_var___ = Repl(" ", 5)
            Ora_var___ = "00:00";

            //   Tipo_rec__ = "A"
            Tipo_rec__ = "A";

            //   If m_rsVarPos!pos_tip = "A" Then
            //      Tipo_var__ = "R"                 'CANCELLAZIONE
            //   Else
            //      Tipo_var__ = Repl(" ", 1)
            //   End If
            if (false)          //(string)rowPos["pos_   "] == "A")
                Tipo_var__ = "R";                 //'CANCELLAZIONE
            else
                Tipo_var__ = " ";

            //   Cod_art___ = Repl(" ", 20)
            Cod_art___ = new string(' ', 20);

            //   Cod_art_s_ = Repl(" ", 20)
            Cod_art_s_ = new string(' ', 20);

            //   Cod_rio___ = Repl(" ", 6)
            Cod_rio___ = new string(' ', 6);

            //   If Len(Trim(m_rsVarPos("pos_ean"))) = 13 And Left(m_rsVarPos("pos_ean"), 1) = "2" Then

            //      'Articoli con barcode a peso
            //      Tipo_bila_ = "L"
            //      Cod_bar___ = FillChar(Left(m_rsVarPos("pos_ean"), 12), 20, "0", "S")

            //   ElseIf Len(Trim(m_rsVarPos("pos_ean"))) <= 4 Then

            //      Tipo_bila_ = " "

            //      Cod_bar___ = FillChar(Trim(m_rsVarPos("pos_ean")), 20, "0", "D")

            //   Else

            //      If False And Nul2Space(m_rsVarPos!pos_bil) = "S" Then
            //         'Articoli pesati alle casse con ceck-out
            //         Tipo_bila_ = "S"
            //         Cod_bar___ = Trim(m_rsVarPos("pos_ean"))
            //      Else
            //         Tipo_bila_ = " "
            //         Cod_bar___ = FillChar(Trim(m_rsVarPos("pos_ean")), 13, "0", "D")
            //      End If

            //      'Eseguo il controllo check digit altrimenti se digit errato si bloccano le casse
            //      If Nul2Zero(Cod_bar___) > 999999 Then
            //         If Cod_bar___ <> CtrlDigit(Left(Cod_bar___, 12)) Then
            //            Print #3, "Articolo " + m_rsVarPos!pos_art + " non inviato perchè BARCODE (" + Cod_bar___ + ")errato"
            //            bOk = False
            //         End If
            //      End If

            //      Cod_bar___ = FillChar(Trim(m_rsVarPos("pos_ean")), 20, "0", "D")

            //   End If

            //Tipo_bila_ = "";

            if ((string)rowPos["pos_art"] == "0000175")
                Console.WriteLine("aaaaa");

            Cod_bar___ = CodEan((string)rowPos["pos_ean"], ref Tipo_bila_, (Boolean)rowPos["bil_bil"]);

            //   Descr_____ = Repl(" ", 40)
            Descr_____ = (((string)rowPos["pos_ard"]) + new string(' ', 40)).Substring(0, 40);

            //   Descr_cas_ = Left(Nul2Space(m_rsVarPos!pos_des + Space(21)), 21)
            Descr_cas_ = (((string)rowPos["pos_arb"]) + new string(' ', 21)).Substring(0,21);

            //   Cod_rep___ = FillChar(Nul2Zero(m_rsVarPos!pos_rev), 3, " ", "D")
            Cod_rep___ = ((string)rowPos["pos_rep"]).Trim().PadLeft(3, Convert.ToChar("0")); 

            //   Cod_rep_ex = Repl(" ", 5)
            Cod_rep_ex = new string(' ', 5);

            //   Cod_fam___ = Repl(" ", 4)
            Cod_fam___ = new string(' ', 4);

            //   Cod_gam___ = Repl(" ", 4)
            Cod_gam___ = new string(' ', 4);

            //   Cod_vuoto_ = " 0"
            Cod_vuoto_ = " 0";

            //   If Val(Nul2Zero(m_rsVarPos!pos_iva)) < 50 Then
            //      Cod_iva___ = FillChar(Nul2Zero(m_rsVarPos!pos_iva), 2, "0", "D")
            //   Else
            //      Cod_iva___ = Space(2)
            //   End If
            if (Convert.ToInt16(rowPos["pos_iva"]) < 50)
                Cod_iva___ = ((string)rowPos["pos_iva"]).Substring(1, 2);
            else
                Cod_iva___ = "  ";

            //   Cod_iva_ex = Repl(" ", 5)
            Cod_iva_ex = new string(' ', 5);

            //   Prz_costo_ = Repl(" ", 7) + "0"
            Prz_costo_ = new string(' ', 7) + "0";

            //   Prz_vend__ = FillChar(Multi(Nul2Zero(m_rsVarPos!pos_prz), 100), 8, " ", "D")
            Prz_vend__ = Convert.ToInt32((decimal)rowPos["pos_prv"] * 100).ToString().PadLeft(8, Convert.ToChar(" "));

            //   Un_misura_ = Repl(" ", 2)
            Un_misura_ = ((string)rowPos["pos_umi"] + "  ").Substring(0, 2);

            //   Un_vendita = Repl(" ", 2)
            Un_vendita = new string(' ', 2);

            //   Peso______ = Repl(" ", 7)
            Peso______ = "  0.000";
            //   Tara______ = Repl(" ", 5)
            Tara______ = new string(' ', 4) + "0";
            //   Corsia____ = Repl(" ", 10)
            Corsia____ = new string(' ', 10);
            //   Scafale___ = Repl(" ", 10)
            Scafale___ = new string(' ', 10);
            //   Fila______ = Repl(" ", 10)
            Fila______ = new string(' ', 10);
            //   File_scaf_ = Repl(" ", 3)
            File_scaf_ = new string(' ', 2) + "0";
            //   Sc_minima_ = Repl(" ", 5)
            Sc_minima_ = new string(' ', 4) + "0";
            //   Qta_rio___ = Repl(" ", 5)
            Qta_rio___ = new string(' ', 4) + "0";
            //   Qta_x_cart = Repl(" ", 5)
            Qta_x_cart = new string(' ', 4) + "0";
            //   Fornitore_ = Repl(" ", 6)
            Fornitore_ = new string(' ', 6);
            //   Prz_rich__ = Repl(" ", 1)
            Prz_rich__ = new string(' ', 1);
            //   Art_n_rend = Repl(" ", 1)
            Art_n_rend = new string(' ', 1);
            //   Art_n_sc__ = Repl(" ", 1)
            Art_n_sc__ = new string(' ', 1);
            //   Qta_n_per_ = Repl(" ", 1)
            Qta_n_per_ = new string(' ', 1);
            //   Over_n_per = Repl(" ", 1)
            Over_n_per = new string(' ', 1);
            //   No_boll___ = Repl(" ", 1)
            No_boll___ = new string(' ', 1);
            //   Sc_man_lim = Repl(" ", 1)
            Sc_man_lim = new string(' ', 1);
            //   Per_max_sc = Repl(" ", 6)
            Per_max_sc = "  0.00";
            //   Art_ris___ = Repl(" ", 1)
            Art_ris___ = new string(' ', 1);

            //   Tipo_off__ = Repl(" ", 3)           'Tipo Sconto
            Tipo_off__ = new string(' ', 3);
            //   Codice_off = Repl(" ", 3)
            Codice_off = "  0";
            //   Paniere___ = Repl(" ", 2)
            Paniere___ = " 0";
            //   Codice_mix = Repl(" ", 5)
            Codice_mix = "    0";

            //   Sconto_per = Repl(" ", 6)           'Sconto %
            Sconto_per ="  0.00";

            //   Sconto_val = Repl(" ", 8)
            Sconto_val = new string(' ', 7) + "0";
            //   NxM_______ = Repl(" ", 6)
            NxM_______ = new string(' ', 6);
            //   Qta_sc_1__ = Repl(" ", 3)
            Qta_sc_1__ = "  0";
            //   Per_sc_1__ = Repl(" ", 9)
            Per_sc_1__ = new string(' ', 5) + "0.00";
            //   Val_sc_1__ = Repl(" ", 9)
            Val_sc_1__ = new string(' ', 8) + "0";
            //   Ricalc_1__ = Repl(" ", 1)
            Ricalc_1__ = new string(' ', 1);
            //   Qta_sc_2__ = Repl(" ", 5)
            Qta_sc_2__ = new string(' ', 4) + "0";
            //   Per_sc_2__ = Repl(" ", 9)
            Per_sc_2__ = new string(' ', 5) + "0.00";
            //   Val_sc_2__ = Repl(" ", 9)
            Val_sc_2__ = new string(' ', 8) + "0"; ;
            //   Ricalc_2__ = Repl(" ", 1)
            Ricalc_2__ = new string(' ', 1);
            //   Qta_sc_3__ = Repl(" ", 3)
            Qta_sc_3__ = new string(' ', 2) + "0";
            //   Max_qta_a_ = Repl(" ", 5)
            Max_qta_a_ = new string(' ', 4) + "0";
            //   Num_boll__ = Repl(" ", 5)
            Num_boll__ = " 0.00";
            //   Num_art_b_ = Repl(" ", 3)
            Num_art_b_ = new string(' ', 2) + "0";
            //   Off_boll__ = Repl(" ", 5)
            Off_boll__ = new string(' ', 4) + "0";
            //   Pnt_prem1_ = Repl(" ", 5)
            Pnt_prem1_ = new string(' ', 4) + "0";
            //   Imp_prem1_ = Repl(" ", 9)
            Imp_prem1_ = new string(' ', 8) + "0";
            //   Pnt_prem2_ = Repl(" ", 5)
            Pnt_prem2_ = new string(' ', 4) + "0";
            //   Imp_prem2_ = Repl(" ", 9)
            Imp_prem2_ = new string(' ', 8) + "0";

            ////   PosFts = Nul2Space(m_rsVarPos!pos_ovt)
            //PosFts = (string)rowPos["pos_tva"];
            ////   PosFvs = Nul2Zero(m_rsVarPos!pos_ova)
            //PosFvs = (decimal)rowPos["pos_prv"];

            ////   ProCat = Space(5)                      'Linea/categoria blank x tutti, 2 cat A, 3 cat B
            //ProCat = new string(' ', 5);

            ////   If Nul2Space(PosFts) = "P" And PosFvs > 0 Then         'Tipo "P" come variazione
            //if (PosFts == _clsDef.OFAPRZ)

            //   Prz_vend__ = FillChar(Multi(Nul2Zero(PosFvs), 100), 8, " ", "D")
            //Prz_vend__ = Convert.ToInt32((decimal)rowPos["pos_prv"] * 100).ToString().PadLeft(8, Convert.ToChar("0"));
            //   End If

            //   If Val(Nul2Space(Cod_bar___)) = 0 Then
            //      Print #3, "Articolo " + m_rsVarPos!pos_art + " non inviato perchè BARCODE non corretto"
            //      bOk = False
            //   ElseIf Nul2Zero(Prz_vend__) = 0 Then
            //      Print #3, "Articolo " + m_rsVarPos!pos_art + " non inviato perchè manca il PREZZO"
            //      bOk = False
            //   ElseIf Nul2Zero(Cod_rep___) = 0 Then
            //      Print #3, "Articolo " + m_rsVarPos!pos_art + " non inviato perchè manca il REPARTO"
            //      bOk = False
            //   End If

            //   If bOk Then

            //      Print #1, Err_______ + Cod_fil___ + Data_var__ + Ora_var___ + _
            //                Tipo_rec__ + Tipo_var__ + Cod_art___ + Cod_art_s_ + _
            //                Cod_bar___ + Cod_rio___ + Tipo_bila_ + Descr_____ + _
            //                Descr_cas_ + Cod_rep___ + Cod_rep_ex + Cod_fam___ + _
            //                Cod_gam___ + Cod_vuoto_ + Cod_iva___ + Cod_iva_ex + _
            //                Prz_costo_ + Prz_vend__ + Un_misura_ + Un_vendita + _
            //                Peso______ + Tara______ + Corsia____ + Scafale___ + _
            //                Fila______ + File_scaf_ + Sc_minima_ + Qta_rio___ + _
            //                Qta_x_cart + Fornitore_ + Prz_rich__ + Art_n_rend + _
            //                Art_n_sc__ + Qta_n_per_ + Over_n_per + No_boll___ + _
            //                Sc_man_lim + Per_max_sc + Art_ris___ + Tipo_off__ + _
            //                Codice_off + Paniere___ + Codice_mix + Sconto_per + _
            //                Sconto_val + NxM_______ + Qta_sc_1__ + Per_sc_1__ + _
            //                Val_sc_1__ + Ricalc_1__ + Qta_sc_2__ + Per_sc_2__ + _
            //                Val_sc_2__ + Ricalc_2__ + Qta_sc_3__ + Max_qta_a_ + _
            //                Num_boll__ + Num_art_b_ + Off_boll__ + Pnt_prem1_ + _
            //                Imp_prem1_ + Pnt_prem2_ + Imp_prem2_

            if (bOk)
            {

                sRig += Err_______ + Cod_fil___ + Data_var__ + Ora_var___ ;
                sRig += Tipo_rec__ + Tipo_var__ + Cod_art___ + Cod_art_s_ ;
                sRig += Cod_bar___ + Cod_rio___ + Tipo_bila_ + Descr_____ ;
                sRig += Descr_cas_ + Cod_rep___ + Cod_rep_ex + Cod_fam___ ;
                sRig += Cod_gam___ + Cod_vuoto_ + Cod_iva___ + Cod_iva_ex ;
                sRig += Prz_costo_ + Prz_vend__ + Un_misura_ + Un_vendita ;
                sRig += Peso______ + Tara______ + Corsia____ + Scafale___ ;
                sRig += Fila______ + File_scaf_ + Sc_minima_ + Qta_rio___ ;
                sRig += Qta_x_cart + Fornitore_ + Prz_rich__ + Art_n_rend ;
                sRig += Art_n_sc__ + Qta_n_per_ + Over_n_per + No_boll___ ;
                sRig += Sc_man_lim + Per_max_sc + Art_ris___ + Tipo_off__ ;
                sRig += Codice_off + Paniere___ + Codice_mix + Sconto_per ;
                sRig += Sconto_val + NxM_______ + Qta_sc_1__ + Per_sc_1__ ;
                sRig += Val_sc_1__ + Ricalc_1__ + Qta_sc_2__ + Per_sc_2__ ;
                sRig += Val_sc_2__ + Ricalc_2__ + Qta_sc_3__ + Max_qta_a_ ;
                sRig += Num_boll__ + Num_art_b_ + Off_boll__ + Pnt_prem1_;
                sRig += Imp_prem1_ + Pnt_prem2_ + Imp_prem2_;
                sRig += "        "; //         0           0";
                sRig += _clsDef.CRLF;

            }

            return sRig;
        }

        public string PosOffArtRow(DataRow rowPos)
        {
            string sRig = "";

            string Cod_art___ = "";
            string Cod_bar___ = "";
            string Tipo_bila_ = "";

            //Dim ProInz As String          'Prefisso
            string ProInz = "";
            //Dim ProTyp As String          'Tipo operazione
            string ProTyp = "";
            //Dim ProFi1 As String          'Filler
            string ProFi1 = "";
            //Dim ProTip As String          'Tipo articolo
            string ProTip = "";
            //Dim ProCod As String          'Codice iniziativa
            string ProCod = "";
            //Dim ProSco As String          'Aplicazione sconti
            string ProSco = "";
            //Dim ProCat As String          'Linea
            string ProCat = "";
            //Dim ProClu As String          'Cluster
            string ProClu = "";
            //Dim ProDay As String          'Data inizio + fine + ora inizio + fine
            string ProDay = "";
            //Dim ProSet As String          ''Giorni attivi sulla settimana
            string ProSet = "";
            //Dim ProTin As String          'Tipologia iniziativa
            string ProTin = "";
            //Dim ProLIn As String          'Limite inferiore
            string ProLIn = "";
            //Dim ProLSu As String          'Limite superiore
            string ProLSu = "";
            //Dim ProLPs As String          'Limite passo
            string ProLPs = "";
            //Dim ProLVa As String          'Limite valore
            string ProLVa = "";
            //Dim ProLim As String          'Limiti rimanaeti
            string ProLim = "";
            //Dim ProFlg As String          'Flag applicazione formula
            string ProFlg = "";
            //Dim ProPag As String          'Tipo pagamento
            string ProPag = "";

            //Dim PosFts As String       'Field tipo sconto
            string PosFts = "";
            //Dim PosFvs As String       'Field valore sconto
            string PosFvs = "";
            //Dim PosFfp As String       'Field Fidelity punti
            string PosFfp = "";
            //Dim PosCat As String       'Categoria di riferimento per fidelity
            string PosCat = "";
            //Dim PosPun As String       'Punti dell'articolo
            string PosPun = "";

            Cod_art___ = (string)rowPos["pos_art"];
            Cod_bar___ = CodEan((string)rowPos["pos_ean"], ref Tipo_bila_, (Boolean)rowPos["bil_bil"]);

            //      'Promozioni
            //      If TipVar <> "Z" Then            'Impianto
            if (!_bolImpianto)
            {
                //         ProInz = "INZ"                      'Prefisso
                ProInz = "INZ";
                //         ProTyp = "DT"                       'Tipo operazione
                ProTyp = "DT";
                //         ProFi1 = Space(10)                  'Filler
                ProFi1 = new string(' ', 10);
                //         ProTip = "ART"                      'Tipo articolo
                ProTip = "ART";
                //         ProCod = FillChar(1, 10, " ", "S")  'Codice iniziativa ( sempre 1 )
                ProCod = "1";
                //         ProSco = " "                        'Aplicazione sconti
                ProSco = " ";
                //         ProCat = Space(5)                   'Linea/categoria    1 x tutti, 2 cat A, 3 cat B
                ProCat = "     ";
                //         ProClu = Repl("0", 10)              'Cluster
                ProClu = new string('0', 10);
                //         ProDay = Repl(" ", 24)              'Data inizio + fine + ora inizio + fine
                ProDay = new string(' ', 24);
                //         ProSet = Repl("1", 7)               'Giorni attivi sulla settimana
                ProSet = new string('1', 7);
                //         ProTin = "  "                       'Tipologia iniziativa
                ProTin = "  ";
                //         ProLIn = Repl("0", 8)               'Limite inferiore
                ProLIn = new string('0', 8);
                //         ProLSu = Repl("9", 8)               'Limite superiore
                ProLSu = new string('9', 8);
                //         ProLPs = Repl("0", 7) + "1"         'Limite passo
                ProLPs = new string('0', 7) + "1";
                //         ProLVa = Repl("0", 8)               'Limite quantità valore
                ProLVa = new string('0', 8);
                //         ProLim = Repl(Repl("0", 32), 9)     'Limiti rimanenti ( altri 9 )
                ProLim = new string('0', 288);
                //         ProFlg = "0"                        'Flag applicazione formula
                ProFlg = "0";
                //         ProPag = "0"                        'Tipo pagamento
                ProPag = "0";

                //         'Cancellazione promozione
                //         Print #2, ProInz + ProTyp + ProFi1 + ProTip + _
                //                   Cod_bar___ + ProCod + ProSco + ProCat + _
                //                   ProClu + ProDay + ProSet + ProTin + ProLIn + _
                //                   ProLSu + ProLPs + ProLVa + ProLim + _
                //                   Cod_art___ + ProFlg + ProPag

                sRig += ProInz + ProTyp + ProFi1 + ProTip ;
                sRig += Cod_bar___ + ProCod + ProSco + ProCat ;
                sRig += ProClu + ProDay + ProSet + ProTin + ProLIn ;
                sRig += ProLSu + ProLPs + ProLVa + ProLim ;
                sRig += Cod_art___ + ProFlg + ProPag;
                sRig += _clsDef.CRLF;

                //         'X TUTTI: Field Tipo Sconto articolo ( Offerta ) + Valore sconto
                //         '
                //         PosFts = Nul2Space(m_rsVarPos!pos_ovt)
                PosFts = (string)rowPos["pos_tva"];
                //         PosFvs = Nul2Zero(m_rsVarPos!pos_ova)
                //PosFvs = (decimal)rowPos["pos_prv"];
                //         ProCat = Space(5)                      'Linea/categoria blank x tutti, 2 cat A, 3 cat B
                ProCat = new string(' ', 5);
                //         If Not IsVuoto(PosFts) And PosFts <> "P" Then         'Tipo "P" come variazione
                if ((string)rowPos["pos_tva"] != _clsDef.OFAPRZ)
                {
                    //            If PosFts = "X" Then
                    if ((string)rowPos["pos_tva"] == _clsDef.OFAMXN)
                    {
                        //               Cod_art___ = FillChar(Nul2Space(m_rsVarPos!pos_mix), 20, " ", "S")
                        Cod_art___ = ((string)rowPos["pos_tvd"] + new string(' ', 20)).Substring(0, 20);
                        //            End If

                        PosFvs = ((decimal)rowPos["pos_xem"]).ToString("0") + ((decimal)rowPos["pos_xen"]).ToString("0");
                    }

                    //            sRiga = UGalaPro("", PosFts, PosFvs, ProInz, ProTyp, ProFi1, ProTip, _
                    //                             Cod_bar___, ProCod, ProSco, ProCat, ProClu, _
                    //                             ProDay, ProSet, ProTin, ProLIn, ProLSu, ProLPs, _
                    //                             ProLVa, ProLim, Cod_art___, ProFlg, ProPag)

                    sRig = UGalaPro("", PosFts, PosFvs, ProInz, ProTyp, ProFi1, ProTip, 
                                        Cod_bar___, ProCod, ProSco, ProCat, ProClu, 
                                        ProDay, ProSet, ProTin, ProLIn, ProLSu, ProLPs, 
                                        ProLVa, ProLim, Cod_art___, ProFlg, ProPag);

                        //            If Not IsVuoto(sRiga) Then
                        //               Print #2, sRiga
                        //            End If
                }
                    //         End If
                    //         'X FIDELITY A: Field Tipo Sconto articolo ( Offerta ) + Valore sconto
                    //         '
                    //         ProCat = FillChar(1, 5, " ", "S")      'Linea/categoria    1 x tutti, 2 cat A, 3 cat B
                    //         PosFts = Nul2Space(m_rsVarPos!pos_fta)
                    //         PosFvs = Nul2Zero(m_rsVarPos!pos_fva)
                    //         PosPun = FillChar(Nul2Zero(m_rsVarPos!pos_fpa), 8, "0", "D")

                    //         If Not IsVuoto(PosFts) Then
                if(PosFts != "")
                {

                    //            sRiga = UGalaPro("", PosFts, PosFvs, ProInz, ProTyp, ProFi1, ProTip, _
                    //                             Cod_bar___, ProCod, ProSco, ProCat, ProClu, _
                    //                             ProDay, ProSet, ProTin, ProLIn, ProLSu, ProLPs, _
                    //                             ProLVa, ProLim, Cod_art___, ProFlg, ProPag)

                    //            If Not IsVuoto(sRiga) Then
                    //               Print #2, sRiga
                    //            End If

                    //         End If

                    //         If Val(PosPun) > 0 Then
                    //            sRiga = UGalaPro(PosPun, PosFts, PosFvs, ProInz, ProTyp, ProFi1, ProTip, _
                    //                             Cod_bar___, ProCod, ProSco, ProCat, ProClu, _
                    //                             ProDay, ProSet, ProTin, ProLIn, ProLSu, ProLPs, _
                    //                             ProLVa, ProLim, Cod_art___, ProFlg, ProPag)
                    //            If Not IsVuoto(sRiga) Then
                    //               Print #2, sRiga
                    //            End If
                    //         End If

                    //         'X FIDELITY B: Field Tipo Sconto articolo ( Offerta ) + Valore sconto
                    //         '
                    //         ProCat = FillChar(2, 5, " ", "S")      'Linea/categoria    1 x tutti, 2 cat A, 3 cat B
                    //         PosFts = Nul2Space(m_rsVarPos!pos_ftb)
                    //         PosFvs = Nul2Zero(m_rsVarPos!pos_fvb)
                    //         PosPun = FillChar(Nul2Zero(m_rsVarPos!pos_fpb), 8, "0", "D")

                    //         If Not IsVuoto(PosFts) Then

                    //            sRiga = UGalaPro("", PosFts, PosFvs, ProInz, ProTyp, ProFi1, ProTip, _
                    //                             Cod_bar___, ProCod, ProSco, ProCat, ProClu, _
                    //                             ProDay, ProSet, ProTin, ProLIn, ProLSu, ProLPs, _
                    //                             ProLVa, ProLim, Cod_art___, ProFlg, ProPag)

                    //            If Not IsVuoto(sRiga) Then
                    //               Print #2, sRiga
                    //            End If

                    //         End If

                    //         If Val(PosPun) > 0 Then
                    //            sRiga = UGalaPro(PosPun, PosFts, PosFvs, ProInz, ProTyp, ProFi1, ProTip, _
                    //                             Cod_bar___, ProCod, ProSco, ProCat, ProClu, _
                    //                             ProDay, ProSet, ProTin, ProLIn, ProLSu, ProLPs, _
                    //                             ProLVa, ProLim, Cod_art___, ProFlg, ProPag)
                    //            If Not IsVuoto(sRiga) Then
                    //               Print #2, sRiga
                    //            End If
                    //         End If
                }
                //      End If
            }
            //   End If

            return sRig;

            //End Sub
        }

        //'Genera il record Ugalaxy Promozioni
        //Private Function UGalaPro(PosPun, PosFts, PosFvs, ProInz, ProTyp, ProFi1, ProTip, _
        //                          Cod_bar___, ProCod, ProSco, ProCat, ProClu, _
        //                          ProDay, ProSet, ProTin, ProLIn, ProLSu, ProLPs, _
        //                          ProLVa, ProLim, Cod_art___, ProFlg, ProPag)
        private string UGalaPro(string PosPun, string PosFts, string PosFvs, string ProInz, string ProTyp, string ProFi1,
                                string ProTip, string Cod_bar___, string ProCod, string ProSco, string ProCat, string ProClu, 
                                string ProDay, string ProSet, string ProTin, string ProLIn, string ProLSu, string ProLPs, 
                                string ProLVa, string ProLim, string Cod_art___, string ProFlg, string ProPag)
        {
            //Dim sRiga As String
            string sRig = "";
            //Dim bOk As Boolean
            Boolean bOk = false;

            //   ProTyp = "MO"                       'Tipo operazione modifica
            ProTyp = "MO";
            //   If Val(PosPun) > 0 Then             'Punti
            //      ProTin = "PQ"
            //      bOk = True
            //      ProLVa = FillChar(Multi(PosPun, 100), 8, "0", "D")
            //   ElseIf PosFts = "P" Then            'Taglio prezzo
            //      ProLVa = FillChar(Multi(PosFvs, 10000), 8, "0", "D")
            //      ProTin = "LQ"
            //      bOk = True
            //   ElseIf PosFts = "S" Then            'Sconto su ammontare
            //      ProLVa = FillChar(Multi(PosFvs, 10000), 8, "0", "D")
            //      ProTin = "BV"
            //'      ProLVa = Multi(ProLVa, 100)
            //      bOk = True
            //   ElseIf PosFts = "%" Then            'Sconto in %
            //      ProLVa = FillChar(Multi(PosFvs, 100), 8, "0", "D")
            //      ProTin = "SV"
            //      bOk = True
            //   ElseIf PosFts = "X" Then            'Sconto M x N
            //      ProLPs = FillChar(Mid(Trim(Str(PosFvs)), 1, 1), 8, "0", "D")
            //      ProLVa = FillChar(Multi(Sottrai(Val(Mid(Trim(Str(PosFvs)), 1, 1)), Val(Mid(Trim(Str(PosFvs)), 2, 1))), 100), 8, "0", "D")
            //      ProTin = "AQ"
            //      bOk = True
            //   End If

            if(Convert.ToInt16(PosPun) > 0)
            {
                bOk = true;
                ProTin = "PQ";
                ProLVa = Convert.ToInt16(PosPun).ToString("00000000");   
            }
            else if (PosFts == "P")
            {
                bOk = true;
                ProTin = "LQ";
                ProLVa = Convert.ToInt16(PosFvs).ToString("00000000");   
            }
            else if (PosFts == "S")
            {
                bOk = true;
                ProLVa = Convert.ToInt16(PosFvs).ToString("00000000");   
                ProTin = "BV";
            }
            else if (PosFts == "%")
            {
                bOk = true;
                ProLVa = Convert.ToInt16(PosFvs).ToString("00000000");
                ProTin = "SV";
            }
            else if (PosFts == "X")
            {
                bOk = true;
                int iXem = Convert.ToInt16(PosFvs.Substring(0,1));
                int iXen = Convert.ToInt16(PosFvs.Substring(1,1));
                ProLPs = iXem.ToString("00000000");
                ProLVa = ((iXem - iXen) * 100).ToString("00000000");
                ProTin = "AQ";
            }

            //   If bOk Then
            //      sRiga = ProInz + ProTyp + ProFi1 + ProTip + _
            //              Cod_bar___ + ProCod + ProSco + ProCat + _
            //              ProClu + ProDay + ProSet + ProTin + ProLIn + _
            //              ProLSu + ProLPs + ProLVa + ProLim + _
            //              Cod_art___ + ProFlg + ProPag
            //   End If
         
            if(bOk)
            {
                sRig = ProInz + ProTyp + ProFi1 + ProTip;
                sRig += Cod_bar___ + ProCod + ProSco + ProCat;
                sRig += ProClu + ProDay + ProSet + ProTin + ProLIn;
                sRig += ProLSu + ProLPs + ProLVa + ProLim;
                sRig += Cod_art___ + ProFlg + ProPag;
            }

            //   UGalaPro = sRiga
            return sRig;
            //End Function
        }

        private string CodEan(string strEan, ref string Tipo_bila_, Boolean bolBil)
        {
            string sEan = "";

            string s = strEan.Trim();
            if (s.Substring(0, 1) == "2" && s.Length == 13 && s.Substring(7, 6) == "000000")
            {
                Tipo_bila_ = "L";
                sEan = s.PadLeft(13, Convert.ToChar("0")).Substring(0, 12); // FillChar(Left(m_rsVarPos("pos_ean"), 12), 20, "0", "S")
            }
            else if (s.Length <= 4)
            {
                Tipo_bila_ = " ";
                sEan = s.PadLeft(20, Convert.ToChar("0"));
            }
            else
            {
                if (false)
                {
                    //         'Articoli pesati alle casse con ceck-out
                    Tipo_bila_ = "S";
                    sEan = s;
                }
                else
                {
                    Tipo_bila_ = " ";
                    sEan = s.PadLeft(13, Convert.ToChar("0"));        // FillChar(Trim(m_rsVarPos("pos_ean")), 13, "0", "D")
                }

            }
            sEan = s.PadLeft(20, Convert.ToChar("0"));

            return sEan;
        }

    }
}
