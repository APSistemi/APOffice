using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.IO;
using System.Globalization;

namespace APOffice
{
    class clsQuery
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private const string TABANACNF = "AnaConfigurazione";
        private const string TABANAART = "AnaArticoli";
        private const string TABANAEAN = "AnaBarcode";
        private const string TABLISACQ = "GesLisAcquisto";
        private const string TABLISVEN = "GesLisVendita";
        private const string TABGESOFA = "GesOffArticoli";
        private const string TABANACLI = "AnaClienti";
        private const string TABANAFOR = "AnaFornitori";
        private const string TABTABFIM = "TabForImport";
        private const string TABANATES = "AnaTessere";

        private static Dictionary<string, string> _cacheTabListiniIva = new Dictionary<string, string>();

        private string _strConSql = "";
        private string _strConSqlSta = "";
        private string _strConSqlLog = "";

        public string _strPar031Lotti2Pos = "";

        public clsQuery()
        {
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
            _strConSqlLog = _clsFun.ConSql("4");
        }

        public DataTable ConfAzienda(string strCod)
        {
            string s = "SELECT * FROM AnaConfigurazione WHERE cnf_cod='" + strCod + "'";
            DataTable t = _clsFun.FillTabSql(TABANACNF, s, false, _strConSql);

            if (t.Rows.Count == 0)
                MessageBox.Show("Configurazione azienda mancante o accesso al database non riuscito!", "CONTROLLO ACCESSO DATI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return t;
        }

        public DataTable ConfSeek(string strCod, string strTip)
        {
            string sCod = strCod;
            if (sCod == "")
            {
                sCod = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
                if (sCod == "")
                    sCod = "001";
            }
            string s = "SELECT cnf_cod, cnf_pos, cnf_bil FROM AnaConfigurazione WHERE cnf_cod='" + sCod + "'";
            DataTable t = _clsFun.FillTabSql(TABANACNF, s, false, _strConSql);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "CnfVar",
                Caption = "Path i/o",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "CnfVab",
                Caption = "Path i/o",
                MaxLength = 100,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "CnfVao",
                Caption = "Path i/o offerte",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "CnfVen",
                Caption = "Path i/o offerte",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            if (t.Rows.Count == 0)
                MessageBox.Show("Configurazione server mancante!");
            else
            {
                if (strTip == "POS")
                {
                    if ((string)t.Rows[0]["cnf_pos"] == "")
                        MessageBox.Show("Definizione cassa mancante in configurazione!");
                    else
                    {
                        string p = "TabPos";
                        s = "SELECT * FROM " + p + " WHERE tab_cod='" + t.Rows[0]["cnf_pos"] + "'";
                        DataTable tTmp = _clsFun.FillTabSql(p, s, true, _strConSql);
                        if (tTmp.Rows.Count > 0)
                        {
                            t.Rows[0]["CnfVar"] = ((string)tTmp.Rows[0]["tab_var"]).Trim();
                            t.Rows[0]["CnfVao"] = ((string)tTmp.Rows[0]["tab_vao"]).Trim();
                        }
                        else
                            MessageBox.Show("Path variazioni casse non definita in tabella casse!");
                    }

                    if ((string)t.Rows[0]["cnf_bil"] != "")
                    {
                        string p = "TabBilance";
                        s = "SELECT * FROM " + p + " WHERE tab_ann=0 AND tab_cod='" + t.Rows[0]["cnf_bil"] + "'";
                        DataTable tTmp = _clsFun.FillTabSql(p, s, true, _strConSql);
                        if (tTmp.Rows.Count > 0)
                            t.Rows[0]["CnfVab"] = ((string)tTmp.Rows[0]["tab_var"]).Trim();
                        else
                            MessageBox.Show("Path variazioni bilance non definita in tabella bilance!");
                    }

                }

                if (strTip == "VEN")
                {
                    if ((string)t.Rows[0]["cnf_pos"] == "")
                        MessageBox.Show("Definizione cassa mancante in configurazione!");
                    else
                    {
                        string p = "TabPos";
                        s = "SELECT * FROM " + p + " WHERE tab_cod='" + t.Rows[0]["cnf_pos"] + "'";
                        DataTable tTmp = _clsFun.FillTabSql(p, s, true, _strConSql);
                        if (tTmp.Rows.Count > 0)
                            t.Rows[0]["CnfVen"] = ((string)tTmp.Rows[0]["tab_ven"]).Trim();
                        else
                            MessageBox.Show("Path venduto casse non definita in tabella casse!");
                    }
                }

                //if (strTip == "BIL")
                //{
                //    if ((string)t.Rows[0]["cnf_bil"] == "")
                //        MessageBox.Show("Definizione bilancia mancante in configurazione!");
                //    else
                //    {
                //        string p = "TabBilance";
                //        s = "SELECT * FROM " + p + " WHERE tab_cod='" + t.Rows[0]["cnf_pos"] + "'";
                //        DataTable tTmp = _clsFun.FillTabSql(p, s, true, _strConSql);
                //        if (tTmp.Rows.Count > 0)
                //            t.Rows[0]["CnfVar"] = ((string)tTmp.Rows[0]["tab_var"]).Trim();
                //        else
                //            MessageBox.Show("Path variazioni bilance non definita in tabella bilance!");
                //    }
                //}

            }

            return t;
        }

        public DataTable DocCfo(string strCfo, string strCod)
        {
            DataTable t = new clsGenTabTmp().TabTmpDocMov("DocTes");

            string s = "SELECT * FROM " + TABANACLI + " WHERE cli_cod='" + strCod + "'";
            if(strCfo == "FOR")
                s = "SELECT * FROM " + TABANAFOR + " WHERE for_cod='" + strCod + "'";
            DataTable tCfo = _clsFun.FillTabSql(TABANACLI, s, false, _strConSql);
            if(tCfo.Rows.Count > 0)
            {
                DataRow x = t.NewRow();
                if (strCfo == _clsDef.TIPCLI)
                {
                    x["tmp_cfo"] = tCfo.Rows[0]["cli_cod"];
                    x["tmp_rag"] = tCfo.Rows[0]["cli_des"];
                    x["tmp_ra2"] = tCfo.Rows[0]["cli_de2"];
                    x["tmp_piv"] = tCfo.Rows[0]["cli_piv"];
                    x["tmp_cfi"] = tCfo.Rows[0]["cli_cfi"];
                    x["tmp_ind"] = tCfo.Rows[0]["cli_ind"];
                    x["tmp_ncv"] = "";
                    if (!DBNull.Value.Equals(tCfo.Rows[0]["cli_ncv"]))
                        x["tmp_ncv"] = tCfo.Rows[0]["cli_ncv"];
                    x["tmp_loc"] = ((string)tCfo.Rows[0]["cli_cap"]).Trim() + " " + ((string)tCfo.Rows[0]["cli_loc"]).Trim() + " (" + (string)tCfo.Rows[0]["cli_prv"] + ")";
                }
                else if (strCfo == _clsDef.TIPFOR)
                {
                    x["tmp_cfo"] = tCfo.Rows[0]["for_cod"];
                    x["tmp_rag"] = tCfo.Rows[0]["for_des"];
                    x["tmp_piv"] = tCfo.Rows[0]["for_piv"];
                    x["tmp_cfi"] = tCfo.Rows[0]["for_cfi"];
                    x["tmp_ind"] = tCfo.Rows[0]["for_ind"];

                    x["tmp_ncv"] = "";
                    if (!DBNull.Value.Equals(tCfo.Rows[0]["for_ncv"]))
                        x["tmp_ncv"] = tCfo.Rows[0]["for_ncv"];

                    x["tmp_loc"] = ((string)tCfo.Rows[0]["for_cap"]).Trim() + " " + ((string)tCfo.Rows[0]["for_loc"]).Trim() + " (" + (string)tCfo.Rows[0]["for_prv"] + ")";
                }

                x["tmp_ndo"] = "";
                //x["tmp_ddo"] = t.Rows[0][""];
                t.Rows.Add(x);
            }
            else
            {
                DataRow x = t.NewRow();
                x["tmp_cfo"] = "";
                x["tmp_rag"] = "";
                x["tmp_piv"] = "";
                x["tmp_cfi"] = "";
                x["tmp_ind"] = "";
                x["tmp_loc"] = "";
                x["tmp_ndo"] = "";
                //x["tmp_ddo"] = t.Rows[0][""];
                t.Rows.Add(x);
            }

            return t;
        }

        public DataTable ArtSeek(string strKey, string strSeek)
        {
            string s = "";
            string sKey = "";
            string sSql = "";
            string sFor = "";
            DataTable tTmp;
            DataTable t;
            Boolean bBil = false;

            if(strSeek.Length > 3 && strSeek.Substring(0,3) == "FOR")
            {
                sFor = strSeek.Substring(3, 5);
                strSeek = "SEEK";
            }

            if(_clsFun.Numerico(strKey))
            {
                if (strSeek == "PLU")
                {
                    string sReb = strKey.Substring(0, 1);
                    string sPlu = strKey.Substring(1);

                    if (_clsFun.Numerico(sPlu))
                    {
                        sPlu = Convert.ToInt32(sPlu).ToString();

                        sSql = "SELECT * FROM AnaArticoli ";
                        sSql += "WHERE ";
                        sSql += "art_reb='" + sReb + "' AND ";
                        sSql += "(art_plu='" + sPlu + "' OR art_plu='" + sPlu.PadLeft(4, '0') + "')";
                    }
                }
                else if(strSeek.Length > 3 && strSeek.Substring(0,3) == "ARF")
                {
                    sFor = strSeek.Substring(3, 5);
                    sSql = "SELECT TOP 1 lia_art FROM GesLisAcquisto WHERE lia_arf='" + strKey + "' AND lia_for='" + sFor + "'";
                    t = _clsFun.FillTabSql(TABLISACQ, sSql, true, _strConSql);

                    if (t.Rows.Count > 0)
                    {
                        sKey = ((string)t.Rows[0]["lia_art"]).PadLeft(7, Convert.ToChar("0"));
                        sSql = "SELECT * FROM AnaArticoli WHERE art_cod='" + sKey + "'";
                    }
                }
                else if(strKey.Length < 8) 
                {
                    sKey = strKey.PadLeft(7,Convert.ToChar("0"));
                    sSql = "SELECT * FROM AnaArticoli WHERE art_cod='" + sKey + "'";
                }
                else if(strKey.Length <= 13) 
                {
                    bBil = true;
                    sKey = strKey.Trim();
                    sSql = "SELECT ean_art, ean_ean, ean_tip, ean_ecp, AnaArticoli.* ";
                    sSql += "FROM AnaBarcode ";
                    sSql += "LEFT OUTER JOIN AnaArticoli ON AnaBarcode.ean_art = AnaArticoli.art_cod ";
                    sSql += "WHERE ";
                    
                    string sTrimKey = sKey.TrimStart('0');
                    if (!string.IsNullOrEmpty(sTrimKey) && sTrimKey != sKey)
                        sSql += "(ean_ean='" + sKey + "' OR ean_ean='" + sTrimKey + "')";
                    else
                        sSql += "ean_ean='" + sKey + "'";
                }
            }
            else
            {
                //sTip = "ART";
                sKey = _clsFun.FaiLApice(strKey.Trim());
                sSql = "SELECT * FROM AnaArticoli WHERE art_des LIKE '%" + sKey + "%'";
            }

            t = _clsFun.FillTabSql(TABANAART, sSql, false, _strConSql);

            if(strSeek == "SEEK" && t.Rows.Count != 1)
            {
                frmSeekArt f = new frmSeekArt();
                f._bolValori = false;
                f._strSeek = strKey;
                f.ShowDialog();
                if (f._tabArt != null && f._tabArt.Rows.Count > 0)
                {
                    sSql = "SELECT * FROM AnaArticoli WHERE art_cod='" + (string)f._tabArt.Rows[0]["tmp_art"] +  "'";
                    t = _clsFun.FillTabSql(TABANAART, sSql, false, _strConSql);
                }
                else
                    t.Clear();
                
            }
            if(t.Rows.Count > 0)
            {
                tTmp = new clsGenTabTmp().TabTmpArt("ArtCod");
                DataRow x = tTmp.NewRow();
                x["tmp_art"] = t.Rows[0]["art_cod"];
                x["tmp_ard"] = t.Rows[0]["art_des"];
                x["tmp_arb"] = t.Rows[0]["art_deb"];
                x["tmp_sta"] = t.Rows[0]["art_sta"];
                x["tmp_rep"] = t.Rows[0]["art_rep"];
                x["tmp_iva"] = t.Rows[0]["art_iva"];
                x["tmp_umi"] = t.Rows[0]["art_umi"];
                x["tmp_bil"] = t.Rows[0]["art_bil"];
                x["tmp_ecr"] = t.Rows[0]["art_ec1"] + (string)t.Rows[0]["art_ec2"] + (string)t.Rows[0]["art_ec3"];

                x["tmp_pxc"] = t.Rows[0]["art_pxc"];
                x["tmp_tgr"] = t.Rows[0]["art_tgr"];
                x["tmp_pne"] = t.Rows[0]["art_pne"];
                x["tmp_sfr"] = t.Rows[0]["art_sfr"];
                x["tmp_tar"] = t.Rows[0]["art_tar"];
                x["tmp_ori"] = t.Rows[0]["art_ori"];
                x["tmp_cal"] = t.Rows[0]["art_cal"];
                x["tmp_cat"] = t.Rows[0]["art_cat"];
                x["tmp_reb"] = t.Rows[0]["art_reb"];
                x["tmp_gsc"] = t.Rows[0]["art_gsc"];
                x["tmp_plu"] = ((string)t.Rows[0]["art_plu"]).Trim();

                x["tmp_tas"] = "";
                if (((string)t.Rows[0]["art_tas"]).Trim() != "0")
                    x["tmp_tas"] = ((string)t.Rows[0]["art_tas"]).Trim();

                x["tmp_bpz"] = t.Rows[0]["art_bpz"];  //a pezzo in bilancia
                x["tmp_tra"] = t.Rows[0]["art_tra"];  //tracciabilità lotto 'S'
                x["tmp_eqp"] = t.Rows[0]["art_eqp"];  //Equivalenza prezzo
                if (bBil && t.Columns.Contains("ean_tip") && !DBNull.Value.Equals(t.Rows[0]["ean_tip"]) && (string)t.Rows[0]["ean_tip"] == "L")
                    x["tmp_lot"] = ((string)t.Rows[0]["ean_ean"]).Substring(2, 3);  //Equivalenza prezzo
                if (bBil && t.Columns.Contains("ean_ecp") && !DBNull.Value.Equals(t.Rows[0]["ean_ecp"]) && (Boolean)t.Rows[0]["ean_ecp"])
                    x["tmp_ecp"] = (Boolean)t.Rows[0]["ean_ecp"];  //Ean a peso


                tTmp.Rows.Add(x);
                t = tTmp.Copy();
            }

            if(t.Rows.Count > 0)
            {
                s = "";
                if (strSeek.Length > 3 && strSeek.Substring(0, 3) == "LNE")
                    s = strSeek.Substring(3);

                t = ArtPrezzo(t, s);
                t = ArtCosto(sFor,"", t, DateTime.Today);
            }

            return t;
        }

        public string SeekArfEan(string strFor, string strArf, DataRow[] jEan)
        {
            string sArt = "";
            string s = "SELECT TOP 1 lia_art, lia_ann FROM GesLisAcquisto WHERE lia_for='" + strFor + "' AND lia_arf='" + strArf + "' ORDER BY lia_dti DESC";
            DataTable t = _clsFun.FillTabSql(TABLISACQ, s, true, _strConSql);

            if (t.Rows.Count > 0)
            {
                if((Boolean)t.Rows[0]["lia_ann"] == false)
                    sArt = (string)t.Rows[0]["lia_art"];
            }
            else
            {
                for (int i = 0; i < jEan.Length; i++)
                {
                    s = "SELECT TOP 1 ean_art FROM AnaBarcode WHERE ean_ann=0 AND ean_ean='" + (string)jEan[i]["die_ean"] + "' ";
                    t = _clsFun.FillTabSql(TABLISACQ, s, true, _strConSql);

                    if (t.Rows.Count > 0)
                    {
                        sArt = (string)t.Rows[0]["ean_art"];
                        break;
                    }
                }
            }

            return sArt;
        }

        public DataTable SeekArtEan(string strArt, string strPar)
        {
            string s = "";
            string sArt = strArt.Trim().PadLeft(7, '0');
            string sSql = "";

            string sLisNeg = "";
            if(strPar.Length >3 && strPar.Substring(0,3) == "LNE")
                sLisNeg = strPar.Substring(3);

            string sLottoEanPref = "";
            int iLotGGscadenza = 0;

            if(_strPar031Lotti2Pos != "")
            {
                string[] a = _strPar031Lotti2Pos.Split(',');
                if (a.Length > 2)
                    sLottoEanPref = a[1];
                if (a.Length > 3 && _clsFun.Numerico(a[3].Trim(), "0123456789"))
                    iLotGGscadenza = Convert.ToInt16(a[3]);
            }

            sSql = "SELECT ";            
            sSql += "AnaArticoli.art_cod, "; 
            sSql += "AnaArticoli.art_des, "; 
            sSql += "AnaArticoli.art_deb, "; 
            sSql += "AnaArticoli.art_sta, "; 
            sSql += "AnaArticoli.art_iva, "; 
            sSql += "AnaArticoli.art_umi, "; 
            sSql += "AnaArticoli.art_pxc, ";
            sSql += "AnaArticoli.art_tgr, ";
            sSql += "AnaArticoli.art_pne, ";
            sSql += "AnaArticoli.art_rep, "; 
            sSql += "AnaArticoli.art_ec1, ";
            sSql += "AnaArticoli.art_ec2, ";
            sSql += "AnaArticoli.art_ec3, ";
            sSql += "AnaArticoli.art_ori, "; 
            sSql += "AnaArticoli.art_cal, "; 
            sSql += "AnaArticoli.art_cat, "; 
            sSql += "AnaArticoli.art_reb, "; 
            sSql += "AnaArticoli.art_tra, "; 
            sSql += "AnaArticoli.art_gsc, ";
            sSql += "AnaArticoli.art_plu, ";
            sSql += "AnaArticoli.art_tas, ";        //Tasto
            sSql += "AnaArticoli.art_tar, ";
            sSql += "AnaArticoli.art_sfr, ";
            sSql += "AnaArticoli.art_bpz, ";
            sSql += "AnaArticoli.art_eti, ";
            sSql += "AnaArticoli.art_img, ";
            sSql += "AnaBarcode.ean_bil, ";
            sSql += "AnaBarcode.ean_tip, ";
            sSql += "AnaBarcode.ean_ean, ";
            sSql += "AnaBarcode.ean_qta, ";
            sSql += "AnaBarcode.ean_ecp, ";
            sSql += "AnaBarcode.ean_prv, ";
            sSql += "AnaBarcode.ean_dti, ";
            sSql += "AnaBarcode.ean_dtm, ";
            sSql += "AnaBarcode.ean_pun, ";
            sSql += "AnaBarcode.ean_ann "; 
            sSql += "FROM AnaArticoli LEFT OUTER JOIN AnaBarcode ON AnaArticoli.art_cod = AnaBarcode.ean_art ";
            sSql += "WHERE art_cod='" + sArt + "' ";
            sSql += "ORDER BY AnaBarcode.ean_dti DESC";
            DataTable tArt = _clsFun.FillTabSql(TABANAART, sSql, false, _strConSql);

            DataTable t = new clsGenTabTmp().TabTmpArt("ArtCod");

            Boolean b = true;

            for(int i=1; i <= 2; i++)           //Meno priorità ai barcode brevi
            {
                foreach (DataRow y in tArt.Rows)
                {
                    b= false;

                    if (((string)y["art_sta"]).Trim() == _clsDef.STAREP || ((string)y["art_sta"]).Trim() == _clsDef.STAPRE)
                    {
                        if (i == 1)
                            b = true;
                    }
                    else
                    {
                        s = Convert.ToString(y["ean_ean"]).Trim();
                        if (s == "" && i == 2)
                            s = "0";    //Facciamo passare anche se senza barcode

                        if (s != "" && i == 1 && _clsFun.Numerico(s,true) && Convert.ToInt64(s) > 29000000)
                            b = true;
                        else if (i == 2 && _clsFun.Numerico(s, true) && Convert.ToInt64(s) <= 29000000)
                            b = true;
                    }

                    if ((string)y["ean_tip"] == "L" && s.Length == 13 && s.Substring(0, 2) == sLottoEanPref && !DBNull.Value.Equals(y["ean_bil"]) && (Boolean)y["ean_bil"] == false)
                    {
                        b = false;

                        if (!(Boolean)y["ean_ann"] && iLotGGscadenza > 0 && ((DateTime)y["ean_dtm"]).AddDays(iLotGGscadenza) < DateTime.Today)
                        {
                            s = "UPDATE AnaBarcode SET ean_ann=1, ean_dtm=" + _clsFun.DaySql(DateTime.Today) + " WHERE ean_ean='" + Convert.ToString(y["ean_ean"]).Trim() + "'";
                            _clsFun.SqlWrite(s, _strConSql);
                            y["ean_ann"] = true;
                            b = true;
                        }
                        else if (DateTime.Compare((DateTime)y["ean_dtm"], DateTime.Today) == 0)
                            b = true;
                    }

                    if (b)
                    {
                        DataRow x = t.NewRow();
                        x["tmp_art"] = y["art_cod"];
                        x["tmp_ard"] = y["art_des"];
                        x["tmp_arb"] = y["art_deb"];
                        x["tmp_sta"] = y["art_sta"];
                        x["tmp_rep"] = y["art_rep"];

                        x["tmp_ecr"] = ((string)y["art_ec1"] + new string(' ', 3)).Substring(0, 3);
                        x["tmp_ecr"] += ((string)y["art_ec2"] + new string(' ', 3)).Substring(0, 3);
                        x["tmp_ecr"] += ((string)y["art_ec3"] + new string(' ', 3)).Substring(0, 3);

                        x["tmp_iva"] = y["art_iva"];
                        x["tmp_umi"] = y["art_umi"];

                        x["tmp_pxc"] = y["art_pxc"];
                        x["tmp_tgr"] = y["art_tgr"];
                        x["tmp_pne"] = y["art_pne"];
                        x["tmp_ori"] = y["art_ori"];
                        x["tmp_cal"] = y["art_cal"];
                        x["tmp_cat"] = y["art_cat"];
                        x["tmp_tar"] = y["art_tar"];
                        x["tmp_sfr"] = y["art_sfr"];
                        x["tmp_reb"] = y["art_reb"];
                        x["tmp_gsc"] = y["art_gsc"];
                        x["tmp_plu"] = ((string)y["art_plu"]).Trim();

                        x["tmp_tas"] = "";
                        if (!DBNull.Value.Equals(y["art_tas"]))
                            x["tmp_tas"] = ((string)y["art_tas"]).Trim();

                        x["tmp_bpz"] = y["art_bpz"];
                        x["tmp_tra"] = y["art_tra"];  //tracciabilità da definire
                        x["tmp_eti"] = y["art_eti"];  //tipo etichetta
                        x["tmp_tas"] = y["art_tas"];  //tasto bilancia
                        x["tmp_img"] = y["art_img"];  //File immagine

                        x["tmp_bil"] = y["ean_bil"];
                        x["tmp_ean"] = y["ean_ean"];
                        x["tmp_eaq"] = y["ean_qta"];
                        x["tmp_eap"] = y["ean_prv"];
                        x["tmp_ecp"] = y["ean_ecp"];
                        x["tmp_eaa"] = y["ean_ann"];
                        x["tmp_epu"] = y["ean_pun"];        //Punti
                        t.Rows.Add(x);
                    }
                }
            }
            if (t.Rows.Count > 0)
            {
                t = ArtPrezzo(t, sLisNeg);
                t = ArtCosto("", "", t, DateTime.Today);
            }

            return t;
        }

        public DataTable SeekArtEanIn(string strArtIn, string strPar)
        {
            // Versione ottimizzata per caricamento bulk
            string sSql = "";
            string sLisNeg = "";
            if (strPar.Length > 3 && strPar.Substring(0, 3) == "LNE")
                sLisNeg = strPar.Substring(3);

            sSql = "SELECT ";
            sSql += "AnaArticoli.*, AnaBarcode.* ";
            sSql += "FROM AnaArticoli LEFT OUTER JOIN AnaBarcode ON AnaArticoli.art_cod = AnaBarcode.ean_art ";
            sSql += "WHERE art_cod IN (" + strArtIn + ") ";
            sSql += "ORDER BY AnaBarcode.ean_dti DESC";
            
            DataTable tArt = _clsFun.FillTabSql(TABANAART, sSql, false, _strConSql);
            return tArt;
        }

        public DataTable SeekEanArt(string strEan)
        {
            DataTable t = new clsGenTabTmp().TabTmpArt("ArtCod");

            string s = "";
            string sEan = strEan.Trim();

            s = "SELECT * FROM AnaBarcode WHERE ean_ean='" + strEan + "'";
            DataTable tEan = _clsFun.FillTabSql(TABANAEAN, s, false, _strConSql);

            if(tEan.Rows.Count > 0)
            {
                s = "SELECT * FROM AnaArticoli WHERE art_cod='" + tEan.Rows[0]["ean_art"] + "'";
                DataTable tTmp = _clsFun.FillTabSql(TABANAEAN, s, false, _strConSql);
                if (tTmp.Rows.Count > 0)
                {
                    DataRow x = t.NewRow();
                    x["tmp_art"] = tTmp.Rows[0]["art_cod"];
                    x["tmp_ard"] = tTmp.Rows[0]["art_des"];
                    x["tmp_arb"] = tTmp.Rows[0]["art_deb"];
                    x["tmp_sta"] = tTmp.Rows[0]["art_sta"];
                    x["tmp_rep"] = tTmp.Rows[0]["art_rep"];

                    x["tmp_ecr"] = ((string)tTmp.Rows[0]["art_ec1"] + new string(' ', 3)).Substring(0, 3);
                    x["tmp_ecr"] += ((string)tTmp.Rows[0]["art_ec2"] + new string(' ', 3)).Substring(0, 3);
                    x["tmp_ecr"] += ((string)tTmp.Rows[0]["art_ec3"] + new string(' ', 3)).Substring(0, 3);

                    x["tmp_iva"] = tTmp.Rows[0]["art_iva"];
                    x["tmp_umi"] = tTmp.Rows[0]["art_umi"];

                    x["tmp_pxc"] = tTmp.Rows[0]["art_pxc"];
                    x["tmp_tgr"] = tTmp.Rows[0]["art_tgr"];
                    x["tmp_pne"] = tTmp.Rows[0]["art_pne"];
                    x["tmp_ori"] = tTmp.Rows[0]["art_ori"];
                    x["tmp_cal"] = tTmp.Rows[0]["art_cal"];
                    x["tmp_cat"] = tTmp.Rows[0]["art_cat"];
                    x["tmp_reb"] = tTmp.Rows[0]["art_reb"];
                    x["tmp_gsc"] = tTmp.Rows[0]["art_gsc"];
                    x["tmp_plu"] = tTmp.Rows[0]["art_plu"];
                    x["tmp_bpz"] = tTmp.Rows[0]["art_bpz"];
                    x["tmp_tra"] = tTmp.Rows[0]["art_tra"];  //tracciabilità da definire

                    x["tmp_bil"] = tEan.Rows[0]["ean_bil"];
                    x["tmp_ean"] = tEan.Rows[0]["ean_ean"];
                    x["tmp_eaq"] = tEan.Rows[0]["ean_qta"];
                    x["tmp_eap"] = tEan.Rows[0]["ean_prv"];
                    x["tmp_ecp"] = tEan.Rows[0]["ean_ecp"];
                    x["tmp_eaa"] = tEan.Rows[0]["ean_ann"];
                    x["tmp_epu"] = tEan.Rows[0]["ean_pun"];
                    t.Rows.Add(x);
                }
            }
            if (t.Rows.Count > 0)
            {
                t = ArtPrezzo(t, "");
                t = ArtCosto("", "", t, DateTime.Today);
            }

            return t;
        }

        public Decimal ArtCostoUltimo(string strFor, string strTip, string strArt, DateTime dayLim)
        {
            decimal d = 0;
            string s = "SELECT TOP 1 lia_cos FROM GesLisAcquisto ";
            s += "WHERE lia_art='" + strArt + "' AND lia_ann=0 ";
            //if (strFor != "" || strTip != "")
            //    s += "WHERE ... ";
            if (strFor != "")
                s += " AND lia_for='" + strFor + "'";
            s += "AND lia_dti <= " + _clsFun.DaySql(dayLim) + " ";
            s += "ORDER BY lia_dti DESC ";

            DataTable t = _clsFun.FillTabSql(TABLISACQ, s, true, _strConSql);

            if (t.Rows.Count > 0)
            {
                d = (decimal)t.Rows[0]["lia_cos"];
            }

            return d;
        }

        public DataTable ArtCosto(string strFor, string strTip, DataTable tabTmp, DateTime dayLim)
        {
            string s = "SELECT TOP 1 lia_cos, lia_for, lia_arf, for_des FROM GesLisAcquisto ";
            s += "LEFT JOIN AnaFornitori ON AnaFornitori.for_cod = GesLisAcquisto.lia_for ";
            s += "WHERE lia_art='" + tabTmp.Rows[0]["tmp_art"] + "' AND (lia_ann=0 OR lia_ann is null)  ";
            //if (strFor != "" || strTip != "")
            //    s += "WHERE ... ";
            if (strFor != "")
                s += " AND lia_for='" + strFor + "'" ;
            s += "AND lia_dti <= " + _clsFun.DaySql(dayLim) + " ";
            s += "ORDER BY lia_dti DESC ";

            DataTable t = _clsFun.FillTabSql(TABLISACQ, s, true, _strConSql);

            if (t.Rows.Count > 0)
            {
                tabTmp.Rows[0]["tmp_for"] = t.Rows[0]["lia_for"];
                tabTmp.Rows[0]["tmp_fod"] = t.Rows[0]["for_des"];
                tabTmp.Rows[0]["tmp_arf"] = t.Rows[0]["lia_arf"];
                tabTmp.Rows[0]["tmp_cos"] = t.Rows[0]["lia_cos"];
            }

            return tabTmp;
        }

        public DataTable ArtPrezzo(DataTable tabTmp, string strLne)
        {
            decimal dPrv = 0;   //Prezzo in vigore
            decimal dPpv = 0;   //Prezzo pieno
            string sLis = "";   //Listino
            string sLiv = "";   //Listino ivato
            string sLva = "";   //Listino annullato

            if ((string)tabTmp.Rows[0]["tmp_art"] == "0009931")
                Console.WriteLine("aaaaaaaa");

            string s = "SELECT TOP 100 liv_prv, liv_lis, liv_dti, liv_dtf, liv_ann FROM GesLisVendita WHERE ";
            s += "liv_art='" + tabTmp.Rows[0]["tmp_art"] + "' AND ";
            s += "liv_dti <= " + _clsFun.DaySql(DateTime.Today) + " ";
            //s += "liv_ann=0 ";
            s += "ORDER BY liv_dti DESC";

            s = "";
            s += "SELECT * FROM (";
            s += "SELECT ";
            s += "ROW_NUMBER() OVER (PARTITION BY liv_art, liv_lis ORDER BY liv_art, liv_dti DESC) AS ROW, ";
            s += "liv_art, ";
            s += "liv_lis, ";
            s += "liv_prv, ";
            s += "liv_dti, ";
            s += "liv_dtf, ";
            s += "liv_ann  ";
            s += "FROM GesLisVendita ";
            s += "WHERE liv_art='" + tabTmp.Rows[0]["tmp_art"] + "' AND liv_dti <= " + _clsFun.DaySql(DateTime.Today) + "";
            s += ") AS A ";
            s += "WHERE ROW = 1";
            DataTable t = _clsFun.FillTabSql(TABLISVEN, s, false, _strConSql);

            DataView v = new DataView();

            if (t != null && t.Columns.Contains("liv_lis") && t.Rows.Count > 0)
            {
                if (strLne != "")
                {
                    s = "liv_lis='" + strLne + "' AND liv_dti <= " + _clsFun.DayMdb(DateTime.Today) + " AND liv_dtf >= " + _clsFun.DayMdb(DateTime.Today) + "";
                    v = new DataView(t, s, "liv_dti DESC", DataViewRowState.CurrentRows);
                }

                if (v.Count > 0 && strLne == _clsDef.LISPOS && !DBNull.Value.Equals(v[0]["liv_ann"]) && Convert.ToBoolean(v[0]["liv_ann"]))
                    sLva = strLne;

                if (v.Count == 0 || (!DBNull.Value.Equals(v[0]["liv_ann"]) && Convert.ToBoolean(v[0]["liv_ann"])))
                {
                    s = "liv_lis='" + _clsDef.LISPRO + "' AND liv_dti <= " + _clsFun.DayMdb(DateTime.Today) + " AND liv_dtf >= " + _clsFun.DayMdb(DateTime.Today) + "";
                    v = new DataView(t, s, "liv_dti DESC", DataViewRowState.CurrentRows);
                }
                if (v.Count > 0 && (DBNull.Value.Equals(v[0]["liv_ann"]) || !Convert.ToBoolean(v[0]["liv_ann"])))
                {
                    dPrv = (decimal)v[0]["liv_prv"];
                    sLis = (string)v[0]["liv_lis"];
                }
                s = "liv_lis='" + _clsDef.LISPOS + "'";
                v = new DataView(t, s, "liv_dti DESC", DataViewRowState.CurrentRows);

                if (v.Count == 0 || (!DBNull.Value.Equals(v[0]["liv_ann"]) && Convert.ToBoolean(v[0]["liv_ann"])))
                {
                    s = "liv_lis='" + _clsDef.LISFOR + "'";
                    v = new DataView(t, s, "liv_dti DESC", DataViewRowState.CurrentRows);
                }
                if (v.Count > 0 && (DBNull.Value.Equals(v[0]["liv_ann"]) || !Convert.ToBoolean(v[0]["liv_ann"])))
                {
                    dPpv = (decimal)v[0]["liv_prv"];
                }
                if (dPrv == 0)
                {
                    dPrv = dPpv;
                    if (v.Count > 0 && (DBNull.Value.Equals(v[0]["liv_ann"]) || !Convert.ToBoolean(v[0]["liv_ann"])))
                        sLis = (string)v[0]["liv_lis"];
                }
            }

            if (!string.IsNullOrEmpty(sLis))
            {
                if (_cacheTabListiniIva.ContainsKey(sLis))
                {
                    sLiv = _cacheTabListiniIva[sLis];
                }
                else
                {
                    s = "SELECT tab_cod, tab_iva FROM TabListini WHERE tab_cod='" + sLis + "'";
                    t = _clsFun.FillTabSql("TabListini", s, true, _strConSql);
                    if (t.Rows.Count > 0 && !DBNull.Value.Equals(t.Rows[0]["tab_iva"]))
                    {
                        sLiv = (string)t.Rows[0]["tab_iva"];
                        _cacheTabListiniIva[sLis] = sLiv;
                    }
                }
            }

            foreach (DataRow y in tabTmp.Rows)
            {
                y["tmp_prv"] = dPrv;
                y["tmp_ppv"] = dPpv;
                y["tmp_lis"] = sLis;
                y["tmp_liv"] = sLiv;
                y["tmp_lva"] = sLva;            //Listino annullato
            }

            if (t != null)
            {
                t.Dispose();
                t = null;
            }

            return tabTmp;
        }

        public bool AggiornaPrezzoVenditaConStorico(string strArt, string strLis, decimal decPrv, string strConSql, DateTime? dtDecorrenza = null)
        {
            if (string.IsNullOrEmpty(strArt) || string.IsNullOrEmpty(strLis))
                return false;

            if (string.IsNullOrEmpty(strConSql))
                strConSql = _strConSql;

            DateTime dToday = dtDecorrenza.HasValue ? dtDecorrenza.Value.Date : DateTime.Today;

            // 1. Cerca record attivi per questo articolo e listino ordinati per data inizio decrescente
            string sQry = "SELECT * FROM GesLisVendita WHERE liv_art='" + strArt.Replace("'", "''") + "' AND liv_lis='" + strLis.Replace("'", "''") + "' AND (liv_ann=0 OR liv_ann IS NULL) ORDER BY liv_dti DESC";
            DataTable t = _clsFun.FillTabSql(TABLISVEN, sQry, false, strConSql);

            var batchSql = new List<string>();

            if (t != null && t.Rows.Count > 0)
            {
                DataRow rTop = t.Rows[0];
                DateTime rDti = Convert.ToDateTime(rTop["liv_dti"]).Date;
                int rIdx = !DBNull.Value.Equals(rTop["liv_idx"]) ? Convert.ToInt32(rTop["liv_idx"]) : 0;

                if (rDti == dToday)
                {
                    // Se esiste già una riga con data inizio oggi, aggiorna il prezzo odierno
                    string sqlUpd = "UPDATE GesLisVendita SET liv_prv = " + decPrv.ToString(CultureInfo.InvariantCulture) + ", liv_dtf = " + _clsFun.DaySql(_clsDef.DAYOUT) + " WHERE liv_idx = " + rIdx;
                    batchSql.Add(sqlUpd);
                }
                else if (rDti < dToday)
                {
                    // Chiude il record precedente a ieri (o data antecedente a dToday)
                    DateTime dIeri = dToday.AddDays(-1);
                    string sqlClose = "UPDATE GesLisVendita SET liv_dtf = " + _clsFun.DaySql(dIeri) + " WHERE liv_idx = " + rIdx;
                    batchSql.Add(sqlClose);

                    // Inserisce nuova riga per oggi
                    string sqlIns = "INSERT INTO GesLisVendita (liv_lis, liv_art, liv_prv, liv_dti, liv_dtf, liv_day, liv_sta, liv_ann) VALUES (" +
                                    "'" + strLis.Replace("'", "''") + "', " +
                                    "'" + strArt.Replace("'", "''") + "', " +
                                    decPrv.ToString(CultureInfo.InvariantCulture) + ", " +
                                    _clsFun.DaySql(dToday) + ", " +
                                    _clsFun.DaySql(_clsDef.DAYOUT) + ", " +
                                    _clsFun.DaySql(DateTime.Today) + ", " +
                                    "'A', 0)";
                    batchSql.Add(sqlIns);
                }
                else
                {
                    // rDti > dToday (decorrenza futura): aggiorna il record
                    string sqlUpd = "UPDATE GesLisVendita SET liv_prv = " + decPrv.ToString(CultureInfo.InvariantCulture) + " WHERE liv_idx = " + rIdx;
                    batchSql.Add(sqlUpd);
                }
            }
            else
            {
                // Nessun record presente: crea il primo record
                string sqlIns = "INSERT INTO GesLisVendita (liv_lis, liv_art, liv_prv, liv_dti, liv_dtf, liv_day, liv_sta, liv_ann) VALUES (" +
                                "'" + strLis.Replace("'", "''") + "', " +
                                "'" + strArt.Replace("'", "''") + "', " +
                                decPrv.ToString(CultureInfo.InvariantCulture) + ", " +
                                _clsFun.DaySql(dToday) + ", " +
                                _clsFun.DaySql(_clsDef.DAYOUT) + ", " +
                                _clsFun.DaySql(DateTime.Today) + ", " +
                                "'A', 0)";
                batchSql.Add(sqlIns);
            }

            // Se il listino è LISPOS ("001"), aggiorna anche AnaArticoli.art_prv
            if (strLis == _clsDef.LISPOS)
            {
                string sqlArt = "UPDATE AnaArticoli SET art_prv = " + decPrv.ToString(CultureInfo.InvariantCulture) + " WHERE art_cod = '" + strArt.Replace("'", "''") + "'";
                batchSql.Add(sqlArt);
            }

            if (batchSql.Count > 0)
            {
                _clsFun.SqlWriteBatch(batchSql, strConSql);
                return true;
            }

            return false;
        }

        public bool AggiornaPrezzoVenditaConStoricoBatch(Dictionary<string, decimal> artPrezzi, string strLis, string strConSql, DateTime? dtDecorrenza = null)
        {
            if (artPrezzi == null || artPrezzi.Count == 0 || string.IsNullOrEmpty(strLis))
                return false;

            if (string.IsNullOrEmpty(strConSql))
                strConSql = _strConSql;

            DateTime dToday = dtDecorrenza.HasValue ? dtDecorrenza.Value.Date : DateTime.Today;
            var batchSql = new List<string>();

            // Raggruppa gli articoli in blocchi da 500 per query IN(...) ottimizzata
            var artList = artPrezzi.Keys.ToList();
            const int chunkSize = 500;
            for (int i = 0; i < artList.Count; i += chunkSize)
            {
                var chunk = artList.Skip(i).Take(chunkSize).ToList();
                string inClause = "'" + string.Join("','", chunk.Select(a => a.Replace("'", "''"))) + "'";

                string sQry = "SELECT * FROM GesLisVendita WHERE liv_art IN (" + inClause + ") AND liv_lis='" + strLis.Replace("'", "''") + "' AND (liv_ann=0 OR liv_ann IS NULL) ORDER BY liv_art, liv_dti DESC";
                DataTable t = _clsFun.FillTabSql(TABLISVEN, sQry, false, strConSql);

                var groupedByArt = new Dictionary<string, List<DataRow>>();
                if (t != null)
                {
                    foreach (DataRow r in t.Rows)
                    {
                        string art = r["liv_art"].ToString().Trim();
                        if (!groupedByArt.ContainsKey(art))
                            groupedByArt[art] = new List<DataRow>();
                        groupedByArt[art].Add(r);
                    }
                }

                foreach (var artCod in chunk)
                {
                    decimal decPrv = artPrezzi[artCod];
                    if (groupedByArt.ContainsKey(artCod) && groupedByArt[artCod].Count > 0)
                    {
                        DataRow rTop = groupedByArt[artCod][0]; // già ordinato per liv_dti DESC
                        DateTime rDti = Convert.ToDateTime(rTop["liv_dti"]).Date;
                        int rIdx = !DBNull.Value.Equals(rTop["liv_idx"]) ? Convert.ToInt32(rTop["liv_idx"]) : 0;

                        if (rDti == dToday)
                        {
                            batchSql.Add("UPDATE GesLisVendita SET liv_prv = " + decPrv.ToString(CultureInfo.InvariantCulture) + ", liv_dtf = " + _clsFun.DaySql(_clsDef.DAYOUT) + " WHERE liv_idx = " + rIdx);
                        }
                        else if (rDti < dToday)
                        {
                            DateTime dIeri = dToday.AddDays(-1);
                            batchSql.Add("UPDATE GesLisVendita SET liv_dtf = " + _clsFun.DaySql(dIeri) + " WHERE liv_idx = " + rIdx);
                            batchSql.Add("INSERT INTO GesLisVendita (liv_lis, liv_art, liv_prv, liv_dti, liv_dtf, liv_day, liv_sta, liv_ann) VALUES (" +
                                         "'" + strLis.Replace("'", "''") + "', '" + artCod.Replace("'", "''") + "', " +
                                         decPrv.ToString(CultureInfo.InvariantCulture) + ", " +
                                         _clsFun.DaySql(dToday) + ", " +
                                         _clsFun.DaySql(_clsDef.DAYOUT) + ", " +
                                         _clsFun.DaySql(DateTime.Today) + ", 'A', 0)");
                        }
                        else
                        {
                            batchSql.Add("UPDATE GesLisVendita SET liv_prv = " + decPrv.ToString(CultureInfo.InvariantCulture) + " WHERE liv_idx = " + rIdx);
                        }
                    }
                    else
                    {
                        batchSql.Add("INSERT INTO GesLisVendita (liv_lis, liv_art, liv_prv, liv_dti, liv_dtf, liv_day, liv_sta, liv_ann) VALUES (" +
                                     "'" + strLis.Replace("'", "''") + "', '" + artCod.Replace("'", "''") + "', " +
                                     decPrv.ToString(CultureInfo.InvariantCulture) + ", " +
                                     _clsFun.DaySql(dToday) + ", " +
                                     _clsFun.DaySql(_clsDef.DAYOUT) + ", " +
                                     _clsFun.DaySql(DateTime.Today) + ", 'A', 0)");
                    }

                    if (strLis == _clsDef.LISPOS)
                    {
                        batchSql.Add("UPDATE AnaArticoli SET art_prv = " + decPrv.ToString(CultureInfo.InvariantCulture) + " WHERE art_cod = '" + artCod.Replace("'", "''") + "'");
                    }
                }
            }

            if (batchSql.Count > 0)
            {
                return _clsFun.SqlWriteBatch(batchSql, strConSql);
            }
            return false;
        }

        public DataTable LivPromo()
        {
            string s = "";

            s = "SELECT * FROM GesLisVendita WHERE ";
            s += "liv_lis = '" + _clsDef.LISPRO + "' AND ";
            s += "liv_ann = 0 AND ";
            //s += "liv_dti <= " + _clsFun.DaySql(dDay) + " AND ";
            //s += "liv_dtf >= " + _clsFun.DaySql(dDay) + " AND ";
            s += "(liv_sta='' OR liv_sta='" + _clsDef.STAATT + "' OR liv_sta='" + _clsDef.STADAA + "')";
            DataTable t = _clsFun.FillTabSql(TABLISVEN, s, false, _strConSql);

            return t;
        }

        public DataTable OfaSeek(string strArt, DateTime dayDti, DateTime dayDtf)
        {
            string s = "";
            s += "SELECT ";
            s += "GesOffArticoli.ofa_yea, ";
            s += "GesOffArticoli.ofa_cod, ";
            s += "GesOffArticoli.ofa_art, ";
            s += "GesOffTestate.oft_dti, ";
            s += "GesOffTestate.oft_dtf, ";
            s += "GesOffTestate.oft_cod ";
            s += "FROM GesOffArticoli ";
            s += "INNER JOIN GesOffTestate ON GesOffArticoli.ofa_yea = GesOffTestate.oft_yea AND GesOffArticoli.ofa_cod = GesOffTestate.oft_cod ";
            s += "WHERE ";
            s += "(GesOffArticoli.ofa_art = '" + strArt + "') AND (GesOffArticoli.ofa_ann = 0) AND ";
            s += "((GesOffTestate.oft_dti <= " + _clsFun.DaySql(dayDti) + " AND ";
            s += "GesOffTestate.oft_dtf >= " + _clsFun.DaySql(dayDti) + ") OR ";
            s += "(GesOffTestate.oft_dti <= " + _clsFun.DaySql(dayDtf) + " AND ";
            s += "GesOffTestate.oft_dtf >= " + _clsFun.DaySql(dayDtf) + "))";
            DataTable t = _clsFun.FillTabSql(TABGESOFA, s, true, _strConSql);

            return t;
        }

        public DataTable ArtInOff(string strArt, DateTime dayDti, DateTime dayDtf)
        {
            string s = "";
            s += "SELECT ";
            s += "GesOffArticoli.ofa_yea, ";
            s += "GesOffArticoli.ofa_cod, ";
            s += "GesOffArticoli.ofa_art, ";
            s += "GesOffArticoli.ofa_val, ";
            s += "GesOffArticoli.ofa_tip, ";
            s += "GesOffTestate.oft_des, ";
            s += "GesOffTestate.oft_dti, ";
            s += "GesOffTestate.oft_dtf, ";
            s += "GesOffTestate.oft_cod ";
            s += "FROM GesOffArticoli ";
            s += "INNER JOIN GesOffTestate ON GesOffArticoli.ofa_yea = GesOffTestate.oft_yea AND GesOffArticoli.ofa_cod = GesOffTestate.oft_cod ";
            s += "WHERE ";
            s += "GesOffArticoli.ofa_art = '" + strArt + "' AND GesOffArticoli.ofa_ann = 0 AND ";
            s += "GesOffTestate.oft_dti <= " + _clsFun.DaySql(dayDti) + " AND ";
            s += "GesOffTestate.oft_dtf >= " + _clsFun.DaySql(dayDti) + " ";
            DataTable t = _clsFun.FillTabSql(TABGESOFA, s, true, _strConSql);

            return t;
        }

        public DataTable OfsSeek(DateTime dayDay)
        {
            string s = "";
            s = "SELECT * ";
            s += "FROM GesOffArticoli ";
            s += "INNER JOIN GesOffTestate ON GesOffArticoli.ofa_yea = GesOffTestate.oft_yea AND GesOffArticoli.ofa_cod = GesOffTestate.oft_cod ";
            s += "WHERE ";
            s += "GesOffArticoli.ofa_ann = 0 AND ";
            s += "GesOffTestate.oft_dti <= " + _clsFun.DaySql(dayDay) + " AND ";
            s += "GesOffTestate.oft_dtf >= " + _clsFun.DaySql(dayDay) + "";
            DataTable t = _clsFun.FillTabSql(TABGESOFA, s, false, _strConSql);

            return t;
        }

        public DataTable OfsSeekEti(DateTime dayDay)
        {
            //DateTime d = dayDay.AddDays(-180);

            string s = "";
            s = "SELECT * ";
            s += "FROM GesOffArticoli ";
            s += "INNER JOIN GesOffTestate ON GesOffArticoli.ofa_yea = GesOffTestate.oft_yea AND GesOffArticoli.ofa_cod = GesOffTestate.oft_cod ";
            s += "WHERE ";
            s += "GesOffTestate.oft_sta = '" + _clsDef.STANOA + "' OR ";
            s += "GesOffTestate.oft_sta = '" + _clsDef.STADAA + "' OR ";
            s += "GesOffTestate.oft_sta = '" + _clsDef.STAATT + "'";
            //s += "GesOffTestate.oft_dti >= " + _clsFun.DaySql(d) + " ";
            //s += "GesOffTestate.oft_dtf >= " + _clsFun.DaySql(dayDay) + "";
            DataTable t = _clsFun.FillTabSql(TABGESOFA, s, false, _strConSql);

            return t;
        }

        public DataTable StaRep(DateTime dayDti)
        {
            string s = "";
            s += "SELECT ven_neg, ven_day, ven_rep, SUM(ven_qta) AS ven_qta, SUM(ven_qkg) AS ven_qkg, SUM(ven_ven) AS ven_ven, SUM(ven_pun) AS ven_pun ";
            s += "FROM  GesNegVen ";
            s += "WHERE ven_day = " + _clsFun.DaySql(dayDti) + " ";
            s += "GROUP BY ven_neg, ven_day, ven_rep";
            DataTable t = _clsFun.FillTabSql(TABGESOFA, s, false, _strConSqlSta);

            return t;
        }

        public ArrayList MixMatch()
        {
            string s = "SELECT DISTINCT ofa_mix FROM GesOffArticoli WHERE ofa_mix > 0 ORDER BY ofa_mix";

            DateTime dLim = DateTime.Now.AddDays(-30);

            s = "SELECT A.ofa_yea, A.ofa_cod, A.ofa_mix, GesOffTestate.oft_dtf ";
            s += "FROM (SELECT ofa_yea, ofa_cod, ofa_mix ";
            s += "FROM GesOffArticoli ";
            s += "WHERE (ofa_mix > 0) ";
            s += "GROUP BY ofa_yea, ofa_cod, ofa_mix ";
            s += ") AS A LEFT OUTER JOIN ";
            s += "GesOffTestate ON A.ofa_yea = GesOffTestate.oft_yea AND A.ofa_cod = GesOffTestate.oft_cod ";
            s += "WHERE (GesOffTestate.oft_dtf > " + _clsFun.DaySql(dLim) + ") ";
            s += "ORDER BY A.ofa_mix ";

            DataTable t = _clsFun.FillTabSql(TABGESOFA, s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = t.Columns["ofa_mix"];
           // t.PrimaryKey = keys;  modifica per errore su chiavi dupplicate

            s = "SELECT otr_mix FROM GesOffTransazione WHERE otr_mix > 0 ORDER BY otr_mix";
            DataTable tOtr = _clsFun.FillTabSql("GesOffTransazione", s, false, _strConSql);
            if(tOtr.Rows.Count > 0)
            {
                foreach(DataRow y in tOtr.Rows)
                {
                    DataRow x = t.NewRow();
                    x["ofa_mix"] = (decimal)y["otr_mix"];
                    t.Rows.Add(x);
                }
            }

            ArrayList a = new ArrayList();
            DataRow[] j;

            for (int i = 1; i <= 999; i++)
            {
                j = t.Select("ofa_mix=" + i.ToString());
                if (j.Length == 0)
                    a.Add(i);
            }

            return a;
        }

        public DataTable Campagna()
        {
            string s = "SELECT * FROM TabFidCampagne WHERE tab_ann = 0 ORDER BY tab_dti DESC";
            DataTable t = _clsFun.FillTabSql("TabFidCampagne", s, false, _strConSql);
            return t;
        }

        public DataTable TabIva()
        {
            string s = "SELECT * FROM TabIva";
            DataTable t = _clsFun.FillTabSql("TabIva", s, false, _strConSql);
            return t;
        }

        public DataTable AssFor(string strFor, string strArt)
        {
            string s = "SELECT ";
            s += "GesLisAcquisto.lia_arf, ";
            s += "AnaBarcode.ean_ean, ";
            s += "AnaBarcode.ean_dti, ";
            s += "GesLisAcquisto.lia_art, ";
            s += "GesLisAcquisto.lia_dti, ";
            s += "GesLisAcquisto.lia_cos, ";
            s += "GesLisAcquisto.lia_pxc, ";
            s += "GesLisAcquisto.lia_cxp ";
            s += "FROM GesLisAcquisto ";
            s += "LEFT OUTER JOIN AnaBarcode ON AnaBarcode.ean_art = GesLisAcquisto.lia_art ";
            s += "WHERE ";
            s += "(GesLisAcquisto.lia_ann is null OR GesLisAcquisto.lia_ann = 0) AND ";
            s += "GesLisAcquisto.lia_for = N'" + strFor + "' AND  ";
            s += "GesLisAcquisto.lia_art = N'" + strArt + "' ";
            s += "ORDER BY GesLisAcquisto.lia_dti DESC, AnaBarcode.ean_dti DESC";
            DataTable t = _clsFun.FillTabSql("TabAss", s, false, _strConSql);
            return t;
        }

        public DataTable DivTerm()
        {
            DataTable tTmp = new clsGenTabTmp().TabTmpDivTerm("TrmFor");

            string p = "TabTerm";
            string s = "SELECT * FROM TabTerm WHERE tab_ann=0 AND tab_var=1";
            DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
            DataRow x;
            if (t.Rows.Count > 0)
            {
                foreach (DataRow y in t.Rows)
                {
                    x = tTmp.NewRow();
                    x["trm_tip"] = "T";
                    x["trm_cod"] = y["tab_cod"];
                    x["trm_bat"] = y["tab_bat"];
                    x["trm_div"] = y["tab_div"];
                    tTmp.Rows.Add(x);
                }
            }

            //p = "TabForImport";
            s = "SELECT * FROM TabForImport WHERE tab_ann=0";
            t = _clsFun.FillTabSql(TABTABFIM, s, false, _strConSql);
            if (t.Rows.Count > 0)
            {
                foreach (DataRow y in t.Rows)
                {
                    if (!"TMPDIVFTPDIV".Contains((string)y["tab_tip"]))
                    {
                        x = tTmp.NewRow();
                        x["trm_tip"] = "F";
                        x["trm_cod"] = y["tab_cod"];
                        tTmp.Rows.Add(x);
                    }
                }
            }

            return tTmp;
        }

        public DataTable DivArtApShop(string strArt, string strLis)
        {
            string s = "";
            string sArt = strArt.PadLeft(7, Convert.ToChar("0"));
            string sSql = "";

            if (sArt == "0010984")
                Console.WriteLine("aaaa");

            sSql = "SELECT ";
            sSql += "AnaArticoli.art_cod, ";
            sSql += "AnaArticoli.art_des, ";
            sSql += "AnaArticoli.art_deb, ";
            sSql += "AnaArticoli.art_sta, ";
            sSql += "AnaArticoli.art_iva, ";
            sSql += "AnaArticoli.art_umi, ";
            sSql += "AnaArticoli.art_rep, ";
            sSql += "AnaArticoli.art_reb, ";
            sSql += "AnaArticoli.art_plu, ";
            sSql += "AnaArticoli.art_tar, ";
            sSql += "AnaArticoli.art_bpz, ";
            sSql += "AnaArticoli.art_cel, ";
            sSql += "AnaBarcode.ean_ean, ";
            sSql += "AnaBarcode.ean_qta, ";
            sSql += "AnaBarcode.ean_prv, ";
            sSql += "AnaBarcode.ean_ecp, ";
            sSql += "AnaBarcode.ean_bil, ";
            sSql += "AnaBarcode.ean_pun, ";
            sSql += "AnaBarcode.ean_tip, ";
            sSql += "AnaBarcode.ean_dtm, ";
            sSql += "AnaBarcode.ean_ann ";
            sSql += "FROM AnaArticoli LEFT OUTER JOIN AnaBarcode ON AnaArticoli.art_cod = AnaBarcode.ean_art ";
            sSql += "WHERE art_cod='" + strArt + "' ";
            sSql += "ORDER BY AnaBarcode.ean_dti DESC";
            DataTable tArt = _clsFun.FillTabSql(TABANAART, sSql, false, _strConSql);

            DataTable t = new clsGenTabTmp().TabTmpDivArtApShop("ArtCod");

            Boolean b = true;


            int iLotGGscadenza = 0;
            string sLottoEanPref = "";
            if (_strPar031Lotti2Pos != "")
            {
                string[] a = _strPar031Lotti2Pos.Split(',');
                if (a.Length > 2)
                    sLottoEanPref = a[1];
                if (a.Length > 3 && _clsFun.Numerico(a[3].Trim(), "0123456789"))
                    iLotGGscadenza = Convert.ToInt16(a[3]);
            }


            for (int i = 1; i <= 2; i++)
            {
                foreach (DataRow y in tArt.Rows)
                {
                    b = false;

                    s = Convert.ToString(y["ean_ean"]).Trim();

                    if (s == "" && i == 2)
                        s = "0";    //Facciamo passare anche se senza barcode



                    if ((string)y["ean_tip"] == "L" && s.Length == 13 && s.Substring(0, 2) == sLottoEanPref && !DBNull.Value.Equals(y["ean_bil"]) && (Boolean)y["ean_bil"] == false)
                    {
                        b = false;

                        if (!(Boolean)y["ean_ann"] && iLotGGscadenza > 0 && ((DateTime)y["ean_dtm"]).AddDays(iLotGGscadenza) < DateTime.Today)
                        {
                            //s = "UPDATE AnaBarcode SET ean_ann=1, ean_dtm=" + _clsFun.DaySql(DateTime.Today) + " WHERE ean_ean='" + Convert.ToString(y["ean_ean"]).Trim() + "'";
                            //_clsFun.SqlWrite(s, _strConSql);
                            //y["ean_ann"] = true;
                            b = true;
                        }
                        else if (DateTime.Compare((DateTime)y["ean_dtm"], DateTime.Today) == 0)
                            b = true;
                    }


                    if ((string)y["ean_tip"] == "L" && !b && !(Boolean)y["ean_bil"])
                        b = false;
                    else if (s != "" && i == 1 && _clsFun.Numerico(s, true) && Convert.ToInt64(s) > 999999)
                        b = true;
                    else if (i == 2 && _clsFun.Numerico(s, true) && Convert.ToInt64(s) <= 999999)
                        b = true;

                    if (b)
                    {
                        DataRow x = t.NewRow();
                        x["tmp_art"] = y["art_cod"];
                        x["tmp_ard"] = y["art_des"];
                        x["tmp_arb"] = y["art_deb"];
                        x["tmp_sta"] = y["art_sta"];

                        x["tmp_iva"] = y["art_iva"];
                        x["tmp_umi"] = y["art_umi"];
                        x["tmp_rep"] = y["art_rep"];
                        x["tmp_reb"] = y["art_reb"];
                        x["tmp_plu"] = y["art_plu"];
                        x["tmp_tar"] = y["art_tar"];
                        x["tmp_bpz"] = y["art_bpz"];
                        x["tmp_cel"] = y["art_cel"];

                        x["tmp_ean"] = y["ean_ean"];
                        x["tmp_eaq"] = y["ean_qta"];
                        x["tmp_eap"] = y["ean_prv"];
                        x["tmp_ecp"] = y["ean_ecp"];
                        x["tmp_eaa"] = y["ean_ann"];
                        x["tmp_eab"] = y["ean_bil"];
                        x["tmp_epu"] = y["ean_pun"];
                        t.Rows.Add(x);
                    }
                }
            }
            if (tArt != null)
            {
                tArt.Dispose();
                tArt = null;
            }

            if (t.Rows.Count > 0)
            {
                t = ArtPrezzo(t, strLis);
                //t = ArtCosto("", "", t, DateTime.Today);
            }

            return t;
        }

        public string DivFilArtApShop(int intPos, string strFilVar)
        {
            string s = "";
            Boolean b = true;
            string sFil = "";
            string sPath = Path.GetDirectoryName(strFilVar) + "\\";
            string sFile = Path.GetFileNameWithoutExtension(strFilVar);
            foreach (string sFi in Directory.GetFiles(sPath))
            {
                s = Path.GetFileNameWithoutExtension(sFi);
                if (s.Length > 6)
                {
                    if (s.Substring(0, 6) == sFile && s.Length == 10)
                    {
                        b = true;
                        for (int n = 1; n <= intPos; n++)
                        {
                            if (File.Exists(sFi + "_" + n.ToString("0") + ".csv"))
                            {
                                b = false;
                                break;
                            }
                        }
                        if (b)
                        {
                            string sOld = Path.GetDirectoryName(sFi) + "\\Old\\" + Path.GetFileName(sFi) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                            if (File.Exists(sOld))
                                File.Delete(sOld);
                            File.Move(sFi, sOld);
                        }
                    }
                }
            }

            for (int n = 1; n < 1000; n++)
            {
                sFil = sPath + sFile + "_" + n.ToString("000");
                if (!File.Exists(sFil))
                    break;
            }
            return sFil;
        }

        public DataTable tabEcr()
        {
            string p = "TabEcr";
            string s = "";
            s += "SELECT ";
            s += "TabEcrLv1.tab_cod AS l1c, ";
            s += "TabEcrLv1.tab_des AS l1d, ";
            s += "TabEcrLv2.tab_cod AS l2c, ";
            s += "TabEcrLv2.tab_des AS l2d, ";
            s += "TabEcrLv3.tab_cod AS l3c, ";
            s += "TabEcrLv3.tab_des AS l3d ";
            s += "FROM TabEcrLv2 ";
            s += "INNER JOIN TabEcrLv1 ON TabEcrLv2.tab_lv1 = TabEcrLv1.tab_cod ";
            s += "INNER JOIN TabEcrLv3 ON TabEcrLv2.tab_cod = TabEcrLv3.tab_lv2 AND TabEcrLv2.tab_lv1 = TabEcrLv3.tab_lv1";
            DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);

            return t;
        }

        public string NewPlu(string strReb)
        {
            string s = "";

            try
            {
                DataRow[] j;
                //int i = 0;
                s = "SELECT art_plu ";
                s += "FROM AnaArticoli ";
                s += "WHERE art_reb='" + strReb + "' AND art_plu <>'' ";
                s += "ORDER BY art_plu DESC";

                s = "SELECT CONVERT(int, art_plu) AS art_plu ";
                s += "FROM AnaArticoli ";
                s += "WHERE art_reb='" + strReb + "' AND art_plu <>'' ";
                s += "ORDER BY CONVERT(int, art_plu) DESC";

                DataTable t = _clsFun.FillTabSql("ArtPlu", s, false, _strConSql);
                DataColumn[] keys = new DataColumn[1];
                keys[0] = t.Columns["art_plu"];
                t.PrimaryKey = keys;

                s = "1";

                if (t.Rows.Count > 0)
                {
                    //i = Convert.ToInt16(t.Rows[0]["art_plu"]);
                    for (int n = 1; n <= 9999; n++)
                    {
                        //s = n.ToString("0000");
                        //j = t.Select("VAL(art_plu)='" + s + "'");
                        j = t.Select("art_plu=" + n.ToString() + "");
                        if (j.Length == 0)
                        {
                            //break;
                            s = n.ToString();
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                s = "NON UNIVOCI";

                Console.WriteLine(ex.Message);

            }

            return s;
        }

        public DataTable PluExist(string strReb, string strPlu, string strArt)
        {
            //DataRow[] j;

            string s = "SELECT art_plu, art_cod, art_des ";
            s += "FROM AnaArticoli ";
            s += "WHERE art_reb='" + strReb + "' AND art_plu='" + strPlu + "' AND art_cod <>'" + strArt + "' ";
            //s += "ORDER BY art_plu DESC";
            DataTable t = _clsFun.FillTabSql("ArtPlu", s, false, _strConSql);

            return t;
        }

        public DataTable SeekArtEanCosMin(string[] arsFor, DataRow rowOrd, DataTable tabMsg)
        {
            string sArt = ("" + rowOrd["ord_art"]);
            string sEan = ("" + rowOrd["ord_ean"]);

            string s = "";

            s = "SELECT ";
            s += "AnaBarcode.ean_ean, ";
            s += "GesLisAcquisto.lia_art, ";
            s += "GesLisAcquisto.lia_for, ";
            s += "GesLisAcquisto.lia_dti, ";
            s += "GesLisAcquisto.lia_dtf, ";
            s += "GesLisAcquisto.lia_arf, ";
            s += "GesLisAcquisto.lia_cos, ";
            s += "GesLisAcquisto.lia_pxc, ";
            s += "GesLisAcquisto.lia_cxp, ";
            s += "AnaArticoli.art_des, ";
            s += "AnaArticoli.art_umi ";
            s += "FROM AnaBarcode ";
            s += "LEFT OUTER JOIN GesLisAcquisto ON AnaBarcode.ean_art = GesLisAcquisto.lia_art ";
            s += "LEFT OUTER JOIN AnaArticoli ON AnaBarcode.ean_art = AnaArticoli.art_cod ";
            s += "WHERE ";
            if(sEan != "")
                s +="ean_ean='" + sEan + "' AND ";
            else if(sArt != "")
                s += "ean_art='" + sArt + "' AND ";

            if ((string)rowOrd["ord_for"] != "")
            {
                string s2 = "";

                foreach (string sa in arsFor)
                {
                    if (sa == (string)rowOrd["ord_for"])
                        s2 = (string)rowOrd["ord_for"];
                }

                if(s2 != "")
                    s += "lia_for='" + s2 + "' ";
                else
                    s += "lia_for='" + "ZZZZZZZZZZZZZZ" + "' ";
            }
            else
            {
                s += "(";

                foreach (string sa in arsFor)
                {
                    s += "lia_for='" + sa + "' OR ";
                }

                s = s.Substring(0, s.Length - 4);

                s += ") ";
            }

            s += "ORDER BY lia_dti DESC";

            DataTable t = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);

            string sFor = "";

            if(t.Rows.Count > 0)
            {
                decimal dCos = 0;
                ArrayList aFor = new ArrayList();

                dCos = (decimal)t.Rows[0]["lia_cos"];
                sFor = (string)t.Rows[0]["lia_for"];
                aFor.Add(sFor);

                foreach (DataRow y in t.Rows)
                {
                    if (aFor.IndexOf((string)y["lia_for"]) < 0)
                    {
                        if ((decimal)y["lia_cos"] < dCos)
                        {
                            sFor = (string)y["lia_for"];
                        }
                        aFor.Add((string)y["lia_for"]);

                    }
                }
            }
            else
            {
                string sDes = "";
                s = "";

                if(sEan != "")
                    s = sEan;
                else if(sArt != "")
                    s = sArt;
                if(s != "")
                {
                    DataTable tTmp = ArtSeek(s, "");
                    if (tTmp.Rows.Count > 0)
                        sDes = (string)tTmp.Rows[0]["tmp_art"] + " " + (string)tTmp.Rows[0]["tmp_ard"];
                    if (sDes.Length > 50)
                        sDes = sDes.Substring(0, 50);
                }

                DataRow x = tabMsg.NewRow();
                if (sEan != "")
                {
                    x["tmp_ean"] = sEan;
                    x["tmp_msg"] = sDes + " barcode non trovato assortimento fornitori";
                }
                else if (sArt != "")
                {
                    x["tmp_ean"] = sArt;
                    x["tmp_msg"] = sDes + " articolo non trovato assortimento fornitori";
                }
                tabMsg.Rows.Add(x);
            }

            DataTable tLia = new DataView(t, "lia_for='" + sFor + "'", "", DataViewRowState.CurrentRows).ToTable();

            return tLia;
        }

        public string DefCnfPar(string strTip)
        {
            string sPar = "";
            string s = "";

            if (strTip == "POS")
            {
                s = _clsFun.ParGet(clsDefine.enuParametri.ParDivPrezziAZero, _strConSql);
                if (s == "S")
                    sPar += s;
                else
                    sPar += "_";
            }

            return sPar;
        }

        public DataTable ArtFornitori(string strArt)
        {
            string s = "SELECT lia_for FROM " + TABLISACQ + " WHERE lia_art='" + strArt + "' ";
            s += "GROUP BY lia_for ";

            s = "SELECT AnaFornitori.for_des ";
            s += "FROM GesLisAcquisto ";
            s += "INNER JOIN AnaFornitori ON GesLisAcquisto.lia_for = AnaFornitori.for_cod ";
            s += "WHERE (GesLisAcquisto.lia_art = '" + strArt + "')";
            s += "GROUP BY AnaFornitori.for_des ";
            s += "ORDER BY AnaFornitori.for_des ";

            DataTable t = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);

            return t;
        }

        public DataTable TessSeek(string strCod)
        {
            string s = "SELECT * FROM " + TABANATES + " WHERE tes_cod='" + strCod + "' ";
            DataTable t = _clsFun.FillTabSql(TABANATES, s, false, _strConSql);

            return t;
        }

        public String ApPhoneDiv2TmpDiv(string strDivPath, ArrayList aryPre)
        {
            string[] sFils = Directory.GetFiles(strDivPath);
            string s = "";
            string sRig = "";
            string sFi = "";
            int i = 0;

            foreach (string sFil in sFils)
            {
                foreach (string sPar in aryPre)
                {
                    i++;

                    sFi = Path.GetFileName(sFil);

                    if(sFi.Substring(0,4) == sPar)
                    {
                        using (StreamReader sr = new StreamReader(sFil))
                        {
                            while ((sRig = sr.ReadLine()) != null)
                            {
                                s = "INSERT INTO TmpDiv (tmp_tip, tmp_fil, tmp_tmp, tmp_sta, tmp_day) VALUES (";
                                s += "'" + sFi.Substring(0,3).ToUpper() + "',";
                                s += "'" + sFi + "',";
                                s += "'" + sRig + "',";
                                s += "'S',";
                                s += "'" + DateTime.Now.ToString("yyyyMMddHHmmss") + "'";
                                s += ")";
                                _clsFun.SqlWrite(s, _strConSql);
                            }
                        }
                    }

 
                }

                s = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil) + DateTime.Today.ToString("yyyyMMddHHmmss");
                if (File.Exists(s))
                    File.Delete(s);
                File.Move(sFil, s);
            }

            return "NUM,"+i.ToString();
        }

        public void ArtAttiva(string strArt)
        {
            string s = "UPDATE AnaArticoli SET art_sta='" + _clsDef.STAATT + "' WHERE art_cod='" +  strArt + "'";
            _clsFun.SqlWrite(s, _strConSql);
        }

        public string SmfCtrl(string strTip)
        {
            string s = "";
            string sRes = "";

            string sTerm = _clsFun.FileIni("R", clsDefine.enuIni.Ini10NumeroTerminale, "");
            string sPath = _clsFun.FileIni("R", clsDefine.enuIni.Ini07PathDivCasse, "");

            if (sPath != "")
            {

                int i = sPath.IndexOf("ApProject", StringComparison.OrdinalIgnoreCase);
                if (i < 0)
                    i = sPath.IndexOf("APproject", StringComparison.OrdinalIgnoreCase);

                if (i > 0)
                    sPath = sPath.Substring(0, i);
                else
                    sPath = Path.GetPathRoot(sPath);

                //_clsFun.ErrorLog(sPath, i.ToString());

                string sFil = _clsDef.SMFVARPOS + "_" + sTerm + ".smf";

                //_clsFun.ErrorLog(sPath, sFil);

                s = Path.GetDirectoryName(_clsDef.SMFVARPOS);

                // Remove drive letter / root from SMFVARPOS-based path to avoid duplicating it
                string sSmfRoot = Path.GetPathRoot(sFil);
                if (!string.IsNullOrEmpty(sSmfRoot))
                    sFil = sFil.Substring(sSmfRoot.Length);

                sFil = Path.Combine(sPath, sFil);

                //_clsFun.ErrorLog(sPath, sFil);

                if (sFil != "")
                {
                    if (strTip == "ENTRA")
                    {
                        if (!File.Exists(sFil))
                        {
                            var f = File.Create(sFil);
                            f.Close();
                        }

                        s = Path.GetFileName(_clsDef.SMFVARPOS) + "*.smf";
                        sPath = Path.GetDirectoryName(sFil) + "\\";
                        string[] sFils = Directory.GetFiles(Path.GetDirectoryName(sPath), s);
                        foreach (string sFi in sFils)
                        {
                            s = Path.GetFileName(sFi);

                            string ss = s.Substring(s.IndexOf("_") + 1, 3);

                            if (ss != sTerm)
                            {
                                sRes = ss;
                                break;
                            }
                        }
                    }
                    else
                    {
                        //Cancella
                        if (File.Exists(sFil))
                            File.Delete(sFil);
                    }
                }

                s = sFil + "|" + sRes;
            }

            return s;
        }

        public void LogSql(string[] strLog)
        {
            if (strLog == null || strLog.Length < 8) return;
            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    string day = DateTime.Today.ToString("yyyy-MM-dd");
                    string ora = DateTime.Now.ToString("HHmm");
                    string tip = (strLog[0] ?? "").Replace("'", "''");
                    string art = (strLog[1] ?? "").Replace("'", "''");
                    string ean = (strLog[2] ?? "").Replace("'", "''");
                    string des = (strLog[3] ?? "").Replace("'", "''");
                    string fil = (strLog[4] ?? "").Replace("'", "''");
                    decimal acq = 0; if (!string.IsNullOrEmpty(strLog[5])) try { acq = Convert.ToDecimal(strLog[5]); } catch { }
                    decimal prv = 0; if (!string.IsNullOrEmpty(strLog[6])) try { prv = Convert.ToDecimal(strLog[6]); } catch { }
                    string msg = (strLog[7] ?? "").Replace("'", "''");

                    string sql = $"INSERT INTO LogLog (log_day, log_ora, log_tip, log_art, log_ean, log_des, log_fil, log_acq, log_prv, log_msg) " +
                                 $"VALUES ('{day}', '{ora}', '{tip}', '{art}', '{ean}', '{des}', '{fil}', {acq.ToString(System.Globalization.CultureInfo.InvariantCulture)}, {prv.ToString(System.Globalization.CultureInfo.InvariantCulture)}, '{msg}')";

                    _clsFun.SqlWrite(sql, _strConSqlLog);
                }
                catch (Exception ex)
                {
                    _clsFun.ErrorLog(ex.Message, "LogSql");
                }
            });
        }

        public void LogSqlBulk(List<string[]> listLog)
        {
            if (listLog == null || listLog.Count == 0) return;
            try
            {
                string p = "LogLog";
                string s = "SELECT * FROM LogLog WHERE 1=0";
                DataTable t = _clsFun.FillTabSql(p, s, false, _strConSqlLog);

                foreach (string[] strLog in listLog)
                {
                    DataRow x = t.NewRow();
                    x["log_day"] = DateTime.Today;
                    x["log_ora"] = DateTime.Now.ToString("HHmm");
                    x["log_tip"] = strLog[0];
                    x["log_art"] = strLog[1];
                    x["log_ean"] = strLog[2];
                    x["log_des"] = strLog[3];
                    x["log_fil"] = strLog[4];

                    x["log_acq"] = 0;
                    if (strLog.Length > 5 && strLog[5] != "")
                        x["log_acq"] = _clsFun.Txt2Dec(strLog[5]);

                    x["log_prv"] = 0;
                    if (strLog.Length > 6 && strLog[6] != "")
                        x["log_prv"] = _clsFun.Txt2Dec(strLog[6]);

                    x["log_msg"] = strLog.Length > 7 ? strLog[7] : "";

                    t.Rows.Add(x);
                }

                using (SqlConnection cnLog = new SqlConnection(_strConSqlLog))
                {
                    cnLog.Open();
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(cnLog))
                    {
                        bulkCopy.DestinationTableName = "LogLog";
                        bulkCopy.WriteToServer(t);
                    }
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "LogSqlBulk");
            }
        }

        public RichTextBox sql2Rtf(string strArt)
        {
            RichTextBox rtxBox = new RichTextBox();

            SqlConnection cn = new SqlConnection(_strConSql);
            cn.Open();
            SqlCommand cmd = new SqlCommand("SELECT ing_img FROM AnaArtIngredienti2 WHERE ing_art='" + strArt + "'", cn);
            SqlDataReader rdr = cmd.ExecuteReader();
            rdr.Read();
            if (rdr.HasRows)
            {
                if (!rdr.IsDBNull(0))
                {
                    Byte[] rtf = new Byte[Convert.ToInt32((rdr.GetBytes(0, 0, null, 0, Int32.MaxValue)))];
                    long bytesReceived = rdr.GetBytes(0, 0, rtf, 0, rtf.Length);

                    ASCIIEncoding encoding = new ASCIIEncoding();
                    rtxBox.Rtf = encoding.GetString(rtf, 0, Convert.ToInt32(bytesReceived));
                }
            }

            return rtxBox;
        }

        public string GetFileIngredienti(String strArt)
        {
            string sIng = "";
            String s = _clsDef.PATHINGREDIENTI + "et01_" + strArt + ".rtf";
            if (File.Exists(s))
            {
                RichTextBox rch = new RichTextBox();
                rch.LoadFile(s);
                sIng = rch.Text;
                //sIng = _clsFun.StrSplit(sIng, 40);
            }

            return sIng;
        }

        public string ParForm(Form frm, string strTip, string strPar)
        {
            string p = "TabForm";
            string s = "SELECT * FROM TabForm";
            DataRow x;
            DataRow[] j;
            string sRes = "";

            DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);

            if (strTip == "W")
            {
                ArrayList aWhe = new ArrayList();
                aWhe.Add("tab_cod");
                ArrayList aExl = new ArrayList();

                x = t.NewRow();
                x["tab_cod"] = frm.Name;
                x["tab_szw"] = frm.Size.Width;
                x["tab_szh"] = frm.Size.Height;
                x["tab_par"] = strPar;

                j = t.Select("tab_cod='" + frm.Name + "'");
                if (j.Length > 0)
                    s = _clsFun.SqlUpdRow(p, t, j[0], x, aWhe, aExl);
                else
                    s = _clsFun.SqlInsertRow(p, t, x);

                if (s != "")
                    _clsFun.SqlWrite(s, _strConSql);
            }
            else
            {
                j = t.Select("tab_cod ='" + frm.Name + "'");
                if (j.Length > 0 && (int)j[0]["tab_szw"] > 100 && (int)j[0]["tab_szh"] > 100)
                {
                    frm.Size = new System.Drawing.Size((int)j[0]["tab_szw"], (int)j[0]["tab_szh"]);
                    if (!DBNull.Value.Equals(j[0]["tab_par"]))
                        sRes = (string)j[0]["tab_par"];
                }
            }

            return sRes;
        }

        //public string TogliereNewLotto(string strYea, string strTip)
        //{
        //    string s = "SELECT * FROM GesDocLotti WHERE lot_yea='" + strYea + "' AND lot_tip='" + strTip + "'";
        //    return "";
        //}

        public string ArtLotto(string strArt)
        {
            /*Se nella sottofamiglia dell'articolo è previsto un tipo prendo l'ultimo lotto inserito del 'tipo' altrimenti non restituisco niente è viene definito un lotto con la settimana*/

            string sLot = "";
            string s = "";
            s = "SELECT AnaArticoli.art_cod, AnaArticoli.art_des, AnaArticoli.art_ec3, TabEcrLv3.tab_tip ";
            s += "FROM AnaArticoli INNER JOIN TabEcrLv3 ON AnaArticoli.art_ec1 = TabEcrLv3.tab_lv1 AND AnaArticoli.art_ec2 = TabEcrLv3.tab_lv2 AND AnaArticoli.art_ec3 = TabEcrLv3.tab_cod ";
            s += "WHERE (AnaArticoli.art_cod = '" + strArt + "')";
            DataTable t = _clsFun.FillTabSql("ART", s, false, _strConSql);
            if(t.Rows.Count > 0)
            {
                s = (string)t.Rows[0]["tab_tip"];

                if (s != "")
                {
                    s = "SELECT TOP(5) lot_tip, lot_cod, lot_yea FROM GesDocLotti WHERE lot_tip='" + s + "' ORDER BY lot_ddo DESC";
                    t = _clsFun.FillTabSql("LOT", s, false, _strConSql);
                    if(t.Rows.Count > 0)
                    {
                        //sYea = (string)t.Rows[0]["lot_yea"];
                        sLot = ((string)t.Rows[0]["lot_tip"]).Trim() + "-" + (string)t.Rows[0]["lot_cod"] + "-" + (string)t.Rows[0]["lot_yea"];
                    }
                }
            }

            if (sLot == "")
                sLot = DateLotto();

            return sLot;
        }

        private string DateLotto()
        {
            string s = DateTime.Today.ToString("ddd").Substring(0, 2).ToUpper();

            var d = DateTime.Today;
            CultureInfo cul = CultureInfo.CurrentCulture;

            var firstDayWeek = cul.Calendar.GetWeekOfYear(
                d,
                CalendarWeekRule.FirstDay,
                DayOfWeek.Monday);

            int weekNum = cul.Calendar.GetWeekOfYear(
                d,
                CalendarWeekRule.FirstDay,
                DayOfWeek.Monday);
            weekNum -= 1;
            int year = weekNum == 52 && d.Month == 1 ? d.Year - 1 : d.Year;
            //Console.WriteLine("Year: {0} Week: {1}", year, weekNum);

            s += weekNum.ToString("00") + "-" + d.Year.ToString();

            return s;
        }

        public DataTable LotCarico(DataRow rowLot)
        {
            string s = "";

            if ((string)rowLot["lot_dot"] == "M")
            {
                s = "SELECT ";
                s += "GesMovTestate.mot_ddo AS DocDdo, ";
                s += "GesMovTestate.mot_neg AS DocNeg, ";
                s += "GesMovimenti.mov_qta, ";
                s += "GesMovimenti.mov_qkg ";
                s += "FROM GesMovTestate LEFT OUTER JOIN ";
                s += "GesMovimenti ON GesMovTestate.mot_ymo = GesMovimenti.mov_ymo AND GesMovTestate.mot_nmo = GesMovimenti.mov_nmo ";
                //s += "LEFT OUTER JOIN ";
                s += "WHERE ";
                s += "mov_ymo='" + rowLot["lot_doy"] + "' AND ";
                s += "mov_nmo='" + rowLot["lot_dom"] + "' AND ";
                s += "mov_rmo='" + rowLot["lot_dor"] + "'";
            }
            else
            {
                s = "SELECT ";
                s += "GesFatTestate.fat_ddo AS DocDdo, ";
                s += "GesFatTestate.fat_neg AS DocNeg, ";
                s += "GesMovimenti.mov_qta, ";
                s += "GesMovimenti.mov_qkg ";
                s += "FROM GesMovimenti ";
                s += "INNER JOIN GesFatTestate ON GesMovimenti.mov_yfa = GesFatTestate.fat_yfa AND GesMovimenti.mov_nfa = GesFatTestate.fat_nfa ";
                s += "WHERE ";
                s += "mov_yfa='" + rowLot["lot_doy"] + "' AND ";
                s += "mov_nfa='" + rowLot["lot_dom"] + "' AND ";
                s += "mov_rfa='" + rowLot["lot_dor"] + "'";
            }

            DataTable tTmp = _clsFun.FillTabSql("TabDoc", s, false, _strConSql);

            return tTmp;
        }

        public decimal LotVenduto(DataTable tabMov, string strElo)
        {
            decimal d = 0;
            string sNeg = (string)tabMov.Rows[0]["DocNeg"];
            DateTime dIni = (DateTime)tabMov.Rows[0]["DocDdo"];
            if (sNeg == "")
                sNeg = "001";

            string s = "SELECT ven_art, ven_qta, ven_qkg FROM GesNegVen WHERE ";
            s += "ven_neg = '" + sNeg + "' AND ";
            //s += "ven_cau = '" + sCau + "' AND ";
            s += "ven_day >= " + _clsFun.DaySql(dIni) + "  AND ";
            s += "LEFT(ven_ean, 5) = '" + strElo + "'";

            DataTable t = _clsFun.FillTabSql("TabDoc", s, false, _strConSqlSta);

            foreach(DataRow y in t.Rows)
            {
                d += (decimal)y["ven_qkg"];
            }

            string sLot = "00" + strElo.Substring(2);

            s = "SELECT ";
            s += "GesMovTestate.mot_ddo AS DocDdo, ";
            s += "GesMovTestate.mot_neg AS DocNeg, ";
            s += "GesMovimenti.mov_qta, ";
            s += "GesMovimenti.mov_qkg ";
            s += "FROM GesMovTestate ";
            s += "LEFT OUTER JOIN GesMovimenti ON GesMovTestate.mot_ymo = GesMovimenti.mov_ymo AND GesMovTestate.mot_nmo = GesMovimenti.mov_nmo ";
            s += "LEFT OUTER JOIN TabMovCausali ON GesMovTestate.mot_cau = TabMovCausali.tab_cod ";
            s += "WHERE ";
            s += "tab_cfo <> '" + _clsDef.TIPFOR + "'  AND ";
            s += "mot_neg ='" + sNeg + "'  AND ";
            s += "mot_neg ='" + sNeg + "'  AND ";
            s += "mov_day >=" + _clsFun.DaySql(dIni) + "  AND ";
            s += "mov_lot='" + sLot + "' ";

            t = _clsFun.FillTabSql("TabMov", s, false, _strConSql);

            foreach (DataRow y in t.Rows)
            {
                d += (decimal)y["mov_qkg"];
            }

            return d;
        }

        public string Lotto2Ean(string strLot, string strArt, string strPes)
        {
            string s = _clsFun.ParGet(clsDefine.enuParametri.Par031Lotti2Pos, _strConSql);
            string sPrf = "";
            string[] a = s.Split(',');
            if (a.Length > 1 && a[1] != "")
                sPrf = a[1];
            //string sLot = strLot;
            //string sLeg = "";
            //s = "SELECT lot_art TOP(5) lot_tip, lot_cod, lot_yea FROM GesDocLotti WHERE lot_tip='" + s + "' ORDER BY lot_ddo DESC";
            //DataTable t = _clsFun.FillTabSql("GesDocLotti", s, false, _strConSql);
            //string sEan = sPrf + sLot + sLeg + new string('0', 6);
            //s = "SELECT * FROM AnaBarcode WHERE ean_ean='" + sEan + "'";
            //DataTable t = _clsFun.FillTabSql("AnaBarcode", s, true, _strConSql);
            //if (t.Rows.Count > 0)
            //{
            //}

            string sLot = "";
            a = strLot.Split('-');
            if (a.Length > 1 && a[1] != "" && a[1].Length == 5)
                sLot = a[1].Substring(2);

            string sEan = "";

            s = "SELECT ean_ean FROM AnaBarcode WHERE ean_art='" + strArt + "' ORDER BY ean_dtm DESC";
            DataTable t = _clsFun.FillTabSql("GesDocLotti", s, false, _strConSql);
            foreach(DataRow y in t.Rows)
            {
                s = ((string)y["ean_ean"]).Trim();

                if(s.Length == 13 && s.Substring(0,2) == sPrf && s.Substring(2,3) == sLot)
                {
                    decimal dPes = 0;
                    if (_clsFun.Numerico(strPes, "0123456789,."))
                        dPes = Convert.ToDecimal(strPes.Replace(",", "."));

                    string sPes = "00000";
                    if (dPes > 0)
                        //sPes = (Convert.ToInt64(strPes) / 10).ToString("00000");
                        sPes = Convert.ToInt64(dPes / 10).ToString("00000");

                        s = s.Substring(0, 7) + sPes;
                        s = s + new clsCtrlCodici().FindMod10Digit(s);

                        sEan = s;

                    break;
                }
            }

            return sEan;
        }

        public string GenLotto2Ean(string strLot, string strArt, string strPes)
        {
            string s = _clsFun.ParGet(clsDefine.enuParametri.Par031Lotti2Pos, _strConSql);
            string sPrf = "";
            string sEanPes = "";
            string sLot = "";
            string[] a = s.Split(',');
            if (a.Length > 1 && a[1] != "")
            {
                sPrf = a[1];
                if (a.Length > 4 && a[4] != "")
                    sEanPes = a[4];
            }

            if (strLot.Length == 5)
                sLot = strLot.Substring(2);
            else
            {
                a = strLot.Split('-');
                sLot = a[1].Substring(2);
            }

            DateTime dDay = DateTime.Today.AddDays(-30);

            string sEan = "";
            string sCod = "00";
            s = "SELECT * FROM AnaBarcode WHERE LEFT(ean_ean, 5) ='" + sPrf + sLot + "' AND ean_dti > " + _clsFun.DaySql(dDay) + " ORDER BY ean_ean";
            DataTable tTmp = _clsFun.FillTabSql(TABANAEAN, s, false, _strConSql);
            foreach(DataRow y in tTmp.Rows)
            {
                s = ((string)y["ean_ean"]).Trim();
                sCod = s.Substring(5, 2);

                if ((string)y["ean_art"] == strArt)
                {
                    sEan = ((string)y["ean_ean"]).Trim();
                    break;
                }
            }

            if (sEan == "")
            {
                sCod = (Convert.ToInt16(sCod) + 1).ToString("00");
                sEan = sPrf + sLot + sCod + new string('0', 6);
            }

            s = "SELECT * FROM AnaBarcode WHERE ean_ean='" + sEan + "'";
            tTmp = _clsFun.FillTabSql("AnaBarcode", s, true, _strConSql);
            if (tTmp.Rows.Count > 0)
            {
                DataRow x = tTmp.NewRow();

                foreach (DataColumn c in tTmp.Columns)
                    x[c.ColumnName] = tTmp.Rows[0][c.ColumnName];

                x["ean_art"] = strArt;
                x["ean_tip"] = "L";             //Lotto
                x["ean_dti"] = DateTime.Today;
                x["ean_dtm"] = DateTime.Today;
                x["ean_bil"] = true;
                x["ean_ann"] = false;
                if (sEanPes == "EANPESO")
                    x["ean_ecp"] = true;

                s = _clsFun.SqlUpdRowIdx("AnaBarcode", tTmp, tTmp.Rows[0], x, null);
            }
            else
            {
                DataRow x = tTmp.NewRow();
                x["ean_art"] = strArt;
                x["ean_ean"] = sEan;
                x["ean_ann"] = false;
                x["ean_qta"] = 0;
                x["ean_prv"] = 0;
                x["ean_dti"] = DateTime.Today;
                x["ean_dtm"] = DateTime.Today;
                x["ean_bil"] = true;
                x["ean_ecp"] = false;

                if (sEanPes == "EANPESO")
                    x["ean_ecp"] = true;

                x["ean_pun"] = 0;
                x["ean_tip"] = "L";

                s = _clsFun.SqlInsertRow("AnaBarcode", tTmp, x);
            }

            _clsFun.FileLog("LOTTO EAN", "", s);

            if (s != "")
                _clsFun.SqlWrite(s, _strConSql);

            sEan = sEan.Substring(0, 7) + strPes.PadLeft(5,Convert.ToChar('0')) + new clsCtrlCodici().FindMod10Digit(sEan);

            return sEan;
        }

        public DataTable DdtStato(string strYfa, string strNfa, string strSta)
        {
            decimal d = 0;

            string s = "";
            s = "SELECT mot_nmo, mot_sta, mot_ndo, mot_ddo FROM GesMovTestate WHERE mot_yfa='" + strYfa + "' AND mot_nfa='" + strNfa + "' AND mot_sta='" + strSta + "'";
            DataTable t = _clsFun.FillTabSql("MOTT", s, false, _strConSql);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "TmpImp",
                Caption = "Importo",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            return t;
        }

        public string GetDatiLotto(string strParLot, string strEan)
        {
            string s = "";
            string sRes = "";
            string[] a = strParLot.Split(',');
            if(a.Length > 2 && a[2] == "S" && strEan.Substring(0,2) == a[1])
            {
                string sLot = "00" + strEan.Substring(2, 3);     //Codice lotto MASSIMO 3 CIFRE

                s = "SELECT * FROM GesDocLotti WHERE lot_cod='" + sLot + "' ORDER BY lot_ddo DESC";
                DataTable t = _clsFun.FillTabSql("GesDocLotti", s, true, _strConSql);
                if (t.Rows.Count > 0)
                {
                    string sYea = (string)t.Rows[0]["lot_yea"];

                    string sNat = "";
                    if (!DBNull.Value.Equals(t.Rows[0]["lot_nat"]))
                        sNat = ((string)t.Rows[0]["lot_nat"]).Trim();

                    string sLof = "";       //Lotto del fornitore
                    if (!DBNull.Value.Equals(t.Rows[0]["lot_lot"]))
                        sLof = ((string)t.Rows[0]["lot_lot"]).Trim();

                    string sAll = "";
                    if (!DBNull.Value.Equals(t.Rows[0]["lot_all"]))
                        sAll = ((string)t.Rows[0]["lot_all"]).Trim();

                    //Macellato
                    string sMac = "";
                    if (!DBNull.Value.Equals(t.Rows[0]["lot_mna"]) && !DBNull.Value.Equals(t.Rows[0]["lot_mac"]))
                        sMac = ((string)t.Rows[0]["lot_mna"] + new string(' ', 10)).Substring(0, 10) + " " + ((string)t.Rows[0]["lot_mac"]).Trim() + "|";

                    //Sezionato
                    string sSez = "";
                    if (!DBNull.Value.Equals(t.Rows[0]["lot_sez"]) && !DBNull.Value.Equals(t.Rows[0]["lot_sbo"]))
                        sSez = ((string)t.Rows[0]["lot_sez"] + new string(' ', 10)).Substring(0, 10) + " " + ((string)t.Rows[0]["lot_sbo"]).Trim() + "|";

                    //s = "Lotto:         " + sLot + "-" + sYea + "|";
                    s = "Lotto:         " + sLof + "-" + sYea + "|";
                    sRes += s; // +_clsDef.CRLF;

                    if (sNat != "")
                    {
                        s = "Nato in:       " + sNat + "|";
                        sRes += s; // +_clsDef.CRLF;
                    }
                    if (sAll != "")
                    {
                        s = "Allevato in :  " + sAll + "|";
                        sRes += s; // +_clsDef.CRLF;
                    }
                    if (sMac != "")
                    {
                        s = "Macellato in : " + sMac;
                        sRes += s; // +_clsDef.CRLF;

                        //if (sSez == "")
                        //{
                        //    sSez = ((string)tCnf.Rows[0]["cnf_v01"]).Trim();
                        //    int pos = sSez.IndexOf(' ');
                        //    sSez = sSez.Substring(pos);
                        //}

                        s = "Sezionato in:  " + sSez;
                        sRes += s; // +_clsDef.CRLF;
                    }
                }
            }

            return sRes;
        }

    }
}


    
