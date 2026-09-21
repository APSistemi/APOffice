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

namespace APOffice
{
    class clsVariazioni
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        //private const string TABANAART = "AnaArticoli";
        private const string TABANAEAN = "AnaBarcode";
        //private const string TABLISACQ = "GesLisAcquisto";
        //private const string TABLISVEN = "GesLisVendita";
        //private const string TABGESOFA = "GesOffArticoli";

        private const string TABVARPOS = "GesVarPos";
        private const string TABVARBIL = "GesVarBilance";
        private const string TABVARETI = "GesVarEtichette";
        private const string TABGESVAR = "GesVariazioni";

        private string _strConSql = "";

        public clsVariazioni()
        {
            _strConSql = _clsFun.ConSql("");
        }

        public void Variazioni(string strArt, string strSql, string strOri, string strTip)
        {
            if (string.IsNullOrEmpty(strArt) || string.IsNullOrEmpty(strSql)) return;
            strArt = strArt.Trim().PadLeft(7, '0');

            Boolean b = !string.IsNullOrEmpty(strSql);
            if (!b) return;

            StringBuilder sb = new StringBuilder();

            if (strTip == _clsDef.VARPOS || strTip == _clsDef.VARALL)
            {
                sb.Append("IF EXISTS (SELECT 1 FROM GesVariazioni WHERE var_art='").Append(strArt).Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='POS') ")
                  .Append("UPDATE GesVariazioni SET var_dti=GETDATE(), var_dtv=GETDATE(), var_inv='0', var_ori='").Append(strOri.Replace("'", "''")).Append("' WHERE var_art='").Append(strArt).Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='POS' ")
                  .Append("ELSE ")
                  .Append("INSERT INTO GesVariazioni (var_tip, var_art, var_num, var_dti, var_dtv, var_ori, var_inv, var_off) VALUES ('POS', '").Append(strArt).Append("', '").Append(_clsDef.COD03Z).Append("', GETDATE(), GETDATE(), '").Append(strOri.Replace("'", "''")).Append("', '0', ''); ");
            }

            /*** Etichetta ***/

            if (strTip == _clsDef.VARETI || strTip == _clsDef.VARALL)
            {
                string varOff = "";
                if (strSql.Length > 3 && strSql.Substring(0, 3) == "OFF") varOff = strSql.Substring(3);
                if (strSql.Length >= 3 && strSql.Substring(0, 3) == "PRO") varOff = "PRO";

                sb.Append("IF EXISTS (SELECT 1 FROM GesVariazioni WHERE var_art='").Append(strArt).Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='ETI') ")
                  .Append("UPDATE GesVariazioni SET var_dti=GETDATE(), var_dtv=GETDATE(), var_inv='0', var_ori='").Append(strOri.Replace("'", "''")).Append("', var_off='").Append(varOff.Replace("'", "''")).Append("' WHERE var_art='").Append(strArt).Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='ETI' ")
                  .Append("ELSE ")
                  .Append("INSERT INTO GesVariazioni (var_tip, var_art, var_num, var_dti, var_dtv, var_ori, var_inv, var_off) VALUES ('ETI', '").Append(strArt).Append("', '").Append(_clsDef.COD03Z).Append("', GETDATE(), GETDATE(), '").Append(strOri.Replace("'", "''")).Append("', '0', '").Append(varOff.Replace("'", "''")).Append("'); ");
            }

            if (sb.Length > 0)
            {
                _clsFun.SqlWrite(sb.ToString(), _strConSql);
            }
        }

        public void VariazioniBatch(IEnumerable<Tuple<string, string, string, string>> items)
        {
            if (items == null) return;
            StringBuilder sb = new StringBuilder();
            int count = 0;

            foreach (var item in items)
            {
                string strArt = item.Item1;
                string strSql = item.Item2;
                string strOri = item.Item3;
                string strTip = item.Item4;

                if (string.IsNullOrEmpty(strArt) || string.IsNullOrEmpty(strSql)) continue;
                strArt = strArt.Trim().PadLeft(7, '0');

                if (strTip == _clsDef.VARPOS || strTip == _clsDef.VARALL)
                {
                    sb.Append("IF EXISTS (SELECT 1 FROM GesVariazioni WHERE var_art='").Append(strArt).Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='POS') ")
                      .Append("UPDATE GesVariazioni SET var_dti=GETDATE(), var_dtv=GETDATE(), var_inv='0', var_ori='").Append(strOri.Replace("'", "''")).Append("' WHERE var_art='").Append(strArt).Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='POS' ")
                      .Append("ELSE ")
                      .Append("INSERT INTO GesVariazioni (var_tip, var_art, var_num, var_dti, var_dtv, var_ori, var_inv, var_off) VALUES ('POS', '").Append(strArt).Append("', '").Append(_clsDef.COD03Z).Append("', GETDATE(), GETDATE(), '").Append(strOri.Replace("'", "''")).Append("', '0', ''); ");
                    count++;
                }

                if (strTip == _clsDef.VARETI || strTip == _clsDef.VARALL)
                {
                    string varOff = "";
                    if (strSql.Length > 3 && strSql.Substring(0, 3) == "OFF") varOff = strSql.Substring(3);
                    if (strSql.Length >= 3 && strSql.Substring(0, 3) == "PRO") varOff = "PRO";

                    sb.Append("IF EXISTS (SELECT 1 FROM GesVariazioni WHERE var_art='").Append(strArt).Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='ETI') ")
                      .Append("UPDATE GesVariazioni SET var_dti=GETDATE(), var_dtv=GETDATE(), var_inv='0', var_ori='").Append(strOri.Replace("'", "''")).Append("', var_off='").Append(varOff.Replace("'", "''")).Append("' WHERE var_art='").Append(strArt).Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='ETI' ")
                      .Append("ELSE ")
                      .Append("INSERT INTO GesVariazioni (var_tip, var_art, var_num, var_dti, var_dtv, var_ori, var_inv, var_off) VALUES ('ETI', '").Append(strArt).Append("', '").Append(_clsDef.COD03Z).Append("', GETDATE(), GETDATE(), '").Append(strOri.Replace("'", "''")).Append("', '0', '").Append(varOff.Replace("'", "''")).Append("'); ");
                    count++;
                }

                if (count >= 100)
                {
                    _clsFun.SqlWrite(sb.ToString(), _strConSql);
                    sb.Clear();
                    count = 0;
                }
            }

            if (sb.Length > 0)
            {
                _clsFun.SqlWrite(sb.ToString(), _strConSql);
                sb.Clear();
            }
        }

        /// <summary>
        /// Controllo e consolidamento automatico offerte all'avvio:
        /// 1. Offerte in partenza o attive oggi (oft_dti <= Oggi <= oft_dtf) con stato 'N' o 'D':
        ///    - Accoda variazioni POS (Casse/Bilance) ed ETI (Frontalini con codice offerta) in GesVariazioni con var_num='000'.
        ///    - Aggiorna lo stato su GesOffTestate a 'D' (STADAA).
        /// 2. Offerte scadute (oft_dtf < Oggi) con stato 'A' o 'D':
        ///    - Accoda variazioni POS ed ETI con prezzo di listino ripristinato (var_off='') in GesVariazioni.
        ///    - Aggiorna lo stato su GesOffTestate a 'C' (STACLO).
        /// </summary>
        public void ConsolidaOfferteAllAvvio(out int countPos, out int countEti)
        {
            countPos = 0;
            countEti = 0;

            try
            {
                StringBuilder sbSql = new StringBuilder();
                int pendingStatements = 0;

                Action flushSql = () =>
                {
                    if (sbSql.Length > 0)
                    {
                        _clsFun.SqlWrite(sbSql.ToString(), _strConSql);
                        sbSql.Clear();
                        pendingStatements = 0;
                    }
                };

                // A. OFFERTE IN ATTIVAZIONE / VALIDE OGGI (da attivare)
                string sSqlAtt = @"SELECT oft_yea, oft_cod, oft_des, oft_dti, oft_dtf, oft_sta 
                                   FROM GesOffTestate 
                                   WHERE (oft_sta = 'N' OR oft_sta = 'D') 
                                     AND CONVERT(date, oft_dti) <= CONVERT(date, GETDATE()) 
                                     AND CONVERT(date, oft_dtf) >= CONVERT(date, GETDATE())";
                DataTable dtAtt = _clsFun.FillTabSql("GesOffTestate", sSqlAtt, false, _strConSql);

                if (dtAtt != null && dtAtt.Rows.Count > 0)
                {
                    foreach (DataRow rOft in dtAtt.Rows)
                    {
                        string sYea = rOft["oft_yea"].ToString().Trim();
                        string sCod = rOft["oft_cod"].ToString().Trim();
                        string sDes = rOft["oft_des"] != DBNull.Value ? rOft["oft_des"].ToString().Trim() : "";
                        string sOri = ("Offerta " + sCod + " " + sDes).Trim();
                        if (sOri.Length > 50) sOri = sOri.Substring(0, 50);

                        // Aggiorna stato offerta a 'D' (STADAA) se era 'N'
                        sbSql.Append("UPDATE GesOffTestate SET oft_sta='").Append(_clsDef.STADAA)
                             .Append("' WHERE oft_yea='").Append(sYea)
                             .Append("' AND oft_cod='").Append(sCod)
                             .Append("' AND oft_sta='").Append(_clsDef.STANOA).Append("'; ");
                        pendingStatements++;

                        // Seleziona tutti gli articoli attivi dell'offerta
                        string sSqlArt = "SELECT ofa_art FROM GesOffArticoli WHERE ofa_yea='" + sYea + "' AND ofa_cod='" + sCod + "' AND ofa_ann = 0";
                        DataTable dtArt = _clsFun.FillTabSql("GesOffArticoli", sSqlArt, false, _strConSql);

                        if (dtArt != null)
                        {
                            foreach (DataRow rArt in dtArt.Rows)
                            {
                                if (rArt["ofa_art"] == DBNull.Value) continue;
                                string strArt = rArt["ofa_art"].ToString().Trim().PadLeft(7, '0');
                                if (string.IsNullOrEmpty(strArt) || strArt == "0000000") continue;

                                // Variazione POS per Casse e Bilance
                                sbSql.Append("IF EXISTS (SELECT 1 FROM GesVariazioni WHERE var_art='").Append(strArt)
                                     .Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='POS') ")
                                     .Append("UPDATE GesVariazioni SET var_dti=GETDATE(), var_dtv=GETDATE(), var_inv='0', var_ori='").Append(sOri.Replace("'", "''"))
                                     .Append("', var_off='").Append(sCod.Replace("'", "''"))
                                     .Append("' WHERE var_art='").Append(strArt).Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='POS' ")
                                     .Append("ELSE ")
                                     .Append("INSERT INTO GesVariazioni (var_tip, var_art, var_num, var_dti, var_dtv, var_ori, var_inv, var_off) VALUES ('POS', '")
                                     .Append(strArt).Append("', '").Append(_clsDef.COD03Z).Append("', GETDATE(), GETDATE(), '").Append(sOri.Replace("'", "''"))
                                     .Append("', '0', '").Append(sCod.Replace("'", "''")).Append("'); ");
                                countPos++;
                                pendingStatements++;

                                // Variazione ETI per Stampa Etichette / Frontalini Offerta
                                sbSql.Append("IF EXISTS (SELECT 1 FROM GesVariazioni WHERE var_art='").Append(strArt)
                                     .Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='ETI') ")
                                     .Append("UPDATE GesVariazioni SET var_dti=GETDATE(), var_dtv=GETDATE(), var_inv='0', var_ori='").Append(sOri.Replace("'", "''"))
                                     .Append("', var_off='").Append(sCod.Replace("'", "''"))
                                     .Append("' WHERE var_art='").Append(strArt).Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='ETI' ")
                                     .Append("ELSE ")
                                     .Append("INSERT INTO GesVariazioni (var_tip, var_art, var_num, var_dti, var_dtv, var_ori, var_inv, var_off) VALUES ('ETI', '")
                                     .Append(strArt).Append("', '").Append(_clsDef.COD03Z).Append("', GETDATE(), GETDATE(), '").Append(sOri.Replace("'", "''"))
                                     .Append("', '0', '").Append(sCod.Replace("'", "''")).Append("'); ");
                                countEti++;
                                pendingStatements++;

                                if (pendingStatements >= 100) flushSql();
                            }
                        }
                    }
                }

                // B. OFFERTE SCADUTE / DA CHIUDERE (dtf < oggi)
                string sSqlSca = @"SELECT oft_yea, oft_cod, oft_des, oft_dti, oft_dtf, oft_sta 
                                   FROM GesOffTestate 
                                   WHERE (oft_sta = 'A' OR oft_sta = 'D') 
                                     AND CONVERT(date, oft_dtf) < CONVERT(date, GETDATE())";
                DataTable dtSca = _clsFun.FillTabSql("GesOffTestate", sSqlSca, false, _strConSql);

                if (dtSca != null && dtSca.Rows.Count > 0)
                {
                    foreach (DataRow rOft in dtSca.Rows)
                    {
                        string sYea = rOft["oft_yea"].ToString().Trim();
                        string sCod = rOft["oft_cod"].ToString().Trim();
                        string sOri = ("CHIUSURA OFFERTA " + sCod).Trim();
                        if (sOri.Length > 50) sOri = sOri.Substring(0, 50);

                        // Aggiorna stato offerta a 'C' (STACLO)
                        sbSql.Append("UPDATE GesOffTestate SET oft_sta='").Append(_clsDef.STACLO)
                             .Append("' WHERE oft_yea='").Append(sYea)
                             .Append("' AND oft_cod='").Append(sCod).Append("'; ");
                        pendingStatements++;

                        // Seleziona tutti gli articoli attivi dell'offerta chiusa
                        string sSqlArt = "SELECT ofa_art FROM GesOffArticoli WHERE ofa_yea='" + sYea + "' AND ofa_cod='" + sCod + "' AND ofa_ann = 0";
                        DataTable dtArt = _clsFun.FillTabSql("GesOffArticoli", sSqlArt, false, _strConSql);

                        if (dtArt != null)
                        {
                            foreach (DataRow rArt in dtArt.Rows)
                            {
                                if (rArt["ofa_art"] == DBNull.Value) continue;
                                string strArt = rArt["ofa_art"].ToString().Trim().PadLeft(7, '0');
                                if (string.IsNullOrEmpty(strArt) || strArt == "0000000") continue;

                                // Variazione POS per ripristino prezzo normale su Casse e Bilance (var_off = '')
                                sbSql.Append("IF EXISTS (SELECT 1 FROM GesVariazioni WHERE var_art='").Append(strArt)
                                     .Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='POS') ")
                                     .Append("UPDATE GesVariazioni SET var_dti=GETDATE(), var_dtv=GETDATE(), var_inv='0', var_ori='").Append(sOri.Replace("'", "''"))
                                     .Append("', var_off='' WHERE var_art='").Append(strArt).Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='POS' ")
                                     .Append("ELSE ")
                                     .Append("INSERT INTO GesVariazioni (var_tip, var_art, var_num, var_dti, var_dtv, var_ori, var_inv, var_off) VALUES ('POS', '")
                                     .Append(strArt).Append("', '").Append(_clsDef.COD03Z).Append("', GETDATE(), GETDATE(), '").Append(sOri.Replace("'", "''"))
                                     .Append("', '0', ''); ");
                                countPos++;
                                pendingStatements++;

                                // Variazione ETI per Stampa Etichette con prezzo normale ripristinato (var_off = '')
                                sbSql.Append("IF EXISTS (SELECT 1 FROM GesVariazioni WHERE var_art='").Append(strArt)
                                     .Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='ETI') ")
                                     .Append("UPDATE GesVariazioni SET var_dti=GETDATE(), var_dtv=GETDATE(), var_inv='0', var_ori='").Append(sOri.Replace("'", "''"))
                                     .Append("', var_off='' WHERE var_art='").Append(strArt).Append("' AND var_num='").Append(_clsDef.COD03Z).Append("' AND var_tip='ETI' ")
                                     .Append("ELSE ")
                                     .Append("INSERT INTO GesVariazioni (var_tip, var_art, var_num, var_dti, var_dtv, var_ori, var_inv, var_off) VALUES ('ETI', '")
                                     .Append(strArt).Append("', '").Append(_clsDef.COD03Z).Append("', GETDATE(), GETDATE(), '").Append(sOri.Replace("'", "''"))
                                     .Append("', '0', ''); ");
                                countEti++;
                                pendingStatements++;

                                if (pendingStatements >= 100) flushSql();
                            }
                        }
                    }
                }

                flushSql();
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("clsVariazioni.ConsolidaOfferteAllAvvio", ex.Message);
            }
        }

        //        public void Var2Var(string strTip, DataRow rowVar)
        //        {
        //            string sTip = "";
        //            string sTip = "";

        //var_idx	int	Unchecked
        //var_tip	nchar(3)	Unchecked
        //var_num	nchar(3)	Unchecked
        //var_art	nchar(7)	Unchecked
        //var_inv	nchar(1)	Unchecked
        //var_dti	datetime	Unchecked
        //var_dtv	datetime	Unchecked
        //var_ori	nchar(50)	Unchecked
        //var_off	nchar(3)	Unchecked
        //var_etq	decimal(4, 0)	Checked
        //var_ean	nchar(13)	Checked
        //        Unchecked
        //        }

        public void DivNegArticoli(string strPar, ArrayList aryArt, ProgressBar pBar, string strNeg, string strFor)
        {
            if (string.IsNullOrEmpty(strPar)) return;
            string sParNorm = strPar.Trim().ToUpper();
            if (sParNorm == "N" || sParNorm == "NO" || sParNorm == "0" || sParNorm == "FALSE") return;

            if (aryArt == null || aryArt.Count == 0) return;

            string[] a = strPar.Split('|');
            string sPath = a[0].Trim();
            if (string.IsNullOrEmpty(sPath) || sPath.ToUpper() == "N" || sPath.ToUpper() == "NO" || sPath.ToUpper() == "0") return;

            if (!sPath.EndsWith("\\") && !sPath.EndsWith("/"))
                sPath += "\\";

            if (!Directory.Exists(sPath))
            {
                try { Directory.CreateDirectory(sPath); } catch { }
            }

            string s = "";
            DataRow[] j;

            s = "SELECT * FROM TabNegozi WHERE tab_ann=0 AND tab_tip='L'";
            if(strNeg != "")
                s = "SELECT * FROM TabNegozi WHERE tab_cod='" + strNeg + "'";
            DataTable tNeg = _clsFun.FillTabSql("TabNegozi", s, false, _strConSql);

            if(tNeg.Rows.Count == 0)
                MessageBox.Show("Parametro 'E' in tabella negozi non presente!", "Divulgazione a negozi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                DataTable tTmp = new clsGenTabTmp().TabTmpArt("ArtCod");
                DataTable t = new DataTable();
                DataTable tEan = new DataTable();

                if (aryArt.Count <= 100)
                {
                    string inArt = string.Join(",", aryArt.Cast<object>().Select(x => "'" + Convert.ToString(x).Replace("'", "''") + "'"));
                    s = "SELECT gia_neg, gia_art, gia_gia FROM AnaArtGiacenza WHERE gia_art IN (" + inArt + ")";
                }
                else
                {
                    s = "SELECT gia_neg, gia_art, gia_gia FROM AnaArtGiacenza";
                }
                DataTable tGia = _clsFun.FillTabSql("AnaArtGiacenza", s, false, _strConSql);
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tGia.Columns["gia_neg"];
                keys[1] = tGia.Columns["gia_art"];
                tGia.PrimaryKey = keys;

                string inArtAll = string.Join(",", aryArt.Cast<object>().Select(x => "'" + Convert.ToString(x).Replace("'", "''") + "'"));
                DataTable tEanAll = _clsFun.FillTabSql("AnaBarcode", "SELECT ean_art, ean_ean, ean_qta, ean_prv, ean_bil, ean_ecp, ean_ann FROM AnaBarcode WHERE ean_art IN (" + inArtAll + ")", false, _strConSql);
                Dictionary<string, List<DataRow>> dictEan = new Dictionary<string, List<DataRow>>();
                if (tEanAll != null)
                {
                    foreach (DataRow re in tEanAll.Rows)
                    {
                        if (re["ean_art"] != DBNull.Value)
                        {
                            string eArt = re["ean_art"].ToString().Trim().PadLeft(7, '0');
                            if (!dictEan.ContainsKey(eArt)) dictEan[eArt] = new List<DataRow>();
                            dictEan[eArt].Add(re);
                        }
                    }
                }

                // Pre-caricamento offerte attive O(1) per evitare migliaia di query individuali in loop
                DataTable tOffAll = _clsQry.OfsSeek(DateTime.Today);
                Dictionary<string, DataRow> dictOff = new Dictionary<string, DataRow>();
                if (tOffAll != null)
                {
                    foreach (DataRow ro in tOffAll.Rows)
                    {
                        if (ro["ofa_art"] != DBNull.Value)
                        {
                            string artCode = ro["ofa_art"].ToString().Trim().PadLeft(7, '0');
                            dictOff[artCode] = ro;
                        }
                    }
                }

                foreach(DataRow k in tNeg.Rows)
                {
                    tTmp.Rows.Clear();
                    string sLisNeg = "LNE" + (string)k["tab_lis"];

                    if (pBar != null)
                    {
                        pBar.Value = 0;
                        pBar.Maximum = aryArt.Count;
                        pBar.Minimum = 0;
                    }

                    int idxArt = 0;
                    foreach(string sa in aryArt)
                    {
                        idxArt++;
                        if (pBar != null && (idxArt % 25 == 0 || idxArt == aryArt.Count))
                        {
                            pBar.Value = Math.Min(idxArt, pBar.Maximum);
                            Application.DoEvents();
                        }

                        if(strFor != "")
                            t = _clsQry.ArtSeek(sa, "FOR" + strFor);
                        else
                            t = _clsQry.ArtSeek(sa, sLisNeg);

                        if (t != null && t.Rows.Count > 0)
                            tTmp.ImportRow(t.Rows[0]);
                    }

                    string sFilPref = "dar_" + (string)k["tab_cod"] + "_";
                    string sTimeTag = DateTime.Now.ToString("yyyyMMddHHmmssff");
                    string sFilTemp = sPath + "_" + sFilPref + sTimeTag + ".csv";
                    string sFilFinal = sPath + sFilPref + sTimeTag + ".csv";

                    try
                    {
                        using (StreamWriter sw = new StreamWriter(sFilTemp, false))
                        {
                            string sRig = "T|";
                            foreach (DataColumn c in tTmp.Columns)
                                sRig += c.ColumnName + "|";
                            sRig += "ean_ean";
                            sw.Write(sRig + _clsDef.CRLF);

                            if (pBar != null)
                            {
                                pBar.Value = 0;
                                pBar.Maximum = tTmp.Rows.Count;
                                pBar.Minimum = 0;
                            }

                            int idxRow = 0;
                            foreach (DataRow y in tTmp.Rows)
                            {
                                idxRow++;
                                if (pBar != null && (idxRow % 25 == 0 || idxRow == tTmp.Rows.Count))
                                {
                                    pBar.Value = Math.Min(idxRow, pBar.Maximum);
                                    Application.DoEvents();
                                }

                                string yArtNorm = ((string)y["tmp_art"]).Trim().PadLeft(7, '0');

                                //Giacenza O(1) con PrimaryKey
                                DataRow giaRow = tGia.Rows.Find(new object[] { (string)k["tab_cod"], (string)y["tmp_art"] });
                                if (giaRow != null && giaRow["gia_gia"] != DBNull.Value)
                                    y["tmp_gia"] = (decimal)giaRow["gia_gia"];

                                //Offerta O(1) da cache pre-caricata
                                if (dictOff.TryGetValue(yArtNorm, out DataRow tOffRow))
                                {
                                    y["tmp_tof"] = (string)tOffRow["ofa_tip"];
                                    y["tmp_pro"] = (decimal)tOffRow["ofa_val"];
                                }

                                sRig = "R|";
                                foreach (DataColumn c in tTmp.Columns)
                                {
                                    if (c.DataType.Name == "Boolean")
                                    {
                                        if (Convert.ToString(y[c]) == "False")
                                            sRig += "0|";
                                        else
                                            sRig += "1|";
                                    }
                                    else
                                        sRig += Convert.ToString(y[c]) + "|";
                                }

                                s = "";
                                if (dictEan.TryGetValue(yArtNorm, out List<DataRow> eanList))
                                {
                                    foreach (DataRow yy in eanList)
                                    {
                                        s += (string)yy["ean_ean"] + ";";
                                        s += _clsFun.Dec2Txt(Convert.ToInt16(yy["ean_qta"]), 0) + ";";
                                        s += _clsFun.Dec2Txt(Convert.ToInt16(yy["ean_prv"]), 2) + ";";
                                        s += Convert.ToInt16(yy["ean_bil"]) + ";";
                                        s += Convert.ToInt16(yy["ean_ecp"]) + ";";
                                        s += Convert.ToInt16(yy["ean_ann"]) + "?";
                                    }
                                }

                                sRig += s;

                                sw.Write(sRig + _clsDef.CRLF);
                            }

                            sw.Flush();
                        }

                        if (File.Exists(sFilTemp))
                        {
                            if (File.Exists(sFilFinal))
                                File.Delete(sFilFinal);

                            File.Move(sFilTemp, sFilFinal);
                        }
                    }
                    catch (Exception ex)
                    {
                        _clsFun.ErrorLog("clsVariazioni.DivNegArticoli", ex.Message);
                    }
                }
            }
        }

        public void DivNegOfferta(string strPar, string strOfff, ProgressBar pBar)
        {
            string s = "";

            string[] a = strOfff.Split(',');

            string sOffYea = a[0];
            string sOffCod = a[1];

            s = "SELECT * FROM TabNegozi WHERE tab_ann=0 AND tab_tip='E'";
            DataTable tNeg = _clsFun.FillTabSql("TabNegozi", s, false, _strConSql);

            if (tNeg.Rows.Count == 0)
                MessageBox.Show("Parametro 'E' in tabella negozi non presente!", "Divulgazione a negozi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                DataTable tTmp = new clsGenTabTmp().TabTmpArtOff("ArtOff");
                DataTable t = new DataTable();
                DataTable tEan = new DataTable();

                a = strPar.Split('|');
                string sPath = a[0];

                s = "SELECT * FROM GesOffArticoli ";
                s += "LEFT OUTER JOIN GesOffTestate ON GesOffArticoli.ofa_yea = GesOffTestate.oft_yea AND GesOffArticoli.ofa_cod = GesOffTestate.oft_cod "; 
                s += "WHERE ";
                s += "ofa_yea='" + sOffYea + "' AND ";
                s += "ofa_cod='" + sOffCod + "' ";
                s += "ORDER BY ofa_rig";
                DataTable tOff = _clsFun.FillTabSql("", s, false, _strConSql);

                foreach (DataRow k in tNeg.Rows)
                {
                    string sLisNeg = "LNE" + (string)k["tab_lis"];

                    if (pBar != null)
                    {
                        pBar.Value = 0;
                        pBar.Maximum = tOff.Rows.Count;
                        pBar.Minimum = 0;
                    }

                    tTmp.Clear();

                    foreach (DataRow y in tOff.Rows)
                    {
                        if (pBar != null)
                        {
                            pBar.Increment(1);
                            Application.DoEvents();
                        }

                        if ((string)y["ofa_art"] == "0000703")
                            Console.WriteLine("xxxxxxx");

                        t = _clsQry.ArtSeek((string)y["ofa_art"], sLisNeg);

                        DataRow x = tTmp.NewRow();

                        foreach (DataColumn c in t.Columns)
                            x[c.ColumnName] = t.Rows[0][c.ColumnName];

                        x["oft_dti"] = y["oft_dti"];
                        x["oft_dtf"] = y["oft_dtf"];
                        x["oft_yea"] = y["oft_yea"];
                        x["oft_cod"] = y["oft_cod"];
                        x["oft_des"] = y["oft_des"];
                        x["ofa_rig"] = y["ofa_rig"];
                        x["ofa_tip"] = y["ofa_tip"];
                        x["ofa_val"] = y["ofa_val"];
                        x["ofa_xem"] = y["ofa_xem"];
                        x["ofa_xen"] = y["ofa_xen"];
                        x["ofa_mix"] = y["ofa_mix"];
                        x["ofa_pun"] = y["ofa_pun"];
                        x["ofa_ann"] = y["ofa_ann"];
                        x["ofa_not"] = y["ofa_not"];

                        tTmp.Rows.Add(x);
                    }

                    string sFilPref = "dof_" + (string)k["tab_cod"] + "_";

                    string sFil = sPath + "_" + sFilPref + DateTime.Now.ToString("yyyyMMddHHmmssff") + ".csv";

                    StreamWriter sw = new StreamWriter(sFil, true);

                    string sRig = "T|";

                    //if (a.Length == 0)
                    //{
                    foreach (DataColumn c in tTmp.Columns)
                        sRig += c.ColumnName + "|";
                    sRig += "ean_ean";
                    sw.Write(sRig + _clsDef.CRLF);
                    //}

                    if (pBar != null)
                    {
                        pBar.Value = 0;
                        pBar.Maximum = tTmp.Rows.Count;
                        pBar.Minimum = 0;
                    }

                    foreach (DataRow y in tTmp.Rows)
                    {
                        if (pBar != null)
                        {
                            pBar.Increment(1);
                            Application.DoEvents();
                        }
                        sRig = "R|";
                        foreach (DataColumn c in tTmp.Columns)
                        {
                            if (c.DataType.Name == "Boolean")
                            {
                                if (Convert.ToString(y[c]) == "False")
                                    sRig += "0|";
                                else
                                    sRig += "1|";
                            }
                            else if (c.DataType.Name == "DateTime")
                                sRig += Convert.ToDateTime(y[c]).ToString("yyyyMMdd") + "|";
                            else
                                sRig += Convert.ToString(y[c]) + "|";
                        }
                        s = "SELECT * FROM AnaBarcode WHERE ean_art='" + (string)y["tmp_art"] + "'";
                        tEan = _clsFun.FillTabSql(TABANAEAN, s, false, _strConSql);
                        s = "";
                        foreach (DataRow yy in tEan.Rows)
                        {
                            s += (string)yy["ean_ean"] + ";";
                            s += _clsFun.Dec2Txt(Convert.ToInt16(yy["ean_qta"]), 0) + ";";
                            s += _clsFun.Dec2Txt(Convert.ToInt16(yy["ean_prv"]), 2) + ";";
                            s += Convert.ToInt16(yy["ean_bil"]) + ";";
                            s += Convert.ToInt16(yy["ean_ecp"]) + ";";
                            s += Convert.ToInt16(yy["ean_ann"]) + "?";
                        }

                        sRig += s;

                        sw.Write(sRig + _clsDef.CRLF);
                    }

                    ((TextWriter)sw).Flush();
                    sw.Close();
                    sw.Dispose();

                    s = sPath + Path.GetFileName(sFil).Substring(1);

                    File.Move(sFil, s);
                }
            }
        }

        public void DivNegTabelle(string strPar, DataTable tabTab)
        {
            string s = "";

            s = "SELECT * FROM TabNegozi WHERE tab_ann=0 AND tab_tip='E'";
            DataTable tNeg = _clsFun.FillTabSql("TabNegozi", s, false, _strConSql);

            string[] a = strPar.Split('|');
            string sPath = a[0];

            foreach(DataRow k in tNeg.Rows)
            {
                string sFilPref = "tab_" + (string)k["tab_cod"] + "_";

                string sFil = sPath + "_" + sFilPref + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";

                a = Directory.GetFiles(sPath, sFilPref + "*.csv");

                if (a.Length > 0)
                {
                    sFil = a[0];

                    s = sPath + "_" + Path.GetFileName(sFil);

                    File.Move(sFil, s);
                    sFil = s;
                }

                StreamWriter sw = new StreamWriter(sFil, true);

                string sRig = "T|tab_tab|";

                foreach (DataColumn c in tabTab.Columns)
                    sRig += c.ColumnName + "|";
                sw.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in tabTab.Rows)
                {
                    sRig = "R|" + tabTab.TableName + "|";
                    foreach (DataColumn c in tabTab.Columns)
                    {
                        if (c.DataType.Name == "Boolean")
                        {
                            if (Convert.ToString(y[c]) == "False")
                                sRig += "0|";
                            else
                                sRig += "1|";
                        }
                        else
                            sRig += Convert.ToString(y[c]) + "|";
                    }

                    sw.Write(sRig + _clsDef.CRLF);
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                s = sPath + Path.GetFileName(sFil).Substring(1);

                File.Move(sFil, s);

            }
        }

        public void DivNegDocumenti(string strPar)
        {
            string s = "";

            string[] a = strPar.Split(';');

            string sTip = a[0];
            string sYea = a[1];
            string sNum = a[2];

            string sTab = "GesFatTestate";
            if(sTip == "M")
                sTab = "GesMovTestate";

            string sNeg = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");

            //string sPath = _clsFun.ParGet(clsDefine.enuParametri.Par016PathDivNegozi, _strConSql);
            s = _clsFun.ParGet(clsDefine.enuParametri.Par036Div2Sede, _strConSql);
            a = s.Split(',');
            string sPath = a[0];

            string sFil = sPath + "_" + "doc_" + sNeg + "_" + DateTime.Now.ToString("yyyyMMddHHmmssff") + ".csv";
            StreamWriter sw = new StreamWriter(sFil, true);

            if (sTip == "F")
                s = "SELECT * FROM GesFatTestate WHERE fat_yfa='" + sYea + "' AND  fat_nfa='" + sNum + "'";
            else
                s = "SELECT * FROM GesMovTestate WHERE mot_ymo='" + sYea + "' AND  mot_nmo='" + sNum + "'";

            DataTable t = _clsFun.FillTabSql("DOC", s, false, _strConSql);

            //string sRig = "TF|" + sTip + "|";
            string sRig = "T|tab_tab|" + sTip + "|";

            foreach (DataColumn c in t.Columns)
                sRig += c.ColumnName + "|";
            sw.Write(sRig + _clsDef.CRLF);

            sRig = "R|" + sTab + "|" + sTip + "|";
            foreach (DataColumn c in t.Columns)
            {
                if (c.DataType.Name == "Boolean")
                {
                    if (Convert.ToString(t.Rows[0][c]) == "False")
                        sRig += "0|";
                    else
                        sRig += "1|";
                }
                else
                    sRig += Convert.ToString(t.Rows[0][c]) + "|";
            }
            sw.Write(sRig + _clsDef.CRLF);

            s = "SELECT * FROM GesMovimenti WHERE ";
            if (sTip == "F")
            {
                s += "mov_yfa = '" + sYea + "' AND ";
                s += "mov_nfa = '" + sNum + "' ";
                s += "ORDER BY mov_rfa";
            }
            else
            {
                s += "mov_ymo = '" + sYea + "' AND ";
                s += "mov_nmo = '" + sNum + "' ";
                s += "ORDER BY mov_rmo";
            }
            t = _clsFun.FillTabSql("MOV", s, false, _strConSql);

            //sRig = sRig = "MF|" + sTip + "|";
            sRig = sRig = "T|tab_tab|" + sTip + "|";

            foreach (DataColumn c in t.Columns)
                sRig += c.ColumnName + "|";
            sw.Write(sRig + _clsDef.CRLF);

            foreach (DataRow y in t.Rows)
            {
                sRig = "R|GesMovimenti|" + sTip + "|";
                foreach (DataColumn c in t.Columns)
                {
                    if (c.DataType.Name == "Boolean")
                    {
                        if (Convert.ToString(y[c]) == "False")
                            sRig += "0|";
                        else
                            sRig += "1|";
                    }
                    else
                        sRig += Convert.ToString(y[c]) + "|";
                }

                sw.Write(sRig + _clsDef.CRLF);
            }

            ((TextWriter)sw).Flush();
            sw.Close();
            sw.Dispose();

            s = sPath + Path.GetFileName(sFil).Substring(1);

            File.Move(sFil, s);
        }

        public string DivNeg2SedeApPhoneTmp(string strPar, DataRow rowTmp)
        {
            string s = "";
            string sMsg = "";

            string sDay = (string)rowTmp["tmp_day"];

            s = "SELECT * FROM TmpDiv WHERE tmp_day='" + sDay + "'";
            DataTable t = _clsFun.FillTabSql("TmpDiv", s, false, _strConSql);

            string[] a = strPar.Split('|');
            string sPath = a[0];
            string sNeg = a[1];

            if (!Directory.Exists(sPath))
                sMsg = "Percorso di rete non trovato: " + sPath;
            else
            {
                string sFilPref = "apf_" + sNeg + "_";

                string sFil = sPath + "_" + sFilPref + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";

                a = Directory.GetFiles(sPath, sFilPref + "*.csv");

                if (a.Length > 0)
                {
                    sFil = a[0];

                    s = sPath + "_" + Path.GetFileName(sFil);

                    File.Move(sFil, s);
                    sFil = s;
                }

                StreamWriter sw = new StreamWriter(sFil, true);

                string sRig = "T|tab_tab|";

                foreach (DataColumn c in t.Columns)
                    sRig += c.ColumnName + "|";
                sw.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in t.Rows)
                {
                    sRig = "R|" + t.TableName + "|";
                    foreach (DataColumn c in t.Columns)
                    {
                        if (c.DataType.Name == "Boolean")
                        {
                            if (Convert.ToString(y[c]) == "False")
                                sRig += "0|";
                            else
                                sRig += "1|";
                        }
                        else
                            sRig += Convert.ToString(y[c]) + "|";
                    }

                    sw.Write(sRig + _clsDef.CRLF);
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                s = sPath + Path.GetFileName(sFil).Substring(1);

                File.Move(sFil, s);
            }

            return sMsg;
        }

        public string DivNeg2SedeInventario(string strPar)
        {
            string s = "";
            string sMsg = "";

            string[] a = strPar.Split(',');
            string sPath = a[0];
            string sNum = a[1];
            string sNeg = a[2];

            if (!Directory.Exists(sPath))
                sMsg = "Percorso di rete non trovato: " + sPath;
            else
            {
                string sFilPref = "inv_" + sNeg + "_";

                string sFil = sPath + "_" + sFilPref + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";

                a = Directory.GetFiles(sPath, sFilPref + "*.csv");

                if (a.Length > 0)
                {
                    sFil = a[0];

                    s = sPath + "_" + Path.GetFileName(sFil);

                    File.Move(sFil, s);
                    sFil = s;
                }

                StreamWriter sw = new StreamWriter(sFil, true);

                //Testata

                s = "SELECT * FROM GesInvTestate WHERE int_neg='" + sNeg + "' AND int_num='" + sNum + "'";
                DataTable t = _clsFun.FillTabSql("GesInvTestate", s, false, _strConSql);

                string sRig = "T|tab_tab|";

                foreach (DataColumn c in t.Columns)
                    sRig += c.ColumnName + "|";
                sw.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in t.Rows)
                {
                    sRig = "R|" + t.TableName + "|";
                    foreach (DataColumn c in t.Columns)
                    {
                        if (c.DataType.Name == "Boolean")
                        {
                            if (Convert.ToString(y[c]) == "False")
                                sRig += "0|";
                            else
                                sRig += "1|";
                        }
                        else
                            sRig += Convert.ToString(y[c]) + "|";
                    }

                    sw.Write(sRig + _clsDef.CRLF);
                }

                //Righe

                s = "SELECT * FROM GesInventario WHERE inv_neg='" + sNeg + "' AND inv_num='" + sNum + "'";
                t = _clsFun.FillTabSql("GesInventario", s, false, _strConSql);

                sRig = "T|tab_tab|";

                foreach (DataColumn c in t.Columns)
                    sRig += c.ColumnName + "|";
                sw.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in t.Rows)
                {
                    sRig = "R|" + t.TableName + "|";
                    foreach (DataColumn c in t.Columns)
                    {
                        if (c.DataType.Name == "Boolean")
                        {
                            if (Convert.ToString(y[c]) == "False")
                                sRig += "0|";
                            else
                                sRig += "1|";
                        }
                        else
                            sRig += Convert.ToString(y[c]) + "|";
                    }

                    sw.Write(sRig + _clsDef.CRLF);
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                s = sPath + Path.GetFileName(sFil).Substring(1);

                File.Move(sFil, s);
            }

            return sMsg;
        }

        public void AggApOfficeTab(string strFil)
        {
            DataRow x;
            DataRow[] j;
            string s = "";
            string sRig = "";

            ArrayList aWhe = new ArrayList();
            aWhe.Add("tab_cod");
            ArrayList aExl = new ArrayList();

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                string[] aTes = sRig.Split('|');
                DataTable tTab = new DataTable();
                string sTab = "";

                while ((sRig = sr.ReadLine()) != null)
                {
                    if(sRig.Substring(0,1) == "T")
                        aTes = sRig.Split('|');
                    else if(sRig.Substring(0,1) == "R")
                    {
                        string[] aRig = sRig.Split('|');
                        int i = Array.IndexOf(aTes, "tab_tab");

                        //i -= 1;

                        if (sTab != aRig[i])
                        {
                            sTab = aRig[i];
                            s = "SELECT * FROM " + sTab;
                            tTab = _clsFun.FillTabSql(sTab, s, false, _strConSql);

                            if (sTab == "TabEcrLv2")
                            {
                                aWhe = new ArrayList();
                                aWhe.Add("tab_cod");
                                aWhe.Add("tab_lv1");
                                aExl = new ArrayList();
                            }
                            else if (sTab == "TabEcrLv3")
                            {
                                aWhe = new ArrayList();
                                aWhe.Add("tab_cod");
                                aWhe.Add("tab_lv1");
                                aWhe.Add("tab_lv2");
                                aExl = new ArrayList();
                            }

                        }
                        if (sTab != "TabParametri")
                        {
                            i = Array.IndexOf(aTes, "tab_cod");
                            string sCod = aRig[i];

                            x = tTab.NewRow();
                            foreach (DataColumn c in tTab.Columns)
                            {
                                if (c.ColumnName != "tab_idx")
                                {
                                    i = Array.IndexOf(aTes, c.ColumnName);
                                    //x[c.ColumnName] = aRig[i].Trim();
                                    if (c.DataType.Name == "Boolean")
                                        x[c.ColumnName] = Convert.ToInt16(aRig[i]);
                                    else
                                        x[c.ColumnName] = aRig[i].Trim();
                                }
                            }

                            //s = "tab_cod='" + sCod + "'";
                            //if (sTab == "TabEcrLv2")
                            //    s = "tab_cod='" + sCod + "' AND tab_lv1='" + x["tab_lv1"] + "'";
                            //else if (sTab == "TabEcrLv3")
                            //    s = "tab_cod='" + sCod + "' AND tab_lv1='" + x["tab_lv1"] + "' AND tab_lv2='" + x["tab_lv2"] + "'";

                            s = "";
                            foreach(string k in aWhe)
                                s += k + "='" + x[k] + "' AND ";

                            s = s.Substring(0, s.Length - 4);

                            j = tTab.Select(s);

                            if (j.Length == 0)
                                s = _clsFun.SqlInsertRow(sTab, tTab, x);
                            else
                                s = _clsFun.SqlUpdRow(sTab, tTab, j[0], x, aWhe, aExl);

                            if (s != "")
                                _clsFun.SqlWrite(s, _strConSql);
                        }
                    }
                }

                sr.Close();
                sr.Dispose();

                s = Path.GetDirectoryName(strFil) + "\\Old\\" + Path.GetFileName(strFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                File.Move(strFil, s);
            }
        }

        private bool _isExecutingEtiBatch = false;

        public void EtiElettroniche(string strCod, DataTable tabEti, Label lblPrg = null, ProgressBar prgBar = null, IWin32Window owner = null)
        {
            if (_isExecutingEtiBatch) return;
            _isExecutingEtiBatch = true;
            try
            {
                string s = "";
                decimal d = 0;

            s = "SELECT * FROM TabEtichette WHERE tab_cod='" + strCod + "'";
            DataTable t = _clsFun.FillTabSql("TabEtichette", s, false, _strConSql);

            if(t.Rows.Count > 0)
            {
                s = (string)t.Rows[0]["tab_par"];

                string[] a = s.Split(';');

                if (s == "" || a.Length < 2 || a[1].Length ==  0)
                {
                    frmWait.CloseWait();
                    MessageBox.Show(owner ?? Form.ActiveForm, "Configurazione path mancante!", "CONTROLLO PARAMETRI ETICHETTA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                { 
                    string sPthEsl = a[1].Trim();
                    if (!sPthEsl.EndsWith("\\")) sPthEsl += "\\";

                    string sPth = "\\ApProject\\Temp\\DivEtichette\\";

                    if (!Directory.Exists(sPth))
                        Directory.CreateDirectory(sPth + "\\Old");

                    string sFil = sPth + "\\varEti.csv";
                    string sRig = "";

                    if (!File.Exists(sFil))
                        sRig = "#ACTION;Article No.;EAN/UPC;Description;Group;Standard price;Sell price;Discount %;Content;Unit;Quantity (GR/ML);Std. Quantity (GR/ML);Price/Std. Qnt.;Std. Unit;Location X (cm);Location Y (cm);Alt. Image;Data Estrazione;Data Variazione;Data Inizio Offerta;Data Fine Offerta;";    
            
                    // Pre-cache: verifica quali file ingredienti esistono
                    string sIngDir = _clsDef.PATHINGREDIENTI;
                    if (!string.IsNullOrEmpty(sIngDir) && !sIngDir.EndsWith("\\")) sIngDir += "\\";
                    HashSet<string> ingFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    if (!string.IsNullOrEmpty(sIngDir) && Directory.Exists(sIngDir))
                    {
                        foreach (string f in Directory.GetFiles(sIngDir, "et01_*.rtf"))
                        {
                            ingFiles.Add(Path.GetFullPath(f));
                        }
                    }
                    Dictionary<string, string> ingTextCache = new Dictionary<string, string>();

                    string sDataOggi = DateTime.Now.ToString("dd/MM/yyyy");
                    bool hasEtiCol = (tabEti.Rows.Count > 0 && tabEti.Columns.Contains("eti_eti"));

                    using (BufferedStream bs = new BufferedStream(new FileStream(sFil, FileMode.Append, FileAccess.Write, FileShare.None), 65536))
                    using (StreamWriter sw = new StreamWriter(bs, System.Text.Encoding.Default))
                    {
                        if (sRig.Length > 0)
                            sw.Write(sRig + _clsDef.CRLF);

                        System.Text.StringBuilder sb = new System.Text.StringBuilder(512);

                        int iRowCsv = 0;
                        int iTotCsv = tabEti.Rows.Count;

                        foreach (DataRow y in tabEti.Rows)
                        {
                            iRowCsv++;
                            if (lblPrg != null && prgBar != null && (iRowCsv % 250 == 0 || iRowCsv == iTotCsv))
                            {
                                int iPrc = (iTotCsv > 0) ? (iRowCsv * 100 / iTotCsv) : 0;
                                if (prgBar.Maximum != iTotCsv) prgBar.Maximum = iTotCsv > 0 ? iTotCsv : 1;
                                prgBar.Value = iRowCsv;
                                lblPrg.Text = "Generazione file CSV...  " + iPrc.ToString() + "%  (" + iRowCsv.ToString() + " / " + iTotCsv.ToString() + ")";
                                Application.DoEvents();
                            }

                            sb.Clear();
                            sb.Append("I;");
                            sb.Append(Convert.ToString(y["eti_art"]).Trim()).Append(";");
                                sb.Append(((string)y["eti_ean"]).Trim()).Append(";");
                                sb.Append(((string)y["eti_ard"]).Trim()).Append(";");
                                sb.Append(((string)y["eti_rep"]).Trim()).Append(";");
                                sb.Append(Convert.ToString(y["eti_pve"])).Append(";");
                                sb.Append(Convert.ToString(y["eti_prv"])).Append(";");
                                sb.Append(";");                                                //Sconto %
                                sb.Append(Convert.ToInt32(y["eti_pne"]).ToString()).Append(";");     //Contenuto
                                sb.Append(Convert.ToString(y["eti_tgr"])).Append(";");               //Tipo grammatura
                                sb.Append(Convert.ToInt32(y["eti_pne"]).ToString()).Append(";");     //Contenuto

                                sb.Append(Convert.ToString(y["eti_tgv"])).Append(";");               //Moltiplicatore GR ML = 1000

                                if ((string)y["eti_tgr"] != "PZ" &&
                                    Convert.ToDecimal(y["eti_prv"]) > 0 &&
                                    Convert.ToDecimal(y["eti_pne"]) > 0 &&
                                    Convert.ToDecimal(y["eti_tgv"]) > 0)
                                {
                                    d = Convert.ToDecimal(y["eti_prv"]) / (Convert.ToDecimal(y["eti_pne"]) / Convert.ToDecimal(y["eti_tgv"]));
                                    if (d > 500m)
                                    {
                                        d = 0m;
                                    }
                                }
                                else
                                    d = 0;

                                sb.Append(d.ToString()).Append(";");                                            //Prezzo a KG
                                sb.Append(Convert.ToString(y["eti_umi"])).Append(";");           //UMI

                                // Lettura ingredienti da file RTF senza RichTextBox (con cache in memoria)
                                string sIng = "";
                                string sArtCode = Convert.ToString(y["eti_art"]).Trim();
                                string sIngFile = sIngDir + "et01_" + sArtCode + ".rtf";
                                string fullIngPath = Path.GetFullPath(sIngFile);
                                if (ingFiles.Contains(fullIngPath))
                                {
                                    if (ingTextCache.TryGetValue(fullIngPath, out string cachedText))
                                    {
                                        sIng = cachedText;
                                    }
                                    else
                                    {
                                        try
                                        {
                                            string rtfContent = File.ReadAllText(fullIngPath);
                                            sIng = ExtractTextFromRtf(rtfContent);
                                            ingTextCache[fullIngPath] = sIng;
                                        }
                                        catch { }
                                    }
                                }

                                sb.Append(sIng).Append(";");                                       //Ingredienti    ex Posizione x
                                sb.Append(((string)y["eti_plu"]).Trim()).Append(";;");              //Posizione Y
                                sb.Append(sDataOggi).Append(";");
                                if (DBNull.Value.Equals(y["eti_day"]))
                                {
                                    sb.Append(";");
                                }
                                else
                                {
                                    sb.Append(Convert.ToDateTime(y["eti_day"]).ToString("dd/MM/yyyy")).Append(";");
                                }

                                if (DBNull.Value.Equals(y["eti_odi"]))
                                {
                                    sb.Append(";");
                                }
                                else
                                {
                                    sb.Append(Convert.ToDateTime(y["eti_odi"]).ToString("dd/MM/yyyy")).Append(";");
                                }

                                if (DBNull.Value.Equals(y["eti_odf"]))
                                {
                                    sb.Append(";");
                                }
                                else
                                {
                                    sb.Append(Convert.ToDateTime(y["eti_odf"]).ToString("dd/MM/yyyy")).Append(";");
                                }

                                sw.WriteLine(sb.ToString());
                        }
                    }

                    if (!Directory.Exists(sPthEsl))
                        Directory.CreateDirectory(sPthEsl);

                    string s1 = sPthEsl + "dbase.csv";
                    string s2 = sPthEsl + "dbase.csv.tmp";

                    if (File.Exists(s2))
                        File.Delete(s2);
                    File.Copy(sFil, s2);

                    if (File.Exists(s1))
                        File.Delete(s1);
                    File.Move(s2, s1);

                    s2 = sPth + "Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (File.Exists(sFil))
                        File.Move(sFil, s2);

                    // 1. Funzione helper: lancia il batch con cmd.exe /c su thread separato
                    //    con guardia anti-doppia esecuzione e attesa reattiva sulla UI
                    bool isBatchRunning = false;
                    Action executeBatch = () =>
                    {
                        if (isBatchRunning) return;
                        isBatchRunning = true;
                        try
                        {
                            if (a.Length >= 3 && !string.IsNullOrEmpty(a[2].Trim()))
                            {
                                string rawCmd = a[2].Trim().Replace("\"", "");
                                string exePath = rawCmd;
                                string args = "";
                                bool isVisible = (a.Length >= 4 && a[3].Trim().ToUpper() == "V");

                                if (!File.Exists(exePath) && rawCmd.Contains(" "))
                                {
                                    int spaceIdx = rawCmd.IndexOf(' ');
                                    exePath = rawCmd.Substring(0, spaceIdx).Trim();
                                    args = rawCmd.Substring(spaceIdx + 1).Trim();
                                }

                                if (File.Exists(exePath))
                                {
                                    var evtDone = new System.Threading.ManualResetEventSlim(false);
                                    Exception thrEx = null;

                                    var th = new System.Threading.Thread(() =>
                                    {
                                        try
                                        {
                                            var psi = new System.Diagnostics.ProcessStartInfo();
                                            string ext = Path.GetExtension(exePath).ToLower();
                                            if (ext == ".bat" || ext == ".cmd")
                                            {
                                                psi.FileName = "cmd.exe";
                                                psi.Arguments = "/c \"" + exePath + "\"" + (string.IsNullOrEmpty(args) ? "" : " " + args);
                                                psi.UseShellExecute = false;
                                                psi.CreateNoWindow = !isVisible;
                                                if (isVisible)
                                                    psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
                                            }
                                            else
                                            {
                                                psi.FileName = exePath;
                                                psi.Arguments = args;
                                                psi.UseShellExecute = true;
                                                if (!isVisible)
                                                    psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                                            }
                                            psi.WorkingDirectory = Path.GetDirectoryName(exePath);

                                            using (var p = System.Diagnostics.Process.Start(psi))
                                            {
                                                if (p != null)
                                                {
                                                    if (!p.WaitForExit(60000))
                                                    {
                                                        try { p.Kill(); } catch { }
                                                        thrEx = new Exception("Il processo batch ha superato il tempo limite di attesa (60s) ed è stato terminato.");
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            thrEx = ex;
                                        }
                                        finally
                                        {
                                            evtDone.Set();
                                        }
                                    });
                                    th.IsBackground = true;
                                    th.Name = "BatchWaiter";
                                    th.Start();

                                    // Thread UI libero: ogni 100ms pompa il message loop
                                    while (!evtDone.Wait(100))
                                    {
                                        Application.DoEvents();
                                    }

                                    if (thrEx != null)
                                        _clsFun.ErrorLog("EtiElettroniche.ExecCmd", thrEx.Message);
                                }
                                else
                                {
                                    frmWait.CloseWait();
                                    MessageBox.Show(owner ?? Form.ActiveForm, "File comando batch non trovato al percorso:\n" + exePath + "\n\nVerificare la configurazione in TabEtichette (tab_par).", "COMANDO BATCH NON TROVATO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                            }
                        }
                        finally
                        {
                            isBatchRunning = false;
                        }
                    };

                    // Lancia il batch ed attende la fine senza bloccare la UI
                    executeBatch();

                    // Chiude la form di attesa non appena il batch si conclude
                    frmWait.CloseWait();

                    // Chiede all'utente se desidera segnare le etichette come stampate/inviate per cambiarne lo stato
                    DialogResult askState = MessageBox.Show(
                        owner ?? Form.ActiveForm,
                        "Generazione ed esportazione file CSV per etichette elettroniche completata.\n\nSi desidera segnare le etichette come stampate/inviate per cambiarne lo stato?",
                        "ESPORTAZIONE ETICHETTE ELETTRONICHE",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button1);

                    if (askState == DialogResult.Yes)
                    {
                        SegnaEtichetteStampate(tabEti, hasEtiCol, strCod);
                        MessageBox.Show(owner ?? Form.ActiveForm, "Le etichette sono state contrassegnate come stampate/inviate.", "STATO AGGIORNATO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show(owner ?? Form.ActiveForm, "Le etichette sono state esportate senza modificarne lo stato.", "ESPORTAZIONE COMPLETATA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            }
            finally
            {
                _isExecutingEtiBatch = false;
            }
        }

        private void SegnaEtichetteStampate(DataTable tabEti, bool hasEtiCol, string strCod)
        {
            try
            {
                HashSet<string> exportedArts = new HashSet<string>();
                foreach (DataRow r in tabEti.Rows)
                {
                    string artCode = "";
                    if (r.Table.Columns.Contains("eti_art") && !DBNull.Value.Equals(r["eti_art"]))
                        artCode = r["eti_art"].ToString().Trim();
                    else if (r.Table.Columns.Contains("var_art") && !DBNull.Value.Equals(r["var_art"]))
                        artCode = r["var_art"].ToString().Trim();

                    if (!string.IsNullOrEmpty(artCode))
                    {
                        exportedArts.Add(artCode.PadLeft(7, '0'));
                    }
                }

                if (exportedArts.Count == 0) return;

                int iNum = 0;
                string sSel = "SELECT TOP 1 var_num FROM GesVariazioni WHERE var_tip='ETI' AND var_num <> '000' ORDER BY var_dtv DESC, var_num DESC";
                DataTable tVarExist = _clsFun.FillTabSql(TABGESVAR, sSel, true, _strConSql);
                if (tVarExist != null && tVarExist.Rows.Count > 0 && !DBNull.Value.Equals(tVarExist.Rows[0]["var_num"]))
                {
                    int.TryParse(tVarExist.Rows[0]["var_num"].ToString(), out iNum);
                }
                iNum++;
                if (iNum > 999) iNum = 1;
                string strNumBatch = iNum.ToString("000");

                List<string> artList = exportedArts.ToList();
                for (int i = 0; i < artList.Count; i += 500)
                {
                    var chunk = artList.Skip(i).Take(500);
                    string inClause = string.Join(",", chunk.Select(c => "'" + c.Replace("'", "''") + "'"));

                    string sUpd = "UPDATE GesVariazioni SET var_num='" + strNumBatch + "', var_inv='1', var_dtv=" + _clsFun.DaySql(DateTime.Today) + " WHERE var_art IN (" + inClause + ") AND var_tip='ETI'";
                    _clsFun.SqlWrite(sUpd, _strConSql);
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("SegnaEtichetteStampate", ex.Message);
            }
        }

        private string ExtractTextFromRtf(string rtf)
        {
            if (string.IsNullOrEmpty(rtf)) return "";
            StringBuilder sb = new StringBuilder(rtf.Length);
            bool inTag = false;
            for (int i = 0; i < rtf.Length; i++)
            {
                char c = rtf[i];
                if (c == '\\')
                {
                    // Skip RTF control word
                    while (i + 1 < rtf.Length && char.IsLetter(rtf[i + 1])) i++;
                    if (i + 1 < rtf.Length && rtf[i + 1] == ' ') i++; // skip trailing space of control word
                }
                else if (c == '{' || c == '}')
                {
                    continue;
                }
                else if (c != '\r' && c != '\n')
                {
                    sb.Append(c);
                }
            }
            return sb.ToString().Trim();
        }
    }
}
