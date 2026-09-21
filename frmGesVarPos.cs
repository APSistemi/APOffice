using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace APOffice
{
    public partial class frmGesVarPos : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public DataTable _tabPos = new DataTable();
        public DataTable _tabCli = new DataTable();
        public DataTable _tabFid = new DataTable();
        public DataTable _tabOtr = new DataTable();         //Offerte transazione
        private DataTable _tabNeg = new DataTable();

        public string _strTip = "";

        private const string TABLISACQ = "GesLisAcquisto";

        private string _strConSql = "";
        private string _strCnfPar = "";    //Parametri

        //public string _strRet = "";

        public frmGesVarPos()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);

            _tabNeg = _clsFun.FillTabSql("TabNegozi", "SELECT * FROM TabNegozi WHERE tab_ann=0 AND tab_tip='L'", false, _strConSql);
        }

        private void frmGesVarPos_Load(object sender, EventArgs e)
        {
            frmWait.ShowWait("INVIO IN CORSO A CASSE E BILANCE... ATTENDERE.");
            try
            {
                _strConSql = _clsFun.ConSql("");
                _strCnfPar = _clsQry.DefCnfPar("POS");
                string s = "";
                string s2 = "";
                DataTable tCnf = _clsQry.ConfSeek("", "POS");

                if (_strTip != "SOLOTERM" && tCnf.Rows.Count > 0)
                {
                    s = Convert.ToInt16((object)clsDefine.enuPos.posDitron).ToString("00");
                    /*** Variazioni cassa ***/
                    if ((string)tCnf.Rows[0]["cnf_pos"] != "" && _tabPos != null && _tabPos.Rows.Count > 0)
                    {
                        if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16((object)clsDefine.enuPos.posDitron).ToString("00"))
                            PosDitron((string)tCnf.Rows[0]["CnfVar"], (string)tCnf.Rows[0]["CnfVao"]);
                        if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16((object)clsDefine.enuPos.posBrainpos).ToString("00"))
                            PosBrainpos((string)tCnf.Rows[0]["CnfVar"], (string)tCnf.Rows[0]["CnfVao"]);
                        if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16((object)clsDefine.enuPos.posNcr745x).ToString("00"))
                            PosNcr745x((string)tCnf.Rows[0]["CnfVar"]);
                        if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16((object)clsDefine.enuPos.posUgaSid20).ToString("00"))
                            PosUgaSid20((string)tCnf.Rows[0]["CnfVar"], (string)tCnf.Rows[0]["CnfVao"]);
                        if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16((object)clsDefine.enuPos.posApShop01).ToString("00"))
                            PosApShop01((string)tCnf.Rows[0]["CnfVar"], (string)tCnf.Rows[0]["CnfVao"]);
                    }
                    /*** Variazioni clienti fidelity ***/
                    if ((string)tCnf.Rows[0]["cnf_pos"] != "" && _tabPos != null && _tabFid.Rows.Count > 0)
                    {
                        if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16(clsDefine.enuPos.posDitron).ToString("00"))
                        {
                            s = Path.GetDirectoryName((string)tCnf.Rows[0]["CnfVar"]) + "\\" + "CLIENTI.OK";
                            s2 = Path.GetDirectoryName((string)tCnf.Rows[0]["CnfVar"]) + "\\" + "PROGUPD.OK";  //Montante
                            PosFidDitron(s, s2);
                        }
                        if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16(clsDefine.enuPos.posBrainpos).ToString("00"))
                        {
                            s = Path.GetDirectoryName((string)tCnf.Rows[0]["CnfVar"]) + "\\" + "CLIENTI.TXT";
                            s2 = Path.GetDirectoryName((string)tCnf.Rows[0]["CnfVar"]) + "\\" + "PROGUPD.TXT";  //Montante
                            PosFidBrainpos(s, s2);
                        }
                        if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16(clsDefine.enuPos.posApShop01).ToString("00"))
                        {
                            PosTesApShop01((string)tCnf.Rows[0]["CnfVar"]);
                        }
                    }
                    /*** Variazioni anagrafiche clienti ***/
                    if ((string)tCnf.Rows[0]["cnf_pos"] != "" && _tabPos != null && _tabCli.Rows.Count > 0)
                    {
                        if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16(clsDefine.enuPos.posApShop01).ToString("00"))
                        {
                            PosCliApShop01((string)tCnf.Rows[0]["CnfVar"]);
                        }
                    }
                    if ((string)tCnf.Rows[0]["cnf_pos"] != "" && _tabOtr != null && _tabOtr.Rows.Count > 0)
                    {
                        if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16(clsDefine.enuPos.posApShop01).ToString("00"))
                            PosOffTraApShop01((string)tCnf.Rows[0]["CnfVar"]);
                        if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16(clsDefine.enuPos.posDitron).ToString("00"))
                            PosOffTraDitron((string)tCnf.Rows[0]["CnfVao"]);
                    }

                    /*** Variazioni bilancia ***/
                    if ((string)tCnf.Rows[0]["cnf_bil"] != "" && _tabPos != null && _tabPos.Rows.Count > 0)
                    {
                        DataTable tCnb = _clsQry.ConfSeek("", "BIL");

                        if ((string)tCnf.Rows[0]["cnf_bil"] == Convert.ToInt16((object)clsDefine.enuBilance.bilOmega).ToString("00"))
                            BilOmega((string)tCnf.Rows[0]["CnfVab"]);
                        if ((string)tCnf.Rows[0]["cnf_bil"] == Convert.ToInt16((object)clsDefine.enuBilance.bilBizerba).ToString("00"))
                            BilBizerba((string)tCnf.Rows[0]["CnfVab"], "");
                        if ((string)tCnf.Rows[0]["cnf_bil"] == Convert.ToInt16((object)clsDefine.enuBilance.bilBizWinVarp).ToString("00"))
                            BilBizWinvarp((string)tCnf.Rows[0]["CnfVab"]);
                        if ((string)tCnf.Rows[0]["cnf_bil"] == Convert.ToInt16((object)clsDefine.enuBilance.bilBizerba2).ToString("00"))
                            BilBizerba((string)tCnf.Rows[0]["CnfVab"], "2");
                    }

                    foreach (DataRow y in _tabPos.Rows)
                    {
                        if ((string)y["pos_msg"] == "" && (string)y["pos_inv"] == _clsDef.DIVDAD)
                            y["pos_inv"] = _clsDef.DIVDIV;
                    }

                }

                DivTerm();
            }
            catch (Exception ex)
            {
                frmWait.CloseWait();
                MessageBox.Show(this, "Errore in frmGesVarPos_Load: " + ex.Message, "ERRORE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _clsFun.ErrorLog("frmGesVarPos_Load", ex.Message);
            }
            finally
            {
                frmWait.CloseWait();
            }

            Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        public void ChiudiOfferta(DataRow rowOft)
        {
            DataTable tCnf = _clsQry.ConfSeek("", "POS");

            if (tCnf != null && tCnf.Rows.Count > 0 && (string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16((object)clsDefine.enuPos.posDitron).ToString("00"))
                DitronChiudiOff(rowOft, (string)tCnf.Rows[0]["CnfVao"]);
        }

        private void PosDitron(string strFilVar, string strFilVao)
        {
            DataRow[] j;

            clsPosDitron clsPos = new clsPosDitron();

            string sFil = _clsDef.TMPPOSVAR;
            string sFio = _clsDef.TMPPOSOFF;
            Boolean b = false;
            string s = "";

            StreamWriter sw = new StreamWriter(sFil, true);
            StreamWriter so = new StreamWriter(sFio, true);

            int totalRows = _tabPos.Rows.Count;
            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = totalRows > 0 ? totalRows : 1;
            progressBar1.Minimum = 0;

            ArrayList aMix = new ArrayList();
            List<string[]> listLog = new List<string[]>();
            int idx = 0;
            string origTitle = this.Text;

            foreach (DataRow y in _tabPos.Rows)
            {
                idx++;
                if (idx % 25 == 0 || idx == totalRows)
                {
                    int pct = (int)((idx / (double)totalRows) * 100);
                    progressBar1.Value = Math.Min(idx, progressBar1.Maximum);
                    lblMsg.Text = $"Generazione POSVAR...\n\n{pct}%\n({idx} / {totalRows})";
                    this.Text = $"Invio Variazioni - {pct}% ({idx}/{totalRows})";
                    frmWait.UpdateText($"GENERAZIONE FILE DITRON: {pct}% ({idx}/{totalRows})");
                    Application.DoEvents();
                }

                if ((string)y["pos_inv"] == _clsDef.DIVDAD)
                {
                    s = "";
                    if (((string)y["pos_sta"]).Trim() == _clsDef.STAREP)
                    {
                        if (((string)y["pos_rep"]).Trim() == "")
                            s += "Manca Reparto ";
                    }
                    else
                    {
                        if (((string)y["pos_ean"]).Trim() == "" && ((string)y["pos_art"]).Trim() == "")
                            s += "Manca EAN ";
                        if (((string)y["pos_rep"]).Trim() == "")
                            s += "Manca Reparto ";
                        if ((decimal)y["pos_prv"] == 0)
                        {
                            if (_strCnfPar.Length == 0 || _strCnfPar.Substring(0, 1) != "S")
                                s += "Manca prezzo ";
                        }
                    }
                    y["pos_msg"] = s;

                    if (s == "")
                    {
                        string sRig = clsPos.PosArtRow(y);
                        y["pos_inv"] = _clsDef.DIVDIV;
                        sw.Write(sRig + _clsDef.CRLF);

                        b = true;
                        if (Convert.ToInt32(y["pos_mix"]) > 0)
                        {
                            if (aMix.IndexOf(Convert.ToInt32(y["pos_mix"])) < 0)
                            {
                                aMix.Add(Convert.ToInt32(y["pos_mix"]));
                                if ((string)y["pos_tva"] == _clsDef.OFASCO)
                                    sRig = clsPos.PosOffArtRowSco(y);
                                else if ((string)y["pos_tva"] == _clsDef.OFAMXN)
                                    sRig = clsPos.PosOffArtRowMxN(y);
                                so.Write(sRig + "\r\n");
                            }
                        }

                        string[] aLog = { 
                                    "POSART",                           //  log_tip
                                    (string)y["pos_art"],               //  log_art
                                    (string)y["pos_ean"],               //  log_ean
                                    (string)y["pos_ard"],               //  log_des
                                    "",                                 //  log_fil
                                    "",                                 //  log_acq
                                    ((decimal)y["pos_prv"]).ToString(), //  log_prv
                                    (string)y["pos_tvd"]                //  log_msg
                                        };
                        listLog.Add(aLog);
                    }
                }
            }

            ((TextWriter)sw).Flush();
            sw.Close();
            sw.Dispose();

            ((TextWriter)so).Flush();
            so.Close();
            so.Dispose();

            // Bulk Log Write to SQL Server
            if (listLog.Count > 0)
            {
                _clsQry.LogSqlBulk(listLog);
            }

            this.Text = origTitle;

            FileInfo fInfo = new FileInfo(sFio);
            if (fInfo.Length == 0)
                File.Delete(sFio);

            if (b)
            {
                string sOld = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                if (File.Exists(strFilVar))
                    File.Delete(strFilVar);
                File.Copy(sFil, strFilVar);
                File.Move(sFil, sOld);

                if (File.Exists(sFio))
                {
                    sOld = Path.GetDirectoryName(sFio) + "\\Old\\" + Path.GetFileName(sFio) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (File.Exists(strFilVao))
                        File.Delete(strFilVao);
                    File.Copy(sFio, strFilVao);
                    File.Move(sFio, sOld);
                }
            }

            s = Path.GetDirectoryName(strFilVar) + "\\prodotti.err";
            if (File.Exists(s))
            {
                Process.Start("notepad.exe", s);
            }
        }

        private void PosFidDitron(string strFilVar, string strFilVarMon)
        {
            DataTable t = new clsQuery().Campagna();
            string sCam = (string)t.Rows[0]["tab_cod"];

            clsPosDitron clsPos = new clsPosDitron();

            string sFil = _clsDef.TMPPOSVAR;
            string sFilMon = _clsDef.TMPPOSFID;
            Boolean b = false;
            string s = "";

            StreamWriter sw = new StreamWriter(sFil, true);
            StreamWriter swMon = new StreamWriter(sFilMon, true);

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = _tabFid.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in _tabFid.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                if ((string)y["TesInv"] == _clsDef.DIVDAD)
                {
                    s = "";
                    if (((string)y["tes_cod"]).Trim() == "")
                        s += "Manca codice ";

                    if (s == "")
                    {
                        string sRig = clsPos.PosFidCliRow(y);
                        sw.Write(sRig + _clsDef.CRLF);

                        sRig = clsPos.PosFidMonCliRow(y, sCam);
                        swMon.Write(sRig + _clsDef.CRLF);

                        b = true;
                    }
                }
            }

            ((TextWriter)sw).Flush();
            sw.Close();
            sw.Dispose();

            ((TextWriter)swMon).Flush();
            swMon.Close();
            swMon.Dispose();

            if (b)
            {
                string sOld = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                if (File.Exists(strFilVar))
                    File.Delete(strFilVar);
                File.Copy(sFil, strFilVar);
                File.Move(sFil, sOld);

                sOld = Path.GetDirectoryName(sFilMon) + "\\Old\\" + Path.GetFileName(sFilMon) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                if (File.Exists(strFilVarMon))
                    File.Delete(strFilVarMon);
                File.Copy(sFilMon, strFilVarMon);
                File.Move(sFilMon, sOld);
            }
        }

        private void PosOffTraDitron(string strFilVao)
        {
            DataRow[] j;

            clsPosDitron clsPos = new clsPosDitron();

            //string sFil = _clsDef.TMPPOSVAR;
            string sFio = _clsDef.TMPPOSOFF;
            Boolean b = false;
            string s = "";

            //StreamWriter sw = new StreamWriter(sFil, true);
            StreamWriter so = new StreamWriter(sFio, true);

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = _tabPos.Rows.Count;
            progressBar1.Minimum = 0;

            ArrayList aMix = new ArrayList();

            string sRig = "";

            foreach (DataRow y in _tabOtr.Rows)
                sRig += clsPos.PosOffTrans(y);
            so.Write(sRig + _clsDef.CRLF);

            //foreach (DataRow y in _tabPos.Rows)
            //{
            //    progressBar1.Increment(1);
            //    Application.DoEvents();
            //    if ((string)y["pos_inv"] == _clsDef.DIVDAD)
            //    {
            //        s = "";
            //        if (((string)y["pos_ean"]).Trim() == "")
            //            s += "Manca EAN ";
            //        if (((string)y["pos_rep"]).Trim() == "")
            //            s += "Manca Reparto ";
            //        if ((decimal)y["pos_prv"] == 0)
            //            s += "Manca prezzo ";
            //        //{
            //        //    if (_strCnfPar.Length == 0 || _strCnfPar.Substring(0, 1) != "S")
            //        //        s += "Manca prezzo ";
            //        //}
            //        y["pos_msg"] = s;
            //        if (s == "")
            //        {
            //            string sRig = clsPos.PosArtRow(y);
            //            y["pos_inv"] = _clsDef.DIVDIV;
            //            sw.Write(sRig + _clsDef.CRLF);
            //            b = true;
            //            if (Convert.ToInt32(y["pos_mix"]) > 0)
            //            {
            //                if (aMix.IndexOf(Convert.ToInt32(y["pos_mix"])) < 0)
            //                {
            //                    aMix.Add(Convert.ToInt32(y["pos_mix"]));
            //                    if ((string)y["pos_tva"] == _clsDef.OFASCO)
            //                        sRig = clsPos.PosOffArtRowSco(y);
            //                    else if ((string)y["pos_tva"] == _clsDef.OFAMXN)
            //                        sRig = clsPos.PosOffArtRowMxN(y);
            //                    so.Write(sRig + "\r\n");
            //                }
            //            }
            //        }
            //    }
            //}

            //((TextWriter)sw).Flush();
            //sw.Close();
            //sw.Dispose();

            ((TextWriter)so).Flush();
            so.Close();
            so.Dispose();

            FileInfo fInfo = new FileInfo(sFio);
            if (fInfo.Length == 0)
                File.Delete(sFio);

            if (b)
            {
                //string sOld = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                //if (File.Exists(strFilVar))
                //    File.Delete(strFilVar);
                //File.Copy(sFil, strFilVar);
                //File.Move(sFil, sOld);

                if (File.Exists(sFio))
                {
                    string sOld = Path.GetDirectoryName(sFio) + "\\Old\\" + Path.GetFileName(sFio) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (File.Exists(strFilVao))
                        File.Delete(strFilVao);
                    File.Copy(sFio, strFilVao);
                    File.Move(sFio, sOld);
                }
            }

            s = Path.GetDirectoryName(strFilVao) + "\\prodotti.err";
            if (File.Exists(s))
            {
                Process.Start("notepad.exe", s);
            }
        }

        private void DitronChiudiOff(DataRow rowOft, string strFilVao)
        {
            string s = "";
            ArrayList aMix = new ArrayList();
            clsPosDitron clsPos = new clsPosDitron();

            string sFio = _clsDef.TMPPOSOFF;

            string sYea = (string)rowOft["oft_yea"];
            string sCod = (string)rowOft["oft_cod"];

            s = "SELECT * FROM GesOffArticoli WHERE ";
            s += "GesOffArticoli.ofa_yea = '" + sYea + "' AND ";
            s += "GesOffArticoli.ofa_cod = '" + sCod + "' AND ";
            s += "GesOffArticoli.ofa_ann = 0 AND ";
            s += "GesOffArticoli.ofa_mix > 0 ";
            s += "ORDER BY ofa_mix";
            DataTable tOfa = _clsFun.FillTabSql("tOfa", s, false, _strConSql);

            if (tOfa.Rows.Count > 0)
            {
                StreamWriter so = new StreamWriter(sFio, true);
                string sRig = "";

                foreach (DataRow y in tOfa.Rows)
                {
                    if (Convert.ToInt32(y["ofa_mix"]) > 0)
                    {
                        if (aMix.IndexOf(Convert.ToInt32(y["ofa_mix"])) < 0)
                        {

                            aMix.Add(Convert.ToInt32(y["ofa_mix"]));

                            DataRow x = _tabPos.NewRow();
                            x["pos_sta"] = "C";
                            x["pos_tva"] = y["ofa_tip"];
                            x["pos_mix"] = y["ofa_mix"];
                            x["pos_ost"] = _clsDef.STACAN;
                            x["pos_ofd"] = "";
                            x["pos_dti"] = (DateTime)rowOft["oft_dti"];
                            x["pos_dtf"] = (DateTime)rowOft["oft_dti"];
                            x["pos_ova"] = y["ofa_val"];

                            x["pos_xem"] = y["ofa_xem"];
                            x["pos_xen"] = y["ofa_xen"];

                            if ((string)x["pos_tva"] == _clsDef.OFASCO)
                                sRig = clsPos.PosOffArtRowSco(x);
                            else if ((string)x["pos_tva"] == _clsDef.OFAMXN)
                                sRig = clsPos.PosOffArtRowMxN(x);
                            so.Write(sRig + "\r\n");


                        }
                    }
                }

                ((TextWriter)so).Flush();
                so.Close();
                so.Dispose();

                if (File.Exists(sFio))
                {
                    Thread.Sleep(50);

                    string sOld = Path.GetDirectoryName(sFio) + "\\Old\\" + Path.GetFileName(sFio) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (File.Exists(strFilVao))
                        File.Delete(strFilVao);
                    File.Copy(sFio, strFilVao);
                    File.Move(sFio, sOld);
                }

            }
        }

        private void PosBrainpos(string strFilVar, string strFilVao)
        {
            clsPosBrainpos clsPos = new clsPosBrainpos();

            string sFil = _clsDef.TMPPOSVAR;
            string sFio = _clsDef.TMPPOSOFF;
            Boolean b = false;
            string s = "";

            StreamWriter sw = new StreamWriter(sFil, true);
            StreamWriter so = new StreamWriter(sFio, true);

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = _tabPos.Rows.Count;
            progressBar1.Minimum = 0;

            int idxBrain = 0;
            int totalBrain = _tabPos.Rows.Count;

            foreach (DataRow y in _tabPos.Rows)
            {
                idxBrain++;
                if (idxBrain % 25 == 0 || idxBrain == totalBrain)
                {
                    int pct = (int)((idxBrain / (double)(totalBrain > 0 ? totalBrain : 1)) * 100);
                    frmWait.UpdateText($"GENERAZIONE FILE BRAINPOS: {pct}% ({idxBrain}/{totalBrain})");
                    progressBar1.Value = Math.Min(idxBrain, progressBar1.Maximum);
                    Application.DoEvents();
                }

                if ((string)y["pos_inv"] == _clsDef.DIVDAD)
                {
                    s = "";
                    if (((string)y["pos_sta"]).Trim() == _clsDef.STAREP)
                    {
                        if (((string)y["pos_rep"]).Trim() == "")
                            s += "Manca Reparto ";
                    }
                    else
                    {
                        if (((string)y["pos_ean"]).Trim() == "" && ((string)y["pos_art"]).Trim() == "")
                            s += "Manca EAN ";
                        if (((string)y["pos_rep"]).Trim() == "")
                            s += "Manca Reparto ";
                        if ((decimal)y["pos_prv"] == 0)
                            s += "Manca Prezzo ";
                    }
                    y["pos_msg"] = s;

                    if(s == "")
                    {
                        string sRig = clsPos.PosArtRow(y);
                        y["pos_inv"] = _clsDef.DIVDIV;
                        sw.Write(sRig + _clsDef.CRLF);
                        b = true;
                        //if (Convert.ToInt32(y["pos_mix"]) > 0)
                        //{
                        //    if (aMix.IndexOf(Convert.ToInt32(y["pos_mix"])) < 0)
                        //    {
                        //        aMix.Add(Convert.ToInt32(y["pos_mix"]));
                        //        sRig = clsPos.PosOffArtRow(y);
                        //        so.Write(sRig + "\r\n");
                        //    }
                        //}
                    }
                }
            }

            ((TextWriter)sw).Flush();
            sw.Close();
            sw.Dispose();

            ((TextWriter)so).Flush();
            so.Close();
            so.Dispose();

            FileInfo fInfo = new FileInfo(sFio);
            if (fInfo.Length == 0)
                File.Delete(sFio);

            if (b)
            {
                string sOld = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                if (File.Exists(strFilVar))
                    File.Delete(strFilVar);
                File.Copy(sFil, strFilVar);
                File.Move(sFil, sOld);

                if (File.Exists(sFio))
                {
                    sOld = Path.GetDirectoryName(sFio) + "\\Old\\" + Path.GetFileName(sFio) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (File.Exists(strFilVao))
                        File.Delete(strFilVao);
                    File.Copy(sFio, strFilVao);
                    File.Move(sFio, sOld);
                }

                if(File.Exists(strFilVar))
                {
                    Process p = new Process();
                    p.StartInfo.FileName = "C:\\Brainpos\\RUN\\APvariaz.bat";
                    p.StartInfo.Arguments = "-r";
                    p.StartInfo.ErrorDialog = true;
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                    p.Start();
                    p.WaitForExit(1000 * 60 * 1);    // wait up to 1 minutes.

                    sOld = Path.GetDirectoryName(strFilVar) + "\\" + Path.GetFileNameWithoutExtension(strFilVar) + ".OLD";
                    if (File.Exists(sOld))
                        File.Delete(sOld);
                    File.Move(strFilVar, sOld);
                }
            }
        }

        private void PosFidBrainpos(string strFilVar, string strFilVarMon)
        {
            DataTable t = new clsQuery().Campagna();
            string sCam = (string)t.Rows[0]["tab_cod"];

            clsPosBrainpos clsPos = new clsPosBrainpos();

            string sFil = _clsDef.TMPPOSVAR;
            string sFilMon = _clsDef.TMPPOSFID;
            Boolean b = false;
            string s = "";

            StreamWriter sw = new StreamWriter(sFil, true);
            StreamWriter swMon = new StreamWriter(sFilMon, true);

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = _tabFid.Rows.Count;
            progressBar1.Minimum = 0;

            ArrayList aMix = new ArrayList();

            foreach (DataRow y in _tabFid.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                s = "";
                if (((string)y["tes_cod"]).Trim() == "")
                    s += "Manca codice ";

                if (s == "")
                {
                    string sRig = clsPos.PosFidCliRow(y);
                    sw.Write(sRig + _clsDef.CRLF);

                    sRig = clsPos.PosFidMonCliRow(y, sCam);
                    swMon.Write(sRig + _clsDef.CRLF);

                    b = true;
                }
            }

            ((TextWriter)sw).Flush();
            sw.Close();
            sw.Dispose();

            ((TextWriter)swMon).Flush();
            swMon.Close();
            swMon.Dispose();

            if (b)
            {
                string sOld = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                if (File.Exists(strFilVar))
                    File.Delete(strFilVar);
                File.Copy(sFil, strFilVar);
                File.Move(sFil, sOld);

                sOld = Path.GetDirectoryName(sFilMon) + "\\Old\\" + Path.GetFileName(sFilMon) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                if (File.Exists(strFilVarMon))
                    File.Delete(strFilVarMon);
                File.Copy(sFilMon, strFilVarMon);
                File.Move(sFilMon, sOld);
            }
        }

        private void PosNcr745x(string strFilVar)
        {
            clsPosNcr745x clsPos = new clsPosNcr745x();

            string sFil = _clsDef.TMPPOSVAR;
            //string sFio = _clsDef.TMPPOSOFF;
            DataRow[] j;
            Boolean b = false;
            string s = "";
            int i = 0;

            StreamWriter sw = new StreamWriter(sFil, true);
            //StreamWriter so = new StreamWriter(sFio, true);

            DataTable t = _tabPos.Copy();
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "IvaNcr",
                Caption = "IvaNcr",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            DataTable tIva = new clsQuery().TabIva();

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = _tabPos.Rows.Count;
            progressBar1.Minimum = 0;

            ArrayList aMix = new ArrayList();

            foreach (DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                if ((string)y["pos_inv"] == _clsDef.DIVDAD)
                {
                    s = "";
                    if (((string)y["pos_sta"]).Trim() == _clsDef.STAREP)
                    {
                        if (((string)y["pos_rep"]).Trim() == "")
                            s += "Manca Reparto ";
                    }
                    else
                    {
                        if (((string)y["pos_ean"]).Trim() == "" && ((string)y["pos_art"]).Trim() == "")
                            s += "Manca EAN ";
                        if (((string)y["pos_rep"]).Trim() == "")
                            s += "Manca Reparto ";
                        if ((decimal)y["pos_prv"] == 0)
                            s += "Manca Prezzo ";
                    }
                    y["pos_msg"] = s;

                    if (s == "")
                    {
                        j = tIva.Select("tab_cod='" + y["pos_iva"] + "'");
                        if (j.Length > 0)
                        {
                            i = 0;
                            if (!DBNull.Value.Equals(j[0]["tab_pos"]))
                                i = Convert.ToInt16(j[0]["tab_pos"]);
                            if (i == 0)
                                s += "Manca IVA NCR";
                            else
                                y["IvaNcr"] = i;
                        }
                    }

                    if (s == "")
                    {
                        string sRig = clsPos.PosArtRow(y);
                        y["pos_inv"] = _clsDef.DIVDIV;
                        sw.Write(sRig); // + _clsDef.CRLF);
                        b = true;
                        //if (Convert.ToInt32(y["pos_mix"]) > 0)
                        //{
                        //    if (aMix.IndexOf(Convert.ToInt32(y["pos_mix"])) < 0)
                        //    {
                        //        aMix.Add(Convert.ToInt32(y["pos_mix"]));
                        //        sRig = clsPos.PosOffArtRow(y);
                        //        so.Write(sRig + "\r\n");
                        //    }
                        //}
                    }
                }
            }

            foreach (DataRow y in _tabPos.Rows)
            {
                if ((string)y["pos_inv"] == _clsDef.DIVDAD)
                    y["pos_inv"] = _clsDef.DIVDIV;
            }

            ((TextWriter)sw).Flush();
            sw.Close();
            sw.Dispose();

            //((TextWriter)so).Flush();
            //so.Close();
            //so.Dispose();

            //FileInfo fInfo = new FileInfo(sFio);
            //if (fInfo.Length == 0)
            //    File.Delete(sFio);

            if (b)
            {
                string sOld = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                if (File.Exists(strFilVar))
                    File.Delete(strFilVar);
                File.Copy(sFil, strFilVar);
                File.Move(sFil, sOld);

                //if (File.Exists(sFio))
                //{
                //    sOld = Path.GetDirectoryName(sFio) + "\\Old\\" + Path.GetFileName(sFio) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                //    if (File.Exists(strFilVao))
                //        File.Delete(strFilVao);
                //    File.Copy(sFio, strFilVao);
                //    File.Move(sFio, sOld);
                //}

                if (File.Exists(strFilVar))
                {
                    Process p = new Process();
                    p.StartInfo.FileName = Path.GetDirectoryName(strFilVar) + "\\APvariaz.bat";
                    p.StartInfo.Arguments = "-r";
                    p.StartInfo.ErrorDialog = true;
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                    p.Start();
                    p.WaitForExit(1000 * 60 * 1);    // wait up to 1 minutes.

                    //sOld = Path.GetDirectoryName(strFilVar) + "\\" + Path.GetFileNameWithoutExtension(strFilVar) + ".OLD";
                    //if (File.Exists(sOld))
                    //    File.Delete(sOld);
                    //File.Move(strFilVar, sOld);
                }
            }
        }

        private void PosUgaSid20(string strFilVar, string strFilVao)
        {
            clsPosUgaSid20 clsPos = new clsPosUgaSid20();

            string sFil = _clsDef.TMPPOSVAR;
            string sFio = _clsDef.TMPPOSOFF;
            //DataRow[] j;
            Boolean b = false;
            string s = "";
            int i = 0;

            StreamWriter sw = new StreamWriter(sFil, true);
            StreamWriter so = new StreamWriter(sFio, true);

            //DataTable t = _tabPos.Copy();
            //t.Columns.Add(new DataColumn()
            //{
            //    DataType = Type.GetType("System.Decimal"),
            //    ColumnName = "IvaNcr",
            //    Caption = "IvaNcr",
            //    ReadOnly = false,
            //    DefaultValue = (Decimal)0
            //});

            //DataTable tIva = new clsQuery().TabIva();

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = _tabPos.Rows.Count;
            progressBar1.Minimum = 0;

            //ArrayList aMix = new ArrayList();

            foreach (DataRow y in _tabPos.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                if ((string)y["pos_inv"] == _clsDef.DIVDAD)
                {
                    s = "";
                    if (((string)y["pos_sta"]).Trim() == _clsDef.STAREP)
                    {
                        if (((string)y["pos_rep"]).Trim() == "")
                            s += "Manca Reparto ";
                    }
                    else
                    {
                        if (((string)y["pos_ean"]).Trim() == "" && ((string)y["pos_art"]).Trim() == "")
                            s += "Manca EAN ";
                        if (((string)y["pos_rep"]).Trim() == "")
                            s += "Manca Reparto ";
                        if ((decimal)y["pos_prv"] == 0)
                            s += "Manca Prezzo ";
                    }
                    y["pos_msg"] = s;

                    //if (s == "")
                    //{
                    //    j = tIva.Select("tab_cod='" + y["pos_iva"] + "'");
                    //    if (j.Length > 0)
                    //    {
                    //        i = 0;
                    //        if (!DBNull.Value.Equals(j[0]["tab_ncr"]))
                    //            i = Convert.ToInt16(j[0]["tab_ncr"]);
                    //        if (i == 0)
                    //            s += "Manca IVA NCR";
                    //        else
                    //            y["IvaNcr"] = i;
                    //    }
                    //}

                    if (s == "")
                    {
                        string sRig = clsPos.PosArtRow(y);
                        //y["pos_inv"] = _clsDef.DIVDIV;
                        sw.Write(sRig); // + _clsDef.CRLF);
                        b = true;
                        if ((string)(y["pos_tva"]) != "")
                        {
                            //if (aMix.IndexOf(Convert.ToInt32(y["pos_mix"])) < 0)
                            //{
                            //    aMix.Add(Convert.ToInt32(y["pos_mix"]));
                            //}
                            sRig = clsPos.PosOffArtRow(y);
                            so.Write(sRig); // + _clsDef.CRLF);
                        }
                    }
                }
            }

            //foreach (DataRow y in _tabPos.Rows)
            //{
            //    if ((string)y["pos_inv"] == _clsDef.DIVDAD)
            //        y["pos_inv"] = _clsDef.DIVDIV;
            //}

            ((TextWriter)sw).Flush();
            sw.Close();
            sw.Dispose();

            ((TextWriter)so).Flush();
            so.Close();
            so.Dispose();

            FileInfo fInfo = new FileInfo(sFio);
            if (fInfo.Length == 0)
                File.Delete(sFio);

            if (b)
            {
                string sOld = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                if (File.Exists(strFilVar))
                    File.Delete(strFilVar);
                File.Copy(sFil, strFilVar);
                File.Move(sFil, sOld);

                if (File.Exists(sFio))
                {
                    sOld = Path.GetDirectoryName(sFio) + "\\Old\\" + Path.GetFileName(sFio) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (File.Exists(strFilVao))
                        File.Delete(strFilVao);
                    File.Copy(sFio, strFilVao);
                    File.Move(sFio, sOld);
                }

                if (File.Exists(strFilVar))
                {
                    Process p = new Process();
                    p.StartInfo.FileName = Path.GetDirectoryName(strFilVar) + "\\APvariaz.bat";
                    p.StartInfo.Arguments = "-r";
                    p.StartInfo.ErrorDialog = true;
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                    p.Start();
                    p.WaitForExit(1000 * 60 * 1);    // wait up to 1 minutes.

                    sOld = Path.GetDirectoryName(strFilVar) + "\\" + Path.GetFileNameWithoutExtension(strFilVar) + ".OLD";
                    if (File.Exists(sOld))
                        File.Delete(sOld);
                    File.Move(strFilVar, sOld);
                }
            }
        }

        private void PosApShop01(string strFilVar, string strFilVao)
        {
            if (!File.Exists(_clsDef.TMPPOSSTOP))
            {
                var f = File.Create(_clsDef.TMPPOSSTOP);
                f.Close();
            }

            try
            {
                _clsQry._strPar031Lotti2Pos = _clsFun.ParGet(clsDefine.enuParametri.Par031Lotti2Pos, _strConSql);

                int nPos = 1;
                int.TryParse(strFilVar, out nPos);
                if (nPos <= 0) nPos = 1;

                List<string[]> listLogBulk = new List<string[]>();

                foreach (DataRow yy in _tabNeg.Rows)
                {
                    Boolean b = false;
                    string s = "";
                    string sRig = "";

                    string sPath = _clsDef.TMPPOSVAR;
                    s = _clsFun.FileIni("R", clsDefine.enuIni.Ini07PathDivCasse, "");
                    if (s != "")
                        sPath = s;

                    string sFil = _clsQry.DivFilArtApShop(nPos, sPath);
                    StreamWriter sw = new StreamWriter(sFil, false, Encoding.Default);

                    progressBar1.Value = 0;
                    progressBar1.Visible = true;
                    progressBar1.Maximum = Math.Max(1, _tabPos.Rows.Count);
                    progressBar1.Minimum = 0;

                    DataTable tTmpSchema = new clsGenTabTmp().TabTmpDivArtApShop("TabDiv");
                    foreach (DataColumn c in tTmpSchema.Columns)
                        sRig += c.ColumnName + ";";
                    sw.Write(sRig + _clsDef.CRLF);
                    tTmpSchema.Dispose();

                    DataView v = new DataView(_tabPos, "", "pos_art", DataViewRowState.CurrentRows);

                    int iPosCount = 0;
                    int totalPosCount = v.Count;
                    string sArt = "";
                    string sLisCode = (string)yy["tab_lis"];

                    foreach (DataRowView y in v)
                    {
                        iPosCount++;
                        if (iPosCount % 50 == 0 || iPosCount == totalPosCount)
                        {
                            int pct = (int)((iPosCount / (double)(totalPosCount > 0 ? totalPosCount : 1)) * 100);
                            frmWait.UpdateText($"GENERAZIONE FILE CASSA APSHOP: {pct}% ({iPosCount}/{totalPosCount})");
                            progressBar1.Value = Math.Min(iPosCount, progressBar1.Maximum);
                            Application.DoEvents();
                        }

                        if ((string)y["pos_inv"] == _clsDef.DIVDAD)
                        {
                            if ((string)y["pos_art"] != sArt)
                            {
                                sArt = (string)y["pos_art"];

                                s = "";

                                if (((string)y["pos_sta"]).Trim() == _clsDef.STAPRE)
                                {
                                    Console.WriteLine("OK");
                                }
                                else if (((string)y["pos_sta"]).Trim() == _clsDef.STAREP)
                                {
                                    if (((string)y["pos_rep"]).Trim() == "")
                                        s += "Manca Reparto ";
                                }
                                else
                                {
                                    if (((string)y["pos_sta"]).Trim() != _clsDef.STAREP && ((string)y["pos_ean"]).Trim() == "" && ((string)y["pos_art"]).Trim() == "")
                                        s += "Manca EAN ";
                                    if (((string)y["pos_rep"]).Trim() == "")
                                        s += "Manca Reparto ";
                                    if ((decimal)y["pos_prv"] == 0)
                                    {
                                        if (_strCnfPar.Length == 0 || _strCnfPar.Substring(0, 1) != "S")
                                            s += "Manca prezzo ";
                                    }
                                }
                                y["pos_msg"] = s;

                                if (s == "")
                                {
                                    Boolean bb = true;
                                    s = ((string)y["pos_neg"]).Trim();

                                    if (s != "")
                                    {
                                        if (s != (string)yy["tab_cod"])
                                            bb = false;
                                    }
                                    else
                                    {
                                        if (!DBNull.Value.Equals(yy["tab_rep"]) && (string)yy["tab_rep"] != "")
                                        {
                                            bb = false;

                                            string sTip = ((string)yy["tab_rep"]).Substring(0, 1);
                                            string sRep = ((string)yy["tab_rep"]).Substring(1);

                                            if (sRep.Contains((string)y["pos_rep"]))
                                            {
                                                if (sTip == "E")
                                                    bb = false;
                                                else
                                                    bb = true;
                                            }
                                            else
                                            {
                                                if (sTip == "E")
                                                    bb = true;
                                                else
                                                    bb = false;
                                            }
                                        }
                                    }

                                    if (bb)
                                    {
                                        DataTable tTmpMaster = null;
                                        try
                                        {
                                            tTmpMaster = _clsQry.DivArtApShop(sArt, sLisCode);

                                            if (tTmpMaster != null && tTmpMaster.Rows.Count > 0)
                                            {
                                                foreach (DataRow x in tTmpMaster.Rows)
                                                {
                                                    s = Convert.ToString(x["tmp_ean"]);
                                                    if (s.Length == 13 && s.Substring(0, 1) == "2")
                                                        x["tmp_ean"] = s.Substring(0, 12) + "0";

                                                    if ((string)y["pos_off"] != "")
                                                    {
                                                        if ((string)y["pos_cam"] != "")
                                                        {
                                                            s = Convert.ToString(y["pos_cam"]) + "!";
                                                            s += Convert.ToString(y["pos_off"]) + "!";
                                                            s += Convert.ToString(y["pos_tva"]) + "!";
                                                            s += Convert.ToString(y["pos_prv"]).Replace(",", ".") + "!";
                                                            s += Convert.ToString(y["pos_pun"]) + "!";
                                                            s += Convert.ToString(y["pos_xem"]) + "!";
                                                            s += Convert.ToString(y["pos_xen"]) + "!";
                                                            s += ((DateTime)y["pos_dti"]).ToString("yyyyMMdd") + "!";
                                                            s += ((DateTime)y["pos_dtf"]).ToString("yyyyMMdd") + "!";
                                                            s += Convert.ToString(y["pos_neg"]) + "!";
                                                            s += Convert.ToString(y["pos_mix"]) + "!";
                                                            s += Convert.ToString(y["pos_ova"]).Replace(",", ".") + "|";

                                                            x["tmp_cam"] = s;
                                                        }
                                                        else
                                                        {
                                                            x["tmp_off"] = y["pos_off"];
                                                            x["tmp_oft"] = y["pos_tva"];

                                                            x["tmp_ofv"] = y["pos_prv"];
                                                            if ((string)y["pos_tva"] == _clsDef.OFASCO)
                                                            {
                                                                x["tmp_prv"] = y["pos_prv"];
                                                                x["tmp_ofv"] = y["pos_ova"];
                                                            }

                                                            x["tmp_ofm"] = y["pos_xem"];
                                                            x["tmp_ofn"] = y["pos_xen"];
                                                            x["tmp_odi"] = y["pos_dti"];
                                                            x["tmp_odf"] = y["pos_dtf"];
                                                            x["tmp_pun"] = y["pos_pun"];
                                                            x["tmp_mix"] = y["pos_mix"];
                                                        }
                                                    }

                                                    System.Text.StringBuilder sbRig = new System.Text.StringBuilder(256);
                                                    foreach (DataColumn c in tTmpMaster.Columns)
                                                    {
                                                        if (c.DataType.Name == "DateTime" && !DBNull.Value.Equals(x[c.ColumnName]))
                                                            sbRig.Append(((DateTime)x[c.ColumnName]).ToString("dd/MM/yyyy")).Append(';');
                                                        else
                                                            sbRig.Append(Convert.ToString(x[c.ColumnName])).Append(';');
                                                    }
                                                    sw.WriteLine(sbRig.ToString());
                                                }
                                            }
                                        }
                                        finally
                                        {
                                            if (tTmpMaster != null)
                                            {
                                                tTmpMaster.Dispose();
                                                tTmpMaster = null;
                                            }
                                        }

                                        b = true;

                                        string[] aLog = { 
                                                    "POSART",                           //  log_tip
                                                    (string)y["pos_art"],               //  log_art
                                                    (string)y["pos_ean"],               //  log_ean
                                                    (string)y["pos_ard"],               //  log_des
                                                    "",                                 //  log_fil
                                                    "",                                 //  log_acq
                                                    ((decimal)y["pos_prv"]).ToString(), //  log_prv
                                                    (string)y["pos_tvd"]                //  log_msg
                                                        };
                                        listLogBulk.Add(aLog);

                                        if (listLogBulk.Count >= 1000)
                                        {
                                            _clsQry.LogSqlBulk(listLogBulk);
                                            listLogBulk.Clear();
                                        }
                                    }
                                }
                            }
                        }

                        if (iPosCount % 1000 == 0)
                        {
                            GC.Collect(0, GCCollectionMode.Optimized);
                        }
                    }

                    ((TextWriter)sw).Flush();
                    sw.Close();
                    sw.Dispose();

                    if (!b)
                        File.Delete(sFil);
                    else
                    {
                        string sNegFilt = (string)yy["tab_pos"];

                        s = "";
                        for (int n = 1; n <= nPos; n++)
                        {
                            if (sNegFilt.Contains(n.ToString()))
                                File.Copy(sFil, sFil + "_" + n.ToString("0") + ".csv", true);
                        }
                    }

                    if (listLogBulk.Count > 0)
                    {
                        _clsQry.LogSqlBulk(listLogBulk);
                        listLogBulk.Clear();
                    }
                }
            }
            catch (OutOfMemoryException oomEx)
            {
                _clsFun.ErrorLog("PosApShop01_OutOfMemory", oomEx.Message + "\r\n" + oomEx.StackTrace);
                MessageBox.Show("Memoria insufficiente durante la generazione del file cassa ApShop.\r\n" + oomEx.Message,
                    "Errore Memoria", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("PosApShop01", ex.Message + "\r\n" + ex.StackTrace);
                MessageBox.Show("Errore durante la generazione del file cassa ApShop: " + ex.Message,
                    "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (File.Exists(_clsDef.TMPPOSSTOP))
                    File.Delete(_clsDef.TMPPOSSTOP);
            }
        }

        private void PosTesApShop01(string strFilVar)
        {
            if (!File.Exists(_clsDef.TMPPOSSTOP))
            {
                var f = File.Create(_clsDef.TMPPOSSTOP);
                f.Close();
            }

            try
            {
                Boolean b = false;
                string s = "";
                string sRig = "";
                int nPos = 1;
                int.TryParse(strFilVar, out nPos);
                if (nPos <= 0) nPos = 1;

                string sFil = _clsQry.DivFilArtApShop(nPos, _clsDef.TMPPOSFID);
                StreamWriter sw = new StreamWriter(sFil, false, Encoding.Default);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = Math.Max(1, _tabFid.Rows.Count);
                progressBar1.Minimum = 0;

                sRig = "";

                foreach (DataColumn c in _tabFid.Columns)
                    sRig += c.ColumnName + ";";
                sw.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in _tabFid.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    if ((string)y["TesInv"] == _clsDef.DIVDAD)
                    {
                        sRig = "";

                        foreach (DataColumn c in _tabFid.Columns)
                            sRig += Convert.ToString(y[c.ColumnName]) + ";";
                        sw.Write(sRig + _clsDef.CRLF);
                        b = true;
                    }
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                if (!b)
                    File.Delete(sFil);
                else
                {
                    s = "";
                    for (int n = 1; n <= nPos; n++)
                    {
                        File.Copy(sFil, sFil + "_" + n.ToString("0") + ".csv", true);
                    }
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("PosTesApShop01", ex.Message);
            }
            finally
            {
                if (File.Exists(_clsDef.TMPPOSSTOP))
                    File.Delete(_clsDef.TMPPOSSTOP);
            }
        }

        private void PosCliApShop01(string strFilVar)
        {
            if (!File.Exists(_clsDef.TMPPOSSTOP))
            {
                var f = File.Create(_clsDef.TMPPOSSTOP);
                f.Close();
            }

            try
            {
                Boolean b = false;
                string s = "";
                string sRig = "";
                DataRow[] j;
                int nPos = 1;
                int.TryParse(strFilVar, out nPos);
                if (nPos <= 0) nPos = 1;

                s = "SELECT * FROM AnaCliDestinazioni";
                DataTable tCld = _clsFun.FillTabSql("AnaCliDestinazioni", s, false, _strConSql);

                string sFil = _clsQry.DivFilArtApShop(nPos, _clsDef.TMPPOSCLI);
                StreamWriter sw = new StreamWriter(sFil, false, Encoding.Default);

                string sFilCld = _clsQry.DivFilArtApShop(nPos, _clsDef.TMPPOSCLD);
                StreamWriter swCld = new StreamWriter(sFilCld, false, Encoding.Default);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = Math.Max(1, _tabCli.Rows.Count);
                progressBar1.Minimum = 0;

                sRig = "";
                foreach (DataColumn c in _tabCli.Columns)
                    sRig += c.ColumnName + ";";
                sw.Write(sRig + _clsDef.CRLF);

                sRig = "";
                foreach (DataColumn c in tCld.Columns)
                    sRig += c.ColumnName + ";";
                swCld.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in _tabCli.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    if ((string)y["cli_div"] == _clsDef.DIVDAD)
                    {
                        sRig = "";

                        foreach (DataColumn c in _tabCli.Columns)

                            if (c.ColumnName == "cli_ind")
                            {
                                s = "";
                                if (!DBNull.Value.Equals(y["cli_ncv"]))
                                    s = ((string)y["cli_ncv"]).Trim();
                                sRig += Convert.ToString(y[c.ColumnName]) + " " + s + ";";
                            }
                            else
                                sRig += Convert.ToString(y[c.ColumnName]) + ";";
                        sw.Write(sRig + _clsDef.CRLF);

                        j = tCld.Select("cld_cli='" + y["cli_cod"] + "'");

                        if(j.Length > 0)
                        {
                            for(int i = 0; i < j.Length; i++)
                            {
                                sRig = "";

                                foreach (DataColumn c in tCld.Columns)
                                    sRig += Convert.ToString(j[i][c.ColumnName]) + ";";
                                swCld.Write(sRig + _clsDef.CRLF);
                            }
                        }

                        b = true;
                    }
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                ((TextWriter)swCld).Flush();
                swCld.Close();
                swCld.Dispose();

                if (tCld != null)
                {
                    tCld.Dispose();
                    tCld = null;
                }

                if (!b)
                {
                    File.Delete(sFilCld);
                    File.Delete(sFil);
                }
                else
                {
                    s = "";
                    for (int n = 1; n <= nPos; n++)
                    {
                        File.Copy(sFil, sFil + "_" + n.ToString("0") + ".csv", true);
                        File.Copy(sFilCld, sFilCld + "_" + n.ToString("0") + ".csv", true);
                    }
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("PosCliApShop01", ex.Message);
            }
            finally
            {
                if (File.Exists(_clsDef.TMPPOSSTOP))
                    File.Delete(_clsDef.TMPPOSSTOP);
            }
        }

        private void oldPosOffTraApShop01(string strFilVar)
        {
            if (!File.Exists(_clsDef.TMPPOSSTOP))
            {
                var f = File.Create(_clsDef.TMPPOSSTOP);
                f.Close();
            }

            // Apro tabella negozi 

            Boolean b = false;
            string s = "";
            string sRig = "";
            int nPos = Convert.ToInt16(strFilVar);

            foreach (DataRow yy in _tabNeg.Rows)
            {
                string sFil = _clsQry.DivFilArtApShop(nPos, _clsDef.TMPPOSOTR);
                StreamWriter sw = new StreamWriter(sFil, true);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = _tabPos.Rows.Count;
                progressBar1.Minimum = 0;

                sRig = "";

                foreach (DataColumn c in _tabOtr.Columns)
                    sRig += c.ColumnName + ";";
                sw.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in _tabOtr.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    string sNeg = (string)y["otr_neg"];

                    if ((DBNull.Value.Equals(y["otr_neg"]) || ((string)y["otr_neg"]).Trim() == "") || (string)yy["tab_cod"] == (string)y["otr_neg"])
                    {
                        sRig = "";

                        foreach (DataColumn c in _tabOtr.Columns)
                        {
                            Console.WriteLine("xxx");

                            if (c.DataType.Name == "DateTime")
                                sRig += ((DateTime)y[c]).ToShortDateString() + ";";
                            else
                                sRig += Convert.ToString(y[c.ColumnName]) + ";";

                        }
                        sw.Write(sRig + _clsDef.CRLF);
                        b = true;
                    }
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                if (!b)
                    File.Delete(sFil);
                else
                {
                    string sNegFilt = ((string)yy["tab_pos"]).Trim();

                    s = "";
                    for (int n = 1; n <= nPos; n++)
                    {
                        if (sNegFilt.Contains(n.ToString()))
                            File.Copy(sFil, sFil + "_" + n.ToString("0") + ".csv");
                    }
                }
            }

            if (File.Exists(_clsDef.TMPPOSSTOP))
                File.Delete(_clsDef.TMPPOSSTOP);
        }

        private void PosOffTraApShop01(string strFilVar)
        {
            if (!File.Exists(_clsDef.TMPPOSSTOP))
            {
                var f = File.Create(_clsDef.TMPPOSSTOP);
                f.Close();
            }

            // Apro tabella offerte 

            Boolean b = false;
            string s = "";
            string sRig = "";
            int nPos = Convert.ToInt16(strFilVar);

            s = "SELECT * FROM GesOffTranDate";
            DataTable tOtd = _clsFun.FillTabSql("GesOffTranDate", s, false, _strConSql);

            foreach (DataRow yy in _tabNeg.Rows)
            {
                string sFil = _clsQry.DivFilArtApShop(nPos, _clsDef.TMPPOSOTR);
                StreamWriter sw = new StreamWriter(sFil, true);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = _tabPos.Rows.Count;
                progressBar1.Minimum = 0;

                sRig = "";

                foreach (DataColumn c in _tabOtr.Columns)
                    sRig += c.ColumnName + ";";
                sw.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in _tabOtr.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    if (!(Boolean)y["otr_ann"])
                    {
                        string sNeg = (string)y["otr_neg"];

                        if ((DBNull.Value.Equals(y["otr_neg"]) || ((string)y["otr_neg"]).Trim() == "") || (string)yy["tab_cod"] == (string)y["otr_neg"])
                        {
                            foreach (DataRow yyy in tOtd.Rows)
                            {
                                sRig = "";
                                if ((string)yyy["otd_cod"] == (string)y["otr_cod"])
                                {
                                    DateTime dDti = (DateTime)yyy["otd_dti"];
                                    DateTime dDtf = (DateTime)yyy["otd_dtf"];

                                    foreach (DataColumn c in _tabOtr.Columns)
                                    {
                                        Console.WriteLine("xxx");

                                        if (c.DataType.Name == "DateTime")
                                        {
                                            if (c.ColumnName == "otr_dti")
                                                sRig += (dDti).ToShortDateString() + ";";
                                            else if (c.ColumnName == "otr_dtf")
                                                sRig += (dDtf).ToShortDateString() + ";";
                                            else
                                                sRig += ((DateTime)y[c]).ToShortDateString() + ";";
                                        }
                                        else
                                            sRig += Convert.ToString(y[c.ColumnName]) + ";";

                                    }
                                    sw.Write(sRig + _clsDef.CRLF);
                                }
                            }
                            b = true;
                        }
                    }
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                if (!b)
                    File.Delete(sFil);
                else
                {
                    string sNegFilt = ((string)yy["tab_pos"]).Trim();

                    s = "";
                    for (int n = 1; n <= nPos; n++)
                    {
                        if (sNegFilt.Contains(n.ToString()))
                            File.Copy(sFil, sFil + "_" + n.ToString("0") + ".csv");
                    }
                }
            }

            if (File.Exists(_clsDef.TMPPOSSTOP))
                File.Delete(_clsDef.TMPPOSSTOP);
        }

        private void BilOmega(string strFilVar)
        {
            DataRow[] j;

            string sOmegaBarcodeLun = _clsFun.ParGet(clsDefine.enuParametri.ParOmegaBarcodeLun, _strConSql);
            string sParBilance = _clsFun.ParGet(clsDefine.enuParametri.Par028ParBilance, _strConSql);

            clsBilOmega clsBil = new clsBilOmega();
            clsBil._strOmegaBarcodeLun = sOmegaBarcodeLun;
            clsBil._strParBilance = sParBilance;

            string sFil = _clsDef.TMPBILVAR;
            Boolean b = false;
            string s = "";

            s = "SELECT * FROM TabRepBilance";
            DataTable tReb = _clsFun.FillTabSql("TReb", s, false, _strConSql);

            //sFil = _clsDef.TMPBILVAR;

            //StreamWriter sw = new StreamWriter(sFil, false, Encoding.UTF8);

            StreamWriter sw = new StreamWriter(sFil, false, Encoding.GetEncoding(1250));

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = _tabPos.Rows.Count;
            progressBar1.Minimum = 0;

            DataView v = new DataView(_tabPos, "bil_bil=1", "bil_reb, bil_plu", DataViewRowState.CurrentRows);

            if (v.Count > 0)
            {
                DataTable t = v.ToTable();

                foreach (DataRow y in t.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    if ((string)y["pos_inv"] == _clsDef.DIVDAD && (Boolean)y["bil_bil"])
                    {
                        s = "";
                        if (((string)y["bil_plu"]).Trim() == "")
                            s += "Manca PLU ";
                        if (((string)y["pos_ean"]).Trim() == "")
                            s += "Manca EAN ";
                        if (((string)y["bil_reb"]).Trim() == "")
                            s += "Manca Reparto bilancia ";
                        if ((decimal)y["pos_prv"] == 0)
                            s += "Manca prezzo ";

                        y["pos_msg"] = s;

                        string sIng = "";

                        if (s == "")
                        {
                            s = _clsDef.PATHINGREDIENTI + "et01_" + y["pos_art"] + ".rtf";
                            if (File.Exists(s))
                            {
                                RichTextBox rch = new RichTextBox();
                                rch.LoadFile(s);
                                sIng = rch.Text;

                                Console.WriteLine("aaaaaaa");

                                sIng = _clsFun.StrSplit(sIng, 40);
                            }
                            if ((decimal)y["bil_tar"] == 0)
                            {
                                j = tReb.Select("tab_cod='" + y["bil_reb"] + "'");
                                if (j.Length > 0)
                                    y["bil_tar"] = (decimal)j[0]["tab_tar"];
                            }
                            string sRig = clsBil.BilArtRow(y, sIng);
                            y["pos_inv"] = _clsDef.DIVDIV;
                            sw.Write(sRig + _clsDef.CRLF);

                            b = true;
                        }
                    }
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                if (b)
                {
                    string sOld = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (File.Exists(strFilVar))
                        File.Delete(strFilVar);
                    File.Copy(sFil, strFilVar);
                    File.Move(sFil, sOld);
                }

                Process p = new Process();
                p.StartInfo.FileName = "C:\\OMEGA\\AP_OMEGA.BAT";
                p.StartInfo.Arguments = "-r";
                p.StartInfo.ErrorDialog = true;
                p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                p.Start();
                p.WaitForExit(1000 * 60 * 7);    // wait up to 1 minutes.
                Thread.Sleep(8000);
                //for (int i = 0; i <= 10; i++)
                //{
                //    Thread.Sleep(8000);
                //    if (!File.Exists(strFilVar))
                //        break;
                //}

                //if (File.Exists(strFilVar))
                //    MessageBox.Show("Aggiornamento a bilancia non riuscito!");
                if (File.Exists(strFilVar))
                    File.Delete(strFilVar);
            }
        }

        private void BilBizerba(string strFilVar, string strTip)
        {
            DataRow[] j;

            clsBilBizerba clsBil = new clsBilBizerba();

            string sFil = _clsDef.TMPBILVAR;
            string sFilIng = _clsDef.TMPBILING;
            Boolean b = false;
            string s = "";

            StreamWriter sw = new StreamWriter(sFil, true);
            StreamWriter swIng = new StreamWriter(sFilIng, true);

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = _tabPos.Rows.Count;
            progressBar1.Minimum = 0;

            DataView v = new DataView(_tabPos, "bil_bil=1", "bil_reb, bil_plu", DataViewRowState.CurrentRows);

            if (v.Count > 0)
            {
                DataTable t = v.ToTable();

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = t.Rows.Count > 0 ? t.Rows.Count : 1;
                progressBar1.Minimum = 0;

                int totalBil = t.Rows.Count;
                int idxBil = 0;

                foreach (DataRow y in t.Rows)
                {
                    idxBil++;
                    if (idxBil % 25 == 0 || idxBil == totalBil)
                    {
                        progressBar1.Value = Math.Min(idxBil, progressBar1.Maximum);
                        Application.DoEvents();
                    }

                    string posInv = (y["pos_inv"] != DBNull.Value) ? (string)y["pos_inv"] : "";
                    bool isBil = (y["bil_bil"] != DBNull.Value) && Convert.ToBoolean(y["bil_bil"]);

                    if ((posInv == _clsDef.DIVDIV || posInv == _clsDef.DIVDAD) && isBil)
                    {
                        s = "";
                        if (((string)y["bil_plu"]).Trim() == "")
                            s += "Manca PLU ";
                        if (((string)y["pos_ean"]).Trim() == "")
                            s += "Manca EAN ";
                        if (((string)y["bil_reb"]).Trim() == "")
                            s += "Manca Reparto bilancia ";
                        if ((decimal)y["pos_prv"] == 0)
                            s += "Manca prezzo ";

                        y["pos_msg"] = s;

                        if (s == "")
                        {
                            string sRig = "";
                            y["pos_inv"] = _clsDef.DIVDIV;
                            if(strTip == "2")
                                sRig = clsBil.BilArtRow2(y);
                            else
                                sRig = clsBil.BilArtRow(y);
                            sw.Write(sRig + _clsDef.CRLF);

                            s = _clsDef.PATHINGREDIENTI + "et01_" + y["pos_art"] + ".rtf";
                            if (File.Exists(s))
                            {
                                string sIng = "";

                                using (RichTextBox rch = new RichTextBox())
                                {
                                    rch.LoadFile(s);
                                    sIng = rch.Text;
                                }

                                if (sIng.Length > 5)
                                {
                                    sIng = _clsFun.StrSplit(sIng, 50);

                                    string s2 = "";
                                    int i = 0;
                                    string[] a = sIng.Split('|');
                                    foreach (string ss in a)
                                    {
                                        s = (ss.Trim() + new string(' ', 50)).Substring(0, 50);

                                        if (i < 2)
                                            s = "5" + s;
                                        else
                                            s = "1" + s;

                                        s2 += s;
                                        i++;

                                        if (i > 10)
                                            break;
                                    }

                                    s2 = (string)y["bil_plu"] + s2;

                                    s2 = (s2 + new string(' ', 517)).Substring(0,516) + DateTime.Today.ToString("ddMMyy");

                                    swIng.Write(s2 + _clsDef.CRLF);
                                }
                            }

                            b = true;
                        }
                    }
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                ((TextWriter)swIng).Flush();
                swIng.Close();
                swIng.Dispose();

                string sFilBiz = "";

                if (b)
                {
                    string[] a = strFilVar.Split(',');

                    if (a.Length < 2)
                        MessageBox.Show("Configurazione tabella bilancia Bizerba con paramtro mancante (variazioni, ingredienti)!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        string sVar = a[0];
                        string sIng = Path.GetDirectoryName(sVar) + "\\" + a[1];

                        string sOld = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                        if (File.Exists(sVar))
                            File.Delete(sVar);
                        File.Copy(sFil, sVar);
                        File.Move(sFil, sOld);
                        sFilBiz = sVar;

                        sOld = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFilIng) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                        if (File.Exists(sIng))
                            File.Delete(sIng);
                        File.Copy(sFilIng, sIng);
                        File.Move(sFilIng, sOld);
                    }

                    Process p = new Process();
                    p.StartInfo.FileName = "C:\\ApProject\\Temp\\DivBilance\\AP_BIZERBA.BAT";
                    p.StartInfo.Arguments = "-r";
                    p.StartInfo.ErrorDialog = true;
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                    p.Start();
                    p.WaitForExit(1000 * 60 * 7);    // wait up to 1 minutes.
                    Thread.Sleep(8000);
                    for (int i = 0; i <= 10; i++)
                    {
                        Thread.Sleep(8000);
                        if (!File.Exists(strFilVar))
                            break;
                    }

                    if (File.Exists(strFilVar))
                        MessageBox.Show("Aggiornamento a bilancia non riuscito!");
                    if (File.Exists(strFilVar))
                        File.Delete(strFilVar);
                }
            }
        }

        private void BilBizWinvarp(string strFilVar)
        {
            string s = "";
            DataRow[] j;
            Boolean b = true;
            Boolean bIng = false;

            string sPar031LottiPos = _clsFun.ParGet(clsDefine.enuParametri.Par031Lotti2Pos, _strConSql);

            Boolean bIng2Varp = false;
            string[] a = strFilVar.Split(',');
            if (a.Length > 3 && a[3] == "winvarp+ing")
                bIng2Varp = true;
            string sRecRebPlu = "";
            if (a.Length > 4 && a[4] != "")
                sRecRebPlu = a[4].Trim();

            s = a[0];

            if (File.Exists(s))
            {
                try
                {
                    File.Delete(s);
                }
                catch (IOException e)
                {
                    MessageBox.Show("Dati bilancia non inviati, riprovare.", "BIZERBA");
                    b = false;
                }
            }
            if (b)
            {
                s = Path.GetDirectoryName(a[0]) + "\\" + a[1];

                if (File.Exists(s))
                {
                    try
                    {
                        File.Delete(s);
                    }
                    catch (IOException e)
                    {
                        MessageBox.Show("Dati bilancia non inviati, riprovare.", "BIZERBA");
                        b = false;
                    }
                }
            }
            //b = false;
            if (b)
            {
                s = "SELECT * FROM TabRepBilance";
                DataTable tReb = _clsFun.FillTabSql("TabRepBilance", s, false, _strConSql);

                s = "SELECT * FROM TabReparti";
                DataTable tRep = _clsFun.FillTabSql("TabReparti", s, false, _strConSql);

                DataColumn[] keys = new DataColumn[1];
                keys[0] = tReb.Columns["tab_cod"];
                tReb.PrimaryKey = keys;

                clsBilBizWinvarp clsBil = new clsBilBizWinvarp();
                clsBil._bolIng2Varp = bIng2Varp;
                clsBil._strIngTip = _clsFun.ParGet(clsDefine.enuParametri.Par021GesIngredienti, _strConSql);
                clsBil._strRecRebPlu = sRecRebPlu;
                clsBil._strPar031Lotti2Pos = sPar031LottiPos;

                //string sFil = _clsDef.TMPBILVAR;

                string sFil = _clsDef.TMPBILVAR;
                s = _clsFun.FileIni("R", clsDefine.enuIni.Ini08PathDivBilance, "");
                if (s != "")
                    sFil = s;

                string sFilIng = _clsDef.TMPBILING;
                s = _clsFun.FileIni("R", clsDefine.enuIni.Ini08PathDivBilance, "");
                if (s != "")
                {
                    string ss = Path.GetFileName(_clsDef.TMPBILING);
                    s = Path.GetDirectoryName(s) + "\\" + ss;
                    sFilIng = s;
                }

                b = false;

                StreamWriter sw = new StreamWriter(sFil, true);
                StreamWriter swIng = new StreamWriter(sFilIng, true);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = _tabPos.Rows.Count;
                progressBar1.Minimum = 0;

                DataView v = new DataView(_tabPos, "bil_bil=1 AND pos_inv='" + _clsDef.DIVDAD + "'", "bil_reb, bil_plu", DataViewRowState.CurrentRows);

                if (v.Count > 0)
                {
                    //DataTable t = v.Table;
                    DataTable t = v.ToTable();

                    foreach (DataRow y in t.Rows)
                    {
                        progressBar1.Increment(1);
                        Application.DoEvents();

                        j = tReb.Select("tab_cod='" + (string)y["bil_reb"] + "'");
                        if (j.Length > 0 && !DBNull.Value.Equals(j[0]["tab_gru"]) && (string)j[0]["tab_gru"] != "")
                            y["bil_gru"] = (string)j[0]["tab_gru"];

                        j = tRep.Select("tab_cod='" + (string)y["pos_rep"] + "'");
                        if (j.Length > 0 && !DBNull.Value.Equals(j[0]["tab_bil"]) && (string)j[0]["tab_bil"] != "")
                            y["bil_rbi"] = (string)j[0]["tab_bil"];

                        //if ((string)y["pos_inv"] == _clsDef.DIVDIV && (Boolean)y["bil_bil"])
                        if ((string)y["pos_inv"] == _clsDef.DIVDAD && (Boolean)y["bil_bil"])
                        {
                            s = "";
                            if (((string)y["bil_plu"]).Trim() == "")
                                s += "Manca PLU ";
                            if (((string)y["pos_ean"]).Trim() == "")
                                s += "Manca EAN ";
                            if (((string)y["bil_reb"]).Trim() == "")
                                s += "Manca Reparto bilancia ";
                            if ((decimal)y["pos_prv"] == 0)
                                s += "Manca prezzo ";

                            y["pos_msg"] = s;

                            if (s == "")
                            {
                                //y["pos_inv"] = _clsDef.DIVDIV;
                                string sRig = clsBil.BilArtRow(y);
                                sw.Write(sRig + _clsDef.CRLF);

                                s = "SELECT * FROM AnaArtIngredienti WHERE ing_art='" + (string)y["pos_art"] + "' AND ing_ann=0";
                                DataTable tIng = _clsFun.FillTabSql("TabIng", s, false, _strConSql);
                                sRig = "";
                                if (tIng.Rows.Count > 0)
                                {
                                    sRig = clsBil.BilArtIng(y, tIng);
                                    if (sRig.Length > 0)
                                        bIng = true;
                                }
                                swIng.Write(sRig + _clsDef.CRLF);

                                b = true;
                            }
                        }
                    }

                    ((TextWriter)sw).Flush();
                    sw.Close();
                    sw.Dispose();

                    ((TextWriter)swIng).Flush();
                    swIng.Close();
                    swIng.Dispose();

                    if (!bIng)
                        File.Delete(sFilIng);

                    string sFilBiz = "";

                    if (b)
                    {
                        a = strFilVar.Split(',');

                        if (a.Length < 2)
                            MessageBox.Show("Configurazione tabella bilancia Bizerba con paramtro mancante (variazioni, ingredienti, batch)!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (a.Length < 3)
                            MessageBox.Show("Configurazione tabella bilancia Bizerba con paramtro BATCH mancante (variazioni, ingredienti, batch)!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        else if (!Directory.Exists(Path.GetDirectoryName(a[0])))
                            MessageBox.Show("Percorso '" + Path.GetDirectoryName(a[0]) + "' non disponibile!", "CONTROLLO RETE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        else
                        {
                            string sVar = a[0];
                            string sIng = Path.GetDirectoryName(sVar) + "\\" + a[1];
                            string sBath = a[2];

                            string sOld = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                            if (File.Exists(sVar))
                                File.Delete(sVar);
                            File.Copy(sFil, sVar);
                            File.Move(sFil, sOld);
                            sFilBiz = sVar;

                            if (File.Exists(sFilIng))
                            {
                                sOld = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFilIng) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                                if (File.Exists(sIng))
                                    File.Delete(sIng);
                                File.Copy(sFilIng, sIng);
                                File.Move(sFilIng, sOld);
                            }

                            if (sBath == "SI")
                            {
                                Process p = new Process();
                                p.StartInfo.FileName = "C:\\ApProject\\Temp\\DivBilance\\AP_BIZERBA.BAT";
                                p.StartInfo.Arguments = "-r";
                                p.StartInfo.ErrorDialog = true;
                                p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                                p.Start();
                                p.WaitForExit(1000 * 60 * 7);    // wait up to 1 minutes.
                                Thread.Sleep(7000);
                                for (int i = 0; i <= 10; i++)
                                {
                                    Thread.Sleep(7000);
                                    if (!File.Exists(strFilVar))
                                        break;
                                }

                                if (File.Exists(strFilVar))
                                    MessageBox.Show("Aggiornamento a bilancia non riuscito!");
                                if (File.Exists(strFilVar))
                                    File.Delete(strFilVar);
                            }
                        }
                    }
                }
            }
            else
            {
                ////20190515 Seck se ci sono problemi metto da divulgare
                //DataView v = new DataView(_tabPos, "bil_bil=1", "bil_reb, bil_plu", DataViewRowState.CurrentRows);

                //if (v.Count > 0)
                //{
                //    DataTable t = v.Table;

                //    foreach (DataRow y in t.Rows)
                //    {
                //        y["pos_inv"] = _clsDef.DIVDAD;                        
                //    }
                //}
            }

            //_strRet = "NO";
        }

        private void DivTerm()
        {
            string s = "";
            Boolean b = true;
            DataTable tTer = _clsQry.DivTerm();
            DataRow[] j;
            string sWheFor = "";
            ArrayList aTer = new ArrayList();

            if (tTer.Rows.Count == 0)
                b = false;
            else
            {
                foreach (DataRow y in tTer.Rows)
                {
                    if ((string)y["trm_tip"] == "F")
                        sWheFor += "lia_for='" + (string)y["trm_cod"] + "' OR ";
                    else if ((string)y["trm_tip"] == "T")
                    {
                        aTer.Add(((string)y["trm_cod"]).Trim());
                        if (((string)y["trm_div"]).Trim() == "")
                        {
                            b = false;
                           // MessageBox.Show("Nome file terminalino non definito!");
                        }

                        if (((string)y["trm_cod"]).Trim() != _clsDef.TERMTDIV)
                        {
                            if (((string)y["trm_bat"]).Trim() == "")
                            {
                                b = false;
                                MessageBox.Show("Nome batch terminalino non definito!");
                            }
                            else if (!File.Exists(((string)y["trm_bat"]).Trim()))
                            {
                                b = false;
                                MessageBox.Show("Nome batch terminalino '" + ((string)y["trm_bat"]).Trim() + "' non presente!");
                            }
                        }
                    }
                }
                if (sWheFor != "")
                    sWheFor = "WHERE " + sWheFor.Substring(0, sWheFor.Length - 4);
                else
                    b = false;
            }

            if (b)
            {
                s = "SELECT ";
                s += "GesLisAcquisto.lia_art, ";
                s += "GesLisAcquisto.lia_for, ";
                s += "GesLisAcquisto.lia_dti, ";
                s += "GesLisAcquisto.lia_tip, ";
                s += "GesLisAcquisto.lia_cos, ";
                s += "GesLisAcquisto.lia_pxc, ";
                s += "GesLisAcquisto.lia_cxp, ";
                s += "GesLisAcquisto.lia_arf, ";
                s += "AnaFornitori.for_des AS LiaFod ";
                s += "FROM GesLisAcquisto LEFT OUTER JOIN AnaFornitori ON GesLisAcquisto.lia_for = AnaFornitori.for_cod ";
                s += sWheFor + " ";
                s += "ORDER BY lia_dti DESC";

                DataTable tLia = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);

                foreach(string sTer in aTer)
                {
                    if (sTer == _clsDef.TERMMEMOR)
                        DivTermMemor(tTer, sTer, tLia);
                    else if (sTer == _clsDef.TERMCSV1)
                        DivTermCsv1(tTer, sTer, tLia);
                    else if (sTer == _clsDef.TERMTDIV)
                        DivTermApPhone(tTer, sTer, tLia);
                }
            }
        }

        private void DivTermMemor(DataTable tabTer, string strTer, DataTable tabArf)
        {
            clsTermMemor clsTerm = new clsTermMemor();

            string s = "";
            Boolean b = false;
            string sFil = "";
            string sFilVar = "";
            string sFilBatch = "";
            DataRow[] j;

            j = tabTer.Select("trm_tip='T' AND trm_cod='" + strTer + "'");
            if (j.Length > 0)
            {
                sFilVar = ((string)j[0]["trm_div"]).Trim();
                sFil = _clsDef.PATHMEMOR + Path.GetFileName(((string)j[0]["trm_div"]).Trim());
                sFilBatch = ((string)j[0]["trm_bat"]).Trim();
            }

            if (sFilVar != "")
            {
                StreamWriter sw = new StreamWriter(sFil, true);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = _tabPos.Rows.Count;
                progressBar1.Minimum = 0;

                //ArrayList aMix = new ArrayList();

                foreach (DataRow y in _tabPos.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    if ((string)y["pos_sta"] != "R" && ((string)y["pos_inv"] == _clsDef.DIVDIV || (string)y["pos_inv"] == _clsDef.DIVDAD))
                    {
                        s = "";
                        if (((string)y["pos_ean"]).Trim() == "")
                            s += "Manca EAN ";

                        y["pos_msg"] = s;

                        if (s == "")
                        {
                            j = tabArf.Select("lia_art='" + (string)y["pos_art"] + "'");
                            if (j.Length > 0)
                            {
                                string sRig = clsTerm.TermArtRow(y, j);
                                y["pos_inv"] = _clsDef.DIVDIV;
                                sw.Write(sRig + _clsDef.CRLF);
                                b = true;
                            }
                            else
                                Console.WriteLine("aaaa");
                        }
                    }
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                if (b)
                {
                    string sOld = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (File.Exists(sOld))
                        File.Delete(sOld);
                    if (File.Exists(sFilVar))
                        File.AppendAllText(sFilVar, File.ReadAllText(sFil));
                    else
                        File.Copy(sFil, sFilVar);
                    File.Move(sFil, sOld);


                    Directory.SetCurrentDirectory(Path.GetDirectoryName(sFilBatch));

                    Process p = new Process();
                    p.StartInfo.FileName = sFilBatch;
                    p.StartInfo.Arguments = "-r";
                    p.StartInfo.ErrorDialog = true;
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                    p.Start();
                    p.WaitForExit(1000 * 60 * 1);    // wait up to 1 minutes.
                    if (File.Exists(sFilVar))
                        MessageBox.Show("Invio a terminalino non riuscito!");
                }
            }
        }

        private void DivTermCsv1(DataTable tabTer, string strTer, DataTable tabArf)
        {
            clsTermCsv1 clsTerm = new clsTermCsv1();

            string s = "";
            Boolean b = false;
            string sRig = "";
            string sFil = "";
            string sFilVar = "";
            string sFilBatch = "";
            DataRow[] j;

            j = tabTer.Select("trm_tip='T' AND trm_cod='" + strTer + "'");
            if (j.Length > 0)
            {
                sFilVar = ((string)j[0]["trm_div"]).Trim();
                sFil = _clsDef.PATHMEMOR + Path.GetFileName(((string)j[0]["trm_div"]).Trim());
                sFilBatch = ((string)j[0]["trm_bat"]).Trim();
            }

            if (sFilVar != "")
            {
                StreamWriter sw = new StreamWriter(sFil, true);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = _tabPos.Rows.Count;
                progressBar1.Minimum = 0;

                //ArrayList aMix = new ArrayList();

                sRig = clsTerm.TermArtTes();
                sw.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in _tabPos.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    if ((string)y["pos_inv"] == _clsDef.DIVDIV || (string)y["pos_inv"] == _clsDef.DIVDAD)
                    {
                        s = "";
                        if (((string)y["pos_ean"]).Trim() == "")
                            s += "Manca EAN ";

                        y["pos_msg"] = s;

                        if (s == "")
                        {
                            j = tabArf.Select("lia_art='" + (string)y["pos_art"] + "'");
                            if (j.Length > 0)
                            {
                                sRig = clsTerm.TermArtRow(y, j);
                                y["pos_inv"] = _clsDef.DIVDIV;
                                sw.Write(sRig + _clsDef.CRLF);
                                b = true;
                            }
                            else
                                Console.WriteLine("aaaa");
                        }
                    }
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                if (b)
                {
                    string sOld = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (File.Exists(sOld))
                        File.Delete(sOld);
                    if (File.Exists(sFilVar))
                        File.AppendAllText(sFilVar, File.ReadAllText(sFil));
                    else
                        File.Copy(sFil, sFilVar);
                    File.Move(sFil, sOld);

                    Directory.SetCurrentDirectory(Path.GetDirectoryName(sFilBatch));

                    Process p = new Process();
                    p.StartInfo.FileName = sFilBatch;
                    p.StartInfo.Arguments = "-r";
                    p.StartInfo.ErrorDialog = true;
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                    p.Start();
                    p.WaitForExit(1000 * 60 * 1);    // wait up to 1 minutes.
                    if (File.Exists(sFilVar))
                        MessageBox.Show("Invio a phone non riuscito!");
                }
            }

        }

        private void DivTermApPhone(DataTable tabTer, string strTer, DataTable tabLia)
        {
            clsTermCsv2 clsTerm = new clsTermCsv2();

            string s = "";
            Boolean b = false;
            string sRig = "";
            string sArt = "";
            string sPath = "";
            string sOld = "Old\\";
            string sFilArt = ""; 
            string sFilEan = "";
            string sFilLia = "";
            //string sFilBatch = "";
            DataRow[] j;

            DataRow[] jTer = tabTer.Select("trm_tip='T' AND trm_cod='" + strTer + "'");
            if (jTer.Length > 0)
            {
                sPath = Path.GetDirectoryName((string)jTer[0]["trm_div"]) + "\\";

                sFilArt = sPath + sOld + "art_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
                sFilEan = sPath + sOld + "ean_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
                sFilLia = sPath + sOld + "lia_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";

                //sFilBatch = ((string)jTer[0]["trm_bat"]).Trim();

                DataTable tRep = _clsFun.FillTabSql("TabRep", "SELECT * FROM TabReparti", false, _strConSql);

                s = sPath + "Temp\\";

                StreamWriter swArt = new StreamWriter(sFilArt, true);
                StreamWriter swEan = new StreamWriter(sFilEan, true);
                StreamWriter swLia = new StreamWriter(sFilLia, true);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = _tabPos.Rows.Count;
                progressBar1.Minimum = 0;

                sRig = clsTerm.TermArtTes();
                swArt.Write(sRig + _clsDef.CRLF);

                sRig = clsTerm.TermEanTes();
                swEan.Write(sRig + _clsDef.CRLF);

                sRig = clsTerm.TermLiaTes();
                swLia.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in _tabPos.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    if ((string)y["pos_inv"] == _clsDef.DIVDIV || (string)y["pos_inv"] == _clsDef.DIVDAD)
                    {
                        //s = "";
                        //if (((string)y["pos_ean"]).Trim() == "")
                        //    s += "Manca EAN ";
                        //y["pos_msg"] = s;


                        if (sArt != (string)y["pos_art"])
                        {
                            sArt = (string)y["pos_art"];

                            sRig = clsTerm.TermArtRow(y, tRep);
                            y["pos_inv"] = _clsDef.DIVDIV;
                            swArt.Write(sRig + _clsDef.CRLF);
                            b = true;

                            ArrayList aFor = new ArrayList();

                            j = tabLia.Select("lia_art='" + y["pos_art"] + "'", "lia_dti DESC");
                            for (int i = 0; i < j.Length; i++)
                            {
                                if (aFor.IndexOf(j[i]["lia_for"]) < 0)
                                {
                                    aFor.Add(j[i]["lia_for"]);
                                    sRig = clsTerm.TermLiaRow(y, j[i]);
                                    if (sRig.Length > 0)
                                        swLia.Write(sRig + _clsDef.CRLF);
                                }
                            }
                        }

                        sRig = clsTerm.TermEanRow(y);
                        if(sRig.Length > 0)
                            swEan.Write(sRig + _clsDef.CRLF);
                    }
                }

                ((TextWriter)swArt).Flush();
                swArt.Close();
                swArt.Dispose();

                ((TextWriter)swEan).Flush();
                swEan.Close();
                swEan.Dispose();

                ((TextWriter)swLia).Flush();
                swLia.Close();
                swLia.Dispose();

                if (!b)
                {
                    File.Delete(sFilArt);
                    File.Delete(sFilEan);
                    File.Delete(sFilLia);
                }
                else
                {
                    s = "SELECT * FROM TabTerm WHERE tab_cod='" + (string)jTer[0]["trm_cod"] + "'";
                    DataTable t = _clsFun.FillTabSql("TabTerm", s, true, _strConSql);
                    if (t.Rows.Count > 0 && ((string)t.Rows[0]["tab_hst"]).Trim() != "")
                    {
                        s = (string)t.Rows[0]["tab_fpt"];
                        string[] aTer = s.Split(';');
                        s = aTer[1];
                        aTer = s.Split(',');

                        foreach (string sTer in aTer)
                        {
                            s = Path.GetFileName(sFilArt);
                            s = s.Substring(0, 4) + sTer + "_" + s.Substring(4);
                            string sFil = sPath + s;
                            if (File.Exists(sFil))
                                File.Delete(sFil);
                            File.Copy(sFilArt, sFil);
                        }

                        foreach (string sTer in aTer)
                        {
                            s = Path.GetFileName(sFilEan);
                            s = s.Substring(0, 4) + sTer + "_" + s.Substring(4);
                            string sFil = sPath + s;
                            if (File.Exists(sFil))
                                File.Delete(sFil);
                            File.Copy(sFilEan, sFil);
                        }

                        foreach (string sTer in aTer)
                        {
                            s = Path.GetFileName(sFilLia);
                            s = s.Substring(0, 4) + sTer + "_" + s.Substring(4);
                            string sFil = sPath + s;
                            if (File.Exists(sFil))
                                File.Delete(sFil);
                            File.Copy(sFilLia, sFil);
                        }

                        string sFtp = "";
                        sFtp += (string)t.Rows[0]["tab_hst"] + "|";
                        sFtp += (string)t.Rows[0]["tab_usr"] + "|";
                        sFtp += (string)t.Rows[0]["tab_pwd"] + "|";

                        s = (string)t.Rows[0]["tab_fpt"];

                        string[] a = s.Split(';');

                        if (a.Length > 1)
                        {
                            sFtp += a[0] + "|"; ;
                            sFtp += a[1] + "|"; ;
                            InvioFtp(sFtp, sPath);
                        }
                        else
                            MessageBox.Show("Percorso destinazione non definito in TabTerm");
                    }
                }
            }
        }

        private string InvioFtp(string strFtp, string strPathLoc)
        {
            string[] a = strFtp.Split('|');

            //string strHst = ((string)tabDiv.Rows[0]["tab_hst"]).Trim();
            //string strUsr = ((string)tabDiv.Rows[0]["tab_uso"]).Trim();
            //string strPwd = ((string)tabDiv.Rows[0]["tab_pwo"]).Trim();
            //string strPath = ((string)tabDiv.Rows[0]["tab_fpt"]).Trim();

            string strHst = a[0];
            string strUsr = a[1];
            string strPwd = a[2];
            string strPath = a[3];
            string strTerm = a[4];

            string s = "";
            string sMsg = "";
            clsFtp clsFtp = new clsFtp();
            clsFtp._strFtpHost = strHst;
            clsFtp._strFtpUser = strUsr;
            clsFtp._strFtpPswd = strPwd;
            clsFtp._strFtpPath = strPath;
            clsFtp._strLocPath = strPathLoc;

            sMsg = clsFtp.FtpTest();

            if (sMsg != "")
                MessageBox.Show(sMsg, "COLLEGAMENTO FTP " + strHst, MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {

                string[] sDests = strTerm.Split(',');

                string[] sFils = Directory.GetFiles(strPathLoc);
                foreach (string sFil in sFils)
                {
                    foreach (string sDest in sDests)
                    {
                        if (Path.GetFileName(sFil).Substring(4, 3) == sDest)
                        {
                            clsFtp._strFtpPath = strPath + sDest + "/";

                            clsFtp.FtpUpLoad(sFil);

                            if (File.Exists(strPathLoc + "Temp\\" + Path.GetFileName(sFil)))
                                File.Delete(strPathLoc + "Temp\\" + Path.GetFileName(sFil));

                            clsFtp._strLocPath = strPathLoc + "Temp\\";
                            sMsg += clsFtp.FtpDownLoad(sFil);
                            FileInfo fi1 = new FileInfo(strPathLoc + "Temp\\" + Path.GetFileName(sFil));
                            FileInfo fi2 = new FileInfo(strPathLoc + Path.GetFileName(sFil));
                            if (sMsg == "")
                            {
                                if (fi1.Length != fi2.Length)
                                    sMsg += sFil + " non inviato correttamente!";
                            }
                        }
                    }

                    s = Path.GetDirectoryName(strPathLoc) + "\\Save\\" + Path.GetFileName(sFil) + "_Inviato_" + DateTime.Now.ToString("yyyyMMddmmss");
                    File.Move(sFil, s);
                }
            }

            return sMsg;
        }

    }
}
