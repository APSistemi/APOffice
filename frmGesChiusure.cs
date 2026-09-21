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
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace APOffice
{
    public partial class frmGesChiusure : Form
    {
        private const string TABANAEAN = "AnaBarcode";
        private const string TABANAART = "AnaArticoli";
        private const string TABTABREP = "TabReparti";
        private const string TABSTAVET = "GesNegVet";
        private const string TABSTAVEP = "GesNegVep";
        private const string TABSTAVEN = "GesNegVen";

        private const string CAUSCO = "SCO";
        
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strConSqlSta = "";
        
        private string _strConPun = "";

        public frmGesChiusure()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
        }

        private void frmGesChiusure_Load(object sender, EventArgs e)
        {
            dgv1.DataSource = new clsGenTabTmp().TabTmpVenduto("TabArt");

            string s = _clsFun.ParGet(clsDefine.enuParametri.Par013CodRepxDefault, _strConSql);
            if (s == "")
                MessageBox.Show("Parametro reparto di defaul non definito!", "CHIAMA L'ASSISTENZA");

            // Definizione codice tabella pagamenti per i punti


            s = "SELECT tab_cod, tab_key FROM TabCauCassa WHERE tab_key='PUN' OR tab_key='PUM'";
            DataTable tCauPag = _clsFun.FillTabSql("tPag", s, false, _strConSql);

            //string sOra = DateTime.Now.ToString("HHmm");
            //string sCau = "PUN";

            //string sTip = "PUN";
            //string sCod = "";

            DataRow[] j = tCauPag.Select("tab_key='PUN'");
            if (j.Length > 0)
            {
                _strConPun = (string)j[0]["tab_cod"] + ";";

                j = tCauPag.Select("tab_key='PUM'");
                if (j.Length > 0)
                    _strConPun += (string)j[0]["tab_cod"];
            }

        }

        private void frmGesChiusure_KeyDown(object sender, KeyEventArgs e)
        {
            Esci();
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void btnChiudi_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource != null)
                ((DataTable)dgv1.DataSource).Rows.Clear();
            if (dgv2.DataSource != null)
                ((DataTable)dgv2.DataSource).Rows.Clear();

            if (MessageBox.Show("Confermi la chiusura di fine giornata?", "CHIUSURA DI FINE GIORNO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Chiusura();
                //if (dgv2.DataSource != null && ((DataTable)dgv2.DataSource).Rows.Count > 0)
                PrnRep();

                //new clsGenPdfRep().PrnPdfRep((DataTable)dgv2.DataSource);
            }
        }

        private void Chiusura()
        {
            DataTable tCnf = _clsQry.ConfSeek("", "VEN");
            if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16((object)clsDefine.enuPos.posDitron).ToString("00"))
            {
                CloDitron(tCnf);
            }
            else if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16((object)clsDefine.enuPos.posBrainpos).ToString("00"))
            {
                CloBrainPos(tCnf);
            }
            else if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16((object)clsDefine.enuPos.posNcr745x).ToString("00"))
            {
                CloNcr745x(tCnf);
            }
            //else if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16((object)clsDefine.enuPos.posApShop01).ToString("00"))
            //{
            //    CloApShop(tCnf);
            //}
        }

        private void CloDitron(DataTable tabCnf)
        {
            Boolean b = true;
            if (chkVenClo.Checked)
            {
                string s = _clsFun.ParGet(clsDefine.enuParametri.ParRitardoChiusura, _strConSql);
                Int64 iGir = Convert.ToInt64(50000);
                if (_clsFun.Numerico(s))
                    iGir = Convert.ToInt64(s);

                b = false;
                btnChiudi.Enabled = false;
                if (File.Exists("c:\\netpos\\ctrlpos\\end_day.sem"))
                    File.Delete("c:\\netpos\\ctrlpos\\end_day.sem");
                if (!File.Exists("C:\\NETPOS\\CTRLPOS\\STARTEOD.SEM"))
                    File.Copy("C:\\NETPOS\\SH\\STARTEOD.SEM", "C:\\NETPOS\\CTRLPOS\\STARTEOD.SEM");
                BackUp();

                //_clsFun.FileLog("Chiusura", "Inizio", "");

                for (int i = 0; i <= 5; i++)
                {
                    //_clsFun.FileLog("Chiusura", "Giro " + i.ToString(), "1");

                    for (Int64 i2 = 0; i2 <= iGir; i2++)
                    {
                        Console.WriteLine("aaaaaaaaaa");
                    }

                    //_clsFun.FileLog("Chiusura", "Giro " + i.ToString(), "2");

                    s = Path.GetDirectoryName((string)tabCnf.Rows[0]["CnfVen"]);

                    foreach (string sFil in Directory.GetFiles(s, "dca*.*"))
                    {
                        //_clsFun.FileLog("Chiusura", "Giro " + i.ToString(), "3");

                        string strDay = File.GetLastWriteTime(sFil).ToString("yyyyMMdd");
                        if (File.GetLastWriteTime(sFil).ToString("yyyyMMdd").Equals(DateTime.Today.ToString("yyyyMMdd")))
                        {
                            //_clsFun.FileLog("Chiusura", "Giro " + i.ToString(), "4 trovato");

                            b = true;
                            break;
                        }
                    }

                    if (b)
                        break;
                }
            }

            if (b)
            {
                //_clsFun.FileLog("Chiusura", "Lettura venduto", "5");

                ChiuDitron(tabCnf, Path.GetDirectoryName((string)tabCnf.Rows[0]["CnfVen"]), "");
                AggDgv2();
            }
            else
                MessageBox.Show("File del venduto non rilevato!");
        }

        private void CloNcr745x(DataTable tabCnf)
        {
            Boolean b = true;
            string s = "";
            string s2 = "";
            string sPathVen = "C:\\ApProject\\Temp\\Venduto\\ApShop\\";

            if (chkVenClo.Checked)
            {

                //Da fare un loop in caso di più casse su S_IDC***.DAT

                s = _clsFun.ParGet(clsDefine.enuParametri.ParRitardoChiusura, _strConSql);
                Int64 iTime = Convert.ToInt16(5000);
                if (_clsFun.Numerico(s))
                    iTime = Convert.ToInt64(s);

                for (int i = 1; i <= 5; i++)
                {
                    s = "C:\\Server\\Data\\S_IDC" + i.ToString("000") + ".DAT";
                    if (File.Exists(s))
                    {
                        s2 = _clsDef.TMPPOSVEN + "Old\\" + Path.GetFileName(s) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
                        File.Copy(s, s2);

                        s2 = sPathVen + Path.GetFileName(s);

                        if (File.Exists(s2))
                            File.Delete(s2);

                        File.Copy(s, s2);
                    }
                }

                b = false;
                btnChiudi.Enabled = false;
                //if (File.Exists("c:\\netpos\\ctrlpos\\end_day.sem"))
                //    File.Delete("c:\\netpos\\ctrlpos\\end_day.sem");
                //if (!File.Exists("C:\\NETPOS\\CTRLPOS\\STARTEOD.SEM"))
                //    File.Copy("C:\\NETPOS\\SH\\STARTEOD.SEM", "C:\\NETPOS\\CTRLPOS\\STARTEOD.SEM");

                Process p = new Process();
                p.StartInfo.FileName = "C:\\GM\\eod7453.bat";
                p.StartInfo.Arguments = "-r";
                p.StartInfo.ErrorDialog = true;
                p.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                p.Start();
                p.WaitForExit(1000 * 60 * 7);    // wait up to 1 minutes.
                Thread.Sleep(15000);

                BackUp();

                for (int i = 1; i <= 5; i++)
                {
                    for (Int64 i2 = 0; i2 <= iTime; i2++)
                        Console.WriteLine("aaaaaaaaaa");

                    s = "S_IDC" + i.ToString("000") + ".DAT";

                    s2 = sPathVen + s;

                    ////Thread.Sleep(iTime);
                    ////s = Path.GetDirectoryName((string)tabCnf.Rows[0]["CnfVen"]);
                    ////foreach (string sFil in Directory.GetFiles(s, "HOCIDC.*"))
                    //foreach (string sFil in Directory.GetFiles(s, "HOCIDC.*"))
                    //{
                    //    //string strDay = File.GetLastWriteTime(sFil).ToString("yyyyMMdd");
                    //    //if (File.GetLastWriteTime(sFil).ToString("yyyyMMdd").Equals(DateTime.Today.ToString("yyyyMMdd")))
                    //    //{
                    //    //    //_clsFun.FileLog("Chiusura", "Giro " + i.ToString(), "4 trovato");
                    //    b = true;
                    //    break;
                    //    //}
                    //}

                    if (File.Exists(s2))
                    {
                        b = true;
                        break;
                    }
                }
            }

            if (b)
            {
                //ChiuNcr745x(tabCnf, Path.GetDirectoryName((string)tabCnf.Rows[0]["CnfVen"]), "");
                ChiuNcr745x(tabCnf, sPathVen, "");
                AggDgv2();
            }
            else
                MessageBox.Show("File del venduto non rilevato!");
        }

        private void CloBrainPos(DataTable tabCnf)
        {
            string sPathBat = Path.GetDirectoryName(_clsDef.TMPPOSVEN);

            Boolean b = true;
            if (chkVenClo.Checked)
            {
                string s = _clsFun.ParGet(clsDefine.enuParametri.ParRitardoChiusura, _strConSql);
                Int64 iGir = Convert.ToInt64(50000);
                if (_clsFun.Numerico(s))
                    iGir = Convert.ToInt64(s);

                b = false;
                btnChiudi.Enabled = false;
                //if (File.Exists("c:\\netpos\\ctrlpos\\end_day.sem"))
                //    File.Delete("c:\\netpos\\ctrlpos\\end_day.sem");
                //if (!File.Exists("C:\\NETPOS\\CTRLPOS\\STARTEOD.SEM"))
                //    File.Copy("C:\\NETPOS\\SH\\STARTEOD.SEM", "C:\\NETPOS\\CTRLPOS\\STARTEOD.SEM");
                //BackUp();


                //Pulizia
                foreach (string sFil in Directory.GetFiles(sPathBat, "dcascii.00*"))
                {
                    File.Delete(sFil);
                }
                sPathBat += "\\Brainpos\\";
                foreach (string sFil in Directory.GetFiles(sPathBat, "DATADC.00*"))
                {
                    File.Delete(sFil);
                }

                for (int i = 0; i <= 5; i++)
                {
                    for (Int64 i2 = 0; i2 <= iGir; i2++)
                    {
                        Console.WriteLine("aaaaaaaaaa");
                    }

                    s = Path.GetDirectoryName((string)tabCnf.Rows[0]["CnfVen"]);

                    foreach (string sFil in Directory.GetFiles(s, "DATADC.00*"))
                    {
                        string strDay = File.GetLastWriteTime(sFil).ToString("yyyyMMdd");
                        if (true || File.GetLastWriteTime(sFil).ToString("yyyyMMdd").Equals(DateTime.Today.ToString("yyyyMMdd")))
                        {
                            //sPathBat = Path.GetDirectoryName(_clsDef.TMPPOSVEN) + "\\Brainpos\\";

                            s = sPathBat + Path.GetFileName(sFil);
                            File.Copy(sFil, s);

                            string sBat = sPathBat + "bDcascii.bat";

                            //string sFilDcascii = _clsDef.TMPPOSVEN + "dcascii" + Path.GetExtension(sFil);
                            string sFilDcascii = "dcascii" + Path.GetExtension(sFil);

                            StreamWriter sw = new StreamWriter(sBat, false);
                            s = "_DCASCII.EXE /I" + Path.GetFileName(sFil) + " /O" + sFilDcascii;
                            sw.Write(s + _clsDef.CRLF);
                            sw.Flush();
                            sw.Close();
                            sw.Dispose();

                            Directory.SetCurrentDirectory(sPathBat);

                            Process p = new Process();
                            p.StartInfo.FileName = sBat;
                            p.StartInfo.Arguments = "-r";
                            p.StartInfo.ErrorDialog = true;
                            p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                            p.Start();
                            p.WaitForExit(1000 * 60 * 2);    // wait up to 2 minutes.

                            //s = sPathBat + sFilDcascii;

                            b = true;
                        }
                    }

                    if (b)
                        break;
                }
            }

            if (b)
            {
                ChiuDitron(tabCnf, sPathBat, "");
                AggDgv2();
            }
            else
                MessageBox.Show("File del venduto non rilevato!");
        }

        private void BackUp()
        {
            if (File.Exists(_clsDef.BATCHBACKUP))
            {
                Process p = new Process();
                p.StartInfo.FileName = _clsDef.BATCHBACKUP;
                p.StartInfo.Arguments = "-r";
                p.StartInfo.ErrorDialog = true;
                p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                p.Start();
                p.WaitForExit(1000 * 60 * 2);    // wait up to 2 minutes.
            }
        }

        public void ChiuDitron(DataTable tabCnf, string strPath, string strParFat)
        {
            string s = "";

            DateTime dFatDay = _clsDef.DAYOUT;
            string sFatPos = "";
            string sFatSco = "";

            if (strParFat != "")
            {
                string[] aa = strParFat.Split('|');
                dFatDay = _clsFun.Str2Day(aa[0]);
                sFatPos = aa[1];
                sFatSco = aa[2];
            }

            s = "SELECT ean_art, ean_ean, art_sta ";
            s += "FROM AnaBarcode ";
            s += "LEFT JOIN AnaArticoli ON AnaBarcode.ean_art = AnaArticoli.art_cod";
            DataTable tEan = _clsFun.FillTabSql(TABANAEAN, s, false, _strConSql);
            DataColumn[] Key = new DataColumn[1] 
            { 
                tEan.Columns["ean_ean"] 
            };
            tEan.PrimaryKey = Key;

            _clsFun.FileLog("Estrazione", "Ean", s);
            _clsFun.FileLog("Estrazione", "Ean", tEan.Columns.Count.ToString());


            if(sFatSco != "")
                s = "SELECT * FROM GesNegVet WHERE vet_day = " + _clsFun.DaySql(dFatDay);
            else
                s = "SELECT * FROM GesNegVet WHERE vet_day > " + _clsFun.DaySql(DateTime.Today.AddDays(-50));
            DataTable tVet = _clsFun.FillTabSql("GesNegVet", s, false, _strConSqlSta);
            Key = new DataColumn[6]
            {
                tVet.Columns["vet_neg"],
                tVet.Columns["vet_cau"],
                tVet.Columns["vet_day"],
                tVet.Columns["vet_ora"],
                tVet.Columns["vet_pos"],
                tVet.Columns["vet_sco"]
            };
            tVet.PrimaryKey = Key;

            _clsFun.FileLog("Estrazione", "Vet", s);
            _clsFun.FileLog("Estrazione", "Vet", tVet.Columns.Count.ToString());


            if (sFatSco != "")
                s = "SELECT * FROM GesNegVep WHERE vep_day = " + _clsFun.DaySql(dFatDay);
            else
                s = "SELECT * FROM GesNegVep WHERE vep_day > " + _clsFun.DaySql(DateTime.Today.AddDays(-50));
            DataTable tVep = _clsFun.FillTabSql("GesNegVep", s, false, _strConSqlSta);
            Key = new DataColumn[8]
            {
                tVep.Columns["vep_neg"],
                tVep.Columns["vep_cau"],
                tVep.Columns["vep_day"],
                tVep.Columns["vep_ora"],
                tVep.Columns["vep_pos"],
                tVep.Columns["vep_sco"],
                tVep.Columns["vep_cod"],
                tVep.Columns["vep_ean"]
            };
            tVep.PrimaryKey = Key;

            _clsFun.FileLog("Estrazione", "Vep", s);
            _clsFun.FileLog("Estrazione", "Vep", tVep.Columns.Count.ToString());


            if (sFatSco != "")
                s = "SELECT * FROM GesNegVen WHERE ven_neg='001' AND ven_cau='SCO' AND AND ven_day > '01/01/2019'";
            else
                s = "SELECT * FROM GesNegVen WHERE ven_neg='001' AND ven_cau='SCO' AND ven_day > '01/01/2019'"; // +_clsFun.DaySql(DateTime.Today.AddDays(-50));  //20181006 per evitare out off memory da Stocco
            DataTable tVen = this._clsFun.FillTabSql("GesNegVen", s, false, _strConSqlSta);
            Key = new DataColumn[9]
            {
                tVen.Columns["ven_neg"],
                tVen.Columns["ven_cau"],
                tVen.Columns["ven_day"],
                tVen.Columns["ven_ora"],
                tVen.Columns["ven_pos"],
                tVen.Columns["ven_sco"],
                tVen.Columns["ven_sct"],
                tVen.Columns["ven_ean"],
                tVen.Columns["ven_art"]
            };
            tVen.PrimaryKey = Key;

            //_clsFun.FileLog("Estrazione", "Ven", s);
            //_clsFun.FileLog("Estrazione", "Ven", tVen.Columns.Count.ToString());

            DataTable tErr = new clsGenTabTmp().TabTmpVenErrori("TabErr");

            string sNeg = (string)tabCnf.Rows[0]["cnf_cod"];

            s = Path.GetDirectoryName((string)tabCnf.Rows[0]["CnfVen"]);

            if (strParFat != "")
                LeggiDitron(sNeg, "C:\\NetPos\\SH\\Fatt.tmp", tEan, tVet, tVep, tVen, tErr, strParFat);
            else
            {
                foreach (string sFil in Directory.GetFiles(strPath))
                {
                    if (Path.GetFileName(sFil).Substring(0, 3).ToLower() == "dca")
                    {
                        ArrayList ary = LeggiDitron(sNeg, sFil, tEan, tVet, tVep, tVen, tErr, "");

                        foreach (string a in ary)
                        {
                            s = a;
                            DateTime day = new DateTime(Convert.ToInt32(s.Substring(3, 4)), Convert.ToInt32(s.Substring(7, 2)), Convert.ToInt32(s.Substring(9, 2)));
                            s = "SELECT SUM(vet_imp) AS Imp FROM GesNegVet WHERE vet_neg='" + sNeg + "' AND vet_day=" + _clsFun.DaySql(day);
                            DataTable t = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlSta);
                            if (t.Rows.Count > 0)
                            {
                                AggDgv1("VEN", sNeg, "Venduto", day, (Decimal)t.Rows[0]["Imp"]);
                            }
                        }
                        s = _clsDef.TMPPOSVEN + "Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
                        File.Copy(sFil, s);
                    }
                }
                if(tErr.Rows.Count > 0)
                    new clsGenPdfVenErr().PrnPdfVenErr(tErr);
            }
        }

        public bool DitronLeggiScontrino(DataTable tabCnf, string strParFat)
        {
            string sPath = "C:\\NetPos\\SH\\";
            string sFilFat = "Fatt.tmp";
            string sFilBat = "bRecSco.bat";
            string s = "";

            s = sPath + sFilFat;
            if (File.Exists(s))
                File.Delete(s);
            s = sPath + sFilBat;
            if (File.Exists(s))
                File.Delete(s);

            Boolean b = true;

            if (strParFat != "")
            {
                string[] aa = strParFat.Split('|');

                DateTime dFatDay = _clsFun.Str2Day(aa[0]);
                string sFatPos = aa[1];
                string sFatSco = aa[2];

                sFatSco = sFatSco.Substring(1);

                //FATTURA /T002 /S0002 /D230309 /BFATT.BIN
                s = "FATTURA /T0" + sFatPos + " /S" + sFatSco + " /D" + dFatDay.ToString("ddMMyy") + " /BFATT.BIN";
                StreamWriter sw = new StreamWriter(sPath + sFilBat, false);
                sw.Write("CD \\NetPos\\SH " + _clsDef.CRLF);
                sw.Write(s + _clsDef.CRLF);
                //sw.Write("pause" + _clsDef.CRLF);
                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                Process p = new Process();
                p.StartInfo.FileName = "C:\\NetPos\\SH\\bRecSco.bat";
                p.StartInfo.Arguments = "-r";
                p.StartInfo.ErrorDialog = true;
                p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                p.Start();
                p.WaitForExit(1000 * 60);    // wait up to 1 minutes.
                Thread.Sleep(8000);

                if (File.Exists(sPath + sFilFat))
                    ChiuDitron(tabCnf, "C:\\NetPos\\SH", strParFat);
                else
                    MessageBox.Show("Scontrino non trovato!", "ESTRAZIONE SCONTRINO", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }

            return b;
        }

        private ArrayList LeggiDitron(string strNeg, string strFil, DataTable tabEan, DataTable tabVet, DataTable tabVep, DataTable tabVen, DataTable tabErr, string strParFat)
        {
            string s = "";
            Boolean bDec = true;
            ArrayList a = new ArrayList();
            DataTable tTmp = new clsGenTabTmp().TabTmpVet("TabVet");
            DataTable tSco = tabVen.Clone();
            Decimal num1 = new Decimal(0);
            string sPos = "";
            string sSco = "";
            string sOra = "";

            if (!chkPer100.Checked)
                bDec = false;

            string sFil = _clsDef.TMPPOSVEN + Path.GetFileName(strFil);

            if (File.Exists(sFil))
                File.Delete(sFil);
            File.Copy(strFil, sFil);

            using (StreamReader sr = new StreamReader(sFil))
            {
                FileInfo fI = new FileInfo(strFil);
                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fI.Length;
                progressBar1.Minimum = 0;
                //lblNeg.Text = Path.GetFileName(strFil);
                string sRig;
                Boolean b = true;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    Application.DoEvents();
                    if (sRig.Substring(5, 1) == "V")
                    {
                        string sTri = sRig.Substring(1, 3);
                        if (sTri == "100")
                        {
                            if (sRig.Substring(7, 3) == "001")
                            {
                                sPos = sRig.Substring(60, 1).PadLeft(2,Convert.ToChar('0'));
                                sSco = sRig.Substring(115, 5);

                                DateTime dd = new DateTime(2000 + (int)Convert.ToInt16(sRig.Substring(77, 2)), (int)Convert.ToInt16(sRig.Substring(74, 2)), (int)Convert.ToInt16(sRig.Substring(71, 2)));

                                //s = sRig.Substring(12,5);

                                b = true;

                                if(strParFat != "")
                                {
                                    string[] aa = strParFat.Split('|');

                                    DateTime dFatDay = _clsFun.Str2Day(aa[0]);
                                    string sFatPos = aa[1];
                                    string sFatSco = aa[2];

                                    TimeSpan diff = dFatDay.Subtract(dd);

                                    if (diff.Days != 0 || sPos != sFatPos || sSco != sFatSco)
                                    {
                                        b = false;
                                    }
                                }

                                if (b)
                                {
                                    if (tTmp.Rows.Count > 0)
                                        tTmp.Rows.Clear();
                                    DataRow x = tTmp.NewRow();
                                    x["VenNeg"] = strNeg;
                                    x["VenDay"] = DateTime.Today.ToShortDateString();
                                    x["VenOra"] = "";
                                    x["VenPos"] = sPos;
                                    x["VenSco"] = sSco;
                                    x["VenFid"] = "";
                                    tTmp.Rows.Add(x);
                                }

                            }
                        }

                        if (b)
                        {
                            if (sTri == "000" || sTri == "004")
                            {
                                if (sRig.Substring(10, 13).Trim() == "2110003")
                                    Console.WriteLine("0000");
                                if (sRig.Substring(59, 4) != "4220")
                                {
                                    string sEan = sRig.Substring(10, 13).Trim().PadLeft(13, Convert.ToChar("0"));
                                    string sArd = sRig.Substring(24, 20);
                                    Decimal dVen = Convert.ToDecimal(sRig.Substring(45, 13)) / new Decimal(100);
                                    Decimal dQta = Convert.ToDecimal(sRig.Substring(59, 9));
                                    string sIva = sRig.Substring(69, 2).PadLeft(3, Convert.ToChar("0"));
                                    Decimal dQkg = Convert.ToDecimal(sRig.Substring(75, 11).Replace(".", ","));
                                    string sRep = sRig.Substring(90, 3);
                                    Decimal dPun = new Decimal(0);
                                    if (sTri == "000")
                                    {
                                        dPun = Convert.ToDecimal(sRig.Substring(112, 5));
                                        if (dPun > 0 && sRig.Substring(110, 1) == "C")
                                            dPun = dPun * -1;
                                    }
                                    Decimal dPrz = new Decimal(0);
                                    if (sTri == "000")
                                        dPrz = Convert.ToDecimal(sRig.Substring(123, 13)) / new Decimal(100);
                                    //if (chkDecimali.Checked)
                                    if (bDec)
                                    {
                                        dVen /= new Decimal(100);
                                        dPrz /= new Decimal(100);
                                    }

                                    if (!_clsFun.Numerico(sEan))
                                        sEan = "";
                                    else
                                    {

                                    if (sEan.Substring(0, 12) == "294832000000")
                                        Console.WriteLine("aaaaa");
                                    if (Convert.ToInt64(sEan) == Convert.ToInt64(2948320))
                                        Console.WriteLine("aaaaa");

                                        s = sEan;
                                        if (s.Length == 13 && s.Substring(0, 7) == "0000002")
                                            s = s.Substring(6, 7) + "000000";
                                        else if (s.Length == 13 && s.Substring(0, 8) == "00000002")
                                            s = s.Substring(7, 6) + "0000000";
                                        else
                                            s = Convert.ToInt64(s).ToString();
                                        sEan = s;
                                    }

                                    //s = "ven_ean='" + sEan + "'";

                                    DataRow[] j1 = tSco.Select("ven_ean='" + sEan + "'");
                                    if (j1.Length == 0)
                                    {
                                        DataRow x = tSco.NewRow();
                                        x["ven_neg"] = strNeg;
                                        x["ven_cau"] = "SCO";
                                        x["ven_day"] = DateTime.Today.ToShortDateString();
                                        x["ven_pos"] = sPos;
                                        x["ven_ora"] = sOra;
                                        x["ven_art"] = "";
                                        x["ven_ard"] = sArd;

                                        if (sEan == "0000080019428")
                                            Console.WriteLine("aaaa");

                                        s = "ean_ean='" + sEan + "'";
                                        if(sEan.Length == 13 && sEan.Substring(0,1) == "2" && sEan.Substring(7,5)=="00000")
                                            s = "SUBSTRING(ean_ean,1,12)='" + sEan.Substring(0,12) + "'";

                                        //DataRow[] j2 = tabEan.Select("ean_ean='" + sEan + "'");
                                        DataRow[] j2 = tabEan.Select(s);
                                        if (j2.Length > 0)
                                        {
                                            x["ven_art"] = ((string)j2[0]["ean_art"]).Trim();
                                            if (((string)j2[0]["art_sta"]).Trim() == "N")
                                            {
                                                ErrMsg(tabErr, ((string)j2[0]["ean_art"]).Trim(), "Articolo attivato - " + sArd + " Q.tà " + dQta + " Prz " + dVen.ToString());
                                                s = "UPDATE AnaArticoli SET art_sta='A' WHERE art_cod='" + x["ven_art"] + "'";
                                                _clsFun.SqlWrite(s, _strConSql);
                                            }
                                        }
                                        else
                                        {
                                            x["ven_art"] = "REP" + sRep.Substring(1, 2);
                                            ErrMsg(tabErr, sEan, "Non trovato - " + sArd + " Q.tà " + dQta + " Prz " + dVen.ToString());
                                        }

                                        x["ven_iva"] = sIva;
                                        x["ven_sco"] = sSco;
                                        x["ven_ean"] = sEan;
                                        x["ven_rep"] = sRep;
                                        x["ven_prz"] = dPrz;
                                        x["ven_qta"] = 0;
                                        x["ven_qkg"] = 0;
                                        x["ven_ven"] = 0;
                                        x["ven_pun"] = 0;
                                        x["ven_sct"] = "";
                                        x["ven_scn"] = 0;
                                        tSco.Rows.Add(x);
                                        j1 = tSco.Select("ven_ean='" + sEan + "'");
                                    }
                                    if (sTri == "004")
                                        dQta *= -1;
                                    j1[0]["ven_qta"] = ((Decimal)j1[0]["ven_qta"] + dQta);
                                    j1[0]["ven_qkg"] = ((Decimal)j1[0]["ven_qkg"] + dQkg);
                                    j1[0]["ven_ven"] = ((Decimal)j1[0]["ven_ven"] + dVen * dQta);
                                    j1[0]["ven_pun"] = ((Decimal)j1[0]["ven_pun"] + dPun);
                                }
                            }
                            else if (sTri == "002")
                            {
                                tTmp.Rows[0]["VenScp"] = (object)Convert.ToDecimal(sRig.Substring(51, 11).Replace(".", ","));
                                tTmp.Rows[0]["VenSct"] = (object)(Convert.ToDecimal(sRig.Substring(69, 13).Replace(".", ",")) / new Decimal(100));
                            }
                            else if (sTri == "005")
                            {
                                if (sRig.Substring(54, 4) != "1000")        //Punti no Sconto
                                {
                                    //tTmp.Rows[0]["VenScp"] = (object)Convert.ToDecimal(sRig.Substring(51, 11).Replace(".", ","));
                                    //tTmp.Rows[0]["VenSct"] = (object)(Convert.ToDecimal(sRig.Substring(34, 13).Replace(".", ",")) / new Decimal(100));
                                    tTmp.Rows[0]["VenSct"] = (object)(Convert.ToDecimal(sRig.Substring(34, 13).Replace(".", ",")));

                                    Console.WriteLine("aaaa");

                                    if (bDec)
                                        tTmp.Rows[0]["VenSct"] = (decimal)tTmp.Rows[0]["VenSct"] / 100;
                                }
                            }
                            else if (!(sTri == "004") && !(sTri == "005"))
                            {
                                if (sTri == "014")
                                {
                                    s = sRig.Substring(10, 13);
                                    tTmp.Rows[0]["VenFid"] = (object)s;
                                }
                                else if (sTri == "044")
                                {
                                    Decimal d = Convert.ToDecimal(sRig.Substring(101, 5));
                                    tTmp.Rows[0]["VenPun"] = (object)d;
                                }
                                else if (sTri == "200")
                                {
                                    DateTime dDay = new DateTime(2000 + (int)Convert.ToInt16(sRig.Substring(13, 2)), (int)Convert.ToInt16(sRig.Substring(10, 2)), (int)Convert.ToInt16(sRig.Substring(7, 2)));
                                    Decimal dImp = Convert.ToDecimal(sRig.Substring(25, 13).Replace(".", ","));
                                    decimal dPun = 0;

                                    if (tTmp.Rows.Count > 0 && (string)tTmp.Rows[0]["VenFid"] != "")
                                    {
                                        foreach (DataRow y in tSco.Rows)
                                            dPun += (decimal)y["ven_pun"];

                                        tTmp.Rows[0]["VenPun"] = (decimal)tTmp.Rows[0]["VenPun"] + dPun;
                                    }

                                    if (dImp > new Decimal(0) || dPun != 0)
                                    {
                                        //if (chkDecimali.Checked)
                                        if (bDec)
                                            dImp /= new Decimal(100);
                                        if (tSco.Rows.Count > 0)
                                        {
                                            sOra = sRig.Substring(16, 2) + sRig.Substring(19, 2);
                                            if (a.IndexOf(strNeg + dDay.ToString("yyyyMMdd")) < 0)
                                                a.Add(strNeg + dDay.ToString("yyyyMMdd"));
                                            tTmp.Rows[0]["VenDay"] = dDay;
                                            tTmp.Rows[0]["VenImp"] = dImp;
                                            tTmp.Rows[0]["VenOra"] = sOra;
                                            VenScontrino(tabVet, tabVep, tabVen, tSco, tTmp);
                                        }
                                    }

                                    tTmp.Rows.Clear();
                                    tSco.Rows.Clear();
                                }
                            }
                        }
                    }
                }
            }
            s = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(strFil);
            if (File.Exists(s))
                File.Delete(s);
            //if (_bolSposta)
            //    File.Move(strFil, s);
            //else
            //    File.Copy(strFil, s);
            // MessageBox.Show(s);
            File.Move(sFil, s);

            return a;
        }

        private void VenScontrino(DataTable tabVet, DataTable tabVep, DataTable tabVen, DataTable tabSco, DataTable tabTmp)
        {            
            string sNeg = (string)tabTmp.Rows[0]["VenNeg"];
            DateTime dDay = (DateTime)tabTmp.Rows[0]["VenDay"];
            string sPos = (string)tabTmp.Rows[0]["VenPos"];
            string sSco = (string)tabTmp.Rows[0]["VenSco"];
            string sOra = (string)tabTmp.Rows[0]["VenOra"];
            Decimal d2 = 0;
            string s = "ven_neg='" + sNeg + "' AND ven_cau='" + CAUSCO + "' AND ven_day=#" + dDay.ToString("MM/dd/yyyy") + "# AND ven_ora='" + sOra + "' AND ven_pos='" + sPos + "' AND ven_sco='" + sSco + "'";
            DataRow[] j = tabVen.Select(s);

            _clsFun.FileLog("Chiusura", s, "Seek");

            if(j.Length == 0)
            {
                s = "vet_neg='" + sNeg + "' AND vet_cau='" + CAUSCO + "' AND vet_day=#" + dDay.ToString("MM/dd/yyyy") + "# AND vet_ora='" + sOra + "' AND vet_pos='" + sPos + "' AND vet_sco='" + sSco + "'";
                j = tabVet.Select(s);
            }
            if (j.Length > 0)
            {
                string sSql = "DELETE FROM GesNegVen WHERE ven_neg='" + sNeg + "' AND ven_cau='" + CAUSCO + "' AND ven_day=" + _clsFun.DaySql(dDay) + " AND ven_ora='" + sOra + "' AND " + "ven_pos='" + sPos + "' AND ven_sco='" + sSco + "'";
                _clsFun.SqlWrite(sSql, _strConSqlSta);

                sSql = "DELETE FROM GesNegVep WHERE vep_neg='" + sNeg + "' AND vep_cau='" + CAUSCO + "' AND vep_day=" + _clsFun.DaySql(dDay) + " AND vep_ora='" + sOra + "' AND " + "vep_pos='" + sPos + "' AND vep_sco='" + sSco + "'";
                _clsFun.SqlWrite(sSql, _strConSqlSta);

                _clsFun.FileLog("Chiusura", s, "Delete");


                sSql = "DELETE FROM GesNegVet WHERE vet_neg='" + sNeg + "' AND vet_cau='" + CAUSCO + "' AND vet_day=" + _clsFun.DaySql(dDay) + " AND vet_ora='" + sOra + "' AND " + "vet_pos='" + sPos + "' AND vet_sco='" + sSco + "'";
                _clsFun.SqlWrite(sSql, _strConSqlSta);
                for (int i = j.Length - 1; i >= 0; --i)
                    j[i].Delete();
                s = "vet_neg='" + sNeg + "' AND vet_cau='" + CAUSCO + "' AND vet_day=#" + dDay.ToString("MM/dd/yyyy") + "# AND vet_ora='" + sOra + "' AND vet_pos='" + sPos + "' AND vet_sco='" + sSco + "'";
                j = tabVet.Select(s);
                if (j.Length > 0)
                    j[0].Delete();
            }

            DataRow x = tabVet.NewRow();
            x["vet_neg"] = sNeg;
            x["vet_cau"] = CAUSCO;
            x["vet_day"] = dDay;
            x["vet_ora"] = sOra;
            x["vet_pos"] = sPos;
            x["vet_sco"] = sSco;
            x["vet_fid"] = tabTmp.Rows[0]["VenFid"];
            x["vet_imp"] = tabTmp.Rows[0]["VenImp"];
            x["vet_pun"] = tabTmp.Rows[0]["VenPun"];
            x["vet_sct"] = tabTmp.Rows[0]["VenSct"];
            x["vet_scp"] = tabTmp.Rows[0]["VenScp"];
            x["vet_art"] = tabSco.Rows.Count;
            tabVet.Rows.Add(x);
            _clsFun.SqlWrite(_clsFun.SqlInsertRow(TABSTAVET, tabVet, x), _strConSqlSta);
            if ((Decimal)tabTmp.Rows[0]["VenSct"] < 0)
            {
                Decimal d = Math.Abs((Decimal)tabTmp.Rows[0]["VenSct"]) / ((Decimal)x["vet_imp"] + Math.Abs((Decimal)tabTmp.Rows[0]["VenSct"])) * 100;
                foreach (DataRow k in tabSco.Rows)
                    k["ven_scn"] = ((Decimal)k["ven_ven"] * d / 100);
            }

            //x = tabVep.NewRow();
            //x["vet_neg"] = sNeg;
            //x["vet_cau"] = CAUSCO;
            //x["vet_day"] = dDay;
            //x["vet_ora"] = sOra;
            //x["vet_pos"] = sPos;
            //x["vet_sco"] = sSco;
            //x["vet_fid"] = tabTmp.Rows[0]["VenFid"];
            //x["vet_imp"] = tabTmp.Rows[0]["VenImp"];
            //x["vet_pun"] = tabTmp.Rows[0]["VenPun"];
            //x["vet_sct"] = tabTmp.Rows[0]["VenSct"];
            //x["vet_scp"] = tabTmp.Rows[0]["VenScp"];
            //x["vet_art"] = tabSco.Rows.Count;
            //tabVet.Rows.Add(x);
            //_clsFun.SqlWrite(_clsFun.SqlInsertRow(TABSTAVET, tabVet, x), _strConSqlSta);
            //if ((Decimal)tabTmp.Rows[0]["VenSct"] < 0)
            //{
            //    Decimal d = Math.Abs((Decimal)tabTmp.Rows[0]["VenSct"]) / ((Decimal)x["vet_imp"] + Math.Abs((Decimal)tabTmp.Rows[0]["VenSct"])) * 100;
            //    foreach (DataRow k in tabSco.Rows)
            //        k["ven_scn"] = ((Decimal)k["ven_ven"] * d / 100);
            //}

            if (((string)tabTmp.Rows[0]["VenFid"]).Trim() != "" && (decimal)tabTmp.Rows[0]["VenPun"] != 0)
            {
                decimal dPun = (decimal)tabTmp.Rows[0]["VenPun"];
                string sCod = "006";

                string[] a = _strConPun.Split(';');
                string sTip = "PUN";

                if (a.Length > 1)
                {
                    sTip = "PUN";
                    if (dPun < 0)
                    {
                        sTip = "PUM";
                        sCod = a[1];
                        dPun = dPun * -1;
                    }
                    else
                    {
                        sCod = a[0];
                    }
                }

                x = tabVep.NewRow();
                x["vep_cau"] = CAUSCO;
                x["vep_neg"] = sNeg;
                x["vep_day"] = dDay;
                x["vep_ora"] = sOra;
                x["vep_pos"] = sPos;
                x["vep_sco"] = sSco;
                x["vep_cod"] = sCod;
                x["vep_tri"] = "3";
                x["vep_tip"] = sTip;
                x["vep_off"] = "001";
                x["vep_imp"] = dPun;
                x["vep_val"] = 0;
                x["vep_ppo"] = "";
                x["vep_ord"] = "";
                s = _clsFun.SqlInsertRow(TABSTAVEP, tabVep, x);
                _clsFun.SqlWrite(s, _strConSqlSta);

                if ((decimal)tabTmp.Rows[0]["VenImp"] > 0)
                {
                    x = tabVep.NewRow();
                    x["vep_cau"] = CAUSCO;
                    x["vep_neg"] = sNeg;
                    x["vep_day"] = dDay;
                    x["vep_ora"] = sOra;
                    x["vep_pos"] = sPos;
                    x["vep_sco"] = sSco;
                    x["vep_cod"] = "001";
                    x["vep_tri"] = "1";
                    x["vep_tip"] = "PAG";
                    x["vep_off"] = "";
                    x["vep_imp"] = tabTmp.Rows[0]["VenImp"];
                    x["vep_val"] = 0;
                    x["vep_ppo"] = "1";
                    x["vep_ord"] = "";
                    s = _clsFun.SqlInsertRow(TABSTAVEP, tabVep, x);
                    _clsFun.SqlWrite(s, _strConSqlSta);
                }
            }

            foreach (DataRow y in tabSco.Rows)
            {
                //string s = (string)dataRow2["ven_ean"];
                //if (((string)dataRow2["ven_ean"]).Trim().Substring(0, 3) == "211")
                //    Console.WriteLine("aaaaaaaaaaaa");

                x = tabVen.NewRow();
                x["ven_neg"] = sNeg;
                x["ven_cau"] = CAUSCO;
                x["ven_day"] = dDay;
                x["ven_ora"] = sOra;
                x["ven_pos"] = sPos;
                x["ven_sco"] = sSco;
                x["ven_ean"] = y["ven_ean"];
                x["ven_art"] = y["ven_art"];
                x["ven_ard"] = y["ven_ard"];
                x["ven_iva"] = y["ven_iva"];
                x["ven_rep"] = y["ven_rep"];
                x["ven_qta"] = y["ven_qta"];
                x["ven_qkg"] = y["ven_qkg"];
                x["ven_prz"] = y["ven_prz"];
                x["ven_ven"] = ((Decimal)y["ven_ven"] - (Decimal)y["ven_scn"]);
                x["ven_sct"] = "";
                x["ven_scn"] = y["ven_scn"];
                x["ven_pun"] = y["ven_pun"];
                x["ven_res"] = y["ven_res"];
                x["ven_off"] = "";
                x["ven_sta"] = "0";
                x["ven_umi"] = y["ven_umi"];
                tabVen.Rows.Add(x);
                _clsFun.SqlWrite(_clsFun.SqlInsertRow("GesNegVen", tabVen, x), _strConSqlSta);
                d2 += (Decimal)y["ven_ven"] - (Decimal)y["ven_scn"];

                if ((decimal)y["ven_pun"] != 0 && ((string)tabTmp.Rows[0]["VenFid"]).Trim() != "")
                {
                    decimal dPun = (decimal)y["ven_pun"];
                    string sCod = "006";

                    string[] a = _strConPun.Split(';');
                    string sTip = "PUN";

                    if (a.Length > 1)
                    {
                        sTip = "PUN";
                        if (dPun < 0)
                        {
                            sTip = "PUM";
                            sCod = a[1];
                            dPun = dPun * -1;
                        }
                        else
                        {
                            sCod = a[0];
                        }
                    }

                    x = tabVep.NewRow();
                    x["vep_cau"] = CAUSCO;
                    x["vep_neg"] = sNeg;
                    x["vep_day"] = dDay;
                    x["vep_ora"] = sOra;
                    x["vep_pos"] = sPos;
                    x["vep_sco"] = sSco;
                    x["vep_cod"] = sCod;
                    x["vep_tri"] = "3";
                    x["vep_tip"] = sTip;
                    x["vep_off"] = "001";
                    x["vep_imp"] = dPun;
                    x["vep_val"] = 0;
                    x["vep_ppo"] = "";
                    x["vep_ord"] = "";
                    x["vep_ean"] = y["ven_ean"];
                    s = _clsFun.SqlInsertRow(TABSTAVEP, tabVep, x);
                    _clsFun.SqlWrite(s, _strConSqlSta);
                }

            }
            if (sSco == "00013")
                Console.WriteLine("Error");
            if (!(Math.Round(d2, 2) != (Decimal)tabTmp.Rows[0]["VenImp"]))
                return;
            Console.WriteLine("Error");
        }

        private void AggDgv1(string strTip, string strNeg, string strDes, DateTime dayDay, Decimal decVal)
        {
            DataTable t = (DataTable)dgv1.DataSource;
            DataRow[] j = t.Select("ven_tip='" + strTip + "' AND ven_neg='" + strNeg + "' AND ven_day=#" + dayDay.ToString("MM/dd/yyyy") + "#");
            if (j.Length == 0)
            {
                DataRow x = t.NewRow();
                x["ven_tip"] = strTip;
                x["ven_day"] = dayDay;
                x["ven_neg"] = strNeg;
                x["ven_des"] = strDes;
                x["ven_imp"] = decVal;
                t.Rows.Add(x);
            }
            else
                j[0]["ven_imp"] = decVal;
        }

        private void AggDgv2()
        {
            string s = "SELECT * FROM " + TABTABREP;
            DataTable tRep = _clsFun.FillTabSql(TABTABREP, s, false, _strConSql);
            DataTable t = new clsGenTabTmp().TabTmpVenReparto("TabRep");
            DataRow[] j;
            decimal d = 0;


            foreach(DataRow y in ((DataTable)dgv1.DataSource).Rows)
            {
                DataTable tSta = _clsQry.StaRep((DateTime)y["ven_day"]);

                foreach (DataRow k in tSta.Rows)
                {
                    DataRow x = t.NewRow();
                    x["vre_neg"] = k["ven_neg"];
                    x["vre_day"] = k["ven_day"];
                    x["vre_rep"] = k["ven_rep"];

                    x["vre_red"] = "";
                    j = tRep.Select("tab_cod='" + x["vre_rep"] + "'");
                    if(j.Length > 0)
                         x["vre_red"] = j[0]["tab_des"];

                    x["vre_ven"] = k["ven_ven"];
                    d += (decimal)x["vre_ven"];
                    t.Rows.Add(x);
                }

            }
            lblTot.Text = d.ToString();
            dgv2.DataSource = t;
        }

        private void btnPrnRep_Click(object sender, EventArgs e)
        {
            if (dgv2.DataSource == null || ((DataTable)dgv2.DataSource).Rows.Count == 0)
                MessageBox.Show("Non ci sono dati da stampare!");
            else
            {
                //frmGesPrn1 f = new frmGesPrn1();
                //f._tabPrn = (DataTable)dgv2.DataSource;
                //f.ShowDialog();
                //new clsGenPdfRep().PrnPdfRep((DataTable)dgv2.DataSource);
                PrnRep();
            }
        }

        private void btnVen_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource != null)
                ((DataTable)dgv1.DataSource).Rows.Clear();
            if (dgv2.DataSource != null)
                ((DataTable)dgv2.DataSource).Rows.Clear();

            DataRow x = ((DataTable)dgv1.DataSource).NewRow();
            x["ven_tip"] = "VEN";
            x["ven_day"] = dtpVenRep.Value;
            x["ven_neg"] = "001";
            x["ven_des"] = "Venduto";
            x["ven_imp"] = 0;
            ((DataTable)dgv1.DataSource).Rows.Add(x);

            if (dgv2.DataSource != null)
                ((DataTable)dgv2.DataSource).Rows.Clear();

            AggDgv2();
        }

        //private void CloApShop(DataTable tabCnf)
        //{
        //    //* Chiusure
        //    //004 DITRON|C:\ApProject\Terron\FileChiu.sem|C:\ApProject\Terron\
        //    //* Path casse ripetuta per ogni cassa
        //    //007 PCPOS|C:\ApProject\Temp\Venduto\|C:\ApProject\Temp\Venduto\

        //    string s = "";
        //    //string[] a = _strPar04Chiusura.Split('|');

        //    //if (a.Length > 1 && File.Exists(a[1]))
        //    //{

        //    string sPar06Pos = _clsFun.FileIni("R", clsDefine.enuIni.Ini06PathApShopCasse, "");

        //    string[] a = sPar06Pos.Split('|');

        //    Boolean b = true;

        //    ArrayList aPos = new ArrayList();

        //    for (int i = 1; i < a.Length; i++)
        //    {
        //        s = a[i].ToLower();

        //        int n = s.IndexOf("approject");

        //        string sPath = s.Substring(0, n) + "ApProject\\ApShop\\DataBase\\";

        //        string sFil = sPath + "apShopMain.sem";

        //        s = Path.GetDirectoryName(sFil);

        //        if (!Directory.Exists(s))
        //        {
        //            if (MessageBox.Show("Collegamento non attivo su cassa " + i.ToString() + ", annulli la chiusura?", "CONTROLLI CHIUSURA", MessageBoxButtons.YesNo, MessageBoxIcon.Error) == DialogResult.Yes)
        //                b = false;
        //        }
        //        else
        //        {
        //            if (!File.Exists(sPath + "apShopMain.sem"))
        //            {
        //                if (MessageBox.Show("Cassa " + i.ToString() + " non è sulla pagina iniziale, annulli la chiusura?", "CONTROLLI CHIUSURA", MessageBoxButtons.YesNo, MessageBoxIcon.Error) == DialogResult.Yes)
        //                    b = false;
        //            }
        //            if(b)
        //                aPos.Add(s + "\\");
        //        }
        //    }

        //    if (b)
        //    {
        //        foreach (string ss in aPos)
        //        {
        //            string sFil = Path.GetDirectoryName(ss) + "\\apShopChiudi.sem";
        //            string sMain = Path.GetDirectoryName(ss) + "\\apShopMain.sem";

        //            if (File.Exists(sMain) && !File.Exists(sFil))
        //                File.Move(sMain, sFil);
        //        }

        //        //new clsStaDitron().Chiusura(_strPar04Chiusura, _strPar07PosVenduto);
        //        //new clsStaDitron().FidMontanti(_strPar04Chiusura);

        //        //s = Path.GetDirectoryName(a[1]) + "\\Old\\" + Path.GetFileName(a[1]) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
        //        //File.Move(a[1], s);
        //    }
        //    else
        //        MessageBox.Show("Errori, chiusura non eseguita." + _clsDef.CRLF + "fare i controlli e rieseguire la chiusura." + _clsDef.CRLF + "Uscire e rientrare nel programma delle casse.", "CONTROLLI CHIUSURA", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);

        //    //}
        //}

        //private ArrayList LeggiApShop(string strNeg, string strFil, DataTable tabArt, DataTable tabVet, DataTable tabVen)
        //{
        //    string s = "";
        //    //Boolean bDec = true;
        //    ArrayList a = new ArrayList();
        //    DataTable tTmp = new clsGenTabTmp().TabTmpVet("TabVet");
        //    DataTable tSco = tabVen.Clone();
        //    Decimal num1 = new Decimal(0);
        //    string sPos = Path.GetFileName(strFil).Substring(10,1).PadLeft(2,Convert.ToChar("0"));
        //    //if (!chkPer100.Checked)
        //    //    bDec = false;
        //    string sFil = _clsDef.TMPPOSVEN + Path.GetFileName(strFil);
        //    //if (File.Exists(sFil))
        //    //    File.Delete(sFil);
        //    //File.Copy(strFil, sFil);
        //    using (StreamReader sr = new StreamReader(sFil))
        //    {
        //        FileInfo fI = new FileInfo(strFil);
        //        progressBar1.Value = 0;
        //        progressBar1.Maximum = (int)fI.Length;
        //        progressBar1.Minimum = 0;
        //        string sNum = "";
        //        string sRig;
        //        while ((sRig = sr.ReadLine()) != null)
        //        {
        //            progressBar1.Increment(sRig.Length);
        //            Application.DoEvents();
        //            string[] v = sRig.Split(';');
        //            if(v.Length > 0 && (string)v[0] != "sct_day")
        //            {
        //                DateTime dVetDay = Convert.ToDateTime(v[0]);
        //                string sVetOra = Convert.ToString(v[1]).Replace(".", "").Trim();
        //                string sVetNum = Convert.ToString(v[2]).PadLeft(5, Convert.ToChar("0")); ;
        //                string sVetFid = Convert.ToString(v[4]);
        //                decimal dVetQta = Convert.ToDecimal(v[5]);
        //                decimal dVetImp = Convert.ToDecimal(v[6]);
        //                decimal dVetPun = Convert.ToDecimal(v[7]);
        //                decimal dVetScv = Convert.ToDecimal(v[8]);
        //                decimal dVetScp = Convert.ToDecimal(v[9]);
        //                string sVenRig = Convert.ToString(v[10]);
        //                string sVenTri = Convert.ToString(v[11]);
        //                string sVenEan = Convert.ToString(v[12]);
        //                string sVenRep = Convert.ToString(v[13]).PadLeft(3, Convert.ToChar("0"));
        //                string sVenArt = Convert.ToString(v[14]).PadLeft(7, Convert.ToChar("0"));
        //                string sVenArd = Convert.ToString(v[15]);
        //                decimal dVenQta = Convert.ToDecimal(v[16]);
        //                decimal dVenPve = Convert.ToDecimal(v[17]);
        //                string sVenSct = Convert.ToString(v[18]);
        //                decimal dVenScn = Convert.ToDecimal(v[19]);
        //                decimal dVenImp = Convert.ToDecimal(v[20]);
        //                string sVenOff = Convert.ToString(v[21]);
        //                string sVenOft = Convert.ToString(v[22]);
        //                decimal dVenPun = Convert.ToDecimal(v[23]);
        //                if (sVetNum != sNum)
        //                {
        //                    if (tTmp.Rows.Count > 0)
        //                    {
        //                        VenScontrino(tabVet, tabVen, tSco, tTmp);
        //                        tTmp.Rows.Clear();
        //                    }
        //                    DataRow x = tTmp.NewRow();
        //                    x["VenNeg"] = strNeg;
        //                    x["VenDay"] = dVetDay;
        //                    x["VenOra"] = sVetOra;
        //                    x["VenPos"] = sPos;
        //                    x["VenSco"] = sVetNum;
        //                    x["VenFid"] = sVetFid;
        //                    tTmp.Rows.Add(x);
        //                    if (a.IndexOf(strNeg + dVetDay.ToString("yyyyMMdd")) < 0)
        //                        a.Add(strNeg + dVetDay.ToString("yyyyMMdd"));
        //                    sNum = sVetNum;
        //                }
        //                DataRow[] j;
        //                string sArt = "";
        //                if (sVenTri == "R")
        //                    j = tabArt.Select("art_sta='R' AND art_rep='" + sVenRep + "'");
        //                else
        //                {
        //                    j = tabArt.Select("art_cod='" + sVenArt + "'");
        //                    if(j.Length == 0)
        //                    {
        //                        j = tabArt.Select("art_sta='R' AND art_rep='" + sVenRep + "'");
        //                    }
        //                }
        //                if(j.Length > 0)
        //                    sArt = (string)j[0]["art_cod"];
        //                else
        //                {
        //                    sArt = "0000001";
        //                }
        //                j = tSco.Select("ven_art='" + sVenArt + "'");
        //                if (j.Length == 0)
        //                {
        //                    DataRow x = tSco.NewRow();
        //                    x["ven_neg"] = strNeg;
        //                    //DataRow dataRow = x;
        //                    x["ven_day"] = dVetDay;
        //                    x["ven_pos"] = sPos;
        //                    x["ven_ora"] = sVetOra;
        //                    x["ven_art"] = sVenArt;
        //                    x["ven_ard"] = sVenArd;
        //                    if (j.Length > 0)
        //                    {
        //                        x["ven_iva"] = ((string)j[0]["art_iva"]).Trim();
        //                        x["ven_art"] = ((string)j[0]["art_cod"]).Trim();
        //                        x["ven_rep"] = ((string)j[0]["art_rep"]).Trim();
        //                    }
        //                    x["ven_sco"] = sVetNum;
        //                    x["ven_ean"] = "";
        //                    x["ven_prz"] = dVenPve;
        //                    x["ven_qta"] = 0;
        //                    x["ven_qkg"] = 0;
        //                    x["ven_ven"] = 0;
        //                    x["ven_pun"] = dVenPun;
        //                    x["ven_scn"] = 0;
        //                    tSco.Rows.Add(x);
        //                    j = tSco.Select("ven_art='" + x["ven_art"] + "'");
        //                }
        //                j[0]["ven_qta"] = ((Decimal)j[0]["ven_qta"] + dVenQta);
        //                //j[0]["ven_qkg"] = ((Decimal)j[0]["ven_qkg"] + dVenQkg);
        //                j[0]["ven_ven"] = ((Decimal)j[0]["ven_ven"] + dVenImp);
        //                j[0]["ven_pun"] = ((Decimal)j[0]["ven_pun"] + dVenPun);
        //            }
        //        }
        //        VenScontrino(tabVet, tabVen, tSco, tTmp);
        //    }
        //    return a;
        //}

        public void ChiuNcr745x(DataTable tabCnf, string strPath, string strParFat)
        {
            string s = "";
            //s = "SELECT ean_art, ean_ean, art_sta, art_des, art_iva, art_rep, art_umi ";
            s = "SELECT art_cod, ean_ean, art_sta, art_des, art_iva, art_rep, art_umi ";
            s += "FROM AnaBarcode ";
            s += "LEFT JOIN AnaArticoli ON AnaBarcode.ean_art = AnaArticoli.art_cod";
            DataTable tEan = _clsFun.FillTabSql(TABANAEAN, s, false, _strConSql);
            DataColumn[] Key = new DataColumn[1] 
            { 
                tEan.Columns["ean_ean"] 
            };
            tEan.PrimaryKey = Key;

            s = "SELECT art_cod, art_sta, art_des, art_iva, art_rep, art_umi ";
            s += "FROM AnaArticoli ";
            s += "WHERE art_sta = '" + _clsDef.STAREP + "'";
            DataTable tArt = _clsFun.FillTabSql(TABANAART, s, false, _strConSql);
            Key = new DataColumn[1] 
            { 
                tArt.Columns["art_cod"] 
            };
            tArt.PrimaryKey = Key;

            s = "SELECT * FROM GesNegVet";
            DataTable tVet = _clsFun.FillTabSql("GesNegVet", s, false, _strConSqlSta);
            Key = new DataColumn[6]
            {
                tVet.Columns["vet_neg"],
                tVet.Columns["vet_cau"],
                tVet.Columns["vet_day"],
                tVet.Columns["vet_ora"],
                tVet.Columns["vet_pos"],
                tVet.Columns["vet_sco"]
            };
            tVet.PrimaryKey = Key;

            s = "SELECT * FROM GesNegVep";
            DataTable tVep = _clsFun.FillTabSql("GesNegVep", s, false, _strConSqlSta);
            Key = new DataColumn[8]
            {
                tVep.Columns["vep_neg"],
                tVep.Columns["vep_cau"],
                tVep.Columns["vep_day"],
                tVep.Columns["vep_ora"],
                tVep.Columns["vep_pos"],
                tVep.Columns["vep_sco"],
                tVep.Columns["vep_cod"],
                tVep.Columns["vep_ean"]
            };
            tVep.PrimaryKey = Key;

            s = "SELECT * FROM GesNegVen";
            DataTable tVen = this._clsFun.FillTabSql("GesNegVen", s, false, _strConSqlSta);
            Key = new DataColumn[9]
            {
                tVen.Columns["ven_neg"],
                tVen.Columns["ven_cau"],
                tVen.Columns["ven_day"],
                tVen.Columns["ven_ora"],
                tVen.Columns["ven_pos"],
                tVen.Columns["ven_sco"],
                tVen.Columns["ven_sct"],
                tVen.Columns["ven_ean"],
                tVen.Columns["ven_art"]
            };
            tVen.PrimaryKey = Key;

            DataTable tErr = new clsGenTabTmp().TabTmpVenErrori("TabErr");

            s = Path.GetDirectoryName((string)tabCnf.Rows[0]["CnfVen"]);

            foreach (string sFil in Directory.GetFiles(strPath))
            {
                //if (Path.GetFileName(sFil).Substring(0, 6).ToUpper() == "HOCIDC")

                //s = "HOCIDC";
                //if(strParFat != "")
                //    s = "S_IDC001.DAT";
                s = "S_IDC";

                if (Path.GetFileName(sFil).Substring(0,5).ToUpper() == s)
                {
                    string sNeg = (string)tabCnf.Rows[0]["cnf_cod"];
                    ArrayList ary = LeggiNcr745x(sNeg, sFil, tArt, tEan, tVet, tVep, tVen, tErr, strParFat);

                    if (strParFat == "")
                    {
                        foreach (string a in ary)
                        {
                            s = a;
                            DateTime day = new DateTime(Convert.ToInt32(s.Substring(3, 4)), Convert.ToInt32(s.Substring(7, 2)), Convert.ToInt32(s.Substring(9, 2)));
                            s = "SELECT SUM(vet_imp) AS Imp FROM GesNegVet WHERE vet_neg='" + sNeg + "' AND vet_day=" + _clsFun.DaySql(day);
                            DataTable t = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlSta);
                            if (t.Rows.Count > 0)
                            {
                                AggDgv1("VEN", sNeg, "Venduto", day, (Decimal)t.Rows[0]["Imp"]);
                            }
                        }
                        s = _clsDef.TMPPOSVEN + "Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
                        File.Copy(sFil, s);
                    }
                }
            }
            if (tErr.Rows.Count > 0)
                new clsGenPdfVenErr().PrnPdfVenErr(tErr);
        }

        private ArrayList LeggiNcr745x(string strNeg, string strFil, DataTable tabArt, DataTable tabEan, DataTable tabVet, DataTable tabVep, DataTable tabVen, DataTable tabErr, string strParFat)
        {
            string s = "";
            //Boolean bDec = true;
            ArrayList a = new ArrayList();
            DataTable tTmp = new clsGenTabTmp().TabTmpVet("TabVet");
            DataTable tSco = tabVen.Clone();
            Decimal num1 = new Decimal(0);
            //string sPos = "";
            //string sSco = "";
            //string sOra = "";
            DataRow[] j1;
            DataRow[] j2;
            //if (!chkPer100.Checked)
            //    bDec = false;

            string sFil = _clsDef.TMPPOSVEN + Path.GetFileName(strFil);

            string sRepPar = _clsFun.ParGet(clsDefine.enuParametri.Par013CodRepxDefault, _strConSql);
            string sIvaPar = _clsFun.ParGet(clsDefine.enuParametri.Par012CodIVAxDefault, _strConSql);
                
            if (File.Exists(sFil))
                File.Delete(sFil);
            File.Copy(strFil, sFil);

            using (StreamReader sr = new StreamReader(sFil))
            {
                FileInfo fI = new FileInfo(strFil);
                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fI.Length;
                progressBar1.Minimum = 0;
                string sRig;

                Boolean bOk = false;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    Application.DoEvents();

                    //if (sRig.Substring(5, 1) == "V")
                    //{

                    string sTri = sRig.Substring(32, 1);
                    if (sTri == "H")                        // if (sTri == "100") inizio
                    {
                        //sPos = sRig.Substring(60, 1);
                        //sSco = sRig.Substring(126, 5);

                        if (tTmp.Rows.Count > 0)
                            tTmp.Rows.Clear();
                        DataRow x = tTmp.NewRow();
                        x["VenNeg"] = strNeg;
                        x["VenDay"] = DateTime.Today.ToShortDateString();
                        x["VenOra"] = "";
                        x["VenPos"] = "";
                        x["VenSco"] = "";
                        x["VenFid"] = "";
                        tTmp.Rows.Add(x);

                        bOk = true;
                    }

                    else if(bOk)
                    { 

                        if (sTri == "S")
                        {
                            //if (sRig.Substring(10, 13).Trim() == "2110003")
                            //    Console.WriteLine("0000");
                            //if (sRig.Substring(59, 4) != "4220")
                            //{

                            Boolean bPes = false;

                            string sEan = sRig.Substring(46, 13).Trim();   //.Trim().PadLeft(13, Convert.ToChar("0"));

                            if (sEan.Length == 13 && sEan.Substring(0, 1) == "2")
                            {
                                sEan = sEan.Substring(0, 12) + "0";
                                bPes = true;
                            }
                            //else
                            //    s = Convert.ToInt64(s).ToString();
                            //sEan = s;

                            if (sRig.Contains("."))
                                Console.WriteLine("aaaaaaa");

                            string sRep = sRig.Substring(39, 3);
                            //string sArd = sRig.Substring(24, 20);
                            //Decimal dQta = Convert.ToDecimal(sRig.Substring(60, 8)) / 1000;
                            Decimal dQta = Convert.ToDecimal(sRig.Substring(60, 4));
                            if (sRig.Substring(59, 1) == "-")
                                dQta = dQta * -1;

                            decimal dQkg = 0;
                            if (bPes && sRig.Substring(64, 1) == ".")
                            {
                                dQta = 1;
                                dQkg = Convert.ToDecimal(sRig.Substring(60, 8));
                            }

                            Decimal dVen = Convert.ToDecimal(sRig.Substring(69, 9)) / 100;
                            //if (sRig.Substring(68, 1) == "-")
                            //    dVen = dVen * -1;

                            string sTre = sRig.Substring(36, 1);

                            decimal dPrz = dVen / dQta;

                            string sArt = "";
                            string sArd = "";
                            string sIva = "";
                            string sUmi = "";

                            if (sEan == "")
                                j2 = tabArt.Select("art_sta='" + _clsDef.STAREP + "' AND art_rep='" + sRep + "'");
                            else
                                j2 = tabEan.Select("ean_ean='" + sEan + "'");
                            if (j2.Length > 0)
                            {
                                sArt = (string)j2[0]["art_cod"];
                                sArd = (string)j2[0]["art_des"];
                                sIva = (string)j2[0]["art_iva"];
                                sUmi = (string)j2[0]["art_umi"];

                                //x["ven_art"] = ((string)j2[0]["ean_art"]).Trim();
                                //x["ven_ard"] = j2[0]["art_des"];

                                if (((string)j2[0]["art_sta"]).Trim() == "N")
                                {
                                    ErrMsg(tabErr, ((string)j2[0]["ean_art"]).Trim(), "Articolo attivato - " + sArd + " Q.tà " + dQta + " Prz " + dVen.ToString());
                                    s = "UPDATE AnaArticoli SET art_sta='A' WHERE art_cod='" + sArt + "'";
                                    _clsFun.SqlWrite(s, _strConSql);
                                }
                            }
                            else 
                            {
                                ErrMsg(tabErr, "", "Articolo/reparto non trovato - " + sRep +" "+ sEan + " Q.tà " + dQta + " Prz " + dVen.ToString());

                                j2 = tabArt.Select("art_sta='" + _clsDef.STAREP + "' AND art_rep='" + sRepPar + "'");
                                if (j2.Length > 0)
                                {
                                    sArt = (string)j2[0]["art_cod"];
                                    sArd = "REPARTO " + sRep + " NON DEFINITO IN ANAG. ARTICOLI";
                                    sIva = (string)j2[0]["art_iva"];
                                    sUmi = (string)j2[0]["art_umi"];
                                }
                                else
                                {
                                    sArt = "";
                                    sArd = "REPARTO " + sRep + " NON DEFINITO IN ANAG. ARTICOLI";
                                    sIva = sIvaPar;
                                    sUmi = "PZ";
                                }

                            }

                            j1 = tSco.Select("ven_art='" + sArt + "'");
                            if (j1.Length == 0)
                            {
                                DataRow x = tSco.NewRow();
                                x["ven_neg"] = strNeg;
                                x["ven_cau"] = "SCO";
                                x["ven_day"] = DateTime.Today.ToShortDateString();
                                x["ven_pos"] = "";
                                x["ven_sco"] = "";
                                x["ven_ora"] = "";
                                x["ven_art"] = sArt;
                                x["ven_ard"] = sArd;
                                x["ven_umi"] = sUmi;

                                if (sEan == "0000080019428")
                                    Console.WriteLine("aaaa");

                                //if (sEan == "")
                                //    j2 = tabEan.Select("art_sta='" + _clsDef.STAREP + "' AND art_rep='" + sRep + "'");
                                //else
                                //    j2 = tabEan.Select("ean_ean='" + sEan + "'");
                                //if (sEan != "" && j2.Length > 0)
                                //{
                                //    sArd = (string)j2[0]["art_des"];
                                //    sIva = (string)j2[0]["art_iva"];
                                //    x["ven_art"] = ((string)j2[0]["ean_art"]).Trim();
                                //    x["ven_ard"] = j2[0]["art_des"];
                                //    if (((string)j2[0]["art_sta"]).Trim() == "N")
                                //    {
                                //        ErrMsg(tabErr, ((string)j2[0]["ean_art"]).Trim(), "Articolo attivato - " + sArd + " Q.tà " + dQta + " Prz " + dVen.ToString());
                                //        s = "UPDATE AnaArticoli SET art_sta='A' WHERE art_cod='" + x["ven_art"] + "'";
                                //        _clsFun.SqlWrite(s, _strConSql);
                                //    }
                                //}
                                //else
                                //{
                                //    x["ven_art"] = "REP" + sRep.Substring(1, 2);
                                //    x["ven_ard"] = "REP" + sRep.Substring(1, 2);
                                //    ErrMsg(tabErr, sEan, "Non trovato - " + "" + " Q.tà " + dQta + " Prz " + dVen.ToString());
                                //}

                                x["ven_iva"] = sIva;
                                x["ven_ean"] = sEan;
                                x["ven_rep"] = sRep;
                                x["ven_prz"] = dPrz;
                                x["ven_qta"] = 0;
                                x["ven_qkg"] = 0;
                                x["ven_ven"] = 0;
                                x["ven_pun"] = 0;
                                x["ven_sct"] = "";
                                x["ven_scn"] = 0;
                                tSco.Rows.Add(x);
                                j1 = tSco.Select("ven_art='" + sArt + "'");
                            }
                            if ("478".Contains(sTre))
                                dQta *= -1;
                            j1[0]["ven_qta"] = ((Decimal)j1[0]["ven_qta"] + dQta);
                            j1[0]["ven_qkg"] = ((Decimal)j1[0]["ven_qkg"] + dQkg);
                            j1[0]["ven_ven"] = ((Decimal)j1[0]["ven_ven"] + dPrz * dQta);
                            j1[0]["ven_prz"] = ((Decimal)j1[0]["ven_ven"] / dQta);
                            j1[0]["ven_pun"] = ((Decimal)j1[0]["ven_pun"] + 0);
                            //}
                        }
                        else if (sTri == "C")
                        {
                            //Sconto 

                            //tTmp.Rows[0]["VenScp"] = (object)Convert.ToDecimal(sRig.Substring(51, 11).Replace(".", ","));
                            //tTmp.Rows[0]["VenSct"] = (object)(Convert.ToDecimal(sRig.Substring(69, 13).Replace(".", ",")) / new Decimal(100));
                        }
                        else if (sTri == "D")
                        {
                            //Sconto transazione

                            s = sRig.Substring(68,10);
                            if(_clsFun.Numerico(s))
                            {
                                decimal dSco = Convert.ToDecimal(s);
                                tTmp.Rows[0]["VenSct"] = dSco / 100;

                                s = sRig.Substring(50, 6);

                                dSco = Convert.ToDecimal(s);
                                //tTmp.Rows[0]["VenScp"] = dSco / 100;
                                tTmp.Rows[0]["VenScp"] = dSco;
                            }
                        }
                        else if (sTri == "p")
                        {
                            //Punti 
                            //                        Dim VarPlu As String
                            //Dim VarPun As String

                            //Const PosPlu = 47      'Codice plu
                            //Const LenPlu = 13

                            //Const PosPun = 69      'Punti articolo
                            //Const LenPun = 10
                        }
                        else if (sTri == "k")
                        {
                            //Fidelity
                            //Const PosClf = 50      'Codice cliente fidelity
                            //Const LenClf = 10      'Codice cliente fidelity
                            //VarClf = Trim(EanOk("4" + Mid(Riga, PosClf, LenClf)))
                            //If Not IsVuoto(Nul2Space(VarClf)) Then       'Seck 00004 20020515
                            //m_rsTmp.MoveFirst
                            //m_rsTmp!sco_clf = FillChar(Trim(VarClf), 13, "0", "D")
                            //m_rsTmp.Update
                            //End If
                        }
                        else if (sTri == "F")
                        {

                            string sTre = sRig.Substring(34,1);

                            if (sTre == "1")
                            {
                                s = sRig.Substring(9, 6);

                                DateTime dDay = _clsFun.Str2Day(s);
                                string sOra = sRig.Substring(16, 4);
                                string sPos = sRig.Substring(6, 2);
                                string sSco = "0" + sRig.Substring(43, 4);

                                Decimal dImp = Convert.ToDecimal(sRig.Substring(69, 9)) / 100;

                                Boolean b = true;

                                if(strParFat != "")
                                {
                                    string[] aa = strParFat.Split('|');

                                    DateTime dFatDay = _clsFun.Str2Day(aa[0]);
                                    string sFatPos = aa[1];
                                    string sFatSco = aa[2];

                                    TimeSpan diff = dFatDay.Subtract(dDay);

                                    if(diff.Days != 0 || sPos != sFatPos || sSco != sFatSco)
                                        b = false;
                                }

                                if (b)
                                {
                                    if (dImp > 0)
                                    {
                                        if (tSco.Rows.Count > 0)
                                        {
                                            //sOra = sRig.Substring(16, 2) + sRig.Substring(19, 2);
                                            if (a.IndexOf(strNeg + dDay.ToString("yyyyMMdd")) < 0)
                                                a.Add(strNeg + dDay.ToString("yyyyMMdd"));
                                            tTmp.Rows[0]["VenDay"] = dDay;
                                            tTmp.Rows[0]["VenImp"] = dImp;
                                            tTmp.Rows[0]["VenOra"] = sOra;
                                            tTmp.Rows[0]["VenPos"] = sPos;
                                            tTmp.Rows[0]["VenSco"] = sSco;
                                            VenScontrino(tabVet, tabVep, tabVen, tSco, tTmp);
                                        }
                                    }
                                }
                            }
                            tTmp.Rows.Clear();
                            tSco.Rows.Clear();
                            bOk = false;
                        }
                    }
                }
            }
            //s = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(strFil) + DateTime.Now.ToString("yyyyMMddHHmmss");
            //if (File.Exists(s))
            //    File.Delete(s);
            //File.Move(sFil, s);

            s = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(strFil) +"_"+ DateTime.Now.ToString("yyyyMMddHHmmss");
            if (File.Exists(s))
                File.Delete(s);
            File.Move(sFil, s);


            return a;
        }

        private void ErrMsg(DataTable tabErr, string strEan, string strMsg)
        {
            DataRow x = tabErr.NewRow();
            x["tmp_ean"] = strEan;
            x["tmp_msg"] = strMsg;
            tabErr.Rows.Add(x);
        }

        private void PrnRep()
        {
            frmGesStat f = new frmGesStat();
            f._bolPrnAuto = true;
            f.ShowDialog();
            f.Dispose();
        }

    }
}
