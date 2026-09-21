using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace APOffice
{
    public partial class frmGesImpTerm : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        public DataTable _tabImp = new DataTable("Tmp");
        public DataTable _tabTer = new DataTable("TmpTer");
        public string _strTip = "ORD";
        public string _strTer = "";
        public string _strDivTip = "";

        public frmGesImpTerm()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
        }

        private void frmGesImpTerm_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            SetDgv1();
            FillTabs();
            FillDati();

            if (_strTip == "INV")
                CtrlNrv();
            else
                grpInv.Visible = false;
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (cmbTer.SelectedValue != null)
                _strTer = cmbTer.SelectedValue.ToString();
            else
                _strTer = "";

            if (_strTip == "INV")
                _strDivTip = cmbNte.Text + lblNrv.Text;
            this.Close();
        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_msg";
            cTbc.Name = "Messaggio";
            cTbc.Width = 220;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillTabs()
        {
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabTerm";
            s = "SELECT * FROM " + p + " WHERE tab_ann=0";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            if (t.Rows.Count == 0)
                MessageBox.Show("Tabella terminalino non configurata!");
            else
            {
                cmbTer.DataSource = t;
                cmbTer.DisplayMember = "tab_des";
                cmbTer.ValueMember = "tab_cod";
                cmbTer.SelectedIndex = 0;
            }

            if(_strTip == "INV")
            {
                cmbNte.Items.Add("01");
                cmbNte.Items.Add("02");
                cmbNte.Items.Add("03");
                cmbNte.Items.Add("04");
                cmbNte.Items.Add("05");
            }
        }

        private void FillDati()
        {
            DataTable t = new DataTable("Msg");            
            t = new DataTable("Msg")

            {
                Columns = {                
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tab_msg",
                        Caption = "msg",
                        MaxLength = 100,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                }
            };

            dgv1.DataSource = t;
        }

        private void btnImp_Click(object sender, EventArgs e)
        {
            if (cmbTer.SelectedValue == null)
            {
                MessageBox.Show("Nessun terminalino selezionato o configurato!", "ERRORE", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (cmbTer.SelectedValue.ToString() == _clsDef.TERMMEMOR)
            {
                DataRow[] j = ((DataTable)cmbTer.DataSource).Select("tab_cod='" + _clsDef.TERMMEMOR + "'");
                if (((string)j[0]["tab_bao"]).Trim() == "")
                    MessageBox.Show("Batch ricezione ordini non definito in tabella term!");
                else
                    _tabImp = ImportMemor(((string)j[0]["tab_bao"]).Trim());
            }
            else if (cmbTer.SelectedValue.ToString() == _clsDef.TERMCSV1)
            {
                DataRow[] j = ((DataTable)cmbTer.DataSource).Select("tab_cod='" + _clsDef.TERMCSV1 + "'");
                if (((string)j[0]["tab_bao"]).Trim() == "")
                    MessageBox.Show("Batch ricezione ordini non definito in tabella term!");
                else
                    _tabImp = ImportMcSv1(((string)j[0]["tab_bao"]).Trim());
            }
            else if (cmbTer.SelectedValue.ToString() == _clsDef.TERMPDT1)
            {
                DataRow[] j = ((DataTable)cmbTer.DataSource).Select("tab_cod='" + _clsDef.TERMPDT1 + "'");
                if (((string)j[0]["tab_bao"]).Trim() == "")
                    MessageBox.Show("Batch ricezione ordini non definito in tabella term!");
                else
                    _tabImp = ImportPdt1(j[0]);
                //_tabImp = ImportPdt1(((string)j[0]["tab_bao"]).Trim());
            }
            else if (cmbTer.SelectedValue.ToString() == _clsDef.TERMTDIV)
            {
                //DataRow[] j = ((DataTable)cmbTer.DataSource).Select("tab_cod='" + _clsDef.TERMPDT1 + "'");
                //if (((string)j[0]["tab_bao"]).Trim() == "")
                //    MessageBox.Show("Batch ricezione ordini non definito in tabella term!");
                //else
                _tabImp = ImportTmpDiv();
            }
            else if (cmbTer.SelectedValue.ToString() == _clsDef.TERMOPH3)
            {
                DataRow[] j = ((DataTable)cmbTer.DataSource).Select("tab_cod='" + _clsDef.TERMOPH3 + "'");
                if (((string)j[0]["tab_bao"]).Trim() == "")
                    MessageBox.Show("Batch ricezione ordini non definito in tabella term!");
                else
                    _tabImp = ImportOph3000(((string)j[0]["tab_bao"]).Trim());
            }

            btnImp.Enabled = false;
        }

        private void cmbNte_SelectedIndexChanged(object sender, EventArgs e)
        {
            CtrlNrv();
        }

        private void CtrlNrv()
        {
            string s = cmbNte.Text;
            if (s == "")
                s = "01";
            DataRow[] j = _tabTer.Select("tmp_nte='" + s + "'");
            if (j.Length > 0)
            {
                cmbNte.SelectedItem = s;
                s = (Convert.ToInt16(j[0]["tmp_nrv"]) + 1).ToString("00");
                lblNrv.Text = s;
            }
            else
            {
                cmbNte.Text = "01";
                lblNrv.Text = "01";
            }
        }
 
        private DataTable ImportMemor(string strBat)
        {
            string s = "";
            Boolean b = true;
            DataTable tMsg = (DataTable)dgv1.DataSource;
            DataTable t = new clsGenTabTmp().TabTmpOrdTermMemor("TabImp");
            string sFil = "";
            string[] sFils = Directory.GetFiles(Path.GetDirectoryName(strBat));

            string sTip = _strTip;
            if (_strTip == "ETI")
                sTip = "INV";

            foreach (string sFi in sFils)
            {
                s = Path.GetFileName(sFi);

                if (s.Length > 3 && s.Substring(0, 3).ToUpper() == sTip)
                {
                    sFil = sFi;
                    break;
                }
            }

            if (sFil == "" || !File.Exists(sFil))
            {
                Directory.SetCurrentDirectory(Path.GetDirectoryName(strBat));

                try
                {
                    Process p = new Process();
                    p.StartInfo.FileName = strBat;
                    p.StartInfo.Arguments = "-r";
                    p.StartInfo.ErrorDialog = true;
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                    p.Start();
                    p.WaitForExit(1000 * 60 * 2);    // wait up to 2 minutes.
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    MessageBox.Show(ex.Message, "ERRORE");
                }

                sFils = Directory.GetFiles(Path.GetDirectoryName(strBat), _strTip + "*");
                if (sFils.Length > 0)
                    sFil = (string)sFils[0];
                else
                    b = false;

                if(sFil == "" && _strTip == "ORD")
                {
                    sFils = Directory.GetFiles(Path.GetDirectoryName(strBat), "INV" + "*");
                    if (sFils.Length > 0)

                        if (MessageBox.Show("Presente file INVENTARIO, procedi lettura come ordine?", "CARICAMENTO INVENTARIO COME ORDINE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                        {
                            sFil = (string)sFils[0];
                            b = true;
                        }
                        else
                            b = false;
                }

            }

            if (b)
            {
                using (StreamReader sr = new StreamReader(sFil))
                {
                    s = Path.GetFileName(sFil);
                    if(s.Length > 3)
                    {
                        if (s.Substring(0, 3).ToUpper() == "ORD")
                            s = "ordine";
                        if (s.Substring(0, 3).ToUpper() == "INV")
                            s = "inventario";
                        if (s.Substring(0, 3).ToUpper() == "ETI")
                            s = "etichette";
                    }

                    DataRow x = tMsg.NewRow();
                    x["tab_msg"] = "Letto " + s +" "+ Path.GetFileName(sFil); 
                    tMsg.Rows.Add(x);
                        
                    string sRig = "";
                    DataRow y;

                    while ((sRig = sr.ReadLine()) != null)
                    {
                        if (sRig.Length > 10)
                        {
                            y = t.NewRow();
                            s = sRig.Substring(20, 14);
                            if (s.Substring(0, 8) == "00000000")
                            {
                                if (_strDivTip.Trim() == _clsDef.DIVTERRON)
                                    y["ord_art"] = s.Substring(8, 6);
                                else
                                    y["ord_art"] = s.Substring(8, 6);
                            }
                            else
                            {
                                //y["ord_ean"] = Convert.ToInt64(sRig.Substring(21, 13)).ToString();
                                s = sRig.Substring(21, 13);
                                if(_clsFun.Numerico(s))
                                    y["ord_ean"] = Convert.ToInt64(s).ToString();
                                else
                                { 
                                    x = tMsg.NewRow();
                                    x["tab_msg"] = "Barcode non corretto " + s + " " + Path.GetFileName(sFil);
                                    tMsg.Rows.Add(x);
                                }
                            }
                            y["ord_qta"] = Convert.ToDecimal(sRig.Substring(34, 8))/100;                     //Q.tà
                            t.Rows.Add(y);
                        }
                    }
                }

                s = Path.GetDirectoryName(strBat) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                File.Move(sFil, s);
            }

            return t;
        }

        private DataTable ImportMcSv1(string strBat)
        {
            string s = "";
            Boolean b = true;
            DataTable tMsg = (DataTable)dgv1.DataSource;
            DataTable t = new clsGenTabTmp().TabTmpOrdTermMemor("TabImp");
            string sFil = "";
            string[] sFils = Directory.GetFiles(Path.GetDirectoryName(strBat));

            string sDayTim = "0";

            foreach (string sFi in sFils)
            {
                s = Path.GetFileName(sFi);
                if (s.Length > 3 && s.Substring(0, 3).ToUpper() == "MC3")
                {
                    FileInfo fI = new FileInfo(sFi);

                    //string ss = s.Substring(13, 14);
                    string ss = fI.LastWriteTime.ToString("yyyyMMddHHmm");

                    Console.WriteLine("xxxxxxx");

                    if (Convert.ToInt64(ss) > Convert.ToInt64(sDayTim))
                    {
                        sFil = sFi;
                        sDayTim = ss; // s.Substring(13, 14);
                    }
                    //break;
                }
            }

            if (sFil == "" || !File.Exists(sFil))
            {
                //Directory.SetCurrentDirectory(Path.GetDirectoryName(strBat));

                //try
                //{
                //    Process p = new Process();
                //    p.StartInfo.FileName = strBat;
                //    p.StartInfo.Arguments = "-r";
                //    p.StartInfo.ErrorDialog = true;
                //    p.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                //    p.Start();
                //    p.WaitForExit(1000 * 60 * 2);    // wait up to 2 minutes.
                //}
                //catch (Exception ex)
                //{
                //    Console.WriteLine(ex.Message);
                //    MessageBox.Show(ex.Message, "ERRORE");
                //}

                //sFils = Directory.GetFiles(Path.GetDirectoryName(strBat), _strTip + "*");
                //if (sFils.Length > 0)
                //    sFil = (string)sFils[0];
                //else
                //    b = false;
            }

            if (b && sFil != "")
            {
                using (StreamReader sr = new StreamReader(sFil))
                {
                    s = Path.GetFileName(sFil);
                    if (s.Length > 3)
                    {
                        if (s.Substring(0, 4).ToUpper() == "MC3O")
                            s = "ordine";
                        if (s.Substring(0, 4).ToUpper() == "MC3I")
                            s = "inventario";
                    }

                    DataRow x = tMsg.NewRow();
                    x["tab_msg"] = "Letto " + s + " " + Path.GetFileName(sFil);
                    tMsg.Rows.Add(x);

                    string sRig = "";
                    DataRow y;

                    while ((sRig = sr.ReadLine()) != null)
                    {
                        if (sRig.Length > 10)
                        {
                            y = t.NewRow();
                            if(sRig.Substring(0,1) == "M")
                                s = sRig.Substring(30, 13).Trim();
                            else
                                s = sRig.Substring(23, 13).Trim();
                            if (s.Length == 7)
                            {   
                                y["ord_art"] = s;
                            }
                            else
                            {
                                if (_clsFun.Numerico(s))
                                    y["ord_ean"] = Convert.ToInt64(s).ToString();
                                else
                                {
                                    x = tMsg.NewRow();
                                    x["tab_msg"] = "Barcode non corretto " + s + " " + Path.GetFileName(sFil);
                                    tMsg.Rows.Add(x);
                                }
                            }
                            y["ord_qta"] = Convert.ToDecimal(sRig.Substring(43, 10)) / 1000;                     //Q.tà
                            t.Rows.Add(y);
                        }
                    }
                }

                //s = Path.GetDirectoryName(strBat) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                s = "C:\\Mc3000\\Backup\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss"); ;
                File.Move(sFil, s);
            }

            return t;
        }

        private DataTable ImportPdt1(DataRow rowTer)
        {
            string sBat = (string)rowTer["tab_bao"];

            // PDT1 versione GMAX
            // PDT2 versione RInaldo Rovigo
            string sDes = ((string)rowTer["tab_des"]).Trim();

            string sTip = "DATI.TXT";

            if (((string)rowTer["tab_div"]).Trim() != "")
            {
                sTip = Path.GetFileName(((string)rowTer["tab_div"])).Trim();
            }
            string s = "";
            Boolean b = true;
            DataTable tMsg = (DataTable)dgv1.DataSource;
            DataTable t = new clsGenTabTmp().TabTmpOrdTermMemor("TabImp");
            string sFil = "";

            s = Path.GetDirectoryName(sBat);
            string[] sFils = Directory.GetFiles(s);

            foreach (string sFi in sFils)
            {
                s = Path.GetFileName(sFi);

                if(sDes == "PDT4")
                {
                    s = s.ToLower();

                    if (s.Contains(sTip.ToLower()))
                    {
                        sFil = sFi;
                        break;
                    }
                }
                else
                {
                    if (s.Length >= sTip.Length && s.Substring(0, sTip.Length).ToUpper() == sTip.ToUpper())
                    {
                        sFil = sFi;
                        break;
                    }
                }
            }

            if (sFil == "" || !File.Exists(sFil))
            {
                Directory.SetCurrentDirectory(Path.GetDirectoryName(sBat));

                try
                {

                    Process p = new Process();
                    p.StartInfo.FileName = sBat;
                    p.StartInfo.Arguments = "-r";
                    p.StartInfo.ErrorDialog = true;
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                    p.Start();
                    p.WaitForExit(1000 * 60 * 2);    // wait up to 2 minutes.
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    MessageBox.Show(ex.Message, "ERRORE");
                }

                sFils = Directory.GetFiles(Path.GetDirectoryName(sBat), sTip + "*");
                //sFils = Directory.GetFiles(Path.GetDirectoryName(strBat), "DATI" + "*");
                if (sFils.Length > 0)
                    sFil = (string)sFils[0];
                else
                    b = false;
            }

            if (b)
            {
                using (StreamReader sr = new StreamReader(sFil))
                {
                    s = Path.GetFileName(sFil);
                    if (s.Length > 3)
                    {
                        if (s.Substring(0, 3).ToUpper() == "ORD")
                            s = "ordine";
                        if (s.Substring(0, 3).ToUpper() == "INV")
                            s = "inventario";
                        if(_strTip == "SCA")
                            s = "scadenze";
                    }

                    DataRow x = tMsg.NewRow();
                    x["tab_msg"] = "Letto " + s + " " + Path.GetFileName(sFil);
                    tMsg.Rows.Add(x);

                    string sRig = "";
                    DataRow y;

                    if (sDes == "PDT1")
                    {
                        while ((sRig = sr.ReadLine()) != null)
                        {
                            if (sRig.Length > 10)
                            {
                                y = t.NewRow();
                                s = sRig.Substring(25, 13);
                                if (_clsFun.Numerico(s))
                                    y["ord_ean"] = Convert.ToInt64(s).ToString();
                                else
                                {
                                    x = tMsg.NewRow();
                                    x["tab_msg"] = "Barcode non corretto " + s + " " + Path.GetFileName(sFil);
                                    tMsg.Rows.Add(x);
                                }

                                s = sRig.Substring(38, 8);

                                if (_clsFun.Numerico(s))
                                    y["ord_qta"] = Convert.ToDecimal(s) / 100;                     //Q.tà
                                t.Rows.Add(y);
                            }
                        }
                    }
                    if (sDes == "PDT2")
                    {
                        while ((sRig = sr.ReadLine()) != null)
                        {
                            string[] a = sRig.Split('+');

                            foreach (string ss in a)
                            {
                                if (ss.Length > 22)
                                {
                                    y = t.NewRow();
                                    s = ss.Substring(1, 13);
                                    if (_clsFun.Numerico(s))
                                    {
                                        y["ord_ean"] = Convert.ToInt64(s).ToString();

                                        s = (string)y["ord_ean"];

                                        if (s.Length < 8)
                                        {
                                            s = s.PadLeft(7, Convert.ToChar('0'));
                                            y["ord_art"] = s;
                                            y["ord_ean"] = "";
                                        }
                                    }
                                    else
                                    {
                                        x = tMsg.NewRow();
                                        x["tab_msg"] = "Barcode non corretto " + s + " " + Path.GetFileName(sFil);
                                        tMsg.Rows.Add(x);
                                    }

                                    s = ss.Substring(14, 8);

                                    if (_clsFun.Numerico(s))
                                    {
                                        //if (_strTip == "SCA")
                                        //    y["ord_qta"] = Convert.ToDecimal(s) / 100;                     //Scadenza m,d
                                        //else
                                        //    y["ord_qta"] = Convert.ToDecimal(s) / 100;                     //Q.tà
                                        y["ord_qta"] = Convert.ToDecimal(s) / 100;                     //Q.tà
                                    }

                                    t.Rows.Add(y);
                                }
                            }
                        }
                    }
                    if (sDes == "PDT3" || sDes == "PDT4")
                    {
                        while ((sRig = sr.ReadLine()) != null)
                        {
                            if (sRig.Length > 10)
                            {
                                y = t.NewRow();
                                s = sRig.Substring(18, 13);
                                if (_clsFun.Numerico(s))
                                    y["ord_ean"] = Convert.ToInt64(s).ToString();
                                else
                                {
                                    x = tMsg.NewRow();
                                    x["tab_msg"] = "Barcode non corretto " + s + " " + Path.GetFileName(sFil);
                                    tMsg.Rows.Add(x);
                                }

                                s = sRig.Substring(31, 8);

                                if (_clsFun.Numerico(s))
                                    y["ord_qta"] = Convert.ToDecimal(s) / 100;                     //Q.tà
                                t.Rows.Add(y);
                            }
                        }
                    }

                }

                s = Path.GetDirectoryName(sBat) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                File.Move(sFil, s);
            }

            return t;
        }

        private DataTable ImportTmpDiv()
        {
            DataTable tMsg = (DataTable)dgv1.DataSource;
            DataTable t = new clsGenTabTmp().TabTmpOrdTermMemor("TabImp");
            DataRow x;
            string sMsgg = "";
            string sTime = "";

            string s = "SELECT * FROM TmpDiv WHERE tmp_tip='" + _strTip + "' AND tmp_sta='S'";
            DataTable tDiv = _clsFun.FillTabSql("TmpDiv", s, false, _strConSql);
            if (tDiv.Rows.Count > 0)
            {
                foreach(DataRow y in tDiv.Rows)
                {
                    //ord_neg, ord_ter, ord_num, ord_ean, ord_art, ord_ard, ord_qta, ord_for

                    string sRig = (string)y["tmp_tmp"];

                    try
                    {
                        string[] aRig = sRig.Split(';');

                        string sEan = "";
                        string sArt = "";
                        string sArd = "";
                        decimal dQta = 0;
                        string sFor = "";
                        string sMsg = "";

                        if (_strTip == "ADV")
                        {
                            sEan = "";
                            sArt = aRig[2].Trim();
                            sArd = aRig[3].Trim();
                            s = aRig[7].Replace(".", ",");
                            dQta = 0;
                            if (_clsFun.Numerico(s, "0123456789.,"))
                                dQta = Convert.ToDecimal(s);
                        }
                        else
                        {
                            sEan = aRig[3].Trim();
                            sArt = aRig[4].Trim();
                            sArd = aRig[5].Trim();
                            s = aRig[6].Replace(".", ",");
                            dQta = 0;
                            if (s != "")
                                dQta = Convert.ToDecimal(s);
                            if (aRig.Length > 7)
                                sFor = aRig[7].Trim();

                            if (sArt == "SoloEan" || sArt == sEan || sArt.Length > 7)
                            {
                                if (string.IsNullOrEmpty(sEan) && sArt.Length > 7)
                                    sEan = sArt;
                                sMsg = "Solo EAN";
                                sArt = "";
                            }
                        }

                        if (sEan.Length > 13)
                            sEan = sEan.Substring(0, 13);

                        x = t.NewRow();
                        x["ord_ean"] = sEan;
                        x["ord_art"] = sArt;
                        x["ord_qta"] = dQta;
                        x["ord_for"] = sFor;
                        x["ord_msg"] = sMsg;
                        t.Rows.Add(x);

                        if (sTime == "" || Convert.ToDouble(sTime) < Convert.ToDouble((string)y["tmp_day"]))
                            sTime = (string)y["tmp_day"];
                    }
                    catch (Exception ex)
                    {
                        sMsgg += sRig + _clsDef.CRLF;
                    }
                }
            }

            if (sTime != "")
            {
                s = "UPDATE TmpDiv SET tmp_sta='N' WHERE tmp_tip='" + _strTip + "' AND tmp_day <= " + sTime;
                _clsFun.SqlWrite(s, _strConSql);
            }

            x = tMsg.NewRow();
            x["tab_msg"] = "Letto " + t.Rows.Count.ToString() + " righe";
            tMsg.Rows.Add(x);
            //if(sMsgg.Trim() != "")
            //    MessageBox.Show(sMsgg, "Righe con errore", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (sMsgg.Trim() != "")
                MessageBox.Show(sMsgg, "Righe con errore", MessageBoxButtons.OK, MessageBoxIcon.Error);


            return t;
        }

        private DataTable ImportOph3000(string strBat)
        {
            string sReg1 = _clsFun.ParGet(clsDefine.enuParametri.Par032ImportEanizzati, _strConSql);

            string s = "";
            Boolean b = true;
            DataTable tMsg = (DataTable)dgv1.DataSource;
            DataTable t = new clsGenTabTmp().TabTmpOrdTermMemor("TabImp");
            string sFil = "";
            string[] sFils = Directory.GetFiles(Path.GetDirectoryName(strBat));

            string s2 = Path.GetFileName(strBat);

            if (s2 == "")
            {
                MessageBox.Show("Prefisso nome file non definito in tabella!", "TABELLA NON COMPLETA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                b = false;
            }

            if(b)
            {
                foreach (string sFi in sFils)
                {
                    s = Path.GetFileName(sFi);

                    if (s.Length >= s2.Length && s.Substring(0, s2.Length).ToUpper() == s2.ToUpper())
                    {
                        sFil = sFi;
                        using (StreamReader sr = new StreamReader(sFil))
                        {
                            s = Path.GetFileName(sFil);
                            if (s.Length > 3)
                            {
                                if (s.Substring(0, 3).ToUpper() == "ORD")
                                    s = "ordine";
                                if (s.Substring(0, 3).ToUpper() == "INV")
                                    s = "inventario";
                                if (s.Substring(0, 3).ToUpper() == "ETI")
                                    s = "etichette";
                            }

                            DataRow x = tMsg.NewRow();
                            x["tab_msg"] = "Letto " + s + " " + Path.GetFileName(sFil);
                            tMsg.Rows.Add(x);

                            string sRig = "";
                            DataRow y;

                            while ((sRig = sr.ReadLine()) != null)
                            {
                                string sArf = "";

                                if (sRig.Length > 10)
                                {
                                    string[] a = sRig.Split(',');

                                    if (a.Length > 2)
                                    {
                                        string sEan = a[0].Trim();
                                        if (sEan.Length > 0 && sEan.Length <= 13 && _clsFun.Numerico(sEan))
                                        {
                                            sEan = Convert.ToDouble(sEan).ToString();

                                            if (sReg1.Length > 0)
                                            {
                                                String[] aa = sReg1.Split(',');
                                                String sPrf = aa[0];
                                                int iLenEan = Convert.ToInt16(aa[1]);
                                                int iLenCod = Convert.ToInt16(aa[2]);
                                                string sFor = aa[3];

                                                if (sEan.Length > sPrf.Length)
                                                {
                                                    s = sEan.Substring(0, sPrf.Length);

                                                    if (sEan.Length == iLenEan && sEan.Substring(0, sPrf.Length) == sPrf)
                                                    {
                                                        //sArf = sEan.Substring(sPrf.Length, sPrf.Length + iLenCod);
                                                        sArf = sEan.Substring(sPrf.Length, iLenCod);
                                                        sEan = "";
                                                    }
                                                }
                                            }

                                            decimal dQta = 0;
                                            string sMsg = "";

                                            string sDay = _clsFun.StrDayTime(a[2], a[3]);

                                            s = a[1].Trim();
                                            if (_clsFun.Numerico(s))
                                            {
                                                dQta = Convert.ToInt64(s);

                                                if (dQta > 9999)
                                                {
                                                    dQta = 0;
                                                    sMsg = "Quantità eccessiva " + s;
                                                }
                                            }
                                            y = t.NewRow();
                                            y["ord_art"] = sArf;
                                            y["ord_ean"] = sEan;
                                            y["ord_qta"] = dQta;
                                            y["ord_msg"] = sMsg;
                                            y["ord_dtt"] = sDay;
                                            t.Rows.Add(y);
                                        }
                                    }
                                }
                            }

                            if(t.Rows.Count > 0)
                            {
                                DataTable t2 = t.Clone();
                                DataView v = new DataView(t, "", "ord_dtt", DataViewRowState.CurrentRows);
                                foreach(DataRowView r in v)
                                {
                                    t2.ImportRow(r.Row);
                                }
                                t = t2.Copy();
                            }
                        }

                        s = Path.GetDirectoryName(strBat) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                        File.Move(sFil, s);
                    }
                }
            }
            return t;
        }

        private void recuperoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s = "";
            string sPath = "";

            if (cmbTer.SelectedValue.ToString() == _clsDef.TERMMEMOR)
            {
                DataRow[] j = ((DataTable)cmbTer.DataSource).Select("tab_cod='" + _clsDef.TERMMEMOR + "'");
                if (((string)j[0]["tab_bao"]).Trim() == "")
                    MessageBox.Show("Batch ricezione ordini non definito in tabella term!");
                else
                    sPath = Path.GetDirectoryName((string)j[0]["tab_bao"]).Trim();
            }

            s = sPath + "\\Old"; 
            openFileDialog1.InitialDirectory = s; 
            openFileDialog1.FileName = String.Empty;
            DialogResult dr = openFileDialog1.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    string sFil = openFileDialog1.FileName;
                    s = sPath + "\\" + Path.GetFileName(sFil);
                    File.Copy(sFil, s);

                    if (((DataTable)dgv1.DataSource).Rows.Count == 0)
                        btnImp.Enabled = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("File non disponibile: " + ex.Message + s);
                }
            }

        }

        private void ftpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s = "SELECT * FROM TabForImport WHERE tab_tip='FTPDIV'";
            DataTable t = _clsFun.FillTabSql("TabForImp", s, true, _strConSql);
            if (t.Rows.Count > 0)
                new frmGesForDivulgazioni().ForDownLoad();
            else
                MessageBox.Show("Opzione non attiva!", "IMPORTAZIONE DA FTP", MessageBoxButtons.OK, MessageBoxIcon.Question);

            s = "SELECT * FROM TabForImport WHERE tab_tip='FTPDIV'";
            t = _clsFun.FillTabSql("TabForImp", s, true, _strConSql);

            ArrayList aPre = new ArrayList();
            aPre.Add("eti_");

            if (t.Rows.Count > 0)
                new clsQuery().ApPhoneDiv2TmpDiv(((string)t.Rows[0]["tab_pth"]).Trim(), aPre);




        }
    }
}
