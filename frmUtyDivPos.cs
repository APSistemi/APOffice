using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    public partial class frmUtyDivPos : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";

        public frmUtyDivPos()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmGesDivPos_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            SetDgv1();
            SetDgv2();
            dgv2.DataSource = new clsGenTabTmp().TabTmpMsg("TabMsg");
            FillPos();
        }
        private void frmUtyDivPos_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
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
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "pos_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "pos_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "pos_inv";
            cCbc.Name = "Invia";
            cCbc.Width = 45;
            dgv1.Columns.Add(cCbc);
        }

        private void SetDgv2()
        {
            dgv2.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv2.AllowUserToAddRows = false;
            dgv2.ReadOnly = false;
            dgv2.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "msg_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 240;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);
        }

        private void FillPos()
        {
            DataTable tCnf = _clsQry.ConfSeek("", "POS");
            if(tCnf.Rows.Count > 0)
            {
                string s = Convert.ToString(tCnf.Rows[0]["CnfVar"]);

                if (_clsFun.Numerico(s))
                {
                    int n = Convert.ToUInt16(s);

                    DataTable t = new DataTable("TabPos");

                    t.Columns.Add(new DataColumn()
                    {
                        DataType = Type.GetType("System.Boolean"),
                        ColumnName = "pos_inv",
                        Caption = "Invio",
                        ReadOnly = false,
                        DefaultValue = (Boolean)true
                    });
                    t.Columns.Add(new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "pos_cod",
                        Caption = "Pos",
                        MaxLength = 2,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    });
                    t.Columns.Add(new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "pos_des",
                        Caption = "Descrizione",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    });

                    for (int i = 1; i <= n; i++)
                    {
                        DataRow x = t.NewRow();
                        x["pos_inv"] = true;
                        x["pos_cod"] = i.ToString("00");
                        x["pos_des"] = "POS N. " + i.ToString("00");
                        t.Rows.Add(x);
                    }

                    dgv1.DataSource = t;
                }
                else
                    MessageBox.Show("Numero casse non definito in tabella POS!");
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (!File.Exists(_clsDef.TMPPOSSTOP))
            {
                var f = File.Create(_clsDef.TMPPOSSTOP);
                f.Close();
            }

            try
            {
                DataTable t = (DataTable)dgv1.DataSource;

                string sPos = "";
                foreach (DataRow y in t.Rows)
                {
                    if ((Boolean)y["pos_inv"])
                    {
                        Console.WriteLine((string)y["pos_des"]);
                        sPos += (string)y["pos_cod"] + ";";
                    }
                }
                sPos = sPos.Substring(0, sPos.Length - 1);

                if (chkTabUmi.Checked)
                    GenVar(sPos, "UMI");
                if (chkTabIva.Checked)
                    GenVar(sPos, "IVA");
                if (chkTabRep.Checked)
                    GenVar(sPos, "REP");
                if (chkAnaArt.Checked)
                {
                    string s = "SELECT * FROM TabNegozi";
                    DataTable tNeg = _clsFun.FillTabSql("TabNegozi", s, false, _strConSql);
                    foreach(DataRow y in tNeg.Rows)
                    {
                        if ((string)y["tab_pos"] != "")
                            Console.WriteLine("zzzzzzzz");

                        Boolean b = false;
                        string[] a = ((string)y["tab_pos"]).Split(',');

                        foreach(string s1 in a)
                        {
                            string sP = s1.PadLeft(2, Convert.ToChar('0'));

                            string[] sz = sPos.Split(';');

                            foreach(string s2 in sz)
                            {
                                if(sP == s2)
                                {
                                    b = true;
                                    GenVarArt(sPos, "ART", y);
                                    break;
                                }
                            }
                            if (b)
                                break;
                        }
                    }
                }
                if (chkTabPpo.Checked)
                    GenVar(sPos, "PPO");
                if (chkTabUsr.Checked)
                    GenVar(sPos, "USR");
                if (chkTabGrt.Checked)
                    GenVar(sPos, "GRT");
                if (chkTabCam.Checked)
                    GenVar(sPos, "CAM");
                if (chkTabMdo.Checked)
                    GenVarCauMovDoc(sPos);
                if (chkTabPag.Checked)
                    GenVar(sPos, "PAG");
                if (chkTabSta.Checked)
                    GenVar(sPos, "STA");
                if (chkTabNot.Checked)
                    GenVar(sPos, "NOT");
                if (chkTabPos.Checked)
                    GenVar(sPos, "POS");
            }
            catch (OutOfMemoryException oomEx)
            {
                _clsFun.ErrorLog("frmUtyDivPos_btnOk_OutOfMemory", oomEx.Message + "\r\n" + oomEx.StackTrace);
                MessageBox.Show("Memoria insufficiente durante la generazione dei file casse ApShop.\r\n" + oomEx.Message,
                    "Errore Memoria", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmUtyDivPos_btnOk", ex.Message + "\r\n" + ex.StackTrace);
                MessageBox.Show("Errore durante la generazione dei file casse ApShop: " + ex.Message,
                    "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (File.Exists(_clsDef.TMPPOSSTOP))
                    File.Delete(_clsDef.TMPPOSSTOP);
            }
        }

        private void GenVar(string strPos, string strTip)
        {
            try
            {
                string p = "";
                string s = "";
                string sRig = "";

                if (strTip == "UMI")
                {
                    p = "TabUmi";
                    s = "SELECT * FROM " + p;
                }
                if (strTip == "PPO")
                {
                    p = "TabCauCassa";
                    s = "SELECT * FROM " + p;
                }
                if (strTip == "USR")
                {
                    p = "TabUtenti";
                    s = "SELECT * FROM " + p;
                }
                if (strTip == "GRT")
                {
                    p = "TabFidGruppi";
                    s = "SELECT * FROM " + p;
                }
                if (strTip == "CAM")
                {
                    p = "TabFidCampagne";
                    s = "SELECT * FROM " + p;
                }
                if (strTip == "IVA")
                {
                    p = "TabIva";
                    s = "SELECT * FROM " + p;
                }
                if (strTip == "REP")
                {
                    p = "TabReparti";
                    s = "SELECT * FROM " + p;
                }
                if (strTip == "PAG")
                {
                    p = "TabPagamenti";
                    s = "SELECT * FROM " + p;
                }
                if (strTip == "STA")
                {
                    p = "TabStato";
                    s = "SELECT * FROM " + p;
                }
                if (strTip == "NOT")
                {
                    p = "TabNote";
                    s = "SELECT * FROM " + p;
                }
                if (strTip == "POS")
                {
                    p = "TabPos";
                    s = "SELECT * FROM " + p + " WHERE tab_ann=0";
                }
                if (strTip == "ART")
                {
                    p = "AnaArticoli";
                    s = "SELECT * FROM AnaArticoli ";
                    if (rdbAtt.Checked)
                        s += "WHERE art_sta='A' OR art_sta='R'";
                }

                DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);

                string[] aryPos = strPos.Split(';');

                string sFil = Path.GetDirectoryName(_clsDef.TMPPOSVAR) + "\\";
                s = _clsFun.FileIni("R", clsDefine.enuIni.Ini07PathDivCasse, "");
                if (s != "")
                {
                    s = Path.GetDirectoryName(s) + "\\";
                    sFil = s;
                }

                if (strTip == "ART")
                    sFil = _clsQry.DivFilArtApShop(aryPos.Length, _clsDef.TMPPOSVAR);
                else if (p.Substring(0, 3) == "Tab")
                    sFil += "Tab" + strTip;
                else
                    sFil += p.Substring(0, 6);

                StreamWriter sw = new StreamWriter(sFil, false, Encoding.Default);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = Math.Max(1, t.Rows.Count);
                progressBar1.Minimum = 0;

                DataTable tMsg = (DataTable)dgv2.DataSource;
                dgv2.DataSource = null;

                sRig = "";

                if (strTip == "ART")
                {
                    DataTable tTmpSchema = new clsGenTabTmp().TabTmpDivArtApShop("TabDiv");
                    foreach (DataColumn c in tTmpSchema.Columns)
                        sRig += c.ColumnName + ";";
                    sw.Write(sRig + _clsDef.CRLF);
                    tTmpSchema.Dispose();
                }
                else
                {
                    foreach (DataColumn c in t.Columns)
                    {
                        if (c.ColumnName.Substring(4, 3) != "idx")
                            sRig += c.ColumnName + ";";
                    }
                    sw.Write(sRig + _clsDef.CRLF);
                }

                int iCount = 0;
                int totalCount = t.Rows.Count;

                try
                {
                    foreach (DataRow y in t.Rows)
                    {
                        iCount++;
                        if (iCount % 50 == 0 || iCount == totalCount)
                        {
                            progressBar1.Value = Math.Min(iCount, progressBar1.Maximum);
                            Application.DoEvents();
                        }

                        s = "";
                        if (strTip == "ART")
                        {
                            if (((string)y["art_des"]).Trim() == "")
                                s += (string)y["art_cod"] + " Manca descrizione ";
                            if (((string)y["art_rep"]).Trim() == "")
                                s += (string)y["art_cod"] + " Manca Reparto ";
                        }
                        if (s != "")
                        {
                            DataRow x = tMsg.NewRow();
                            x["msg_des"] = s;
                            tMsg.Rows.Add(x);
                        }
                        if (s == "")
                        {
                            sRig = "";
                            if (strTip == "ART")
                            {
                                DataTable tTmp = null;
                                try
                                {
                                    tTmp = _clsQry.DivArtApShop((string)y["art_cod"], "");
                                    if (tTmp.Rows.Count == 0)
                                    {
                                        DataRow x = tMsg.NewRow();
                                        x["msg_des"] = (string)y["art_cod"] + " Non inserito ";
                                        tMsg.Rows.Add(x);
                                    }
                                    else
                                    {
                                        for (int i = 0; i < tTmp.Rows.Count; i++)
                                        {
                                            sRig = "";
                                            foreach (DataColumn c in tTmp.Columns)
                                                sRig += Convert.ToString(tTmp.Rows[i][c.ColumnName]) + ";";
                                            sw.Write(sRig + _clsDef.CRLF);
                                        }
                                    }
                                }
                                finally
                                {
                                    if (tTmp != null)
                                    {
                                        tTmp.Dispose();
                                        tTmp = null;
                                    }
                                }
                            }
                            else
                            {
                                foreach (DataColumn c in t.Columns)
                                {
                                    if (c.ColumnName.Substring(4, 3) != "idx")
                                        sRig += Convert.ToString(y[c.ColumnName]) + ";";
                                }
                                sw.Write(sRig + _clsDef.CRLF);
                            }
                        }

                        if (iCount % 1000 == 0)
                        {
                            GC.Collect(0, GCCollectionMode.Optimized);
                        }
                    }
                }
                finally
                {
                    dgv2.DataSource = tMsg;
                    if (t != null)
                    {
                        t.Dispose();
                        t = null;
                    }
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                for (int n = 0; n <= aryPos.Length - 1; n++)
                {
                    s = sFil + "_" + ((string)aryPos[n]).Substring(1, 1) + ".csv";
                    File.Copy(sFil, s, true);
                }
            }
            catch (OutOfMemoryException oomEx)
            {
                _clsFun.ErrorLog("GenVar_OutOfMemory", oomEx.Message);
                MessageBox.Show("Memoria insufficiente durante l'elaborazione dei dati: " + oomEx.Message,
                    "Errore Memoria", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("GenVar", ex.Message);
                MessageBox.Show("Errore durante l'elaborazione: " + ex.Message,
                    "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GenVarArt(string strPos, string strTip, DataRow rowNeg)
        {
            try
            {
                string p = "";
                string s = "";
                string sRig = "";

                p = "AnaArticoli";
                s = "SELECT * FROM AnaArticoli ";
                if (rdbAtt.Checked)
                    s += "WHERE art_sta='A' OR art_sta='R'";
                DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);

                string[] aryPos = strPos.Split(';');

                string sFil = Path.GetDirectoryName(_clsDef.TMPPOSVAR) + "\\";
                sFil = _clsQry.DivFilArtApShop(aryPos.Length, _clsDef.TMPPOSVAR);

                StreamWriter sw = new StreamWriter(sFil, false, Encoding.Default);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = Math.Max(1, t.Rows.Count);
                progressBar1.Minimum = 0;

                lblMsg.Text = "Negozio " + (string)rowNeg["tab_cod"];

                DataTable tMsg = (DataTable)dgv2.DataSource;
                dgv2.DataSource = null;

                sRig = "";

                DataTable tTmpSchema = new clsGenTabTmp().TabTmpDivArtApShop("TabDiv");
                foreach (DataColumn c in tTmpSchema.Columns)
                    sRig += c.ColumnName + ";";
                sw.Write(sRig + _clsDef.CRLF);
                tTmpSchema.Dispose();

                int iCount = 0;
                int totalCount = t.Rows.Count;

                try
                {
                    foreach (DataRow y in t.Rows)
                    {
                        iCount++;
                        if (iCount % 50 == 0 || iCount == totalCount)
                        {
                            progressBar1.Value = Math.Min(iCount, progressBar1.Maximum);
                            Application.DoEvents();
                        }

                        Boolean bb = false;

                        if (!DBNull.Value.Equals(rowNeg["tab_rep"]) && (string)rowNeg["tab_rep"] != "")
                        {
                            string sTipNeg = ((string)rowNeg["tab_rep"]).Substring(0, 1);
                            string sRep = ((string)rowNeg["tab_rep"]).Substring(1);

                            if (sRep.Contains((string)y["art_rep"]))
                            {
                                if (sTipNeg == "E")
                                    bb = false;
                                else
                                    bb = true;
                            }
                            else
                            {
                                if (sTipNeg == "E")
                                    bb = true;
                                else
                                    bb = false;
                            }
                        }
                        else
                            bb = true;

                        if (bb)
                        { 
                            s = "";
                            if (strTip == "ART")
                            {
                                if (((string)y["art_des"]).Trim() == "")
                                    s += (string)y["art_cod"] + " Manca descrizione ";
                                if (((string)y["art_rep"]).Trim() == "")
                                    s += (string)y["art_cod"] + " Manca Reparto ";
                                if (((string)y["art_des"]).Contains(";"))
                                    y["art_des"] = ((string)y["art_des"]).Replace(";", "");
                                if (((string)y["art_deb"]).Contains(";"))
                                    y["art_deb"] = ((string)y["art_deb"]).Replace(";", "");
                            }
                            if (s != "")
                            {
                                DataRow x = tMsg.NewRow();
                                x["msg_des"] = s;
                                tMsg.Rows.Add(x);
                            }
                            if (s == "")
                            {
                                sRig = "";
                                if (strTip == "ART")
                                {
                                    DataTable tTmp = null;
                                    try
                                    {
                                        tTmp = _clsQry.DivArtApShop((string)y["art_cod"], "");
                                        if (tTmp.Rows.Count == 0)
                                        {
                                            DataRow x = tMsg.NewRow();
                                            x["msg_des"] = (string)y["art_cod"] + " Non inserito ";
                                            tMsg.Rows.Add(x);
                                        }
                                        else
                                        {
                                            for (int i = 0; i < tTmp.Rows.Count; i++)
                                            {
                                                sRig = "";
                                                foreach (DataColumn c in tTmp.Columns)
                                                    sRig += Convert.ToString(tTmp.Rows[i][c.ColumnName]) + ";";
                                                sw.Write(sRig + _clsDef.CRLF);
                                            }
                                        }
                                    }
                                    finally
                                    {
                                        if (tTmp != null)
                                        {
                                            tTmp.Dispose();
                                            tTmp = null;
                                        }
                                    }
                                }
                                else
                                {
                                    foreach (DataColumn c in t.Columns)
                                    {
                                        if (c.ColumnName.Substring(4, 3) != "idx")
                                            sRig += Convert.ToString(y[c.ColumnName]) + ";";
                                    }
                                    sw.Write(sRig + _clsDef.CRLF);
                                }
                            }
                        }

                        if (iCount % 1000 == 0)
                        {
                            GC.Collect(0, GCCollectionMode.Optimized);
                        }
                    }
                }
                finally
                {
                    dgv2.DataSource = tMsg;
                    if (t != null)
                    {
                        t.Dispose();
                        t = null;
                    }
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                string[] aPos = ((string)rowNeg["tab_pos"]).Split(',');
                foreach(string ss in aPos)
                {
                    string sP = ss.PadLeft(2, Convert.ToChar('0'));

                    foreach(string sx in aryPos)
                    {
                        if (sP == sx)
                        {
                            s = sFil + "_" + ss + ".csv";
                            File.Copy(sFil, s, true);
                        }
                    }
                }
            }
            catch (OutOfMemoryException oomEx)
            {
                _clsFun.ErrorLog("GenVarArt_OutOfMemory", oomEx.Message);
                MessageBox.Show("Memoria insufficiente durante l'elaborazione degli articoli: " + oomEx.Message,
                    "Errore Memoria", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("GenVarArt", ex.Message);
                MessageBox.Show("Errore durante l'elaborazione degli articoli: " + ex.Message,
                    "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnRep_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Confermi la generazione degli articoli reparto?", "REPARTI CASSA", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                GenArtRep();
        }

        private void GenArtRep()
        {
            string sMsg = "";
            string s = _clsFun.ParGet(clsDefine.enuParametri.ParEcrDefault, _strConSql);

            if(s == "")
                MessageBox.Show("Parametro reparti di default vuoto!", "GENERAZIONE REPARTI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            { 
                Boolean bTerron = false;
                if (MessageBox.Show("Configurazione Terron? ( Se SI prefisso '999' altrimenti '99')", "BARCODE SUI REPARTI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    bTerron = true;

                string[] aEcr = s.Split(',');

                string p = "TabReparti";
                s = "SELECT * FROM " + p + " ORDER BY tab_cod";
                DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);

                foreach (DataRow y in t.Rows)
                {
                    p = "AnaArticoli";
                    s = "SELECT * FROM AnaArticoli WHERE art_sta='R' AND art_rep='" + y["tab_cod"] + "'";
                    DataTable tArt = _clsFun.FillTabSql(p, s, false, _strConSql);
                    if (tArt.Rows.Count > 0)
                        sMsg += "Articolo reparto già presente " + (string)y["tab_cod"] + " " + (string)y["tab_des"] + _clsDef.CRLF;
                    else
                    {
                        string sArt = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaArticoli, 7, _strConSql);

                        DataRow x = tArt.NewRow();
                        x["art_cod"] = sArt;
                        x["art_des"] = "REPARTO " + (string)y["tab_des"];
                        x["art_deb"] = "";
                        x["art_sta"] = "R";
                        x["art_iva"] = "022";
                        x["art_eqp"] = "";
                        x["art_umi"] = "NR";
                        x["art_tgr"] = "KG";
                        x["art_pxc"] = 1;
                        x["art_pne"] = 1;
                        x["art_rep"] = (string)y["tab_cod"];
                        x["art_ec1"] = "";
                        x["art_ec2"] = "";
                        x["art_ec3"] = "";

                        if (aEcr.Length > 2)
                        {
                            x["art_ec1"] = aEcr[0];
                            x["art_ec2"] = aEcr[1];
                            x["art_ec3"] = aEcr[2];
                        }

                        x["art_ori"] = "";
                        x["art_cal"] = "";
                        x["art_reb"] = "";
                        x["art_cat"] = "";
                        x["art_tra"] = "";
                        x["art_gsc"] = "";
                        x["art_bil"] = false;
                        x["art_plu"] = "";
                        x["art_bpz"] = false;
                        x["art_dtm"] = DateTime.Now.ToString("dd/MM/yyyy");
                        x["art_dti"] = DateTime.Today;

                        s = _clsFun.SqlInsertRow(p, tArt, x);

                        if (s != "")
                        {
                            _clsFun.SqlWrite(s, _strConSql);
                            new clsVariazioni().Variazioni((string)x["art_cod"], s, "Anagrafica articolo", _clsDef.VARPOS);
                        }

                        p = "AnaBarcode";
                        s = "SELECT * FROM " + p + " WHERE ean_ean='" + "99" + (string)y["tab_cod"] + "'";
                        DataTable tEan = _clsFun.FillTabSql(p, s, true, _strConSql);
                        if (tEan.Rows.Count > 0)
                            sMsg += "Barcode " + "99" + (string)y["tab_cod"] + " già presente sul codice articolo " + (string)tEan.Rows[0]["ean_art"] + _clsDef.CRLF;
                        else
                        {
                            x = tEan.NewRow();
                            x["ean_art"] = sArt;
                            if(bTerron)
                                x["ean_ean"] = "999" + ((string)y["tab_cod"]).Substring(1);
                            else
                                x["ean_ean"] = "99" + (string)y["tab_cod"];
                            x["ean_ann"] = false;
                            x["ean_qta"] = 0;
                            x["ean_prv"] = 0;
                            x["ean_dti"] = DateTime.Today;
                            x["ean_dtm"] = DateTime.Today;
                            x["ean_bil"] = false;
                            x["ean_ecp"] = false;

                            s = _clsFun.SqlInsertRow(p, tEan, x);

                            if (s != "")
                            {
                                _clsFun.SqlWrite(s, _strConSql);
                                new clsVariazioni().Variazioni(sArt, s, "Anagrafica barcode", _clsDef.VARPOS);
                            }
                        }

                    }
                }
            }
            if (sMsg != "")
                MessageBox.Show(sMsg, "ERRRORI");
            else
                MessageBox.Show("Fine inserimento");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Console.Beep(5000, 300);
            System.Media.SystemSounds.Beep.Play();
            System.Media.SoundPlayer player = new System.Media.SoundPlayer(@"C:\\ApProject\\Temp\\Suoni\\bip.wav");
        }

        private void GenVarCauMovDoc(string strPos)
        {
            string p = "";
            string s = "";

            string sFil = Path.GetDirectoryName(_clsDef.TMPPOSVAR) + "\\TabDoc";

            StreamWriter sw = new StreamWriter(sFil, false);
            string sRig = "tab_tdo;tab_tip;tab_cod;tab_des;tab_cfo;tab_prn;tab_suf;tab_ann;";
            sw.Write(sRig + _clsDef.CRLF);

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = 2;
            progressBar1.Minimum = 0;

            for (int i = 0; i < 2; i++)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                string sSuf = "";
                p = "TabDocTipo";
                s = "SELECT * FROM TabDocumenti WHERE tab_cod='F'";
                DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
                if (t.Rows.Count > 0 && !DBNull.Value.Equals(t.Rows[0]["tab_suf"]))
                    sSuf = (string)t.Rows[0]["tab_suf"];

                string sTip = "";
                if (i == 0)
                {
                    sTip = "DOC";
                    p = "TabDocTipo";
                    s = "SELECT * FROM TabDocTipo WHERE tab_pos<>''";
                }
                else if (i == 1)
                {
                    sTip = "MOV";
                    p = "TabMovCausali";
                    s = "SELECT * FROM TabMovCausali WHERE tab_pos<>''";
                }

                t = _clsFun.FillTabSql(p, s, false, _strConSql);

                foreach (DataRow y in t.Rows)
                {
                    sRig = sTip + ";";
                    sRig += (string)y["tab_tip"] + ";";
                    sRig += (string)y["tab_cod"] + ";";
                    sRig += (string)y["tab_des"] + ";";
                    sRig += (string)y["tab_cfo"] + ";";
                    sRig += (string)y["tab_pos"] + ";";
                    if(sTip == "DOC")
                        sRig += sSuf + ";";
                    else
                        sRig += "" + ";";

                    sRig += ((Boolean)y["tab_ann"]).ToString() + ";";
                    //if((Boolean)y["tab_ann"])
                    //    sRig += "S;";
                    //else
                    //    sRig += "N;";
                    sw.Write(sRig + _clsDef.CRLF);
                }
            }


            ((TextWriter)sw).Flush();
            sw.Close();
            sw.Dispose();

            string[] aryPos = strPos.Split(';');

            for (int n = 0; n <= aryPos.Length - 1; n++)
            {
                s = sFil + "_" + ((string)aryPos[n]).Substring(1, 1) + ".csv";
                if (File.Exists(s))
                    File.Delete(s);
                File.Copy(sFil, s);
            }
        }

        private void chkTabMdo_CheckedChanged(object sender, EventArgs e)
        {

        }

    }
}
