using System;
using System.Collections;
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
    public partial class frmUtyDivTerm : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";

        public frmUtyDivTerm()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0); ;
            _strConSql = _clsFun.ConSql("");
        }

        private void frmUtyDivTerm_Load(object sender, EventArgs e)
        {
            this.Text = "Divulgazione a Ftp/ApPhone";
            SetDgv1();
            FillPos();
        }
        private void frmUtyDivTerm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
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

        private void btnOk_Click(object sender, EventArgs e)
        {
            Divulga();
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
            cTbc.DataPropertyName = "ter_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ter_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "ter_inv";
            cCbc.Name = "Invia";
            cCbc.Width = 45;
            dgv1.Columns.Add(cCbc);
        }

        private void FillPos()
        {
            string s = "";

            s = "SELECT * FROM TabTerm WHERE tab_cod='" + _clsDef.TERMTDIV + "'";
            DataTable tTer = _clsFun.FillTabSql("TabTerm", s, true, _strConSql);

            s = (string)tTer.Rows[0]["tab_fpt"];
            string[] aTer = s.Split(';');

            if(aTer.Length <= 1)
                MessageBox.Show("Codice terminali non definito!", "CONTROLLO TERMINALI", MessageBoxButtons.OK, MessageBoxIcon.Question);
            else
	        {
                s = aTer[1];
                aTer = s.Split(',');

                if(aTer.Length > 0)
                {
                    if (_clsFun.Numerico(s))
                    {
                        //int n = Convert.ToUInt16(s);

                        DataTable t = new DataTable("TabPos");

                        t.Columns.Add(new DataColumn()
                        {
                            DataType = Type.GetType("System.Boolean"),
                            ColumnName = "ter_inv",
                            Caption = "Invio",
                            ReadOnly = false,
                            DefaultValue = (Boolean)true
                        });
                        t.Columns.Add(new DataColumn()
                        {
                            DataType = Type.GetType("System.String"),
                            ColumnName = "ter_cod",
                            Caption = "Term",
                            MaxLength = 3,
                            ReadOnly = false,
                            DefaultValue = (String)""
                        });
                        t.Columns.Add(new DataColumn()
                        {
                            DataType = Type.GetType("System.String"),
                            ColumnName = "ter_des",
                            Caption = "Descrizione",
                            MaxLength = 50,
                            ReadOnly = false,
                            DefaultValue = (String)""
                        });

                        foreach (string ss in aTer)
                        {
                            DataRow x = t.NewRow();
                            x["ter_inv"] = true;
                            x["ter_cod"] = ss;
                            x["ter_des"] = "TERM. N. " + ss;
                            t.Rows.Add(x);
                        }

                        dgv1.DataSource = t;
                    }
                    else
                        MessageBox.Show("Numero terminali non definito in tabella TERM!");
                }
	        }
        }

        private void Divulga()
        {
            if (chkTabUsr.Checked)
                GenVar("usr");
            if (chkAnaArt.Checked)
                DivTermApPhone();
        }

        private void GenVar(string strTip)
        {
            string p = "";
            string s = "";
            string sRig = "";
            string sSep = "|";

            s = "SELECT * FROM TabTerm WHERE tab_cod='" + _clsDef.TERMTDIV + "'";
            DataTable tTer = _clsFun.FillTabSql("TabTerm", s, true, _strConSql);

            if (tTer.Rows.Count == 0)
                MessageBox.Show("Terminale non configurato!", "TABELLA TERMINALI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                string sPath = Path.GetDirectoryName((string)tTer.Rows[0]["tab_div"]) + "\\";

                //s = ((string)tTer.Rows[0]["tab_bat"]).Trim();
                //string[] aTer = s.Split(',');

                //s = (string)tTer.Rows[0]["tab_fpt"];
                //string[] aTer = s.Split(';');
                //s = aTer[1];
                //aTer = s.Split(',');

                if (strTip == "usr")
                {
                    p = "TabUtenti";
                    s = "SELECT * FROM " + p;
                }

                DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);

                string sFil = sPath + "Old\\";
                if (strTip == "ART")
                    Console.WriteLine("");
                else if (p.Substring(0, 3) == "Tab")
                    sFil += strTip + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
                else
                    sFil += p.Substring(0, 6);

                StreamWriter sw = new StreamWriter(sFil, false);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = t.Rows.Count;
                progressBar1.Minimum = 0;

                DataTable tMsg = (DataTable)dgv2.DataSource;

                sRig = "";

                //if (strTip == "ART")
                //{
                //    DataTable tTmp = new clsGenTabTmp().TabTmpDivArtApShop("TabDiv");
                //    foreach (DataColumn c in tTmp.Columns)
                //        sRig += c.ColumnName + "|";
                //    sw.Write(sRig + _clsDef.CRLF);
                //}
                //else
                //{

                sRig = "T" + sSep;
                foreach (DataColumn c in t.Columns)
                {
                    if (c.ColumnName.Substring(4, 3) != "idx")
                        sRig += c.ColumnName + sSep;
                }
                sw.Write(sRig + _clsDef.CRLF);
                //}

                foreach (DataRow y in t.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

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
                        sRig = "R" + sSep;
                        if (strTip == "ART")
                        {
                            //DataTable tTmp = _clsQry.DivArtApShop((string)y["art_cod"]);
                            //if (tTmp.Rows.Count == 0)
                            //{
                            //    DataRow x = tMsg.NewRow();
                            //    x["msg_des"] = (string)y["art_cod"] + " Non inserito ";
                            //    tMsg.Rows.Add(x);
                            //}
                            //else
                            //{
                            //    for (int i = 0; i < tTmp.Rows.Count; i++)
                            //    {
                            //        sRig = "";
                            //        foreach (DataColumn c in tTmp.Columns)
                            //            sRig += Convert.ToString(tTmp.Rows[i][c.ColumnName]) + ";";
                            //        sw.Write(sRig + _clsDef.CRLF);
                            //    }
                            //}
                        }
                        else
                        {
                            foreach (DataColumn c in t.Columns)
                            {
                                if (c.ColumnName.Substring(4, 3) != "idx")
                                    sRig += Convert.ToString(y[c.ColumnName]) + sSep;
                            }
                            sw.Write(sRig + _clsDef.CRLF);
                        }
                    }
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

                //foreach (string sTer in aTer)
                foreach(DataRow y in ((DataTable)dgv1.DataSource).Rows)
                {
                    if ((Boolean)y["ter_inv"])
                    {
                        s = Path.GetFileName(sFil);
                        s = s.Substring(0, 4) + y["ter_cod"] + "_" + s.Substring(4);
                        string sF = sPath + s;
                        if (File.Exists(sF))
                            File.Delete(sF);
                        File.Copy(sFil, sF);
                    }
                }
            }
        }

        private void DivTermApPhone()
        {
            clsTermCsv2 clsTerm = new clsTermCsv2();

            string s = "";
            string sOld = "Old\\";
            string sRig = "";

            s = "SELECT * FROM TabTerm WHERE tab_cod='" + _clsDef.TERMTDIV + "'";
            DataTable tTer = _clsFun.FillTabSql("TabTerm", s, true, _strConSql);

            if(tTer.Rows.Count == 0)
                MessageBox.Show("Terminale non configurato!", "TABELLA TERMINALI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
	        {
                string sPath = Path.GetDirectoryName((string)tTer.Rows[0]["tab_div"]) + "\\";

                s = "SELECT ";
                s += "B.liv_art AS pos_art, ";
                s += "AnaArticoli.art_des AS pos_ard, ";
                s += "AnaArticoli.art_umi AS pos_umi, ";
                s += "AnaArticoli.art_tgr AS pos_tgr, ";
                s += "TabReparti.tab_des AS RepDes, ";
                s += "AnaArticoli.art_pxc AS pos_pxc, ";
                s += "AnaArticoli.art_pne AS pos_pne, ";
                s += "AnaArticoli.art_iva AS pos_iva, ";
                s += "B.liv_prv AS pos_prv, ";
                s += "B.liv_lis, ";
                s += "B.liv_dti ";
                s += "FROM (";
                s += "SELECT A.liv_lis, GesLisVendita_1.liv_prv, A.liv_art, A.liv_dti ";
                s += "FROM (SELECT liv_lis, liv_art, MAX(liv_dti) AS liv_dti ";
                s += "FROM GesLisVendita ";
                s += "GROUP BY liv_lis, liv_art) AS A LEFT OUTER JOIN ";
                s += "GesLisVendita AS GesLisVendita_1 ON GesLisVendita_1.liv_lis = A.liv_lis AND GesLisVendita_1.liv_art = A.liv_art AND GesLisVendita_1.liv_dti = A.liv_dti) AS B LEFT OUTER JOIN ";
                s += "AnaArticoli ON AnaArticoli.art_cod = B.liv_art LEFT OUTER JOIN ";
                s += "TabReparti ON TabReparti.tab_cod = AnaArticoli.art_rep";
                DataTable tArt = _clsFun.FillTabSql("AnaArt", s, false, _strConSql);
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tArt.Columns["liv_art"];
                keys[1] = tArt.Columns["liv_lis"];
                tArt.PrimaryKey = keys;

                string sFilArt = sPath + sOld + "art_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";

                s = sPath + "Temp\\";

                StreamWriter swArt = new StreamWriter(sFilArt, true);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = tArt.Rows.Count;
                progressBar1.Minimum = 0;

                sRig = clsTerm.TermArtTes();
                swArt.Write(sRig + _clsDef.CRLF);

                int i = 0;
                ArrayList aFil = new ArrayList();
                aFil.Add(sFilArt);

                foreach (DataRow y in tArt.Rows)
                {
                    i++;

                    progressBar1.Increment(1);
                    Application.DoEvents();

                    sRig = clsTerm.TermArtRow(y, null);
                    swArt.Write(sRig + _clsDef.CRLF);
                }

                ((TextWriter)swArt).Flush();
                swArt.Close();
                swArt.Dispose();

                /** Barcode **/

                string sFilEan = sPath + sOld + "ean_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
                StreamWriter swEan = new StreamWriter(sFilEan, true);
                aFil.Add(sFilEan);

                s = "SELECT ";
                s += "ean_art AS pos_art, ";
                s += "ean_ean AS pos_ean ";
                s += "FROM AnaBarcode ";
                DataTable tEan = _clsFun.FillTabSql("AnaEan", s, false, _strConSql);
                keys = new DataColumn[2];
                keys[0] = tEan.Columns["ean_art"];
                keys[1] = tEan.Columns["ean_ean"];
                tEan.PrimaryKey = keys;

                sRig = clsTerm.TermEanTes();
                swEan.Write(sRig + _clsDef.CRLF);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = tEan.Rows.Count;
                progressBar1.Minimum = 0;

                foreach (DataRow y in tEan.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    sRig = clsTerm.TermEanRow(y);
                    swEan.Write(sRig + _clsDef.CRLF);
                }

                ((TextWriter)swEan).Flush();
                swEan.Close();
                swEan.Dispose();

                /** Costi **/

                string sFilLia = sPath + sOld + "lia_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
                StreamWriter swLia = new StreamWriter(sFilLia, true);
                aFil.Add(sFilLia);

                s = "SELECT ";
                s += "GesLisAcquisto_1.lia_tip, ";
                s += "A.lia_arf, ";
                s += "A.lia_art, ";
                s += "A.lia_for, ";
                s += "AnaFornitori.for_des AS LiaFod, ";
                s += "GesLisAcquisto_1.lia_pxc, ";
                s += "GesLisAcquisto_1.lia_cos, ";
                s += "A.lia_dti ";
                s += "FROM (SELECT lia_for, lia_art, lia_arf, MAX(lia_dti) AS lia_dti ";
                s += "FROM GesLisAcquisto ";
                s += "GROUP BY lia_for, lia_art, lia_arf) AS A LEFT OUTER JOIN ";
                s += "GesLisAcquisto AS GesLisAcquisto_1 ON GesLisAcquisto_1.lia_for = A.lia_for AND GesLisAcquisto_1.lia_art = A.lia_art  AND GesLisAcquisto_1.lia_arf = A.lia_arf AND GesLisAcquisto_1.lia_dti = A.lia_dti LEFT OUTER JOIN ";
                s += "AnaFornitori ON A.lia_for = AnaFornitori.for_cod";
                DataTable tLia = _clsFun.FillTabSql("AnaLia", s, false, _strConSql);

                sRig = clsTerm.TermLiaTes();
                swLia.Write(sRig + _clsDef.CRLF);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = tLia.Rows.Count;
                progressBar1.Minimum = 0;

                foreach (DataRow y in tLia.Rows)
                {            
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    sRig = clsTerm.TermLiaRow(null, y);
                    swLia.Write(sRig + _clsDef.CRLF);
                }

                ((TextWriter)swLia).Flush();
                swLia.Close();
                swLia.Dispose();

                DataTable tAph = (DataTable)dgv1.DataSource;

                if (i < 1000)
                {
                    foreach (string sFi in aFil)
                    {
                        foreach (DataRow y in tAph.Rows)
                        {
                            if ((Boolean)y["ter_inv"])
                            {
                                s = Path.GetFileName(sFi);
                                s = s.Substring(0, 4) + y["ter_cod"] + "_" + s.Substring(4);
                                string sFil = sPath + s;
                                if (File.Exists(sFil))
                                    File.Delete(sFil);
                                File.Copy(sFi, sFil);
                            }
                        }
                    }
                }
                else
                {
                    int iMax = 1000;

                    int ii = 0;

                    foreach (string sFi in aFil)
                    {
                        int iNre = 0;
                        i = 0;
                        while(iNre > -1)
                        {
                            i++;
                            ii++;

                            iNre = WriFile(sPath, sFi, iNre, iMax, i, tAph);
                        }
                    }

                    MessageBox.Show("File creati " + ii.ToString(), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private int WriFile(string strPath, string strFil, int intNre, int intMax, int intNum, DataTable tabAph)
        {
            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                string s = "";
                string sRig = "";

                s = Path.GetDirectoryName(strFil);

                string sFil = Path.GetFileNameWithoutExtension(strFil);

                Console.WriteLine("aaaaaaa");

                sFil = s + "\\" + sFil.Substring(0,4) + "XXX" + "_" + sFil.Substring(4) + intNum.ToString("000") + ".csv";

                StreamWriter sw = new StreamWriter(sFil, true);

                int i = 0;
                int iNre = intNre;

                while ((sRig = sr.ReadLine()) != null)
                {
                    i++;

                    if (i > iNre)
                    {
                        sw.Write(sRig + _clsDef.CRLF);

                        intNre++;

                        if (i >= intMax + iNre)
                            break;
                    }

                }

                sw.Flush();
                sw.Close();
                sw.Dispose();

                if (intNre < intMax + iNre)
                    intNre = -1;


                foreach (DataRow y in tabAph.Rows)
                {
                    if ((Boolean)y["ter_inv"])
                    {
                        s = Path.GetFileName(sFil);
                        s = s.Substring(0, 4) + y["ter_cod"] + "_" + s.Substring(8);
                        string sFi = strPath + s;
                        if (File.Exists(sFi))
                            File.Delete(sFi);
                        File.Copy(sFil, sFi);
                    }
                }

            }

            return intNre;
        }

        private void btnFtp_Click(object sender, EventArgs e)
        {
            string s = "SELECT * FROM TabTerm WHERE tab_cod='" + _clsDef.TERMTDIV + "'";
            DataTable t = _clsFun.FillTabSql("TabTerm", s, true, _strConSql);
            if (t.Rows.Count > 0 && ((string)t.Rows[0]["tab_hst"]).Trim() != "")
            {
                string sFtp = "";
                sFtp += (string)t.Rows[0]["tab_hst"] + "|";
                sFtp += (string)t.Rows[0]["tab_usr"] + "|";
                sFtp += (string)t.Rows[0]["tab_pwd"] + "|";
                //sFtp += (string)t.Rows[0]["tab_fpt"];

                s = (string)t.Rows[0]["tab_fpt"];

                string[] a = s.Split(';');

                if (a.Length > 1)
                {
                    sFtp += a[0] + "|"; ;
                    sFtp += a[1] + "|"; ;

                    Console.WriteLine("aaaa");

                    InvioFtp(sFtp, lblPath.Text);

                }
                else
                    MessageBox.Show("Percorso destinazione non definito in TabTerm");

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
