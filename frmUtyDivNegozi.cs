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
    public partial class frmUtyDivNegozi : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();
        clsVariazioni _clsVar = new clsVariazioni();

        private string _strConSql = "";

        private string _strPar016PathDivNegozi = "";

        public frmUtyDivNegozi()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmUtyDivNegozi_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");

            _strPar016PathDivNegozi = _clsFun.ParGet(clsDefine.enuParametri.Par016PathDivNegozi, _strConSql);

            SetDgv1();
            SetDgv2();
            dgv2.DataSource = new clsGenTabTmp().TabTmpMsg("TabMsg");
            FillNeg();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmUtyDivNegozi_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
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
            cTbc.DataPropertyName = "tab_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "TabInv";
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

        private void FillNeg()
        {
            string s = "SELECT * FROM TabNegozi WHERE tab_ann=0";
            //string s = "SELECT * FROM TabNegozi WHERE tab_ann=0 AND tab_tip='L'";
            DataTable tNeg = _clsFun.FillTabSql("TabNegozi", s, false, _strConSql);

            tNeg.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "TabInv",
                Caption = "Invio",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });

            if (tNeg.Rows.Count == 0)
                MessageBox.Show("Tabella negozi vuota!");
            else
                dgv1.DataSource = tNeg;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (!File.Exists(_clsDef.TMPPOSSTOP))
            {
                var f = File.Create(_clsDef.TMPPOSSTOP);
                f.Close();
            }

            DataTable t = (DataTable)dgv1.DataSource;

            string sNeg = "";
            foreach (DataRow y in t.Rows)
            {
                if ((Boolean)y["TabInv"])
                {
                    Console.WriteLine((string)y["tab_des"]);
                    sNeg += (string)y["tab_cod"] + ";";
                }
            }
            if (sNeg == "")
                MessageBox.Show("Nessun negozio selezionato!", "CONTROLLO NEGOZI");
            else
            {
                sNeg = sNeg.Substring(0, sNeg.Length - 1);
                if (chkAnaArt.Checked)
                    GenImpArt(sNeg);
                if (chkArtVar.Checked)
                    GenArtVar(sNeg);

                //if (chkTabUmi.Checked)
                //    GenVar(sPos, "UMI");
                //if (chkTabIva.Checked)
                //    GenVar(sPos, "IVA");
                if (chkTabRep.Checked)
                    GenVar(sNeg, "REP");
                //if (chkTabPpo.Checked)
                //    GenVar(sPos, "PPO");
                //if (chkTabUsr.Checked)
                //    GenVar(sPos, "USR");
                //if (chkTabGrt.Checked)
                //    GenVar(sPos, "GRT");
                //if (chkTabCam.Checked)
                //    GenVar(sPos, "CAM");
                if (chkTabEcr.Checked)
                    GenVar(sNeg, "ECR");
            }
            if (File.Exists(_clsDef.TMPPOSSTOP))
                File.Delete(_clsDef.TMPPOSSTOP);
        }

        private void GenImpArt(string sNeg)
        {

            string s = "SELECT art_cod, art_sta FROM AnaArticoli ORDER BY art_cod";
            DataTable t = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);

            ArrayList a = new ArrayList();

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                if (!rdbAtt.Checked || ((string)y["art_sta"] == "A" || (string)y["art_sta"] == "R"))
                    a.Add((string)y["art_cod"]);
            }

            _clsVar.DivNegArticoli(_strPar016PathDivNegozi, a, progressBar1, sNeg, "");

            MessageBox.Show("Fine estrazione!");
        }

        private void GenArtVar(string sNeg)
        {
            string s = "SELECT var_art AS art_cod FROM GesVariazioni WHERE var_tip='POS' AND var_dti=" + _clsFun.DaySql(dtpArtVar.Value) + " ORDER BY art_cod";
            DataTable t = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);

            ArrayList a = new ArrayList();

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                a.Add((string)y["art_cod"]);
            }

            _clsVar.DivNegArticoli(_strPar016PathDivNegozi, a, progressBar1, sNeg, "");

            MessageBox.Show("Fine estrazione!");
        }

        private void GenVar(string strPos, string strTip)
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
            if (strTip == "ART")
            {
                p = "AnaArticoli";
                s = "SELECT * FROM AnaArticoli ";
                if (rdbAtt.Checked)
                    s += "WHERE art_sta='A' OR art_sta='R'";
                //s += "WHERE art_cod='0010984'";
            }

            if (p != "")
            {

                DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);

                s = _clsFun.ParGet(clsDefine.enuParametri.Par016PathDivNegozi, _strConSql);
                if (t != null && t.Rows.Count > 0 && s != "" && p != "")
                    new clsVariazioni().DivNegTabelle(s, t);
            }
            else if(strTip == "ECR")
            {
                string sPath = _clsFun.ParGet(clsDefine.enuParametri.Par016PathDivNegozi, _strConSql);

                p = "TabEcrLv1";
                s = "SELECT * FROM TabEcrLv1 ORDER BY tab_cod";
                DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
                if (t != null && t.Rows.Count > 0 && sPath != "" && p != "")
                    new clsVariazioni().DivNegTabelle(sPath, t);

                p = "TabEcrLv2";
                s = "SELECT * FROM TabEcrLv2 ORDER BY tab_cod";
                t = _clsFun.FillTabSql(p, s, false, _strConSql);
                if (t != null && t.Rows.Count > 0 && sPath != "" && p != "")
                    new clsVariazioni().DivNegTabelle(sPath, t);

                p = "TabEcrLv3";
                s = "SELECT * FROM TabEcrLv3 ORDER BY tab_cod";
                t = _clsFun.FillTabSql(p, s, false, _strConSql);
                if (t != null && t.Rows.Count > 0 && sPath != "" && p != "")
                    new clsVariazioni().DivNegTabelle(sPath, t);
            }

            //string[] aryPos = strPos.Split(';');

            //string sFil = Path.GetDirectoryName(_clsDef.TMPPOSVAR) + "\\";
            //if (strTip == "ART")
            //    sFil = _clsQry.DivFilArtApShop(aryPos.Length, _clsDef.TMPPOSVAR);
            //else if (p.Substring(0, 3) == "Tab")
            //    sFil += "Tab" + strTip;
            //else
            //    sFil += p.Substring(0, 6);

            //StreamWriter sw = new StreamWriter(sFil, false);

            //progressBar1.Value = 0;
            //progressBar1.Visible = true;
            //progressBar1.Maximum = t.Rows.Count;
            //progressBar1.Minimum = 0;

            //DataTable tMsg = (DataTable)dgv2.DataSource;

            //sRig = "";

            //if (strTip == "ART")
            //{
            //    DataTable tTmp = new clsGenTabTmp().TabTmpDivArtApShop("TabDiv");
            //    foreach (DataColumn c in tTmp.Columns)
            //        sRig += c.ColumnName + ";";
            //    sw.Write(sRig + _clsDef.CRLF);
            //}
            //else
            //{
            //    foreach (DataColumn c in t.Columns)
            //    {
            //        if (c.ColumnName.Substring(4, 3) != "idx")
            //            sRig += c.ColumnName + ";";
            //    }
            //    sw.Write(sRig + _clsDef.CRLF);
            //}

            //foreach (DataRow y in t.Rows)
            //{
            //    progressBar1.Increment(1);
            //    Application.DoEvents();

            //    s = "";
            //    if (strTip == "ART")
            //    {
            //        if (((string)y["art_des"]).Trim() == "")
            //            s += (string)y["art_cod"] + " Manca descrizione ";
            //        if (((string)y["art_rep"]).Trim() == "")
            //            s += (string)y["art_cod"] + " Manca Reparto ";
            //    }
            //    if (s != "")
            //    {
            //        DataRow x = tMsg.NewRow();
            //        x["msg_des"] = s;
            //        tMsg.Rows.Add(x);
            //    }
            //    if (s == "")
            //    {
            //        sRig = "";
            //        if (strTip == "ART")
            //        {
            //            DataTable tTmp = _clsQry.DivArtApShop((string)y["art_cod"], "");
            //            if (tTmp.Rows.Count == 0)
            //            {
            //                DataRow x = tMsg.NewRow();
            //                x["msg_des"] = (string)y["art_cod"] + " Non inserito ";
            //                tMsg.Rows.Add(x);
            //            }
            //            else
            //            {
            //                for (int i = 0; i < tTmp.Rows.Count; i++)
            //                {
            //                    sRig = "";
            //                    foreach (DataColumn c in tTmp.Columns)
            //                        sRig += Convert.ToString(tTmp.Rows[i][c.ColumnName]) + ";";
            //                    sw.Write(sRig + _clsDef.CRLF);
            //                }
            //            }
            //        }
            //        else
            //        {
            //            foreach (DataColumn c in t.Columns)
            //            {
            //                if (c.ColumnName.Substring(4, 3) != "idx")
            //                    sRig += Convert.ToString(y[c.ColumnName]) + ";";
            //            }
            //            sw.Write(sRig + _clsDef.CRLF);
            //        }
            //    }
            //}

            //((TextWriter)sw).Flush();
            //sw.Close();
            //sw.Dispose();

            //for (int n = 0; n <= aryPos.Length - 1; n++)
            //{
            //    s = sFil + "_" + ((string)aryPos[n]).Substring(1, 1) + ".csv";
            //    if (File.Exists(s))
            //        File.Delete(s);
            //    File.Copy(sFil, s);
            //}
        }



    }
}
