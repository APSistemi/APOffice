using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;

namespace APOffice
{
    public partial class frmGesStatRotazione : Form
    {
        private const string TABSTAVET = "GesNegVet";
        private const string TABSTAVEP = "GesNegVep";
        private const string TABSTAVEN = "GesNegVen";

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strConSqlSta = "";

        CultureInfo ci = System.Globalization.CultureInfo.CreateSpecificCulture("it");
        DateTimeFormatInfo _ciFormat = new DateTimeFormatInfo();
        Calendar _ciCalendar = CultureInfo.InvariantCulture.Calendar;

        public frmGesStatRotazione()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);

            _ciFormat = ci.DateTimeFormat;
            _ciCalendar = ci.Calendar;
        }

        private void frmGesStatRotazione_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");

            dtpIni.Value = new DateTime(DateTime.Today.Year, 1, 1);
            SetDgv1();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmGesStatRotazione_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private void btnEstrai_Click(object sender, EventArgs e)
        {
            FillDati();
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
            cTbc.DataPropertyName = "tmp_ecd";
            cTbc.Name = "ECR";
            cTbc.Width = 180;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_red";
            cTbc.Name = "Reparto";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_ard";
            cTbc.Name = "Descrizione";
            cTbc.Width = 160;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_umi";
            cTbc.Name = "UM";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_iro";
            cTbc.Name = "Indice di rotazione";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_scm";
            cTbc.Name = "Scorta media";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_cve";
            cTbc.Name = "Q.tà/Costo venduto";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

        }

        private void FillDati()
        {
            int iIni = 0;
            int iFin = 0;

            if (rdbSelMon.Checked)
            {
                iIni = dtpIni.Value.Month;
                iFin = dtpFin.Value.Month;
            }
            else
            {
                //CultureInfo ci = new ("it-IT");
                iIni = _ciCalendar.GetWeekOfYear(dtpIni.Value, _ciFormat.CalendarWeekRule, _ciFormat.FirstDayOfWeek);
                iFin = _ciCalendar.GetWeekOfYear(dtpFin.Value, _ciFormat.CalendarWeekRule, _ciFormat.FirstDayOfWeek);
            }

            DataTable t = new clsGenTabTmp().TabTmpArtRotazione("TabRot", iIni, iFin);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = t.Columns["tmp_art"];
            t.PrimaryKey = keys;

            FillInventario(t);
            FillMovimenti(t);
            FillVendite(t);
            FillAnagrafica(t);

            dgv1.DataSource = t;
        }

        private void FillAnagrafica(DataTable tabTmp)
        {
            string s = "";
            DataRow[] j;

            s = "SELECT ";
            s += "AnaArticoli.art_cod, ";
            s += "AnaArticoli.art_des, "; 
            s += "AnaArticoli.art_umi, ";
            s += "AnaArticoli.art_rep, ";
            s += "AnaArticoli.art_ec1, ";
            s += "AnaArticoli.art_ec2, ";
            s += "AnaArticoli.art_ec3, ";
            s += "TabReparti.tab_des AS RepDes, ";
            s += "TabEcrLv3.tab_des AS Lv3Des ";
            s += "FROM AnaArticoli ";
            s += "INNER JOIN TabReparti ON AnaArticoli.art_rep = TabReparti.tab_cod ";
            s += "INNER JOIN TabEcrLv3 ON AnaArticoli.art_ec1 = TabEcrLv3.tab_lv1 AND AnaArticoli.art_ec2 = TabEcrLv3.tab_lv2 AND AnaArticoli.art_ec3 = TabEcrLv3.tab_cod ";
            DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            progressBar1.Value = 0;
            progressBar1.Maximum = tabTmp.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in tabTmp.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                j = tArt.Select("art_cod='" + y["tmp_art"] + "'");
                if(j.Length > 0)
                {
                    y["tmp_ecr"] = (string)j[0]["art_ec2"] + (string)j[0]["art_ec3"];
                    y["tmp_ecd"] = (string)j[0]["Lv3Des"];
                    y["tmp_rep"] = (string)j[0]["art_rep"];
                    y["tmp_red"] = (string)j[0]["RepDes"];
                    y["tmp_ard"] = (string)j[0]["art_des"];
                    y["tmp_umi"] = (string)j[0]["art_umi"];

                    //Scorta media

                    decimal dGia = (decimal)y["inv_000"];       //Giacenza per calcolare la media in base ai periodi
                    decimal dCve = 0;                           //Costo del venduto dato dalle vendite

                    int iPer = 0;
                    if (dGia > 0)
                        iPer = 1;

                    if ((string)y["tmp_art"] == "0006057")
                        Console.WriteLine("zzzz");

                    foreach(DataColumn c in tabTmp.Columns)
                    {
                        if (c.ColumnName.Substring(0, 4) == "acq_")
                        {
                            iPer++;
                            dGia += (decimal)y[c.ColumnName];
                            //dCve += (decimal)y[c.ColumnName];
                        }
                        else if (c.ColumnName.Substring(0, 4) == "ven_")
                        {
                            dCve += (decimal)y[c.ColumnName];
                            dGia -= (decimal)y[c.ColumnName];
                        }
                    }

                    //y["tmp_scm"] = dGia / iPer;

                    if (iPer == 0)
                        iPer = 1;

                    y["tmp_scm"] = Math.Round(dGia / iPer, 2);
                    y["tmp_cve"] = dCve;
                    //y["tmp_iro"] = Math.Round(dCve / dGia, 2);
                    if(dCve != 0 && dGia != 0)
                        y["tmp_iro"] = Math.Round(dCve / dGia, 2);
                    //break;
                }
            }
        }

        private void FillInventario(DataTable tabTmp)
        {
            string s = "";
            DataRow x;
            DataRow[] j;

            DateTime dDti = dtpIni.Value;
            DateTime dDtf = dtpIni.Value.AddDays(5);            //20180627 simuliamo se esiste inventario nella prima settimana

            s = "SELECT inv_art, inv_qta, int_day FROM GesInventario LEFT JOIN GesInvTestate ON GesInventario.inv_num = GesInvTestate.int_num ";
            s += "WHERE int_day >= " + _clsFun.DaySql(dDti) + " AND int_day <= " + _clsFun.DaySql(dDtf) + " ";
            s += "ORDER BY int_day DESC";
            DataTable tInv = _clsFun.FillTabSql("GesInventatio", s, false, _strConSql);

            if (tInv.Rows.Count > 0)
            {
                progressBar1.Value = 0;
                progressBar1.Maximum = tInv.Rows.Count;
                progressBar1.Minimum = 0;

                foreach (DataRow y in tInv.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    j = tabTmp.Select("tmp_art='" + y["inv_art"] + "'");
                    if (j.Length == 0)
                    {
                        tabTmp.Rows.Add(InsArticolo(tabTmp, (string)y["inv_art"]));
                        j = tabTmp.Select("tmp_art='" + y["inv_art"] + "'");
                    }

                    AggValori("INV", j, y);
                }
            }
        }

        private void FillMovimenti(DataTable tabTmp)
        {
            string s = "";
            //string s = "SELECT art_cod, art_umi FROM AnaArticoli WHERE art_sta='" + _clsDef.STAATT + "'";
            //DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);
            //tMov = new DataTable("TabInv");
            DataRow x;
            DataRow[] j;
            //int iSgn = 1;

            s = "SELECT ";
            s += "GesFatTestate.fat_tpd, ";
            s += "GesFatTestate.fat_ddo, ";
            s += "GesMovimenti.mov_art, ";
            s += "SUM(GesMovimenti.mov_qta) AS mov_qta, ";
            s += "SUM(GesMovimenti.mov_qkg) AS mov_qkg ";
            s += "FROM GesMovimenti ";
            s += "INNER JOIN GesFatTestate ON GesMovimenti.mov_yfa = GesFatTestate.fat_yfa AND GesMovimenti.mov_nfa = GesFatTestate.fat_nfa ";
            s += "WHERE ";
            //s += "(GesMovimenti.mov_art = '" + y["art_cod"] + "') AND ";
            s += "(GesFatTestate.fat_ddo >= " + _clsFun.DaySql(dtpIni.Value) + " AND GesFatTestate.fat_ddo <= " + _clsFun.DaySql(dtpFin.Value) + ") AND ";
            s += "(GesMovimenti.mov_ori = '') ";
            s += "GROUP BY GesFatTestate.fat_tpd, GesFatTestate.fat_ddo, GesMovimenti.mov_art";
            DataTable tMov = _clsFun.FillTabSql("GesMovimenti", s, false, _strConSql);

            if (tMov.Rows.Count > 0)
            {
                progressBar1.Value = 0;
                progressBar1.Maximum = tMov.Rows.Count;
                progressBar1.Minimum = 0;

                foreach (DataRow y in tMov.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    j = tabTmp.Select("tmp_art='" + y["mov_art"] + "'");
                    if (j.Length == 0)
                    {
                        tabTmp.Rows.Add(InsArticolo(tabTmp, (string)y["mov_art"]));
                        j = tabTmp.Select("tmp_art='" + y["mov_art"] + "'");
                    }

                    AggValori("FAT", j, y);
                }
            }
        }

        private void FillVendite(DataTable tabTmp)
        {
            string s = "";
            DataRow x;
            DataRow[] j;
            int iSgn = 1;

            s = "SELECT ";
            s += "ven_day, ";
            s += "ven_art, ";
            s += "SUM(ven_qta) AS ven_qta, ";
            s += "SUM(ven_qkg) AS ven_qkg ";
            s += "FROM GesNegVen ";
            s += "WHERE (ven_day >= " + _clsFun.DaySql(dtpIni.Value) + " AND ven_day <= " + _clsFun.DaySql(dtpFin.Value) + ") ";
            s += "GROUP BY ven_day, ven_art";
            DataTable tVen = _clsFun.FillTabSql("GesVendite", s, false, _strConSqlSta);

            progressBar1.Value = 0;
            progressBar1.Maximum = tVen.Rows.Count;
            progressBar1.Minimum = 0;

            if (tVen.Rows.Count > 0)
            {
                foreach(DataRow y in tVen.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    j = tabTmp.Select("tmp_art='" + y["ven_art"] + "'");
                    if (j.Length == 0)
                    {
                        tabTmp.Rows.Add(InsArticolo(tabTmp, (string)y["ven_art"]));
                        j = tabTmp.Select("tmp_art='" + y["ven_art"] + "'");
                    }

                    AggValori("VEN", j, y);
                }
            }
        }

        private void  AggValori(string strTip, DataRow[] rowTmp, DataRow rowMov)
        {
            int iPer = 0;
            string sPre = "";
            string sFld = "";
            decimal dVal = 0;

            if (strTip == "INV")
            {
                sPre = "inv_";

                if (rdbTipVal.Checked)
                    dVal = (decimal)rowMov["inv_qta"] * (decimal)rowMov["inv_cos"];
                else
                    dVal = (decimal)rowMov["inv_qta"];

            }
            else if (strTip == "FAT")
            {
                if (rdbSelMon.Checked)
                    iPer = ((DateTime)rowMov["fat_ddo"]).Month;
                else
                    iPer = _ciCalendar.GetWeekOfYear((DateTime)rowMov["fat_ddo"], _ciFormat.CalendarWeekRule, _ciFormat.FirstDayOfWeek);

                sPre = "acq_";
                if ((string)rowMov["fat_tpd"] == "FV")
                    sPre = "ven_";

                if (rdbTipVal.Checked)
                    dVal = (decimal)rowMov["mov_qta"] * (decimal)rowMov["mov_cos"];
                else
                    dVal = (decimal)rowMov["mov_qta"];

            }
            else if (strTip == "VEN")
            {
                if (rdbSelMon.Checked)
                    iPer = ((DateTime)rowMov["ven_day"]).Month;
                else
                    iPer = _ciCalendar.GetWeekOfYear((DateTime)rowMov["ven_day"], _ciFormat.CalendarWeekRule, _ciFormat.FirstDayOfWeek);

                //sPre = "acq_";
                //if ((string)rowMov["fat_tpd"] == "FV")
                sPre = "ven_";

                if (rdbTipVal.Checked)
                    dVal = (decimal)rowMov["ven_qta"] * (decimal)rowMov["mov_cos"];
                else
                    dVal = (decimal)rowMov["ven_qta"];
            }

            sFld = sPre + iPer.ToString("000");

            rowTmp[0][sFld] = (decimal)rowTmp[0][sFld] + dVal;

            //if (strTip != "INV")
            //{
            //    sFld = "gia_" + iPer.ToString("000");
            //    string sAcq = "acq_" + iPer.ToString("000");
            //    string sVen = "ven_" + iPer.ToString("000");
            //    rowTmp[0][sFld] = (decimal)rowTmp[0][sAcq] - (decimal)rowTmp[0][sVen];
            //}
        }

        private DataRow InsArticolo(DataTable tabTmp, string strArt)
        {
            DataRow x = tabTmp.NewRow();
            x["tmp_ecr"] = "";
            x["tmp_ecd"] = "";
            x["tmp_rep"] = "";
            x["tmp_red"] = "";
            x["tmp_art"] = strArt;
            x["tmp_ard"] = "";
            x["tmp_umi"] = "";
            return x;
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "tmp_ard", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            string s = "";
            decimal d = 0;

            foreach (DataRowView r in v)
            {
                t.ImportRow(r.Row);
            }

            string sTit = DateTime.Today.ToString("dd/MM/yyyy") + " - Rotazione dal " + dtpIni.Value.ToString("dd/MM/yyyy");
            string sFil = "";

            s = "";

            foreach (DataColumn c in t.Columns)
            {
                s +=  c.ColumnName + ", " + c.ColumnName + ", 100, StringLiteral;";
            }

            string sFld = s;

            string sFoo = "";

            (new clsExcel()).exportToCsv1(sFld, t, sFil, sTit, sFoo);

        }
    }
}
