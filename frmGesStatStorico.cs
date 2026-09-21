using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmGesStatStorico : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strConSqlSta = "";
        private string _strConSqlStaSto = "";

        public frmGesStatStorico()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
            _strConSqlStaSto = _clsFun.ConSql("5");
        }

        private void frmGesStatStorico_Load(object sender, EventArgs e)
        {
            dtpDti.Value = new DateTime(DateTime.Today.Year, 1, 1);

            dtpDti.Value = new DateTime(DateTime.Today.Year, 1, 6);
            dtpDtf.Value = new DateTime(DateTime.Today.Year, 1, 12);
            
            FillTab();

            //SetDgv1();

            DataTable t = new clsGenTabTmp().TabTmpStaStorico("TmpSto");

            DataColumn[] keys = new DataColumn[1];
            keys[0] = t.Columns["sta_art"];
            t.PrimaryKey = keys;
            dgv1.DataSource = t;
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmGesStatStorico_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private void chkTipMon_CheckedChanged(object sender, EventArgs e)
        {
            if (!chkTipMon.Checked)
                chkTipMon.Text = "Mensile";
            else
                chkTipMon.Text = "Settimanale";
        }

        private void chkTipArt_CheckedChanged(object sender, EventArgs e)
        {
            if (chkTipArt.Checked)
                chkTipArt.Text = "Articolo";
            else
                chkTipArt.Text = "Reparto";
        }

        private void FillTab()
        {
            string p = "TabNegozi";
            string s = "SELECT * FROM TabNegozi ORDER BY tab_cod";
            DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
            DataRow x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Tutti";
            t.Rows.InsertAt(x, 0);
            cmbNeg.DataSource = t;
            cmbNeg.DisplayMember = "tab_des";
            cmbNeg.ValueMember = "tab_cod";
            cmbNeg.SelectedValue = "";
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            lblAcqQta.Text= "0";
            lblAcqQkg.Text= "0";
            lblAcqVal.Text= "0";
            lblVenQta.Text= "0";
            lblVenQkg.Text= "0";
            lblVenVal.Text= "0";
            lblVenVni.Text= "0";      //Venduto netto IVA dove c'è un costo
            lblVenVcs.Text= "0";      //Costo del venduto

            SetDgv1();

            if(chkTipArt.Checked)
                FillDatiArt();
            else
                FillDatiRep();
        }

        private void SetDgv1()
        {
            //DataTable tSta = _clsFun.FillTabSql("TabStato", "SELECT * FROM TabStato", false, _strConSql);

            dgv1.AutoGenerateColumns = false;
            dgv1.Columns.Clear();
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;
            dgv1.SelectionMode = DataGridViewSelectionMode.CellSelect;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            if (!chkTipArt.Checked)
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "str_rep";
                cTbc.Name = "Reparto";
                cTbc.Width = 30;
                cTbc.ValueType = typeof(string);
                cTbc.ReadOnly = true;
                //cTbc.Selected = true;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "str_des";
                cTbc.Name = "Descrizione";
                cTbc.Width = 100;
                cTbc.ValueType = typeof(string);
                cTbc.ReadOnly = true;
                //cTbc.Selected = true;
                dgv1.Columns.Add(cTbc);
            }
            else
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "art_rep";
                cTbc.Name = "Reparto";
                cTbc.Width = 30;
                cTbc.ValueType = typeof(string);
                cTbc.ReadOnly = true;
                //cTbc.Selected = true;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "art_red";
                cTbc.Name = "Descrizione";
                cTbc.Width = 100;
                cTbc.ValueType = typeof(string);
                cTbc.ReadOnly = true;
                //cTbc.Selected = true;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "sta_art";
                cTbc.Name = "Articolo";
                cTbc.Width = 50;
                cTbc.ValueType = typeof(string);
                cTbc.ReadOnly = true;
                //cTbc.Selected = true;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "sta_des";
                cTbc.Name = "Descrizione";
                cTbc.Width = 120;
                cTbc.ValueType = typeof(string);
                cTbc.ReadOnly = true;
                //cTbc.Selected = true;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "art_umi";
                cTbc.Name = "UM";
                cTbc.Width = 30;
                cTbc.ValueType = typeof(string);
                cTbc.ReadOnly = true;
                //cTbc.Selected = true;
                dgv1.Columns.Add(cTbc);
            }

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "acq_qta";
            cTbc.Name = "Acq. qta";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "acq_qkg";
            cTbc.Name = "Acq. kg";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "acq_val";
            cTbc.Name = "Acq. valore";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_qta";
            cTbc.Name = "Ven. qta";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_qkg";
            cTbc.Name = "Ven. kg";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_val";
            cTbc.Name = "Ven. valore";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_vni";
            cTbc.Name = "Ven. netto IVA sul costo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.ToolTipText = "Dove presente il costo";
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_vcs";
            cTbc.Name = "Ven. al costo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.ToolTipText = "Dove presente il costo";
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_vct";
            cTbc.Name = "Sconti";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.ToolTipText = "Dove presente il costo";
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "off_qta";
            cTbc.Name = "Off. qta";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "off_qkg";
            cTbc.Name = "Off. kg";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "off_val";
            cTbc.Name = "Off. valore";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "fid_qta";
            cTbc.Name = "Fid. qta";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "fid_qkg";
            cTbc.Name = "Fid. kg";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "fid_val";
            cTbc.Name = "Fid. valore";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_qta";
            cTbc.Name = "Mov. qta";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_qkg";
            cTbc.Name = "Mov. kg";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_val";
            cTbc.Name = "Mov. valore";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0.00";
        }

        private void FillDatiRep()
        {
            lblScn.Text = "0";
            lblVenVal.Text = "0";

            DataTable tCal = new DataTable();
            //((DataTable)dgv1.DataSource).Clear();
            DataTable t = new clsGenTabTmp().TabTmpStaStorico("TmpRep");
            DataColumn[] keys = new DataColumn[1];
            keys[0] = t.Columns["str_rep"];
            t.PrimaryKey = keys;
            dgv1.DataSource = t;

            string sNeg = cmbNeg.SelectedValue.ToString();

            string s = "";
            string sTri = "M";
            if (chkTipMon.Checked)
                sTri = "S";

            string sYea = dtpDti.Value.Year.ToString("00");
            string sPerIni = "";
            string sPerFin = "";

            decimal d = 0;

            if (sTri == "M")
            {
                sPerIni = dtpDti.Value.Month.ToString("00");
                sPerFin = dtpDtf.Value.Month.ToString("00");
            }
            else if (sTri == "S")
            {
                s = "SELECT * FROM TabCalendario WHERE ";
                s += "tab_tri='" + sTri + "' ";
                s += "ORDER BY tab_dti";
                tCal = _clsFun.FillTabSql("TabCalendario", s, false, _strConSqlStaSto);

                foreach (DataRow y in tCal.Rows)
                {
                    if ((string)y["tab_per"] == "33")
                        Console.WriteLine("xxxxxxx");

                    if ((DateTime)y["tab_dti"] == new DateTime(2018, 12, 31))
                        Console.WriteLine("xxxxxxx");

                    if (sPerIni == "")
                    {
                        TimeSpan ts = (DateTime)y["tab_dti"] - dtpDti.Value;

                        if (ts.Days >= 0)
                            sPerIni = (string)y["tab_per"];
                    }

                    if (sPerFin == "")
                    {
                        TimeSpan ts = (DateTime)y["tab_dtf"] - dtpDtf.Value;

                        if (ts.Days >= 0)
                            sPerFin = (string)y["tab_per"];
                    }
                }
            }

            Console.WriteLine("xxxxxxx");

            //s = "SELECT art_cod, art_des, art_umi, art_rep FROM AnaArticoli";
            //DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);
            //DataColumn[] keys = new DataColumn[1];
            //keys[0] = tArt.Columns["art_cod"];
            //tArt.PrimaryKey = keys;

            s = "SELECT * FROM TabReparti";
            DataTable tRep = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tRep.Columns["tab_cod"];
            tRep.PrimaryKey = keys;

            s = "SELECT * FROM TabCalendario WHERE ";
            s += "tab_yea='" + sYea + "' AND ";
            s += "tab_tri='" + sTri + "' AND ";
            s += "tab_per >= '" + sPerIni + "' AND ";
            s += "tab_per <= '" + sPerFin + "' ";
            s += "ORDER BY tab_dti";
            tCal = _clsFun.FillTabSql("TabCalendario", s, false, _strConSqlStaSto);

            progressBar1.Value = 0;
            progressBar1.Maximum = tCal.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in tCal.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                FillDatiStaRep(y, tRep, sNeg);
            }
        }

        private void FillDatiArt()
        {
            lblScn.Text = "0";
            lblVenVal.Text = "0";

            DataTable tCal = new DataTable();
            //((DataTable)dgv1.DataSource).Clear();
            DataTable t = new clsGenTabTmp().TabTmpStaStorico("TmpArt");
            DataColumn[] keys = new DataColumn[1];
            keys[0] = t.Columns["sta_art"];
            t.PrimaryKey = keys;
            dgv1.DataSource = t;

            string sNeg = cmbNeg.SelectedValue.ToString();

            string s = "";
            string sTri = "M";
            if (chkTipMon.Checked)
                sTri = "S";

            string sYea = dtpDti.Value.Year.ToString("00");
            string sPerIni = "";
            string sPerFin = "";

            decimal d = 0;

            if (sTri == "M")
            {
                sPerIni = dtpDti.Value.Month.ToString("00");
                sPerFin = dtpDtf.Value.Month.ToString("00");
            }
            else if (sTri == "S")
            {
                s = "SELECT * FROM TabCalendario WHERE ";
                s += "tab_tri='" + sTri + "' ";
                s += "ORDER BY tab_dti";
                tCal = _clsFun.FillTabSql("TabCalendario", s, false, _strConSqlStaSto);

                foreach (DataRow y in tCal.Rows)
                {
                    if ((string)y["tab_per"] == "33")
                        Console.WriteLine("xxxxxxx");

                    if ((DateTime)y["tab_dti"] == new DateTime(2018, 12, 31))
                        Console.WriteLine("xxxxxxx");

                    if (sPerIni == "")
                    {
                        TimeSpan ts = (DateTime)y["tab_dti"] - dtpDti.Value;

                        if (ts.Days >= 0)
                            sPerIni = (string)y["tab_per"];
                    }

                    if (sPerFin == "")
                    {
                        TimeSpan ts = (DateTime)y["tab_dtf"] - dtpDtf.Value;

                        if (ts.Days >= 0)
                            sPerFin = (string)y["tab_per"];
                    }
                }
            }

            Console.WriteLine("xxxxxxx");

            s = "SELECT art_cod, art_des, art_umi, art_rep FROM AnaArticoli";
            DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            s = "SELECT * FROM TabReparti";
            DataTable tRep = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tRep.Columns["tab_cod"];
            tRep.PrimaryKey = keys;

            s = "SELECT * FROM TabCalendario WHERE ";
            s += "tab_yea='" + sYea + "' AND ";
            s += "tab_tri='" + sTri + "' AND ";
            s += "tab_per >= '" + sPerIni + "' AND ";
            s += "tab_per <= '" + sPerFin + "' ";
            s += "ORDER BY tab_dti";
            tCal = _clsFun.FillTabSql("TabCalendario", s, false, _strConSqlStaSto);

            progressBar1.Value = 0;
            progressBar1.Maximum = tCal.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in tCal.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                FillDatiStaArt(y, tArt, tRep, sNeg);
            }
        }

        private void FillDatiStaArt(DataRow rowCal, DataTable tabArt, DataTable tabRep, string strNeg)
        {
            string s = "";
            DataRow x;
            DataRow[] j;

            decimal dAcqQta = Convert.ToDecimal(lblAcqQta.Text.Replace(".", "")); //.Replace(",", "."));
            decimal dAcqQkg = Convert.ToDecimal(lblAcqQkg.Text.Replace(".", "")); //.Replace(",", "."));
            decimal dAcqVal = Convert.ToDecimal(lblAcqVal.Text.Replace(".", "")); //.Replace(",", "."));

            decimal dVenQta = Convert.ToDecimal(lblVenQta.Text.Replace(".", "")); //.Replace(",", "."));
            decimal dVenQkg = Convert.ToDecimal(lblVenQkg.Text.Replace(".", "")); //.Replace(",", "."));
            decimal dVenVal = Convert.ToDecimal(lblVenVal.Text.Replace(".", "")); //.Replace(",", "."));
            decimal dVenVni = Convert.ToDecimal(lblVenVni.Text.Replace(".", "")); //.Replace(",", "."));      //Venduto netto IVA dove c'è un costo
            decimal dVenVcs = Convert.ToDecimal(lblVenVcs.Text.Replace(".", "")); //.Replace(",", "."));      //Costo del venduto

            string sYea = (string)rowCal["tab_yea"];
            string sTri = (string)rowCal["tab_tri"];
            string sPer = (string)rowCal["tab_per"];

            DataTable t = (DataTable)dgv1.DataSource;

            s = "SELECT * FROM StaPerArticoli WHERE ";
            s += "sta_yea='" + sYea + "' AND ";
            s += "sta_tri='" + sTri + "' AND ";
            s += "sta_per='" + sPer + "' ";
            if (strNeg != "")
                s += "AND sta_neg='" + strNeg + "' ";
            s += "ORDER BY sta_art";

            DataTable tSta = _clsFun.FillTabSql("StaPerArticoli", s, false, _strConSqlStaSto);

            foreach (DataRow y in tSta.Rows)
            {
                j = t.Select("sta_art='" + y["sta_art"] + "'");
                if (j.Length == 0)
                {
                    j = tabArt.Select("art_cod='" + (string)y["sta_art"] + "'");

                    x = t.NewRow();

                    x["sta_art"] = y["sta_art"];
                    x["sta_des"] = y["sta_des"];
                    x["art_umi"] = "";
                    x["art_rep"] = "";
                    x["art_red"] = "";
                    x["art_ecr"] = "";
                    x["art_ecd"] = "";

                    if (j.Length > 0)
                    {
                        x["art_umi"] = j[0]["art_umi"];
                        x["art_rep"] = j[0]["art_rep"];
                        x["art_red"] = "";
                        x["art_ecr"] = "";
                        x["art_ecd"] = "";

                        if ((string)x["art_rep"] != "")
                        {
                            j = tabRep.Select("tab_cod='" + (string)x["art_rep"] + "'");
                            if (j.Length > 0)
                                x["art_red"] = (string)j[0]["tab_des"];
                        }
                    }

                    x["acq_qta"] = 0;
                    x["acq_val"] = 0;

                    x["ven_qta"] = 0;
                    x["ven_val"] = 0;

                    x["mov_qta"] = 0;
                    x["mov_val"] = 0;
                    x["off_qta"] = 0;
                    x["off_val"] = 0;
                    x["fid_qta"] = 0;
                    x["fid_val"] = 0;

                    t.Rows.Add(x);

                    j = t.Select("sta_art='" + y["sta_art"] + "'");
                }

                j[0]["acq_qta"] = (decimal)j[0]["acq_qta"] + (decimal)y["acq_qta"];
                j[0]["acq_qkg"] = (decimal)j[0]["acq_qkg"] + (decimal)y["acq_qkg"];
                j[0]["acq_val"] = (decimal)j[0]["acq_val"] + (decimal)y["acq_val"];
                j[0]["ven_qta"] = (decimal)j[0]["ven_qta"] + (decimal)y["ven_qta"];
                j[0]["ven_qkg"] = (decimal)j[0]["ven_qkg"] + (decimal)y["ven_qkg"];
                j[0]["ven_val"] = (decimal)j[0]["ven_val"] + (decimal)y["ven_val"];
                j[0]["ven_vni"] = (decimal)j[0]["ven_vni"] + (decimal)y["ven_vni"];
                j[0]["ven_vcs"] = (decimal)j[0]["ven_vcs"] + (decimal)y["ven_vcs"];
                j[0]["mov_qta"] = (decimal)j[0]["mov_qta"] + (decimal)y["mov_qta"];
                j[0]["mov_qkg"] = (decimal)j[0]["mov_qkg"] + (decimal)y["mov_qkg"];
                j[0]["mov_val"] = (decimal)j[0]["mov_val"] + (decimal)y["mov_val"];
                j[0]["off_qta"] = (decimal)j[0]["off_qta"] + (decimal)y["ofv_qta"];
                j[0]["off_qkg"] = (decimal)j[0]["off_qkg"] + (decimal)y["ofv_qkg"];
                j[0]["off_val"] = (decimal)j[0]["off_val"] + (decimal)y["ofv_val"];
                j[0]["fid_qta"] = (decimal)j[0]["fid_qta"] + (decimal)y["fid_qta"];
                j[0]["fid_qkg"] = (decimal)j[0]["fid_qkg"] + (decimal)y["fid_qkg"];
                j[0]["fid_val"] = (decimal)j[0]["fid_val"] + (decimal)y["fid_val"];

                dAcqQta += (decimal)y["acq_qta"];
                dAcqQkg += (decimal)y["acq_qkg"];
                dAcqVal += (decimal)y["acq_val"];

                dVenQta += (decimal)y["ven_qta"];
                dVenQkg += (decimal)y["ven_qkg"];
                dVenVal += (decimal)y["ven_val"];
                dVenVni += (decimal)y["ven_vni"];
                dVenVcs += (decimal)y["ven_vcs"];
            }

            lblAcqQta.Text = dAcqQta.ToString("#,###,##0");
            lblAcqQkg.Text = dAcqQkg.ToString("#,###,##0.00");
            lblAcqVal.Text = dAcqVal.ToString("#,###,##0.00");

            lblVenQta.Text = dVenQta.ToString("#,###,##0");
            lblVenQkg.Text = dVenQkg.ToString("#,###,##0.00");
            lblVenVal.Text = dVenVal.ToString("#,###,##0.00");
            lblVenVni.Text = dVenVni.ToString("#,###,##0.00");
            lblVenVcs.Text = dVenVcs.ToString("#,###,##0.00");

            decimal dScn = 0;
            s = "SELECT * FROM StaPeriodi WHERE ";
            s += "stp_yea='" + sYea + "' AND ";
            s += "stp_tri='" + sTri + "' AND ";
            s += "stp_per='" + sPer + "' ";
            if (strNeg != "")
                s += "AND stp_neg='" + strNeg + "' ";
            tSta = _clsFun.FillTabSql("StaPeriodi", s, false, _strConSqlStaSto);
            foreach (DataRow y in tSta.Rows)
            {
                dScn += (decimal)y["stp_scn"];
            }

            lblScn.Text = dScn.ToString("#,###,##0");
            lblVenUti.Text = (dVenVni - dVenVcs).ToString("###,##0.00");
            lblVenMrg.Text = _clsFun.Margine(dVenVni, dVenVcs, 0, 0, "P").ToString("###,##0.00");
        }

        private void FillDatiStaRep(DataRow rowCal, DataTable tabRep, string strNeg)
        {
            string s = "";
            DataRow x;
            DataRow[] j;

            decimal dAcqQta = Convert.ToDecimal(lblAcqQta.Text.Replace(".", "").Replace(",", "."));
            decimal dAcqQkg = Convert.ToDecimal(lblAcqQkg.Text.Replace(".", "").Replace(",", "."));
            decimal dAcqVal = Convert.ToDecimal(lblAcqVal.Text.Replace(".", "").Replace(",", "."));

            decimal dVenQta = Convert.ToDecimal(lblVenQta.Text.Replace(".", "").Replace(",", "."));
            decimal dVenQkg = Convert.ToDecimal(lblVenQkg.Text.Replace(".", "").Replace(",", "."));
            decimal dVenVal = Convert.ToDecimal(lblVenVal.Text.Replace(".", "").Replace(",", "."));
            decimal dVenVni = Convert.ToDecimal(lblVenVni.Text.Replace(".", "").Replace(",", "."));      //Venduto netto IVA dove c'è un costo
            decimal dVenVcs = Convert.ToDecimal(lblVenVcs.Text.Replace(".", "").Replace(",", "."));      //Costo del venduto

            string sYea = (string)rowCal["tab_yea"];
            string sTri = (string)rowCal["tab_tri"];
            string sPer = (string)rowCal["tab_per"];

            DataTable t = (DataTable)dgv1.DataSource;

            s = "SELECT * FROM StaPerReparti WHERE ";
            s += "str_yea='" + sYea + "' AND ";
            s += "str_tri='" + sTri + "' AND ";
            s += "str_per='" + sPer + "' ";
            if (strNeg != "")
                s += "AND str_neg='" + strNeg + "' ";
            s += "ORDER BY str_rep";

            DataTable tSta = _clsFun.FillTabSql("StaPerReparti", s, false, _strConSqlStaSto);

            foreach (DataRow y in tSta.Rows)
            {
                j = t.Select("str_rep='" + y["str_rep"] + "'");
                if (j.Length == 0)
                {
                    j = tabRep.Select("tab_cod='" + (string)y["str_rep"] + "'");

                    x = t.NewRow();

                    x["str_rep"] = y["str_rep"];
                    x["str_des"] = "";
                    //x["art_umi"] = "";
                    //x["art_rep"] = "";
                    //x["art_red"] = "";
                    //x["art_ecr"] = "";
                    //x["art_ecd"] = "";

                    if (j.Length > 0)
                    {
                        x["str_des"] = j[0]["tab_des"];
                        //x["art_umi"] = j[0]["art_umi"];
                        //x["art_rep"] = j[0]["art_rep"];
                        //x["art_red"] = "";
                        //x["art_ecr"] = "";
                        //x["art_ecd"] = "";

                        //if ((string)x["art_rep"] != "")
                        //{
                        //    j = tabRep.Select("tab_cod='" + (string)x["art_rep"] + "'");
                        //    if (j.Length > 0)
                        //        x["art_red"] = (string)j[0]["tab_des"];
                        //}
                    }

                    x["acq_qta"] = 0;
                    x["acq_val"] = 0;

                    x["ven_qta"] = 0;
                    x["ven_val"] = 0;

                    x["mov_qta"] = 0;
                    x["mov_val"] = 0;
                    x["off_qta"] = 0;
                    x["off_val"] = 0;
                    x["fid_qta"] = 0;
                    x["fid_val"] = 0;

                    t.Rows.Add(x);

                    j = t.Select("str_rep='" + y["str_rep"] + "'");
                }

                j[0]["acq_qta"] = (decimal)j[0]["acq_qta"] + (decimal)y["acq_qta"];
                j[0]["acq_qkg"] = (decimal)j[0]["acq_qkg"] + (decimal)y["acq_qkg"];
                j[0]["acq_val"] = (decimal)j[0]["acq_val"] + (decimal)y["acq_val"];
                j[0]["ven_qta"] = (decimal)j[0]["ven_qta"] + (decimal)y["ven_qta"];
                j[0]["ven_qkg"] = (decimal)j[0]["ven_qkg"] + (decimal)y["ven_qkg"];
                j[0]["ven_val"] = (decimal)j[0]["ven_val"] + (decimal)y["ven_val"];
                j[0]["ven_vni"] = (decimal)j[0]["ven_vni"] + (decimal)y["ven_vni"];
                j[0]["ven_vcs"] = (decimal)j[0]["ven_vcs"] + (decimal)y["ven_vcs"];
                j[0]["mov_qta"] = (decimal)j[0]["mov_qta"] + (decimal)y["mov_qta"];
                j[0]["mov_qkg"] = (decimal)j[0]["mov_qkg"] + (decimal)y["mov_qkg"];
                j[0]["mov_val"] = (decimal)j[0]["mov_val"] + (decimal)y["mov_val"];
                j[0]["off_qta"] = (decimal)j[0]["off_qta"] + (decimal)y["ofv_qta"];
                j[0]["off_qkg"] = (decimal)j[0]["off_qkg"] + (decimal)y["ofv_qkg"];
                j[0]["off_val"] = (decimal)j[0]["off_val"] + (decimal)y["ofv_val"];
                j[0]["fid_qta"] = (decimal)j[0]["fid_qta"] + (decimal)y["fid_qta"];
                j[0]["fid_qkg"] = (decimal)j[0]["fid_qkg"] + (decimal)y["fid_qkg"];
                j[0]["fid_val"] = (decimal)j[0]["fid_val"] + (decimal)y["fid_val"];

                dAcqQta += (decimal)y["acq_qta"];
                dAcqQkg += (decimal)y["acq_qkg"];
                dAcqVal += (decimal)y["acq_val"];

                dVenQta += (decimal)y["ven_qta"];
                dVenQkg += (decimal)y["ven_qkg"];
                dVenVal += (decimal)y["ven_val"];
                dVenVni += (decimal)y["ven_vni"];
                dVenVcs += (decimal)y["ven_vcs"];
            }

            lblAcqQta.Text = dAcqQta.ToString("#,###,##0");
            lblAcqQkg.Text = dAcqQkg.ToString("#,###,##0.00");
            lblAcqVal.Text = dAcqVal.ToString("#,###,##0.00");

            lblVenQta.Text = dVenQta.ToString("#,###,##0");
            lblVenQkg.Text = dVenQkg.ToString("#,###,##0.00");
            lblVenVal.Text = dVenVal.ToString("#,###,##0.00");
            lblVenVni.Text = dVenVni.ToString("#,###,##0.00");
            lblVenVcs.Text = dVenVcs.ToString("#,###,##0.00");

            decimal dScn = 0;
            s = "SELECT * FROM StaPeriodi WHERE ";
            s += "stp_yea='" + sYea + "' AND ";
            s += "stp_tri='" + sTri + "' AND ";
            s += "stp_per='" + sPer + "' ";
            if (strNeg != "")
                s += "AND stp_neg='" + strNeg + "' ";
            tSta = _clsFun.FillTabSql("StaPeriodi", s, false, _strConSqlStaSto);
            foreach (DataRow y in tSta.Rows)
            {
                dScn += (decimal)y["stp_scn"];
            }

            lblScn.Text = dScn.ToString("#,###,##0");
            lblVenUti.Text = (dVenVni - dVenVcs).ToString("###,##0.00");
            lblVenMrg.Text = _clsFun.Margine(dVenVni, dVenVcs, 0, 0, "P").ToString("###,##0.00");

        }

        private void pdfXRepartoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s = "";

            DataTable t = FillReparto((DataTable)dgv1.DataSource);

            //decimal dTotNiv = 0;    //Somma venduto netto iVA
            //decimal dTotCdv = 0;    //Somma del venduto

            decimal dVenVal = Convert.ToDecimal(lblVenVal.Text.Replace(".", "").Replace(",", "."));
            decimal dVenQta = Convert.ToDecimal(lblVenQta.Text.Replace(".", "").Replace(",", "."));
            decimal dScn = Convert.ToDecimal(lblScn.Text.Replace(".", "").Replace(",", "."));

            DataTable t2 = t.Clone();
            DataView v = new DataView(t, "", "tmp_cod", DataViewRowState.CurrentRows);

            foreach (DataRowView r in v)
            {
                r["tmp_inc"] = (decimal)r["tmp_imp"] / dVenVal * 100;
                r["tmp_sci"] = (decimal)r["tmp_sco"] / dScn * 100;
                if ((decimal)r["tmp_sco"] > 0)
                    r["tmp_scm"] = (decimal)r["tmp_imp"] / (decimal)r["tmp_sco"];
                r["tmp_qti"] = (decimal)r["tmp_qta"] / dVenQta * 100;
                if ((decimal)r["tmp_qta"] > 0)
                    r["tmp_qtm"] = (decimal)r["tmp_imp"] / (decimal)r["tmp_qta"];

                if ((decimal)r["tmp_vcc"] > 0 && (decimal)r["tmp_cdv"] > 0)
                    r["tmp_mav"] = Math.Round(((decimal)r["tmp_vcc"] - (decimal)r["tmp_cdv"]) / (decimal)r["tmp_vcc"], 2) * 100;

                t2.ImportRow(r.Row);

                //dTotCdv = dTotCdv + (decimal)r["tmp_cdv"];       //Costo del venduto netto iVA
                //dTotNiv = dTotNiv + (decimal)r["tmp_vcc"];       //Venduto netto iVA
            }

            string sNeg = "";
            s = cmbNeg.SelectedValue.ToString();
            if (s != "")
                sNeg = cmbNeg.Text;

            clsGenPdfSta1 cls = new clsGenPdfSta1();
            cls._dayIni = dtpDti.Value;
            cls._dayFin = dtpDtf.Value;
            cls.PrnPdfRepSto("ven_rep", t, dScn, 0, sNeg, 0);
        }

        private DataTable FillReparto(DataTable tabSta)
        {
            DataTable t = new clsGenTabTmp().TabTmpStaReparto("TabRep");
            DataRow x;

            ArrayList aScoRep = new ArrayList();

            DataView v = new DataView(tabSta, "", "art_rep", DataViewRowState.CurrentRows);

            foreach (DataRowView r in v)
            {
                string sRep = (string)r["art_rep"];
                decimal d = 0;
                decimal dQta = 0;
                decimal dTot = 0;
                decimal dTls = 0;
                decimal dNiv = 0;

                DataRow[] j;
                //if (sNeg == "")
                j = t.Select("tmp_cod='" + sRep + "'");
                //else
                //    j = t.Select("tmp_cod='" + sRep + "' AND tmp_neg='" + sNeg + "'");

                if (j.Length == 0)
                {
                    x = t.NewRow();
                    x["tmp_cod"] = r["art_rep"];
                    x["tmp_des"] = r["art_red"];
                    x["tmp_ils"] = 0;               //Importo al lordo degli sconti
                    x["tmp_imp"] = 0;
                    x["tmp_inc"] = 0;
                    x["tmp_niv"] = 0;
                    t.Rows.Add(x);

                    j = t.Select("tmp_cod='" + sRep + "'");
                }

                dQta += (decimal)r["ven_qta"];
                dTot += (decimal)r["ven_val"];
                j[0]["tmp_imp"] = (decimal)j[0]["tmp_imp"] + (decimal)r["ven_val"];
                j[0]["tmp_niv"] = (decimal)j[0]["tmp_niv"] + (decimal)r["ven_vni"];
                j[0]["tmp_qta"] = (decimal)j[0]["tmp_qta"] + (decimal)r["ven_qta"];

            }

            return t;
        }

        private void utlitàToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void cancellaStoricoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmGesStatStoCancella f = new frmGesStatStoCancella();
            f.ShowDialog();
            f.Close();
        }

    }
}
