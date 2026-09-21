using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmGesStatArt : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public string _strArtCod = "";
        public string _strArtDes = "";
        public string _strArtUmi = "";

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public frmGesStatArt()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
        }

        private void frmGesStatArt_Load(object sender, EventArgs e)
        {
            this.Text += " " + _strArtCod + " " + _strArtDes;
            dtpIni.Value = new DateTime(DateTime.Now.Year, 1, 1);
            FillTabs();
            SetDgv1();
        }
        private void frmGesStatArt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
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
            cTbc.DataPropertyName = "tmp_neg";
            cTbc.Name = "Negozio";
            cTbc.Width = 35;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_tip";
            cTbc.Name = "Tipo";
            cTbc.Width = 35;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_day";
            cTbc.Name = "Data";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_qta";
            cTbc.Name = "Q.tà";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_qkg";
            cTbc.Name = "Q.peso";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_acq";
            cTbc.Name = "Acquisti";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_prv";
            cTbc.Name = "Prezzo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_ven";
            cTbc.Name = "Vendite";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_mov";
            cTbc.Name = "Movimenti";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 300;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 40;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);
        }

        private void FillTabs()
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

            p = "TabMagazzini";
            s = "SELECT * FROM TabMagazzini ORDER BY tab_cod";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Tutti";
            t.Rows.InsertAt(x, 0);
            cmbMag.DataSource = t;
            cmbMag.DisplayMember = "tab_des";
            cmbMag.ValueMember = "tab_cod";
            cmbMag.SelectedValue = "";
        }

        private void FillDati()
        {
            DataTable t = new clsGenTabTmp().TabTmpStaArt("StaArt");

            string s = "";
            DataRow[] j;
            DataRow x;

            decimal dIQt = 0;
            decimal dIQk = 0;
            decimal dIVa = 0;

            decimal dAQt = 0;
            decimal dAQk = 0;
            decimal dAcq = 0;

            decimal dVQt = 0;
            decimal dVQk = 0;
            decimal dVen = 0;

            decimal dMQt = 0;
            decimal dMQk = 0;
            decimal dMov = 0;

            s = "SELECT for_cod, for_des FROM AnaFornitori";
            DataTable tFor = _clsFun.FillTabSql("For", s, false, _strConSql);

            s = "SELECT cli_cod, cli_des FROM AnaClienti";
            DataTable tCli = _clsFun.FillTabSql("Cli", s, false, _strConSql);

            s = "SELECT ";
            s += "int_neg, ";
            s += "inv_art, ";
            s += "inv_qta, ";
            s += "inv_cos, ";
            s += "int_day ";
            s += "FROM GesInventario LEFT JOIN GesInvTestate ON GesInventario.inv_num = GesInvTestate.int_num ";
            s += "WHERE ";
            s += "int_day >= " + _clsFun.DaySql(dtpIni.Value) + " AND ";
            s += "int_tip<>'G' AND inv_art = " + _strArtCod + " ";
            s += "ORDER BY int_day DESC";
            DataTable tMov = _clsFun.FillTabSql("AnaArtGiacenza", s, false, _strConSql);
            if(tMov.Rows.Count > 0)
            {
                decimal dVal = 0;

                foreach (DataRow y in tMov.Rows)
                {
                    x = t.NewRow();
                    x["tmp_neg"] = y["int_neg"];
                    x["tmp_tip"] = "INV";
                    x["tmp_day"] = y["int_day"];

                    x["tmp_cos"] = y["inv_cos"];

                    if (_strArtUmi == "KG")
                    {
                        x["tmp_qta"] = 1;
                        x["tmp_qkg"] = y["inv_qta"];
                        dVal += (decimal)x["tmp_qkg"] * (decimal)y["inv_cos"];
                    }
                    else
                    {
                        x["tmp_qta"] = y["inv_qta"];
                        dVal += (decimal)x["tmp_qta"] * (decimal)y["inv_cos"];
                    }
                    x["tmp_des"] = "Inventario";
                    t.Rows.Add(x);

                    dIQt += (decimal)x["tmp_qta"];
                    dIQk += (decimal)x["tmp_qkg"];
                    dIVa += dVal;
                }
            }

            s = "SELECT vet_neg, ven_art, ven_day, SUM(ven_qta) AS ven_qta, SUM(ven_qkg) AS ven_qkg, SUM(ven_ven) AS ven_ven ";
            s += "FROM GesNegVen INNER JOIN GesNegVet ON ";
            s += "GesNegVet.vet_neg = GesNegVen.ven_neg AND ";
            s += "GesNegVet.vet_cau = GesNegVen.ven_cau AND ";
            s += "GesNegVet.vet_day = GesNegVen.ven_day AND ";
            s += "GesNegVet.vet_ora = GesNegVen.ven_ora AND ";
            s += "GesNegVet.vet_pos = GesNegVen.ven_pos AND ";
            s += "GesNegVet.vet_sco = GesNegVen.ven_sco ";
            s += "WHERE ";
            s += "GesNegVen.ven_art = '" + _strArtCod + "' AND ";
            s += "GesNegVen.ven_day >= " + _clsFun.DaySql(dtpIni.Value) + " AND ";
            s += "GesNegVen.ven_day <= " + _clsFun.DaySql(dtpFin.Value) + " ";
            
            if(cmbNeg.SelectedValue.ToString() != "")
                s += "AND GesNegVet.vet_neg = '" + cmbNeg.SelectedValue.ToString() + "' ";

            s += "GROUP BY vet_neg, ven_art, ven_day";
            tMov = _clsFun.FillTabSql("Ven", s, false, _strConSqlSta);

            foreach(DataRow y in tMov.Rows)
            {
                x = t.NewRow();
                x["tmp_neg"] = y["vet_neg"];
                x["tmp_tip"] = "VEN";
                x["tmp_day"] = y["ven_day"];
                x["tmp_qta"] = y["ven_qta"];
                x["tmp_qkg"] = y["ven_qkg"];

                if ((decimal)y["ven_qkg"] != 0 && (decimal)y["ven_ven"] != 0)
                    x["tmp_prv"] = (decimal)y["ven_ven"] / (decimal)y["ven_qkg"];
                else if ((decimal)y["ven_qta"] != 0 && (decimal)y["ven_ven"] != 0)
                    x["tmp_prv"] = (decimal)y["ven_ven"] / (decimal)y["ven_qta"];
                
                x["tmp_ven"] = y["ven_ven"];

                x["tmp_des"] = "Venduto casse";
                t.Rows.Add(x);

                dVQt += (decimal)x["tmp_qta"];
                dVQk += (decimal)x["tmp_qkg"];
                dVen += (decimal)x["tmp_ven"];
            }

            s = "SELECT ";
            s += "GesFatTestate.fat_neg, ";
            s += "GesFatTestate.fat_yfa, ";
            s += "GesFatTestate.fat_nfa, ";
            s += "GesFatTestate.fat_tdo, ";
            s += "GesFatTestate.fat_tpd, ";
            s += "GesFatTestate.fat_ddo, ";
            s += "GesFatTestate.fat_ndo, ";
            s += "GesFatTestate.fat_cfo, ";
            s += "GesMovimenti.mov_yfa, ";
            s += "GesMovimenti.mov_nfa, ";
            s += "GesMovimenti.mov_art, ";
            s += "GesMovimenti.mov_umi, ";
            s += "GesMovimenti.mov_qta, ";
            s += "GesMovimenti.mov_qkg, ";
            s += "GesMovimenti.mov_cos, ";
            s += "GesMovimenti.mov_prv, ";
            s += "GesMovimenti.mov_imp, ";
            s += "GesMovimenti.mov_ori, ";
            s += "TabDocTpd.tab_des ";
            s += "FROM GesFatTestate ";
            s += "INNER JOIN GesMovimenti ON GesFatTestate.fat_yfa = GesMovimenti.mov_yfa AND GesFatTestate.fat_nfa = GesMovimenti.mov_nfa ";
            s += "INNER JOIN TabDocTpd ON GesFatTestate.fat_tdo = TabDocTpd.tab_cod ";
            s += "WHERE mov_art='" + _strArtCod + "' AND ";
            s += "GesFatTestate.fat_ddo >= " + _clsFun.DaySql(dtpIni.Value) + " AND ";
            s += "GesFatTestate.fat_ddo <= " + _clsFun.DaySql(dtpFin.Value) + " ";
            
            if (cmbNeg.SelectedValue.ToString() != "")
                s += "AND GesFatTestate.fat_neg = '" + cmbNeg.SelectedValue.ToString() + "' ";

            tMov = _clsFun.FillTabSql("Ven", s, false, _strConSql);

            foreach (DataRow y in tMov.Rows)
            {
                x = t.NewRow();
                x["tmp_neg"] = y["fat_neg"];
                x["tmp_tip"] = "FAT";
                if ((string)y["fat_tdo"] == "NA")
                    x["tmp_tip"] = "NAC";

                x["tmp_day"] = y["fat_ddo"];

                if ((string)y["mov_umi"] == "KG")
                {
                    x["tmp_qta"] = 1;
                    if ((decimal)y["mov_qkg"] > 0)
                        x["tmp_qkg"] = y["mov_qkg"];
                    else
                        x["tmp_qkg"] = (decimal)y["mov_qta"];
                }
                else
                    x["tmp_qta"] = y["mov_qta"];

                x["tmp_cos"] = y["mov_cos"];

                if ((string)y["fat_tpd"] == "FV" || (string)y["fat_tpd"] == "PR")
                {
                    x["tmp_ven"] = y["mov_imp"];
                    x["tmp_prv"] = (decimal)y["mov_prv"];
                }
                else
                {
                    if ((decimal)y["mov_imp"] > 0)
                        x["tmp_acq"] = y["mov_imp"];
                    else
                    {
                        if ((decimal)x["tmp_qkg"] > 0)
                            x["tmp_acq"] = (decimal)x["tmp_qkg"] * (decimal)y["mov_cos"];
                        else
                            x["tmp_acq"] = (decimal)y["mov_qta"] * (decimal)y["mov_cos"];
                    }
                    if ((decimal)y["mov_cos"] == 0)
                    {
                        if (!DBNull.Value.Equals(y["mov_umi"]) && (string)y["mov_umi"] == "KG")
                            x["tmp_cos"] = (decimal)x["tmp_acq"] / (decimal)x["tmp_qkg"];
                        else
                            x["tmp_cos"] = (decimal)x["tmp_acq"] / (decimal)x["tmp_qta"];
                    }
                }

                if ((string)y["fat_tdo"] == "NA")
                {
                    x["tmp_qta"] = (decimal)x["tmp_qta"] * -1;
                    x["tmp_acq"] = (decimal)x["tmp_acq"] * -1;
                }


                if ((string)y["fat_tpd"] == "FV")
                {
                    x["tmp_des"] = "Fatt. vendita " + (string)y["fat_ndo"] + " ";
                    j = tCli.Select("cli_cod='" + y["fat_cfo"] + "'");
                    if (j.Length > 0)
                        x["tmp_des"] = (string)x["tmp_des"] + " di " + (string)j[0]["cli_des"];

                    if (((string)y["mov_ori"]).Trim() == "")
                    {
                        dAQt += (decimal)x["tmp_qta"];
                        dAQk += (decimal)x["tmp_qkg"];
                        dAcq += (decimal)x["tmp_acq"];
                    }
                }
                else
                {
                    x["tmp_des"] = "Fatt. acquisto " + (string)y["fat_ndo"] + " ";
                    j = tFor.Select("for_cod='" + y["fat_cfo"] + "'");
                    if (j.Length > 0)
                        x["tmp_des"] = (string)x["tmp_des"] + " di " + (string)j[0]["for_des"];

                    dAQt += (decimal)x["tmp_qta"];
                    dAQk += (decimal)x["tmp_qkg"];
                    dAcq += (decimal)x["tmp_acq"];
                }
                t.Rows.Add(x);
            }

            /*********/

            s = "SELECT ";
            s += "GesMovTestate.mot_neg, ";
            s += "GesMovTestate.mot_yfa, ";
            s += "GesMovTestate.mot_nfa, ";
            //s += "GesMovTestate.mot_tdo, ";
            //s += "GesMovTestate.mot_tpd, ";
            s += "GesMovTestate.mot_ddo, ";
            s += "GesMovTestate.mot_ndo, ";
            s += "GesMovTestate.mot_cfo, ";
            s += "GesMovimenti.mov_ymo, ";
            s += "GesMovimenti.mov_nmo, ";
            s += "GesMovimenti.mov_art, ";
            s += "GesMovimenti.mov_umi, ";
            s += "GesMovimenti.mov_qta, ";
            s += "GesMovimenti.mov_qkg, ";
            s += "GesMovimenti.mov_cos, ";
            s += "GesMovimenti.mov_prv, ";
            s += "GesMovimenti.mov_imp, ";
            s += "GesMovimenti.mov_ori, ";
            s += "TabMovCausali.tab_sgm, ";
            s += "TabMovCausali.tab_des ";
            s += "FROM GesMovTestate LEFT OUTER JOIN ";

            //s += "TabMovCausali.tab_sgm ";
            //s += "FROM GesMovTestate LEFT OUTER JOIN ";
            s += "GesMovimenti ON GesMovTestate.mot_ymo = GesMovimenti.mov_ymo AND GesMovTestate.mot_nmo = GesMovimenti.mov_nmo LEFT OUTER JOIN ";
            s += "TabMovCausali ON GesMovTestate.mot_cau = TabMovCausali.tab_cod ";

            //s += "INNER JOIN GesMovimenti ON GesMovTestate.fat_yfa = GesMovimenti.mov_yfa AND GesMovTestate.mot_nfa = GesMovimenti.mov_nfa ";
            //s += "INNER JOIN TabDocTpd ON GesMovTestate.mot_tdo = TabDocTpd.tab_cod ";

            s += "WHERE mov_art='" + _strArtCod + "' AND ";
            s += "GesMovTestate.mot_ddo >= " + _clsFun.DaySql(dtpIni.Value) + " AND ";
            s += "GesMovTestate.mot_ddo <= " + _clsFun.DaySql(dtpFin.Value) + " ";

            if (cmbMag.SelectedValue.ToString() != "")
                s += "AND GesMovTestate.mot_mag = '" + cmbMag.SelectedValue.ToString() + "' ";
            if (cmbNeg.SelectedValue.ToString() != "")
                s += "AND GesMovTestate.mot_neg = '" + cmbNeg.SelectedValue.ToString() + "' ";

            tMov = _clsFun.FillTabSql("Mov", s, false, _strConSql);

            foreach (DataRow y in tMov.Rows)
            {
                x = t.NewRow();
                x["tmp_neg"] = y["mot_neg"];
                x["tmp_tip"] = "MOV";
                //if ((string)y["fat_tdo"] == "NA")
                //    x["tmp_tip"] = "NAC";

                x["tmp_day"] = y["mot_ddo"];

                if ((string)y["mov_umi"] == "KG")
                {
                    x["tmp_qta"] = 1;
                    if ((decimal)y["mov_qkg"] > 0)
                        x["tmp_qkg"] = y["mov_qkg"];
                    else
                        x["tmp_qkg"] = (decimal)y["mov_qta"];
                }
                else
                    x["tmp_qta"] = y["mov_qta"];

                x["tmp_cos"] = y["mov_cos"];

                //if ((string)y["fat_tpd"] == "FV")
                //{
                //    x["tmp_ven"] = y["mov_imp"];
                //    x["tmp_prv"] = (decimal)y["mov_prv"];
                //}
                //else
                //{
                //    if ((decimal)y["mov_imp"] > 0)
                //        x["tmp_acq"] = y["mov_imp"];
                //    else
                //    {
                //        if ((decimal)x["tmp_qkg"] > 0)
                //            x["tmp_acq"] = (decimal)x["tmp_qkg"] * (decimal)y["mov_cos"];
                //        else
                //            x["tmp_acq"] = (decimal)y["mov_qta"] * (decimal)y["mov_cos"];
                //    }
                //    if ((decimal)y["mov_cos"] == 0)
                //    {
                //        if ((string)x["mov_umi"] == "KG")
                //            x["tmp_cos"] = (decimal)y["mov_acq"] / (decimal)x["tmp_qkg"];
                //        else
                //            x["tmp_cos"] = (decimal)y["mov_acq"] / (decimal)x["tmp_qta"];
                //    }
                //}

                //if ((string)y["fat_tdo"] == "NA")
                //{
                //    x["tmp_qta"] = (decimal)x["tmp_qta"] * -1;
                //    x["tmp_acq"] = (decimal)x["tmp_acq"] * -1;
                //}

                x["tmp_mov"] = y["mov_imp"];
                x["tmp_prv"] = (decimal)y["mov_prv"];

                //if ((string)y["fat_tpd"] == "FV")
                //{
                //    x["tmp_des"] = "Fatt. vendita " + (string)y["fat_ndo"] + " ";
                //    j = tCli.Select("cli_cod='" + y["fat_cfo"] + "'");
                //    if (j.Length > 0)
                //        x["tmp_des"] = (string)x["tmp_des"] + " di " + (string)j[0]["cli_des"];

                //    if (((string)y["mov_ori"]).Trim() == "")
                //    {
                //        dAQt += (decimal)x["tmp_qta"];
                //        dAQk += (decimal)x["tmp_qkg"];
                //        dAcq += (decimal)x["tmp_acq"];
                //    }
                //}
                //else
                //{
                //    x["tmp_des"] = "Fatt. acquisto " + (string)y["fat_ndo"] + " ";
                //    j = tFor.Select("for_cod='" + y["fat_cfo"] + "'");
                //    if (j.Length > 0)
                //        x["tmp_des"] = (string)x["tmp_des"] + " di " + (string)j[0]["for_des"];

                //    dAQt += (decimal)x["tmp_qta"];
                //    dAQk += (decimal)x["tmp_qkg"];
                //    dAcq += (decimal)x["tmp_acq"];
                //}

                x["tmp_des"] = "Mov " + (string)y["tab_des"] + " ";
                j = tCli.Select("cli_cod='" + y["mot_cfo"] + "'");
                if (j.Length > 0)
                    x["tmp_des"] = (string)x["tmp_des"] + " di " + (string)j[0]["cli_des"];

                //if (((string)y["mov_sgm"]).Trim() == "")
                //{
                //}

                decimal dSgn = 0;
                //if (_clsFun.Numerico(y["tab_sgm"]))
                //    dSgn = Convert.ToDecimal(y["tab_sgm"]);

                if ((string)y["tab_sgm"] == "-")
                    dSgn = -1;
                else if ((string)y["tab_sgm"] == "+")
                    dSgn = 1;

                if (dSgn != 0)
                {
                    if ((decimal)x["tmp_qta"] != 0)
                        dMQt += (decimal)x["tmp_qta"] * dSgn;
                    if ((decimal)x["tmp_qkg"] != 0)
                        dMQk += (decimal)x["tmp_qkg"] * dSgn;
                    if ((decimal)x["tmp_acq"] != 0)
                        dMov += (decimal)x["tmp_mov"] * dSgn;
                }

                t.Rows.Add(x);
            }

            /*********/

            dgv1.DataSource = t;

            lblIQt.Text = dIQt.ToString("###,##0.00");
            lblIQk.Text = dIQk.ToString("###,##0.00");
            lblIVa.Text = dIVa.ToString("###,##0.00");

            lblAQt.Text = dAQt.ToString("###,##0.00");
            lblAQk.Text = dAQk.ToString("###,##0.00");
            lblAcq.Text = dAcq.ToString("###,##0.00");

            lblVQt.Text = dVQt.ToString("###,##0.00");
            lblVQk.Text = dVQk.ToString("###,##0.00");
            lblVen.Text = dVen.ToString("###,##0.00");

            lblMQt.Text = dMQt.ToString("###,##0.00");
            lblMQk.Text = dMQk.ToString("###,##0.00");
            lblMov.Text = dMov.ToString("###,##0.00");

            if (_strArtUmi == "KG")
                lblGiaQta.Text = (dIQk + dAQk - dVQk + dMQk).ToString("###,##0.00");
            else
                lblGiaQta.Text = (dIQt + dAQt - dVQt + dMQt).ToString("###,##0.00");
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "tmp_day, tmp_tip", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            string s = "";
            decimal d = 0;

            foreach (DataRowView r in v)
            {
                t.ImportRow(r.Row);
            }

            DataRow x = t.NewRow();
            //x["tmp_day"] = null;
            x["tmp_tip"] = "INV";
            x["tmp_qta"] = Convert.ToDecimal(lblIQt.Text.Replace(".", ","));
            x["tmp_qkg"] = Convert.ToDecimal(lblIQk.Text.Replace(".", ","));
            x["tmp_cos"] = 0;
            x["tmp_acq"] = Convert.ToDecimal(lblIVa.Text.Replace(".", ","));
            x["tmp_prv"] = 0;
            x["tmp_ven"] = 0;
            x["tmp_mov"] = 0;
            x["tmp_des"] = "TOTALE INVENTARIO";
            t.Rows.Add(x);

            x = t.NewRow();
            //x["tmp_day"] = null;
            x["tmp_tip"] = "ACQ";
            x["tmp_qta"] = Convert.ToDecimal(lblAQt.Text.Replace(".", ","));
            x["tmp_qkg"] = Convert.ToDecimal(lblAQk.Text.Replace(".", ","));
            x["tmp_cos"] = 0;
            x["tmp_acq"] = Convert.ToDecimal(lblAcq.Text.Replace(".", ","));
            x["tmp_prv"] = 0;
            x["tmp_ven"] = 0;
            x["tmp_mov"] = 0;
            x["tmp_des"] = "TOTALE ACQUISTI";
            t.Rows.Add(x);

            x = t.NewRow();
            //x["tmp_day"] = null;
            x["tmp_tip"] = "VEN";
            x["tmp_qta"] = Convert.ToDecimal(lblVQt.Text.Replace(".", ","));
            x["tmp_qkg"] = Convert.ToDecimal(lblVQk.Text.Replace(".", ","));
            x["tmp_cos"] = 0;
            x["tmp_acq"] = 0;
            x["tmp_prv"] = 0;
            x["tmp_ven"] = Convert.ToDecimal(lblVen.Text.Replace(".", ","));
            x["tmp_mov"] = 0;
            x["tmp_des"] = "TOTALE VENDITE";
            t.Rows.Add(x);

            x = t.NewRow();
            //x["tmp_day"] = null;
            x["tmp_tip"] = "MOV";
            x["tmp_qta"] = Convert.ToDecimal(lblMQt.Text.Replace(".", ","));
            x["tmp_qkg"] = Convert.ToDecimal(lblMQk.Text.Replace(".", ","));
            x["tmp_cos"] = 0;
            x["tmp_acq"] = 0;
            x["tmp_prv"] = 0;
            x["tmp_ven"] = Convert.ToDecimal(lblMov.Text.Replace(".", ","));
            x["tmp_mov"] = 0;
            x["tmp_des"] = "TOTALE MOVIMENTI";
            t.Rows.Add(x);

            x = t.NewRow();
            //x["tmp_day"] = null;
            x["tmp_tip"] = "GIA";
            x["tmp_qta"] = Convert.ToDecimal(lblGiaQta.Text.Replace(".", ","));
            x["tmp_qkg"] = 0;
            x["tmp_cos"] = 0;
            x["tmp_acq"] = 0;
            x["tmp_prv"] = 0;
            x["tmp_ven"] = 0;
            x["tmp_mov"] = 0;
            x["tmp_des"] = "TOTALE GIACENZA";
            t.Rows.Add(x);

            //string sTit = "Giacenza - etrazione dati dal " + dtpIni.Value.ToString("dd/MM/yyyy") + " al " + DateTime.Now.ToString("dd/MM/yyyy");
            string sTit = this.Text + " al " + DateTime.Now.ToString("dd/MM/yyyy");
            string sFoo = "";
            //sFoo += "'Totali',";
            //sFoo += ",,,,,,,,," + lblCos.Text.Replace(".", "").Replace(",", ".");
            string sFil = "";

            string sFld = "";
            sFld += "tmp_day, Data, 110, DateTime;";
            sFld += "tmp_tip, Tipo, 40, StringLiteral;";
            sFld += "tmp_qta, Q.ta, 45, Decimal;";
            sFld += "tmp_qkg, Peso, 45, Decimal;";
            sFld += "tmp_cos, Costo, 45, Decimal;";
            sFld += "tmp_acq, Acquisti, 45, Decimal;";
            sFld += "tmp_prv, Prezzo, 45, Decimal;";
            sFld += "tmp_ven, Vendite, 45, Decimal;";
            sFld += "tmp_mov, Movimenti, 45, Decimal;";
            sFld += "tmp_des, Descrizione, 200, StringLiteral;";

            sFil = "ArtGiacenza.xls";

            (new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);
        }

    }
}
