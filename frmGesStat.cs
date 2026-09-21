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
    public partial class frmGesStat : Form
    {
        private const string TABSTAVET = "GesNegVet";
        private const string TABSTAVEP = "GesNegVep";
        private const string TABSTAVEN = "GesNegVen";
        private const string TABMOVFAT = "GesFatTestate";
        private const string TABTABREP = "TabReparti";

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
       private string _strConSqlSta = "";
        private string _strStatVisProf = "";

        public string _strArtCod = "";
        public Boolean _bolAuto = false;
        public Boolean _bolPrnAuto = false;
        public DateTime _dayDay = new DateTime(2050, 1, 1, 0, 0, 0);

        public frmGesStat()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
            _strStatVisProf = _clsFun.ParGet(clsDefine.enuParametri.ParStatVisProf, _strConSql);
        }

        private void frmGesStat_Load(object sender, EventArgs e)
        {
            if (_bolPrnAuto)
            {
                DataTable t = FillVenduto();
                FillPrn("ven_rep", t);
                Esci();
            }
            else
            {
                SetDgv1();
                SetDgv2();

                if (_bolAuto)
                {
                    if (_dayDay != _clsDef.DAYOUT)
                        dtpIni.Value = dtpFin.Value = _dayDay;
                    else if(_strArtCod != "")
                        dtpIni.Value = new DateTime(dtpIni.Value.Year, 1, 1);

                    Estrai();

                    //FillVenduto();
                    //FillFatture();
                    //FillMovimenti();
                }

                dtpIni.Select();
            }
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private void frmGesStat_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void btnEstrai_Click(object sender, EventArgs e)
        {
            Estrai();
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
            cTbc.DataPropertyName = "ven_neg";
            cTbc.Name = "Negozio";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VetCau";
            cTbc.Name = "Causale";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_day";
            cTbc.Name = "Data";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_ora";
            cTbc.Name = "Ora";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_pos";
            cTbc.Name = "Cassa";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_sco";
            cTbc.Name = "Scontrino";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VetFid";
            cTbc.Name = "Tessera";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_ean";
            cTbc.Name = "Ean";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_ard";
            cTbc.Name = "Descrizione";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_iva";
            cTbc.Name = "IVA";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_umi";
            cTbc.Name = "UM";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_rep";
            cTbc.Name = "Reparto";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_qta";
            cTbc.Name = "Q.tà";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_prz";
            cTbc.Name = "Prezzo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_sct";
            cTbc.Name = "T.variazione";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 40;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_scn";
            cTbc.Name = "Sconto";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_ven";
            cTbc.Name = "Importo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_oft";
            cTbc.Name = "Tipo Offerta";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_off";
            cTbc.Name = "Codice Offerta";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);
        }

        private void SetDgv2()
        {
            dgv2.AutoGenerateColumns = false;
            //dgv2.VirtualMode = true;
            //dgv2.Dock = DockStyle.Fill;
            dgv2.AllowUserToAddRows = false;
            dgv2.ReadOnly = false;
            dgv2.AllowUserToDeleteRows = false;
            //dgv2.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_art";
            cTbc.Name = "Codice articolo";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_ard";
            cTbc.Name = "Descrizione";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_rep";
            cTbc.Name = "Rep.";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_red";
            cTbc.Name = "Reparto";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_qve";
            cTbc.Name = "Q.tà vendite";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_qfa";
            cTbc.Name = "Q.tà fatture";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_qmo";
            cTbc.Name = "Q.tà Movimenti";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_qta";
            cTbc.Name = "Q.tà";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_vve";
            cTbc.Name = "Valore venduto";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_vfa";
            cTbc.Name = "Valore fatture";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_vmo";
            cTbc.Name = "Valore movimenti";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_val";
            cTbc.Name = "Valore";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);
        }

        private DataTable OldFillVenduto()
        {
            Color bc = lblMsg.BackColor;
            lblMsg.BackColor = Color.LightCoral;
            lblMsg.Text = "Caricamento in corso ...";
            System.Windows.Forms.Application.DoEvents();

            DataRow[] j;
            DataRow x;

            string s = "";

            s = "SELECT * FROM (";
            s += "SELECT ";
            s += "ROW_NUMBER() OVER (PARTITION BY ean_art ORDER BY ean_art, ean_dtm DESC) AS ROW, ";
            s += "ean_art, ";
            s += "ean_ean  ";
            s += "FROM AnaBarcode) AS A WHERE ROW = 1";
            DataTable tEan = _clsFun.FillTabSql("EAN", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tEan.Columns["ean_art"];
            tEan.PrimaryKey = keys;

            s = "SELECT * FROM (";
            s += "SELECT ";
            s += "ROW_NUMBER() OVER (PARTITION BY lia_art ORDER BY lia_art, lia_dti DESC) AS ROW, ";
            s += "lia_art, ";
            s += "lia_cos ";
            s += "FROM GesLisAcquisto ) AS A WHERE ROW = 1";
            DataTable tPco = _clsFun.FillTabSql("PCO", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tPco.Columns["lia_art"];
            tPco.PrimaryKey = keys;

            s = "SELECT * FROM (";
            s += "SELECT ";
            s += "ROW_NUMBER() OVER (PARTITION BY liv_art ORDER BY liv_art, liv_dti DESC) AS ROW, ";
            s += "liv_art, ";
            s += "liv_prv ";
            s += "FROM GesLisVendita ) AS A WHERE ROW = 1";
            DataTable tPve = _clsFun.FillTabSql("PVE", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tPve.Columns["liv_art"];
            tPve.PrimaryKey = keys;

            s = "SELECT * ";
            s += "FROM GesNegVet INNER JOIN GesNegVen ON ";
            s += "GesNegVet.vet_neg = GesNegVen.ven_neg AND ";
            s += "GesNegVet.vet_cau = GesNegVen.ven_cau AND ";
            s += "GesNegVet.vet_day = GesNegVen.ven_day AND ";
            s += "GesNegVet.vet_ora = GesNegVen.ven_ora AND ";
            s += "GesNegVet.vet_pos = GesNegVen.ven_pos AND ";
            s += "GesNegVet.vet_sco = GesNegVen.ven_sco ";
            s += "WHERE ";
            s += "GesNegVet.vet_day >= " + _clsFun.DaySql(dtpIni.Value) + " AND ";
            s += "GesNegVet.vet_day <= " + _clsFun.DaySql(dtpFin.Value) + " ";
            if (_strArtCod != "")
                s += " AND GesNegVen.ven_art='" + _strArtCod + "' ";
            if (_strStatVisProf == "N")
                s += " AND GesNegVet.vet_cau<>'PRO' ";
            s += "ORDER BY ";
            s += "GesNegVen.ven_day, ";
            s += "GesNegVen.ven_pos, ";
            s += "GesNegVen.ven_sco";
            DataTable t = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlSta);

            DataTable tArt = new clsGenTabTmp().TabTmpArtVenduto("VenArt");
            keys = new DataColumn[1];
            keys[0] = tArt.Columns["tmp_art"];
            tArt.PrimaryKey = keys;

            decimal d = 0;

            progressBar1.Value = 0;
            progressBar1.Maximum = (int)t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in t.Rows)
            {
                if ((string)y["vet_sco"] == "00437")
                    Console.WriteLine("aaaaaaa");

                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                j = tArt.Select("tmp_art='" + y["ven_art"] + "'");
                if(j.Length == 0)
                {
                    x = tArt.NewRow();
                    x["tmp_art"] = y["ven_art"];
                    x["tmp_ard"] = y["ven_ard"];
                    x["tmp_rep"] = y["ven_rep"];

                    x["tmp_ean"] = "";
                    x["tmp_pco"] = 0;
                    x["tmp_pve"] = 0;
                    x["tmp_qve"] = 0;
                    x["tmp_qfa"] = 0;
                    x["tmp_qmo"] = 0;
                    x["tmp_vve"] = 0;
                    x["tmp_vfa"] = 0;
                    x["tmp_vmo"] = 0;

                    j = tEan.Select("ean_art='" + y["ven_art"] + "'");
                    if (j.Length > 0)
                        x["tmp_ean"] = (string)j[0]["ean_ean"];

                    j = tPco.Select("lia_art='" + y["ven_art"] + "'");
                    if (j.Length > 0)
                        x["tmp_pco"] = (decimal)j[0]["lia_cos"];

                    j = tPve.Select("liv_art='" + y["ven_art"] + "'");
                    if (j.Length > 0)
                        x["tmp_pve"] = (decimal)j[0]["liv_prv"]; 

                    tArt.Rows.Add(x);
                }

                j = tArt.Select("tmp_art='" + y["ven_art"] + "'");

                if (((string)y["ven_sct"]).Trim() == "RES")
                {
                    Console.WriteLine("aa");
                    d -= (decimal)y["ven_ven"];
                    j[0]["tmp_qve"] = (decimal)j[0]["tmp_qve"] - (decimal)y["ven_qta"];
                    j[0]["tmp_vve"] = (decimal)j[0]["tmp_vve"] - (decimal)y["ven_ven"];
                }
                else
                {
                    d += (decimal)y["ven_ven"];
                    j[0]["tmp_qve"] = (decimal)j[0]["tmp_qve"] + (decimal)y["ven_qta"];
                    j[0]["tmp_vve"] = (decimal)j[0]["tmp_vve"] + (decimal)y["ven_ven"];
                }
            }

            if (!_bolPrnAuto)
            {
                dgv1.DataSource = t;
                dgv2.DataSource = tArt;

                lblMsg.BackColor = bc;
                lblMsg.Text = "";
            }

            s = "SELECT ";
            s += "SUM(vep_imp) AS VepImp ";
            s += "FROM GesNegVep ";
            s += "WHERE ";
            s += "GesNegVep.vep_day >= " + _clsFun.DaySql(dtpIni.Value) + " AND ";
            s += "GesNegVep.vep_day <= " + _clsFun.DaySql(dtpFin.Value) + " AND ";
            s += "(vep_tip = 'SCP' OR vep_tip = 'SCV') ";
            DataTable tTot = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlSta);
            if (tTot.Rows.Count > 0)
            {
                decimal d2 = Convert.ToDecimal(tTot.Rows[0]["VepImp"]);
                lblTotSct.Text = d2.ToString("#,###,##0.00");
                d = d - d2;

                _clsFun.FileLog("SCT", d2.ToString(), "SCT");
            }

            lblTotVen.Text = d.ToString("#,###,##0.00");

            return t;
        }

        private void Estrai()
        {
            FillVenduto();
            FillFatture();
            FillMovimenti();
        }

        private DataTable FillVenduto()
        {
            Color bc = lblMsg.BackColor;
            lblMsg.BackColor = Color.LightCoral;
            lblMsg.Text = "Caricamento in corso ...";
            System.Windows.Forms.Application.DoEvents();

            DataRow x;
            DataRow[] j;
            DataRow[] j2;

            string s = "";

            s = "SELECT * FROM TabReparti";
            DataTable tRep = _clsFun.FillTabSql(TABTABREP, s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tRep.Columns["tab_cod"];
            tRep.PrimaryKey = keys;

            s = "SELECT art_cod, art_rep FROM AnaArticoli";
            DataTable tAar = _clsFun.FillTabSql("ART", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tAar.Columns["art_cod"];
            tAar.PrimaryKey = keys;

            s = "SELECT * FROM (";
            s += "SELECT ";
            s += "ROW_NUMBER() OVER (PARTITION BY ean_art ORDER BY ean_art, ean_dtm DESC) AS ROW, ";
            s += "ean_art, ";
            s += "ean_ean  ";
            s += "FROM AnaBarcode) AS A WHERE ROW = 1";
            DataTable tEan = _clsFun.FillTabSql("EAN", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tEan.Columns["ean_art"];
            tEan.PrimaryKey = keys;

            s = "SELECT * FROM (";
            s += "SELECT ";
            s += "ROW_NUMBER() OVER (PARTITION BY lia_art ORDER BY lia_art, lia_dti DESC) AS ROW, ";
            s += "lia_art, ";
            s += "lia_cos ";
            s += "FROM GesLisAcquisto) AS A WHERE ROW = 1";
            DataTable tPco = _clsFun.FillTabSql("PCO", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tPco.Columns["lia_art"];
            tPco.PrimaryKey = keys;

            s = "SELECT * FROM (";
            s += "SELECT ";
            s += "ROW_NUMBER() OVER (PARTITION BY liv_art ORDER BY liv_art, liv_dti DESC) AS ROW, ";
            s += "liv_art, ";
            s += "liv_prv ";
            s += "FROM GesLisVendita) AS A WHERE ROW = 1";
            DataTable tPve = _clsFun.FillTabSql("PVE", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tPve.Columns["liv_art"];
            tPve.PrimaryKey = keys;

            s = "SELECT * ";
            s += "FROM GesNegVep ";
            s += "WHERE ";
            s += "GesNegVep.vep_day >= " + _clsFun.DaySql(dtpIni.Value) + " AND ";
            s += "GesNegVep.vep_day <= " + _clsFun.DaySql(dtpFin.Value) + " ";
            //s += "AND (vep_tip = 'SCT' OR vep_tip = 'SCP' OR vep_tip = 'SCV') ";
            DataTable tVep = _clsFun.FillTabSql(TABSTAVEP, s, false, _strConSqlSta);
            keys = new DataColumn[5];
            keys[0] = tPve.Columns["vep_day"];
            keys[1] = tPve.Columns["vep_ora"];
            keys[2] = tPve.Columns["vep_pos"];
            keys[3] = tPve.Columns["vep_sco"];
            keys[4] = tPve.Columns["vep_ean"];
            tVep.PrimaryKey = keys;

            //s = "SELECT * ";
            //s += "FROM GesNegVet INNER JOIN GesNegVen ON ";
            //s += "GesNegVet.vet_neg = GesNegVen.ven_neg AND ";
            //s += "GesNegVet.vet_cau = GesNegVen.ven_cau AND ";
            //s += "GesNegVet.vet_day = GesNegVen.ven_day AND ";
            //s += "GesNegVet.vet_ora = GesNegVen.ven_ora AND ";
            //s += "GesNegVet.vet_pos = GesNegVen.ven_pos AND ";
            //s += "GesNegVet.vet_sco = GesNegVen.ven_sco ";
            //s += "WHERE ";
            //s += "GesNegVet.vet_day >= " + _clsFun.DaySql(dtpIni.Value) + " AND ";
            //s += "GesNegVet.vet_day <= " + _clsFun.DaySql(dtpFin.Value) + " ";
            //if (_strArtCod != "")
            //    s += " AND GesNegVen.ven_art='" + _strArtCod + "' ";
            //if (_strStatVisProf == "N")
            //    s += " AND GesNegVet.vet_cau<>'PRO' ";
            //s += "ORDER BY ";
            //s += "GesNegVen.ven_day, ";
            //s += "GesNegVen.ven_pos, ";
            //s += "GesNegVen.ven_sco";
            //DataTable t = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlSta);

            s = "SELECT * ";
            s += "FROM GesNegVen ";
            s += "WHERE ";
            s += "GesNegVen.ven_day >= " + _clsFun.DaySql(dtpIni.Value) + " AND ";
            s += "GesNegVen.ven_day <= " + _clsFun.DaySql(dtpFin.Value) + " ";
            if (_strArtCod != "")
                s += " AND GesNegVen.ven_art='" + _strArtCod + "' ";
            if (_strStatVisProf == "N")
                s += " AND GesNegVen.ven_cau<>'PRO' ";

            DataTable tVen = _clsFun.FillTabSql(TABSTAVEN, s, false, _strConSqlSta);
            keys = new DataColumn[7];
            keys[0] = tPve.Columns["ven_day"];
            keys[1] = tPve.Columns["ven_ora"];
            keys[2] = tPve.Columns["ven_pos"];
            keys[3] = tPve.Columns["ven_sco"];
            keys[4] = tPve.Columns["ven_sct"];
            keys[5] = tPve.Columns["ven_ean"];
            keys[6] = tPve.Columns["ven_art"];
            tVen.PrimaryKey = keys;

            tVen.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "VetCau",
                Caption = "Vendite",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            tVen.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "VetFid",
                Caption = "Fidelity",
                MaxLength = 13,
                ReadOnly = false,
                DefaultValue = (string)""
            });

            s = "SELECT * ";
            s += "FROM GesNegVet ";
            s += "WHERE ";
            s += "GesNegVet.vet_day >= " + _clsFun.DaySql(dtpIni.Value) + " AND ";
            s += "GesNegVet.vet_day <= " + _clsFun.DaySql(dtpFin.Value) + " ";

            //s += " AND GesNegVet.vet_cau='SCO' AND vet_sco='00001' AND vet_pos='01' ";

            s += "ORDER BY ";
            s += "GesNegVet.vet_day, ";
            s += "GesNegVet.vet_pos, ";
            s += "GesNegVet.vet_sco";
            DataTable tVet = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlSta);

            DataTable tArt = new clsGenTabTmp().TabTmpArtVenduto("VenArt");
            keys = new DataColumn[1];
            keys[0] = tArt.Columns["tmp_art"];
            tArt.PrimaryKey = keys;

            string sSql = "";
            decimal dTot = 0;
            decimal dPro = 0;
            decimal dPag = 0;
            decimal d = 0;
            decimal dValScon = 0;

            progressBar1.Value = 0;
            progressBar1.Maximum = (int)tVet.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in tVet.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if ((string)y["vet_cau"] == "PRO" && (string)y["vet_pos"] == "02" && (string)y["vet_sco"] == "00018")
                    Console.WriteLine("aaaa");

                dValScon = 0;

                if ((string)y["vet_cau"] == "SCO")
                    dTot += (decimal)y["vet_imp"];
                else if ((string)y["vet_cau"] == "PRO")
                    dPro += (decimal)y["vet_imp"];

                sSql = "vep_day=" + _clsFun.DayMdb((DateTime)y["vet_day"]) + " AND ";
                sSql += "vep_cau='" + y["vet_cau"] + "' AND ";
                sSql += "vep_ora='" + y["vet_ora"] + "' AND ";
                sSql += "vep_pos='" + y["vet_pos"] + "' AND ";
                sSql += "vep_sco='" + y["vet_sco"] + "'";

                j = tVep.Select(sSql);
                if (j.Length > 0)
                {
                    for (int i = 0; i < j.Length; i++)
                    {
                        if ((string)j[i]["vep_tip"] == "SCT" || (string)j[i]["vep_tip"] == "SCP" || (string)j[i]["vep_tip"] == "SCV")
                        {
                            dValScon += (decimal)j[i]["vep_imp"];
                        }
                    }
                }

                sSql = "ven_day=" + _clsFun.DayMdb((DateTime)y["vet_day"]) + " AND ";
                sSql += "ven_cau='" + y["vet_cau"] + "' AND ";
                sSql += "ven_ora='" + y["vet_ora"] + "' AND ";
                sSql += "ven_pos='" + y["vet_pos"] + "' AND ";
                sSql += "ven_sco='" + y["vet_sco"] + "'";
                j = tVen.Select(sSql);
                if (j.Length > 0)
                {
                    if (dValScon == 0)
                    {
                        d = 0;
                        for (int i = 0; i < j.Length; i++)
                            d += (decimal)j[i]["ven_ven"];

                        dValScon = d - (decimal)y["vet_imp"];

                    }

                    d = (decimal)y["vet_imp"] + dValScon;

                    decimal dDelta = 0;
                    decimal dd = 0;
                    decimal dRes = 0;

                    sSql = "ven_day=" + _clsFun.DayMdb((DateTime)y["vet_day"]) + " AND ";
                    sSql += "ven_cau='" + y["vet_cau"] + "' AND ";
                    sSql += "ven_ora='" + y["vet_ora"] + "' AND ";
                    sSql += "ven_pos='" + y["vet_pos"] + "' AND ";
                    sSql += "ven_sco='" + y["vet_sco"] + "' AND ";
                    sSql += "ven_sct='RES'";
                    j2 = tVen.Select(sSql);
                    if (j2.Length > 0)
                    {
                        for (int i = 0; i < j2.Length; i++)
                            dRes += (decimal)j[i]["ven_ven"];
                    }

                    d -= dRes;

                    if (dValScon != 0 && d != 0)
                        dDelta = dValScon * 100 / d;

                    Console.WriteLine("aaaa");

                    for (int i = 0; i < j.Length; i++)
                    {
                        j2 = tAar.Select("art_cod='" + j[i]["ven_art"] + "'");
                        if(j2.Length > 0)
                            j[i]["ven_rep"] = j2[0]["art_rep"];

                        d = (decimal)j[i]["ven_ven"];

                        if (dDelta > 0)
                            d = _clsFun.MenoPer(d, dDelta);
                        if ((string)j[i]["ven_sct"] == "RES")
                        {
                            dd -= (decimal)j[i]["ven_ven"];
                            j[i]["ven_ven"] = (decimal)j[i]["ven_ven"] * -1;
                        }
                        else
                        {
                            j[i]["ven_ven"] = d;
                            dd += d;
                        }

                        if (i == j.Length - 1)
                        {
                            if ((decimal)y["vet_imp"] != dd)
                            {
                                d = (decimal)y["vet_imp"] - dd;
                                j[i]["ven_ven"] = (decimal)j[i]["ven_ven"] + d;
                            }
                        }

                        s = ((string)y["vet_cau"]).Trim();
                        if(s== "")
                            Console.WriteLine("aaaaaaaaa");

                        j[i]["VetCau"] = y["vet_cau"];
                        j[i]["VetFid"] = y["vet_fid"];

                        j2 = tArt.Select("tmp_art='" + j[i]["ven_art"] + "'");
                        if (j2.Length == 0)
                        {
                            x = tArt.NewRow();
                            x["tmp_art"] = j[i]["ven_art"];
                            x["tmp_ard"] = j[i]["ven_ard"];
                            x["tmp_rep"] = j[i]["ven_rep"];

                            x["tmp_red"] = "";
                            j2 = tRep.Select("tab_cod='" + x["tmp_rep"] + "'");
                            if (j2.Length > 0)
                                x["tmp_red"] = (string)j2[0]["tab_des"];

                            x["tmp_ean"] = "";
                            x["tmp_pco"] = 0;
                            x["tmp_pve"] = 0;
                            x["tmp_qve"] = 0;
                            x["tmp_qfa"] = 0;
                            x["tmp_qmo"] = 0;
                            x["tmp_vve"] = 0;
                            x["tmp_vfa"] = 0;
                            x["tmp_vmo"] = 0;

                            j2 = tEan.Select("ean_art='" + j[i]["ven_art"] + "'");
                            if (j2.Length > 0)
                                x["tmp_ean"] = (string)j2[0]["ean_ean"];

                            j2 = tPco.Select("lia_art='" + j[i]["ven_art"] + "'");
                            if (j2.Length > 0)
                                x["tmp_pco"] = (decimal)j2[0]["lia_cos"];

                            j2 = tPve.Select("liv_art='" + j[i]["ven_art"] + "'");
                            if (j2.Length > 0)
                                x["tmp_pve"] = (decimal)j2[0]["liv_prv"];

                            tArt.Rows.Add(x);
                        }

                        j2 = tArt.Select("tmp_art='" + j[i]["ven_art"] + "'");

                        if (((string)j[i]["ven_sct"]).Trim() == "RES")
                        {
                            Console.WriteLine("aa");
                            d -= (decimal)j[i]["ven_ven"];
                            j2[0]["tmp_qve"] = (decimal)j2[0]["tmp_qve"] - (decimal)j[i]["ven_qta"];
                            j2[0]["tmp_vve"] = (decimal)j2[0]["tmp_vve"] - (decimal)j[i]["ven_ven"];
                        }
                        else
                        {
                            d += (decimal)j[i]["ven_ven"];
                            j2[0]["tmp_qve"] = (decimal)j2[0]["tmp_qve"] + (decimal)j[i]["ven_qta"];
                            j2[0]["tmp_vve"] = (decimal)j2[0]["tmp_vve"] + (decimal)j[i]["ven_ven"];
                        }
                    }
                }
            }

            if (!_bolPrnAuto)
            {
                dgv1.DataSource = tVen;
                dgv2.DataSource = tArt;

                lblMsg.BackColor = bc;
                lblMsg.Text = "";
            }

            s = "SELECT ";
            s += "SUM(vep_imp) AS VepImp ";
            s += "FROM GesNegVep ";
            s += "WHERE ";
            s += "GesNegVep.vep_day >= " + _clsFun.DaySql(dtpIni.Value) + " AND ";
            s += "GesNegVep.vep_day <= " + _clsFun.DaySql(dtpFin.Value) + " AND ";
            s += "(vep_tip = 'SCP' OR vep_tip = 'SCV') ";
            DataTable tTot = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlSta);
            if (tTot.Rows.Count > 0)
            {
                decimal d2 = Convert.ToDecimal(tTot.Rows[0]["VepImp"]);
                lblTotSct.Text = d2.ToString("#,###,##0.00");
                d = d - d2;

                _clsFun.FileLog("SCT", d2.ToString(), "SCT");
            }

            lblTotVen.Text = dTot.ToString("#,###,##0.00");
            lblTotPro.Text = dPro.ToString("#,###,##0.00");

            return tVen;
        }

        private void FillFatture()
        {
            Color bc = lblMsg.BackColor;
            lblMsg.BackColor = Color.LightCoral;
            lblMsg.Text = "Caricamento in corso ...";
            System.Windows.Forms.Application.DoEvents();

            DataRow[] j;
            DataRow x;

            string s = "";

            s = "SELECT ";
            s += "GesFatTestate.fat_tpd, ";
            s += "GesFatTestate.fat_ann, ";
            s += "GesFatTestate.fat_cfo, ";
            s += "GesMovimenti.mov_rfa, ";
            s += "GesFatTestate.fat_ddo, ";
            s += "GesMovimenti.mov_art, ";
            s += "GesMovimenti.mov_ard, ";
            s += "GesMovimenti.mov_cos, ";
            s += "GesMovimenti.mov_prv, ";
            s += "GesMovimenti.mov_imp, ";
            s += "GesMovimenti.mov_qta ";
            s += "FROM GesFatTestate ";
            s += "LEFT OUTER JOIN GesMovimenti ON GesFatTestate.fat_yfa = GesMovimenti.mov_yfa AND GesFatTestate.fat_nfa = GesMovimenti.mov_nfa ";
            s += "WHERE ";
            s += "GesFatTestate.fat_ddo >= " + _clsFun.DaySql(dtpIni.Value) + " AND ";
            s += "GesFatTestate.fat_ddo <= " + _clsFun.DaySql(dtpFin.Value) + " ";
            if (_strArtCod != "")
                s += " AND GesMovimenti.mov_art='" + _strArtCod + "' ";
            s += "ORDER BY GesMovimenti.mov_art";
            DataTable t = _clsFun.FillTabSql(TABMOVFAT, s, false, _strConSql);

            DataTable tArt = (DataTable)dgv2.DataSource;
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tArt.Columns["tmp_art"];
            tArt.PrimaryKey = keys;

            decimal d = 0; // Convert.ToDecimal(lblTotVen.Text.Replace(",", "."));

            progressBar1.Value = 0;
            progressBar1.Maximum = (int)t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if (((string)y["fat_tpd"]).Trim() == "FA")
                    d += (decimal)y["mov_imp"];
                //else
                //    d += (decimal)y["mov_imp"];

                j = tArt.Select("tmp_art='" + y["mov_art"] + "'");
                if (j.Length == 0)
                {
                    x = tArt.NewRow();
                    x["tmp_art"] = y["mov_art"];
                    x["tmp_ard"] = y["mov_ard"];
                    x["tmp_ean"] = "";
                    x["tmp_pco"] = 0;
                    x["tmp_pve"] = 0;
                    x["tmp_qve"] = 0;
                    x["tmp_qfa"] = 0;
                    x["tmp_qmo"] = 0;
                    x["tmp_vve"] = 0;
                    x["tmp_vfa"] = 0;
                    x["tmp_vmo"] = 0;
                    tArt.Rows.Add(x);
                    j = tArt.Select("tmp_art='" + y["mov_art"] + "'");
                }

                //j = tArt.Select("tmp_art='" + y["mov_art"] + "'");

                if (((string)y["fat_tpd"]).Trim() == "FV")
                    j[0]["tmp_qfa"] = (decimal)j[0]["tmp_qfa"] - (decimal)y["mov_qta"];
                else
                    j[0]["tmp_qfa"] = (decimal)j[0]["tmp_qfa"] + (decimal)y["mov_qta"];


                if (((string)y["fat_tpd"]).Trim() == "FV")
                    j[0]["tmp_vfa"] = (decimal)j[0]["tmp_vfa"] + (decimal)y["mov_imp"];
                else
                    j[0]["tmp_vfa"] = (decimal)j[0]["tmp_vfa"] - (decimal)y["mov_imp"];
            }

            if (!_bolPrnAuto)
            {
                //dgv1.DataSource = t;
                dgv2.DataSource = tArt;

                lblTotAcq.Text = d.ToString("#,###,##0.00");

                lblMsg.BackColor = bc;
                lblMsg.Text = "";
            }
        }

        private void FillMovimenti()
        {
            Color bc = lblMsg.BackColor;
            lblMsg.BackColor = Color.LightCoral;
            lblMsg.Text = "Caricamento in corso ...";
            System.Windows.Forms.Application.DoEvents();

            DataRow[] j;
            DataRow x;

            string s = "";

            s = "SELECT ";
            s += "GesMovTestate.mot_ddo, ";
            s += "GesMovTestate.mot_cau, ";
            s += "GesMovimenti.mov_rfa, ";
            s += "GesMovimenti.mov_art, ";
            s += "GesMovimenti.mov_ard, ";
            s += "GesMovimenti.mov_cos, ";
            s += "GesMovimenti.mov_prv, ";
            s += "GesMovimenti.mov_imp, ";
            s += "GesMovimenti.mov_qta, ";
            s += "TabMovCausali.tab_sgm ";
            s += "FROM GesMovTestate LEFT OUTER JOIN ";
            s += "GesMovimenti ON GesMovTestate.mot_ymo = GesMovimenti.mov_ymo AND GesMovTestate.mot_nmo = GesMovimenti.mov_nmo LEFT OUTER JOIN ";
            s += "TabMovCausali ON GesMovTestate.mot_cau = TabMovCausali.tab_cod ";
            s += "WHERE ";
            s += "(GesMovTestate.mot_ddo >= " + _clsFun.DaySql(dtpIni.Value) + ") AND ";
            s += "(GesMovTestate.mot_ddo <= " + _clsFun.DaySql(dtpFin.Value) + ") ";

            if (_strArtCod != "")
                s += " AND GesMovimenti.mov_art='" + _strArtCod + "' ";
            s += "ORDER BY GesMovimenti.mov_art";

            DataTable t = _clsFun.FillTabSql(TABMOVFAT, s, false, _strConSql);

            DataTable tArt = (DataTable)dgv2.DataSource;
            DataColumn[] keys = new DataColumn[1];
            keys[0] = t.Columns["tmp_art"];
            t.PrimaryKey = keys;

            decimal d = 0;

            progressBar1.Value = 0;
            progressBar1.Maximum = (int)t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if (((string)y["tab_sgm"]).Trim() == "-")
                    d -= (decimal)y["mov_imp"];
                else
                    d += (decimal)y["mov_imp"];

                j = tArt.Select("tmp_art='" + y["mov_art"] + "'");
                if (j.Length == 0)
                {
                    x = tArt.NewRow();
                    x["tmp_art"] = y["mov_art"];
                    x["tmp_ard"] = y["mov_ard"];
                    x["tmp_ean"] = "";
                    x["tmp_pco"] = 0;
                    x["tmp_pve"] = 0;
                    x["tmp_qve"] = 0;
                    x["tmp_qfa"] = 0;
                    x["tmp_qmo"] = 0;
                    x["tmp_vve"] = 0;
                    x["tmp_vfa"] = 0;
                    x["tmp_vmo"] = 0;
                    tArt.Rows.Add(x);

                    j = tArt.Select("tmp_art='" + y["mov_art"] + "'");
                }

                j = tArt.Select("tmp_art='" + y["mov_art"] + "'");

                if (((string)y["tab_sgm"]).Trim() == "-")
                    j[0]["tmp_qmo"] = (decimal)j[0]["tmp_qmo"] - (decimal)y["mov_qta"];
                else
                    j[0]["tmp_qmo"] = (decimal)j[0]["tmp_qmo"] + (decimal)y["mov_qta"];

                if (((string)y["tab_sgm"]).Trim() == "-")
                    j[0]["tmp_vmo"] = (decimal)j[0]["tmp_vmo"] + (decimal)y["mov_imp"];
                else
                    j[0]["tmp_vmo"] = (decimal)j[0]["tmp_vfa"] - (decimal)y["mov_imp"];
            }

            if (!_bolPrnAuto)
            {
                //dgv1.DataSource = t;
                dgv2.DataSource = tArt;

                //lblTotVen.Text = d.ToString("#,###,##0.00");

                lblMsg.BackColor = bc;
                lblMsg.Text = "";
            }
        }

        private void repartoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FillPrn("ven_rep", (DataTable)dgv1.DataSource);
        }

        private void iVAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FillPrn("ven_iva", (DataTable)dgv1.DataSource);
        }

        private void FillPrn(string strFld, DataTable tabSta)
        {
            if (tabSta.Rows.Count == 0)
                MessageBox.Show("Nessun dato estratto!");
            else
            {
                DataRow[] j;
                DataRow x;
                decimal d = 0;
                decimal dSco = 0;
                decimal dQta = 0;
                string s = "";
                string p = "";

                DataTable t = new clsGenTabTmp().TabTmpStaReparto("TabRep");

                if(strFld == "ven_rep")
                    p = "TabReparti";
                else if(strFld == "ven_iva")
                    p = "TabIVA";

                s = "SELECT * FROM " + p;
                DataTable tTmp = _clsFun.FillTabSql(p, s, false, _strConSql);

                string sSco = "";

                ArrayList aScoRep = new ArrayList();

                foreach(DataRow y in tabSta.Rows)
                {
                    if ((string)y["ven_sco"] != sSco)
                    {
                        dSco += 1;
                        sSco = (string)y["ven_sco"];
                        aScoRep.Clear();
                    }

                    j = t.Select("tmp_cod='" + y[strFld] + "'");
                    if (j.Length == 0)
                    {
                        x = t.NewRow();
                        x["tmp_cod"] = y[strFld];
                        x["tmp_des"] = "";

                        j = tTmp.Select("tab_cod='" + (string)y[strFld] + "'");
                        if (j.Length > 0)
                            x["tmp_des"] = j[0]["tab_des"];

                        x["tmp_imp"] = 0;
                        x["tmp_inc"] = 0;
                        t.Rows.Add(x);
                    }
                    j = t.Select("tmp_cod='" + y[strFld] + "'");

                    if (((string)y["ven_sct"]).Trim() == "RES")
                    {
                        d += (decimal)y["ven_ven"];
                        j[0]["tmp_imp"] = (decimal)j[0]["tmp_imp"] + (decimal)y["ven_ven"];
                    }
                    else
                    {
                        d += (decimal)y["ven_ven"];
                        j[0]["tmp_imp"] = (decimal)j[0]["tmp_imp"] + (decimal)y["ven_ven"];

                        if(aScoRep.IndexOf((string)y["ven_sco"]+(string)y["ven_rep"]) < 0)
                        {
                            aScoRep.Add((string)y["ven_sco"] + (string)y["ven_rep"]);
                            j[0]["tmp_sco"] = (decimal)j[0]["tmp_sco"] + 1;
                        }

                        if (!DBNull.Value.Equals(y["ven_qta"]))
                        {
                            j[0]["tmp_qta"] = (decimal)j[0]["tmp_qta"] + (decimal)y["ven_qta"];
                            dQta += (decimal)y["ven_qta"];
                        }
                    }
                }

                DataTable t2 = t.Clone();
                DataView v = new DataView(t, "", "tmp_cod", DataViewRowState.CurrentRows);

                foreach (DataRowView r in v)
                {
                    r["tmp_inc"] = (decimal)r["tmp_imp"] / d * 100;
                    r["tmp_sci"] = (decimal)r["tmp_sco"] / dSco * 100;
                    r["tmp_scm"] = (decimal)r["tmp_imp"] / (decimal)r["tmp_sco"];
                    r["tmp_qti"] = (decimal)r["tmp_qta"] / dQta * 100;
                    if ((decimal)r["tmp_qta"] > 0)
                        r["tmp_qtm"] = (decimal)r["tmp_imp"] / (decimal)r["tmp_qta"];
                    t2.ImportRow(r.Row);
                }
                clsGenPdfSta1 cls = new clsGenPdfSta1();
                cls._dayIni = dtpIni.Value;
                cls._dayFin = dtpFin.Value;
                cls.PrnPdfRep(strFld, t2, dSco, dQta, "", 0);
            }
        }

        private void stampaXArticoloToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void verificaPagamentiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmGesStatCtrlPagamenti f = new frmGesStatCtrlPagamenti();
            f.ShowDialog();
        }

        private void verificaTotaliToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmGesStatCtrlTotali f = new frmGesStatCtrlTotali();
            f.ShowDialog();
        }

        private void verificaTotaliDaDettaglioToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmGesStatCtrlTotVen f = new frmGesStatCtrlTotVen();
            f.ShowDialog();
        }

        private void articoliToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource == null || ((DataTable)dgv1.DataSource).Rows.Count == 0)
                MessageBox.Show("Nessun dato estratto!");
            else
            {
                clsGenPdfSta1 cls = new clsGenPdfSta1();
                cls._dayIni = dtpIni.Value;
                cls._dayFin = dtpFin.Value;
                cls.PrnPdfStaArt((DataTable)dgv2.DataSource);
            }
        }

        private void giacenzaToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 8)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = (string)x["ven_art"];
                    f.ShowDialog();
                }
            }

        }
        private void dgv2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = (string)x["tmp_art"];
                    f.ShowDialog();
                }
            }

        }

    }
}
