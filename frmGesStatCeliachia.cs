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
using System.Diagnostics;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace APOffice
{
    public partial class frmGesStatCeliachia : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public DateTime _dayDti = DateTime.Today;
        public DateTime _dayDtf = DateTime.Today;

        private string _strConSql = "";
        private string _strConSqlSta = "";
        private string _strIvaPar = "022";
        private string _strPar037ParCeliaci = "";

        public frmGesStatCeliachia()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
        }

        private void frmGesStaIVA_Load(object sender, EventArgs e)
        {
            _strIvaPar = _clsFun.ParGet(clsDefine.enuParametri.Par012CodIVAxDefault, _strConSql);
            _strPar037ParCeliaci = _clsFun.ParGet(clsDefine.enuParametri.Par037ParCeliaci, _strConSql);

            if (_strPar037ParCeliaci == "")
                MessageBox.Show("Parametri non impostati!", "CONFIGURAZIONE CELIACI", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            for (int i = DateTime.Today.Year - 4; i < DateTime.Today.Year + 1; i++)
                cmbYea.Items.Add(i.ToString());
            cmbYea.SelectedIndex = 4;

            nudMon.Value = _dayDtf.Month;

            lblMon.Text = String.Format("{0:MMMM}", _dayDtf);

            FillTab();
            SetDgv1();
            SetDgv2();

            dgv2.DataSource = new clsGenTabTmp().TabTmpIvaRiep("TabIva");
        }
        private void frmGesStatIVA_KeyDown(object sender, KeyEventArgs e)
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

        private void nudMon_ValueChanged(object sender, EventArgs e)
        {
            DateTime d = new DateTime(2020, Convert.ToUInt16(nudMon.Value), 1);
            lblMon.Text = String.Format("{0:MMMM}", d);
        }

        private void btnEstrai_Click(object sender, EventArgs e)
        {
            FillDati();
            dgv2.DataSource = FillIva("", "");
        }

        private void cmbNdo_SelectedIndexChanged(object sender, EventArgs e)
        {
            string s = cmbNdo.Text;
            DammiDoc();
        }

        private string DammiDoc()
        {
            string sNdo = "";

            string s = cmbNdo.Text;

            if (s.Length > 1 && _clsFun.Numerico(s.Substring(0,1), "0123456789"))
            {
                sNdo = s.Substring(0, 5);

                lblNdo.Text = sNdo;

                if (s.Length > 15)
                {
                    s = s.Substring(12, 4) + s.Substring(9, 2) + s.Substring(6, 2);

                    dtpDdt.Value = _clsFun.Str2Day(s);

                    nudMon.Value = dtpDdt.Value.Month;
                }
            }
            else
                lblNdo.Text = "";

            return sNdo;
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
            cTbc.DataPropertyName = "vet_cdo";
            cTbc.Name = "Numero documento";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_cdt";
            cTbc.Name = "Data documento";
            cTbc.Width = 75;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_day";
            cTbc.Name = "Data scontrino";
            cTbc.Width = 75;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_pos";
            cTbc.Name = "Cassa";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_sco";
            cTbc.Name = "Scontrino";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "ven_sco";
            //cTbc.Name = "Scontrino";
            //cTbc.Width = 50;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_doc";
            cTbc.Name = "Codice autorizzazione";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_imp";
            cTbc.Name = "Importo scontrino";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_ard";
            cTbc.Name = "Descrizione articolo";
            cTbc.Width = 190;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_qta";
            cTbc.Name = "Q.tà";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_qkg";
            cTbc.Name = "Q.peso";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "ven_cos";
            //cTbc.Name = "Costo";
            //cTbc.Width = 60;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //cTbc.DefaultCellStyle.Format = "###,##0.000";
            //cTbc.ReadOnly = false;
            //dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_ven";
            cTbc.Name = "Importo per riga";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "tmp_mav";
            //cTbc.Name = "Margine Vendita";
            //cTbc.Width = 60;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //cTbc.DefaultCellStyle.Format = "##0.00";
            //cTbc.ReadOnly = false;
            //dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "ven_sco";
            //cTbc.Name = "Scontrino";
            //cTbc.Width = 50;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "ven_ora";
            //cTbc.Name = "Ora";
            //cTbc.Width = 40;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "ven_pos";
            //cTbc.Name = "Cassa";
            //cTbc.Width = 50;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "ven_day";
            //cTbc.Name = "Data";
            //cTbc.Width = 70;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "ven_iva";
            //cTbc.Name = "IVA";
            //cTbc.Width = 50;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "ven_sct";
            //cTbc.Name = "Tipo";
            //cTbc.Width = 50;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_can";
            cTbc.Name = "Stornato";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_ean";
            cTbc.Name = "Barcode";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_iva";
            cTbc.Name = "IVA";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
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
            cTbc.DataPropertyName = "iva_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_ali";
            cTbc.Name = "Aliquota";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_imp";
            cTbc.Name = "Imponbile";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_iva";
            cTbc.Name = "Imposta";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_tot";
            cTbc.Name = "Totale";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "iva_arr";
            cTbc.Name = "Arr";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "0.000";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);
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

            p = "GesNegCel";
            s = "SELECT *, cel_ndo +' '+ CONVERT(varchar, cel_ddo, 103) AS CelDoc FROM GesNegCel WHERE cel_yea='" + cmbYea.Text + "' ORDER BY cel_mon";
            t = _clsFun.FillTabSql(p, s, false, _strConSqlSta);
            x = t.NewRow();
            x["cel_ndo"] = "";
            x["CelDoc"] = "  Nuovo";
            t.Rows.InsertAt(x, 0);
            cmbNdo.DataSource = t;
            cmbNdo.DisplayMember = "CelDoc";
            cmbNdo.ValueMember = "cel_ndo";
            cmbNdo.SelectedIndex = 0;
        }

        private void FillDati()
        {
            string s = "";
            DataRow[] j;
            string sScoKey = "";
            decimal d = 0;
            decimal dTotVen = 0;
            string sNdo = lblNdo.Text;

            string sNeg = cmbNeg.SelectedValue.ToString();
            DateTime dDti = new DateTime(Convert.ToInt16(cmbYea.Text), Convert.ToInt16(nudMon.Value), 1);
            DateTime dDtf = _clsFun.FineMese(dDti);

            s = "SELECT * FROM TabReparti";
            DataTable tRep = _clsFun.FillTabSql("TabReparti", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tRep.Columns["tab_cod"];
            tRep.PrimaryKey = keys;

            s = "SELECT AnaArticoli.art_cod, AnaArticoli.art_des, AnaArticoli.art_iva, AnaArticoli.art_sfr, AnaArticoli.art_sta, AnaArticoli.art_rep, TabReparti.tab_des AS RepDes, TabEcrLv3.tab_des AS EcrDes ";
            s += "FROM AnaArticoli ";
            s += "LEFT OUTER JOIN TabReparti ON AnaArticoli.art_rep = TabReparti.tab_cod ";
            s += "LEFT OUTER JOIN TabEcrLv3 ON AnaArticoli.art_ec1 = TabEcrLv3.tab_lv1 AND AnaArticoli.art_ec2 = TabEcrLv3.tab_lv2 AND AnaArticoli.art_ec3 = TabEcrLv3.tab_cod ";
            DataTable tArt = _clsFun.FillTabSql("Art", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            s = "SELECT ";
            s += "ven_neg, ";
            s += "ven_cau, ";
            s += "ven_day, ";
            s += "ven_ora, ";
            s += "ven_pos, ";
            s += "ven_sco, ";
            s += "ven_art, ";
            s += "ven_ard, ";
            s += "ven_iva, ";
            s += "ven_rep, ";
            s += "ven_sct, ";
            s += "ven_ean, ";
            s += "ven_umi, ";
            s += "ven_qta, ";
            s += "ven_qkg, ";
            s += "ven_prz, ";
            s += "ven_ven, ";
            //s += "CASE WHEN ven_qkg > 0 THEN ven_qkg ELSE ven_qta END * ven_cos AS ven_cos, ";
            s += "vet_doc, ";
            s += "vet_imp, ";
            s += "vet_sct, ";
            s += "vet_cel, ";
            s += "vet_cdo, ";
            s += "vet_cdt, ";
            s += "vet_can ";
            s += "FROM GesNegVen JOIN GesNegVet ON ";
            s += "GesNegVet.vet_neg = GesNegVen.ven_neg AND ";
            s += "GesNegVet.vet_cau = GesNegVen.ven_cau AND ";
            s += "GesNegVet.vet_day = GesNegVen.ven_day AND ";
            s += "GesNegVet.vet_ora = GesNegVen.ven_ora AND ";
            s += "GesNegVet.vet_pos = GesNegVen.ven_pos AND ";
            s += "GesNegVet.vet_sco = GesNegVen.ven_sco ";
            s += "WHERE (ven_day >= " + _clsFun.DaySql(dDti) + ") AND (ven_day <= " + _clsFun.DaySql(dDtf) + ") AND ";
            s += "(GesNegVen.ven_cau = 'SCO') AND ";
            s += "(GesNegVen.ven_cel = 1) ";
            if (sNeg != "")
                s += " AND ven_neg='" + sNeg + "' ";

            if (sNdo != "")             //Filtro su documento già estratto
                s += " AND vet_cdo='" + sNdo + "' ";
            else
                s += " AND vet_cdo='' ";
            //if (_strStatVisProf == "N")
            //    s += " AND GesNegVet.vet_cau<>'PRO' ";

            //s += " AND GesNegVet.vet_cau='SCO' AND vet_sco='00001' AND vet_pos='01' ";
            //s += "ORDER BY ven_sco";
            s += "ORDER BY ven_day,ven_cau,ven_neg,ven_pos,ven_ora,ven_sco";

            DataTable t = _clsFun.FillTabSql("Sta", s, false, _strConSqlSta);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "VenNiv",
                Caption = "Vendita netto IVA",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            keys = new DataColumn[9];
            keys[0] = t.Columns["ven_day"];
            keys[1] = t.Columns["ven_cau"];
            keys[2] = t.Columns["ven_neg"];
            keys[3] = t.Columns["ven_pos"];
            keys[4] = t.Columns["ven_ora"];
            keys[5] = t.Columns["ven_sco"];
            keys[6] = t.Columns["ven_sct"];
            keys[7] = t.Columns["ven_ean"];
            keys[8] = t.Columns["ven_art"];
            t.PrimaryKey = keys;

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ard",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_imp",
                Caption = "Importo per scontrino",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_mav",
                Caption = "Margine",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_iva",
                Caption = "IVA",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_sta",
                Caption = "Stato",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            decimal dScoImp = 0;

            foreach (DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if ((string)y["ven_cau"] == "SCO" && (string)y["ven_pos"] == "01" && (string)y["ven_sco"] == "00033")
                    Console.WriteLine("aaaa");

                s = ((DateTime)y["ven_day"]).ToString("yyyyMMdd") + (string)y["ven_cau"] + (string)y["ven_neg"] + (string)y["ven_pos"] + (string)y["ven_sco"];
                if (s != sScoKey)
                {
                    sScoKey = s;
                    dScoImp = 0;

                    decimal dScn = (decimal)y["vet_sct"];
                    decimal dLor = dScn + (decimal)y["vet_imp"];
                    decimal dDelta = 0;
                    if (dScn > 0 && dLor > 0)
                        dDelta = dScn * 100 / dLor;
                    decimal dd = 0;

                    s = "ven_neg = '" + (string)y["ven_neg"] + "' AND ";
                    s += "ven_cau = '" + (string)y["ven_cau"] + "' AND ";
                    s += "ven_day = " + _clsFun.DayMdb((DateTime)y["ven_day"]) + " AND ";
                    s += "ven_ora = '" + (string)y["ven_ora"] + "' AND ";
                    s += "ven_pos = '" + (string)y["ven_pos"] + "' AND ";
                    s += "ven_sco = '" + (string)y["ven_sco"] + "'";
                    j = t.Select(s);

                    if ((decimal)y["vet_sct"] > 0)
                    {
                        dScn = (decimal)y["vet_sct"];
                        dLor = dScn + (decimal)y["vet_imp"];
                        dDelta = dScn * 100 / dLor;
                        dd = 0;
                    }
                    /* 20171026 Seck serve nel caso ci sia il totale scontrino diverso dal dettaglio */
                    else
                    {
                        d = 0;
                        for (int i = 0; i < j.Length; i++)
                        {
                            if ((string)j[i]["ven_sct"] == "RES")
                                d -= (decimal)j[i]["ven_ven"];
                            else
                                d += (decimal)j[i]["ven_ven"];
                        }

                        if (d != (decimal)y["vet_imp"])
                        {
                            dScn = d - (decimal)y["vet_imp"];
                            dLor = dScn + (decimal)y["vet_imp"];
                            dDelta = dScn * 100 / dLor;
                            dd = 0;
                        }
                    }
                    /**/

                    if (j.Length > 0)
                    {

                        for (int i = 0; i < j.Length; i++)
                        {
                            if (dDelta > 0)
                            {
                                d = _clsFun.MenoPer((decimal)j[i]["ven_ven"], dDelta);

                                j[i]["ven_ven"] = d;

                                dd += d;

                                if (i == j.Length - 1)
                                {
                                    if (dd != (decimal)y["vet_imp"])
                                    {
                                        d = (decimal)y["vet_imp"] - dd;
                                        j[i]["ven_ven"] = (decimal)j[i]["ven_ven"] + d;
                                    }
                                }
                            }

                            dScoImp += (decimal)j[i]["ven_ven"];
                        }
                    }

                }

                y["tmp_imp"] = dScoImp;

                j = tArt.Select("art_cod='" + y["ven_art"] + "'");
                if (j.Length > 0)
                {
                    //y["tmp_rep"] = (string)j[0]["art_rep"];
                    y["tmp_ard"] = (string)j[0]["art_des"];
                    //y["tmp_red"] = (string)j[0]["RepDes"];
                    //y["tmp_ec3"] = (string)j[0]["EcrDes"];
                    y["tmp_sta"] = (string)j[0]["art_sta"];

                    //decimal dIva = 0;
                    //if (_clsFun.Numerico(j[0]["art_iva"], "0123456789"))
                    //    dIva = Convert.ToDecimal(j[0]["art_iva"]);

                    //y["tmp_mav"] = _clsFun.Margine((decimal)y["ven_ven"], (decimal)y["ven_cos"], dIva, (decimal)j[0]["art_sfr"], "P");
                    //y["tmp_iva"] = dIva;
                }
                else
                {
                    //y["tmp_rep"] = (string)y["ven_rep"];
                    y["tmp_ard"] = (string)y["ven_ard"];
                    //y["tmp_red"] = "Non definito";
                    //j = tRep.Select("tab_cod='" + (string)y["ven_rep"] + "'");
                    //if (j.Length > 0)
                    //    y["tmp_red"] = (string)j[0]["tab_des"];
                    //y["tmp_ec3"] = "";
                }

                if ((string)y["ven_sct"] == "RES")
                    dTotVen -= (decimal)y["ven_ven"];
                else
                    dTotVen += (decimal)y["ven_ven"];



            }

            //if (sRep != "")
            //{
            //    dTotVen = 0;
            //    DataTable t2 = t.Clone();
            //    foreach (DataRow y in t.Rows)
            //    {
            //        if ((string)y["tmp_rep"] == sRep)
            //        {
            //            t2.ImportRow(y);
            //            dTotVen += (decimal)y["ven_ven"];
            //        }
            //    }
            //    t = t2.Copy();
            //}

            /* Da togliere inizio fine
            t = FillArfAlba(t);
            */

            dgv1.DataSource = t;

            lblTot.Text = dTotVen.ToString("#,##0.00");
        }

        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string sNeg = cmbNeg.SelectedValue.ToString();

            //DataTable tIva = ((DataView)dgv1.DataSource).ToTable();
            //PrnPdfIva("ven_iva", tIva, sNeg);
        }

        private void btnInvio_Click(object sender, EventArgs e)
        {
            if (lblNdo.Text == "")
                MessageBox.Show("Documento non selezionato!", "GENERAZIONE RENDICONTAZIONE MENSILE", MessageBoxButtons.OK, MessageBoxIcon.Question);
            else
            {
                if (dgv1.DataSource == null || ((DataTable)dgv1.DataSource).Rows.Count == 0)
                    MessageBox.Show("Non sono presenti righe valide!", "GENERAZIONE RENDICONTAZIONE MENSILE", MessageBoxButtons.OK, MessageBoxIcon.Question);
                else
                {
                    DataView v = new DataView((DataTable)dgv1.DataSource, "", "ven_day, ven_pos, ven_sco", DataViewRowState.CurrentRows);
                    if (v.Count == 0)
                        MessageBox.Show("Non sono presenti righe valide!", "GENERAZIONE RENDICONTAZIONE MENSILE", MessageBoxButtons.OK, MessageBoxIcon.Question);
                    else if (MessageBox.Show("Confermi la generazione della rendicontazione di " + lblMon.Text + "/" + cmbYea.Text + "?", "GENERAZIONE RENDICONTAZIONE MENSILE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        InvioRendicontazione();
                    }
                }
            }
        }

        private void btnDoc_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource == null || ((DataTable)dgv1.DataSource).Rows.Count == 0)
                MessageBox.Show("Non sono presenti righe valide!", "GENERAZIONE DOCUMENTO", MessageBoxButtons.OK, MessageBoxIcon.Question);
            else
            {
                DataView v = new DataView((DataTable)dgv1.DataSource, "", "ven_day, ven_pos, ven_sco", DataViewRowState.CurrentRows);
                if (v.Count == 0)
                    MessageBox.Show("Non sono presenti righe valide!", "GENERAZIONE DOCUMENTO", MessageBoxButtons.OK, MessageBoxIcon.Question);
                else if (MessageBox.Show("Confermi la generazione del documento di " + lblMon.Text + "/" + cmbYea.Text + "?", "GENERAZIONE DOCUMENTO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    GenRendicontazione(v);
                }
            }
        }

        private void GenRendicontazione(DataView dtvCel)
        {
            string s = "";

            string sYea = cmbYea.Text;
            string sMon = nudMon.Value.ToString().PadLeft(2,Convert.ToChar('0'));
            string sNdo = "";
            DateTime dDdo = dtpDdt.Value;

            s = "SELECT * FROM GesNegCel WHERE cel_yea='" + sYea + "' AND cel_mon='" + sMon + "'";
            DataTable t = _clsFun.FillTabSql("GesNegCel", s, true, _strConSqlSta);

            DataRow x = t.NewRow();
            x["cel_yea"] = sYea;
            x["cel_mon"] = sMon;
            x["cel_ndo"] = sNdo;
            x["cel_ddo"] = dDdo;

            if (t.Rows.Count > 0)
            {
                sNdo = ((string)t.Rows[0]["cel_ndo"]).Trim();
                if(sNdo == "")
                    sNdo = _clsFun.NewNum(cmbYea.Text, clsDefine.enuNumeratori.NumGesDocCeliaci, 5, _strConSql);
                x["cel_ndo"] = sNdo;
                s = _clsFun.SqlUpdRowIdx("GesNegCel", t, t.Rows[0], x, null);
            }
            else
            {
                sNdo = _clsFun.NewNum(cmbYea.Text, clsDefine.enuNumeratori.NumGesDocCeliaci, 5, _strConSql);
                x["cel_ndo"] = sNdo;
                s = _clsFun.SqlInsertRow("GesNegCel", t, x);
            }
            if(s != "")
                _clsFun.SqlWrite(s, _strConSqlSta);

            string sKey = "";

            foreach (DataRowView r in dtvCel)
            {
                s = (string)r["ven_neg"];
                s += (string)r["ven_cau"];
                s += ((DateTime)r["ven_day"]).ToString("yyyyMMdd");
                s += (string)r["ven_ora"];
                s += (string)r["ven_pos"];
                s += (string)r["ven_sco"];

                if(sKey != s)
                {
                    sKey = s;

                    s = "UPDATE GesNegVet SET ";
                    s += "vet_cdo='" + sNdo + "', ";
                    s += "vet_cdt=" + _clsFun.DaySql(dDdo) + " ";
                    s += "WHERE ";
                    s += "vet_neg='" + (string)r["ven_neg"] + "' AND ";
                    s += "vet_cau='" + (string)r["ven_cau"] + "' AND ";
                    s += "vet_day=" + _clsFun.DaySql((DateTime)r["ven_day"]) + " AND ";
                    s += "vet_ora='" + (string)r["ven_ora"] + "' AND ";
                    s += "vet_pos='" + (string)r["ven_pos"] + "' AND ";
                    s += "vet_sco='" + (string)r["ven_sco"] + "'";

                    _clsFun.SqlWrite(s, _strConSqlSta);
                }
            }
        }

        private void InvioRendicontazione()
        {
            string s = "";
            DataRow[] j;

            DataTable t = ((DataTable)dgv1.DataSource).Clone();
            foreach(DataRow y in ((DataTable)dgv1.DataSource).Rows)
            {
                if((string)y["vet_can"] != "S")
                {
                    s = "ven_neg='" + (string)y["ven_neg"] + "' AND ";
                    s += "ven_cau='" + (string)y["ven_cau"] + "' AND ";
                    s += "ven_day=#" + ((DateTime)y["ven_day"]).ToString("yyyy-MM-dd") + "# AND ";
                    s += "ven_ora='" + (string)y["ven_ora"] + "' AND ";
                    s += "ven_pos='" + (string)y["ven_pos"] + "' AND ";
                    s += "ven_sco='" + (string)y["ven_sco"] + "' AND ";
                    s += "ven_ean='" + (string)y["ven_ean"] + "'";

                    j = t.Select(s);
                    if (j.Length > 0)
                    {
                        j[0]["ven_qta"] = (decimal)j[0]["ven_qta"] + (decimal)y["ven_qta"];
                        j[0]["ven_ven"] = (decimal)j[0]["ven_ven"] + (decimal)y["ven_ven"];
                    }
                    else
                        t.ImportRow(y);
                }
            }

            DataView v = new DataView(t, "vet_can<>'S'", "ven_day, ven_pos, ven_sco", DataViewRowState.CurrentRows);

            if (v.Count > 0)
            {
                string sAsl = "";
                string sPve = "";         //Codice PV per ASL
                string[] a = _strPar037ParCeliaci.Split('-');
                sAsl = a[0];
                sPve = a[1];

                string sFil = "C:\\ApProject\\Temp\\DivCeliachia\\" + "CELIACHIA_" + sPve + "_" + DateTime.Now.ToString("yyyyMM") + ".txt";

                string sYea = cmbYea.Text;
                string sMon = nudMon.Value.ToString("00");
                string sNdo = lblNdo.Text;
                decimal dImp = 0;

                StreamWriter sw = new StreamWriter(sFil, false);

                string sKey = "";
                string sRig = "";

                foreach (DataRowView r in v)
                {
                    //dt.ToString("yyyy-MM-dd

                    //s = (string)r["ven_neg"];
                    //s += (string)r["ven_cau"];
                    ////s += ((DateTime)r["ven_day"]).ToString("yyyy-MM-dd");
                    //s += (string)r["ven_ora"];
                    //s += (string)r["ven_pos"];
                    //s += (string)r["ven_sco"];

                    //if (sKey != s)
                    //{
                    //    sKey = s;

                    //    //DataRow[] j = v.Table.Select("ven_neg + ven_cau + ven_day + ven_ora + ven_pos + ven_sco ='" + s + "'");
                    //    DataRow[] j = v.Table.Select("ven_neg + ven_cau + ven_ora + ven_pos + ven_sco ='" + s + "'");

                    //    Console.WriteLine("xxxx");

                    //    for(int i = 0; i < j.Length; i++)
                    //    {
                    //        if ((DateTime)j[i]["ven_day"] == (DateTime)r["ven_day"])
                    //            dImp += (decimal)j[i]["ven_ven"];
                    //    }
                    //}

                    // 1 - Codice ASL        6
                    sRig = sAsl;

                    // 2 - Anno contabile    4
                    sRig += sYea;

                    // 3 - Mese contabile    2
                    sRig += sMon;

                    // 4 - Numero documento contabile   20
                    sRig += sNdo + new string(' ', 20- sNdo.Length);

                    // 5 - Data documento contabile     8
                    sRig += dtpDdt.Value.ToString("yyyyMMdd");

                    // 6 - Identificativo transazione   12
                    sRig += new string(' ', 12);

                    // 7 - Codice Autorizzazione        20              X
                    s = (string)r["vet_doc"];
                    sRig += s + new string(' ', 20 - s.Length);

                    // 8 - Codice PV                    18
                    sRig += sPve + new string(' ', 18 - sPve.Length);

                    // 9 - Codice cassa                 6
                    s = (string)r["ven_pos"];
                    sRig += s.PadLeft(6, Convert.ToChar('0'));

                    //10 - Numero scontrino             10              X
                    s = (string)r["ven_sco"];
                    sRig += s.PadLeft(10, Convert.ToChar('0'));

                    //11 - Data transazione             8               X
                    sRig += ((DateTime)r["ven_day"]).ToString("yyyyMMdd");

                    //12 - Importo totale               10
                    sRig += ((decimal)r["tmp_imp"] * 100).ToString("0000000000");

                    //13 - Importo autorizzato          10
                    sRig += ((decimal)r["tmp_imp"] * 100).ToString("0000000000");

                    //14 - Codice EAN                   15
                    s = ((string)r["ven_ean"]).Trim();
                    sRig += s + new string(' ', 15 - s.Length);

                    //15 - Quantità                     3
                    sRig += ((decimal)r["ven_qta"]).ToString("000");

                    //16 - Importo                      10
                    sRig += ((decimal)r["ven_ven"] * 100).ToString("0000000000");

                    //17 - Divisa                       3 (fisso EUR)
                    sRig += "EUR";

                    sw.Write(sRig + _clsDef.CRLF);
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();
            }
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s = "";
            decimal d = 0;
            DataRow x;
            DataRow[] j;

            //DataTable t = new DataTable();

            //DataView v = new DataView((DataTable)dgv1.DataSource, "", "DocDdo", DataViewRowState.CurrentRows);
            //t = v.Table.Clone();

            //foreach (DataRowView r in v)
            //{
            //    t.ImportRow(r.Row);
            //}

            DataTable t = ((DataTable)dgv1.DataSource).Copy();
            DataTable tSta = t.Clone();

            tSta.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_key",
                Caption = "Chiave",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            string sSql = "";
            string sKey = "";
            decimal dImp = 0;

            foreach (DataRow y in t.Rows)
            {
                /* Per scontrino */

                sSql = "ven_day=" + _clsFun.DayMdb((DateTime)y["ven_day"]) + " AND ";
                sSql += "ven_pos='" + y["ven_pos"] + "' AND ";
                sSql += "ven_sco='" + y["ven_sco"] + "' AND ";
                sSql += "ven_ora='" + y["ven_ora"] + "'";
                j = tSta.Select(sSql);

                if(j.Length == 0)
                {
                    sKey = ((DateTime)y["ven_day"]).ToString("yyyyMMdd") + "_" + (string)y["ven_pos"] + "_" + (string)y["ven_sco"] + "_" + (string)t.Rows[0]["ven_ora"];

                    x = tSta.NewRow();
                    x["tmp_key"] = sKey;
                    x["ven_neg"] = y["ven_neg"];
                    x["ven_cau"] = y["ven_cau"];
                    x["ven_art"] = "";
                    x["ven_sct"] = "";
                    x["ven_ean"] = "";
                    x["ven_ora"] = y["ven_ora"];
                    x["ven_day"] = y["ven_day"];
                    x["vet_doc"] = y["vet_doc"];
                    x["ven_pos"] = y["ven_pos"];
                    x["ven_sco"] = y["ven_sco"];
                    x["tmp_imp"] = 0;
                    x["vet_can"] = y["vet_can"];
                    tSta.Rows.Add(x);

                    j = tSta.Select(sSql);
                }
                if ((string)y["vet_can"] != "S")
                {
                    j[0]["tmp_imp"] = (decimal)j[0]["tmp_imp"] + (decimal)y["ven_ven"];
                    dImp += (decimal)y["ven_ven"];
                }

                /* Per data */

                sSql = "ven_day=" + _clsFun.DayMdb((DateTime)y["ven_day"]) + "AND ven_ora=''";
                j = tSta.Select(sSql);
                if (j.Length == 0)
                {
                    sKey = ((DateTime)y["ven_day"]).ToString("yyyyMMdd") + "_" + (string)y["ven_pos"] + "_" + "ZZZZZ" + "_" + (string)y["ven_ora"] + "_000";

                    x = tSta.NewRow();
                    x["tmp_key"] = sKey;
                    x["ven_neg"] = "";
                    x["ven_cau"] = "";
                    x["ven_art"] = "";
                    x["ven_sct"] = "";
                    x["ven_ean"] = "";
                    x["ven_ora"] = "";
                    x["ven_day"] = y["ven_day"];
                    x["vet_doc"] = "";
                    x["ven_pos"] = "";
                    x["ven_sco"] = "";
                    x["tmp_imp"] = 0;
                    x["vet_can"] = "";
                    tSta.Rows.Add(x);

                    j = tSta.Select(sSql);
                }

                if ((string)y["vet_can"] != "S")
                    j[0]["tmp_imp"] = (decimal)j[0]["tmp_imp"] + (decimal)y["ven_ven"];

                /* A totale 

                sSql = "tmp_key='ZZZ'";
                j = tSta.Select(sSql);
                if (j.Length == 0)
                {
                    sKey = "ZZZ";

                    x = tSta.NewRow();
                    x["tmp_key"] = sKey;
                    x["ven_neg"] = "";
                    x["ven_cau"] = "";
                    x["ven_art"] = "";
                    x["ven_sct"] = "";
                    x["ven_ean"] = "";
                    x["ven_ora"] = "";
                    x["ven_day"] = _clsDef.DAYOUT;
                    x["vet_doc"] = "";
                    x["ven_pos"] = "";
                    x["ven_sco"] = "";
                    x["tmp_imp"] = 0;
                    x["vet_can"] = "";
                    tSta.Rows.Add(x);

                    j = tSta.Select(sSql);
                }

                if ((string)y["vet_can"] != "S")
                    j[0]["tmp_imp"] = (decimal)j[0]["tmp_imp"] + (decimal)y["ven_ven"];
                 */
            }

            t = new DataView(tSta, "", "tmp_key", DataViewRowState.CurrentRows).ToTable();

            string sTit = DateTime.Now.ToString("dd/MM/yyyy") + " - Riepilogo celiachia nel periodo " + nudMon.Value + "/" + cmbYea.Text;
            string sFoo = "";
            sFoo += "'Totale',,,,," + dImp.ToString().Replace(",", ".")+",";
            string sFil = "";

            string sFld = "";
            sFld += "ven_day, Data, 110, DateTime;";
            sFld += "vet_doc, Autorizzazione, 40, StringLiteral;";
            sFld += "ven_pos, Cassa, 45, StringLiteral;";
            sFld += "ven_ora, Ora, 45, StringLiteral;";
            sFld += "ven_sco, Scontrino, 45, StringLiteral;";
            sFld += "tmp_imp, Importo, 45, Decimal;";
            sFld += "vet_can, Annullato, 45, StringLiteral;";
            //sFld += "tmp_key, Key, 45, StringLiteral;";

            sFil = "CeliachiaRiepilogo.xls";

            (new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);
        }

        private void cmbYea_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillTab();
        }

        private DataTable FillIva(string strRange, string strArr)
        {
            string s = "";
            DataRow[] j;
            DataRow x;
            decimal dQta = 0;
            Boolean b = false;

            //int iRigIni = 0;
            //int iRigFin = 10000;

            //if (strRange != "")
            //{
            //    iRigIni = Convert.ToInt16(strRange.Substring(0, 3));
            //    iRigFin = Convert.ToInt16(strRange.Substring(4, 3));
            //}
            //string sFldRig = "mov_rmo";
            //if (_strMovFat == "F")
            //    sFldRig = "mov_rfa";

            s = "SELECT * FROM TabIva";
            DataTable tTiv = _clsFun.FillTabSql("TabIva", s, false, _strConSql);

            if (dgv2.DataSource == null)
                dgv2.DataSource = new clsGenTabTmp().TabTmpIvaRiep("TabIva");

            DataTable tIva = (DataTable)dgv2.DataSource;
            tIva.Clear();

            DataTable t = (DataTable)dgv1.DataSource;

            decimal dTotScon = 0;           //Seck 20180302 totale valore da scontrini
            ArrayList aryScon = new ArrayList();
            Boolean bTotScon = true;

            foreach (DataRow y in t.Rows)
            {
                if (DBNull.Value.Equals(y["ven_ven"]))
                    y["ven_ven"] = 0;
                //if (DBNull.Value.Equals(y["mov_ann"]))
                //    y["mov_ann"] = false;

                if ((decimal)y["ven_ven"] != 0) // && (strRange == "" || (Convert.ToInt16(y[sFldRig]) >= iRigIni && Convert.ToInt16(y[sFldRig]) <= iRigFin)))
                {
                    b = true;
                    if (((string)y["ven_iva"]).Trim() == "")
                        b = false;
                    if (DBNull.Value.Equals(y["ven_ven"]) || (decimal)y["ven_ven"] == 0)
                        b = false;

                    if (b)
                    {
                        j = tIva.Select("iva_cod='" + (string)y["ven_iva"] + "'");
                        if (j.Length == 0)
                        {
                            j = tTiv.Select("tab_cod='" + ((string)y["ven_iva"]).PadLeft(3, Convert.ToChar('0')) + "'");
                            if (j.Length == 0)
                                j = tTiv.Select("tab_cod='" + _strIvaPar + "'");
                            x = tIva.NewRow();
                            x["iva_cod"] = y["ven_iva"];
                            x["iva_des"] = j[0]["tab_des"];
                            x["iva_ali"] = j[0]["tab_ali"];
                            x["iva_iva"] = 0;
                            x["iva_imp"] = 0;
                            x["iva_tot"] = 0;
                            tIva.Rows.Add(x);
                            j = tIva.Select("iva_cod='" + y["ven_iva"] + "'");
                        }

                        if ((decimal)y["ven_ven"] != 0 && !DBNull.Value.Equals(j[0]["iva_ali"]))
                        {
                            if ((string)y["ven_art"] == "0156195")
                                Console.WriteLine("aaaa");

                            y["VenNiv"] = _clsFun.MenoIva((decimal)y["ven_ven"], (decimal)j[0]["iva_ali"]);

                            //j[0]["iva_iva"] = (decimal)j[0]["iva_iva"] + _clsFun.ValIva((decimal)y["mov_imp"], (decimal)j[0]["iva_ali"]);
                            //j[0]["iva_imp"] = Math.Round((decimal)j[0]["iva_imp"] + (decimal)y["mov_imp"], 2, MidpointRounding.AwayFromZero);
                            //j[0]["iva_imp"] = Math.Round((decimal)j[0]["iva_imp"] + (decimal)y["mov_imp"], 3, MidpointRounding.ToEven);
                            j[0]["iva_imp"] = (decimal)j[0]["iva_imp"] + (decimal)y["VenNiv"];
                        }
                        else
                            Console.WriteLine("aaaaaaaaa");
                    }
                    /*
                    if (!DBNull.Value.Equals(y["mov_ori"]) && ((string)y["mov_ori"]).Trim() != "")
                    {
                        s = (string)y["mov_ori"];

                        string[] a = s.Split(',');

                        if (a.Length > 4)
                        {
                            string sKey = s.Substring(0, 24);

                            if (aryScon.IndexOf(sKey) < 0)
                            {
                                aryScon.Add(sKey);
                                if (a.Length > 5)
                                    dTotScon += Convert.ToDecimal(a[5].Replace(".", ","));
                            }
                        }
                    }
                    else if ((decimal)y["mov_imp"] > 0)     //Seck 20180304 Se ci sono righe non da scontrino con righe da scontrino non faccio arrotondamenti
                        bTotScon = false;
                    */
                }

                if ((string)y["ven_umi"] == "KG")
                    dQta += 1;
                else
                    dQta += (decimal)y["ven_qta"];
            }

            if (!bTotScon)
                dTotScon = 0;

            lblTotQta.Text = _clsFun.Dec2Txt(dQta, 0);

            tIva = Riepilogo(tIva, dTotScon);

            return tIva;
        }

        private DataTable Riepilogo(DataTable tIva, decimal decScoImp)
        {
            decimal dIva = 0;
            decimal dImp = 0;
            decimal dTot = 0;

            //DataView v = new DataView(tIva, "", "tab_ali", DataViewRowState.CurrentRows);
            tIva = new DataView(tIva, "", "iva_ali", DataViewRowState.CurrentRows).Table;

            string[] aArr = { };

            foreach (DataRow y in tIva.Rows)
            {
                if ((decimal)y["iva_imp"] != 0)
                {
                    y["iva_iva"] = Math.Round(_clsFun.ValIva((decimal)y["iva_imp"], (decimal)y["iva_ali"]), 2, MidpointRounding.ToEven);
                    y["iva_iva"] = Math.Round((decimal)y["iva_iva"], 2, MidpointRounding.ToEven);
                    y["iva_tot"] = Math.Round((decimal)y["iva_imp"] + (decimal)y["iva_iva"], 2, MidpointRounding.ToEven);
                    y["iva_imp"] = (decimal)y["iva_tot"] - (decimal)y["iva_iva"];           //20180118 Seck
                }
            }

            foreach (DataRow y in tIva.Rows)
            {
                dImp += (decimal)y["iva_imp"];
                dIva += (decimal)y["iva_iva"];
                dTot += (decimal)y["iva_tot"];
            }

            if (decScoImp > 0 && dTot != decScoImp)
            {
                decimal dDif = decScoImp - dTot;
                if (dDif > 0)
                    tIva.Rows[0]["iva_iva"] = (decimal)tIva.Rows[0]["iva_iva"] + dDif;
                else
                    tIva.Rows[0]["iva_imp"] = (decimal)tIva.Rows[0]["iva_imp"] + dDif;

                tIva.Rows[0]["iva_arr"] = dDif;

                dImp = 0;
                dIva = 0;
                dTot = 0;

                foreach (DataRow y in tIva.Rows)
                {
                    dImp += (decimal)y["iva_imp"];
                    dIva += (decimal)y["iva_iva"];
                    //dTot += (decimal)y["iva_tot"];
                    dTot += (decimal)y["iva_imp"] + (decimal)y["iva_iva"];
                }
            }

            dImp = Math.Round(dImp, 2, MidpointRounding.AwayFromZero); // MidpointRounding.ToEven);

            lblTotIva.Text = _clsFun.Dec2Txt(dIva, 2);
            lblTotImp.Text = _clsFun.Dec2Txt(dImp, 2);
            lblTotTot.Text = _clsFun.Dec2Txt(dTot, 2);

            return tIva;
        }

    }
}
