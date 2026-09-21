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
    public partial class frmGesStatMovCassa : Form
    {
        private const string TABSTAVET = "GesNegVet";
        private const string TABSTAVEP = "GesNegVep";

        public DateTime _dayDti = DateTime.Today;
        public DateTime _dayDtf = DateTime.Today;

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strConSqlStat = "";

        public frmGesStatMovCassa()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlStat = _clsFun.ConSql("3");
        }

        private void frmGesStatMovCassa_Load(object sender, EventArgs e)
        {
            //dtpIni.Value = new DateTime(2019, 3, 2, 0, 0, 0);
            //dtpFin.Value = new DateTime(2019, 3, 2, 0, 0, 0);    

            dtpIni.Value = _dayDti;
            dtpFin.Value = _dayDtf;
            
            FillTabs();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmGesStatMovCassa_KeyDown(object sender, KeyEventArgs e)
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
            dgv1.ReadOnly = true;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;
            
            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;
            dgv1.Columns.Clear();

            if (cmbNeg.SelectedValue != null && cmbNeg.SelectedValue.ToString() != "")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "vet_neg";
                cTbc.Name = "Negozio";
                cTbc.Width = 40;
                cTbc.ValueType = typeof(string);
                cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                cTbc.ReadOnly = true;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "VetNeg";
                cTbc.Name = "Descrizione";
                cTbc.Width = 100;
                cTbc.ValueType = typeof(string);
                //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                cTbc.ReadOnly = true;
                dgv1.Columns.Add(cTbc);
            }

            if (chkUsr.Checked || ( cmbUsr.SelectedValue != null && cmbUsr.SelectedValue.ToString() != ""))
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "vet_usr";
                cTbc.Name = "Utente";
                cTbc.Width = 60;
                cTbc.ValueType = typeof(string);
                cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                cTbc.ReadOnly = true;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "VepUsr";
                cTbc.Name = "Descrizione";
                cTbc.Width = 120;
                cTbc.ValueType = typeof(string);
                //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                cTbc.ReadOnly = true;
                dgv1.Columns.Add(cTbc);
            }

            if (chkPos.Checked || (cmbPos.Text.Length > 0 && cmbPos.Text.Substring(0,1) != " "))
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.DataPropertyName = "vet_pos";
                cTbc.Name = "Cassa";
                cTbc.Width = 60;
                cTbc.ValueType = typeof(string);
                cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                cTbc.ReadOnly = true;
                dgv1.Columns.Add(cTbc);
            }

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_tip";
            cTbc.Name = "Tipo riga";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "PagDes";
            cTbc.Name = "Descrizione";
            cTbc.Width = 140;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "vep_imp";
            //cTbc.Name = "Importi";
            //cTbc.Width = 55;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //cTbc.DefaultCellStyle.Format = "###,##0.00";
            //cTbc.ReadOnly = false;
            //dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "TotSco";
            cTbc.Name = "Importi da scontrino";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "TotVer";
            cTbc.Name = "Importi versamenti";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "TotPre";
            cTbc.Name = "Importi prelievi";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "TotDif";
            cTbc.Name = "Saldo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "vep_val";
            //cTbc.Name = "Valori";
            //cTbc.Width = 70;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //cTbc.DefaultCellStyle.Format = "###,##0.00";
            //cTbc.ReadOnly = false;
            //dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "VepArt";
            //cTbc.Name = "Articoli";
            //cTbc.Width = 300;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = false;
            //dgv1.Columns.Add(cTbc);
        }

        private void FillTabs()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabUtenti";
            s = "SELECT * FROM TabUtenti WHERE tab_ann=0 ORDER BY tab_des";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbUsr.DataSource = t;
            cmbUsr.DisplayMember = "tab_des";
            cmbUsr.ValueMember = "tab_cod";
            cmbUsr.SelectedValue = "";

            //p = "TabCauCassa";
            //s = "SELECT * FROM TabCauCassa WHERE tab_tip='PAG'";
            //t = _clsFun.FillTabSql(p, s, false, _strConSql);
            //x = t.NewRow();
            //x["tab_cod"] = "";
            ////x["tab_pos"] = "";
            //x["tab_des"] = "  Non definito";
            //t.Rows.InsertAt(x, 0);
            //cmbTip.DataSource = t;
            //cmbTip.DisplayMember = "tab_des";
            //cmbTip.ValueMember = "tab_cod";
            //cmbTip.SelectedValue = "";

            p = "TabNegozi";
            s = "SELECT * FROM TabNegozi WHERE tab_tip='L'";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_pos"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);

            if(t.Rows.Count > 0)
            {
                cmbPos.Items.Add(" Non Definito");

                foreach (DataRow y in t.Rows)
                {
                    s = (string)y["tab_pos"];

                    string[] a = s.Split(',');

                    foreach (string ss in a)
                    {
                        if (ss.Trim() != "")
                            cmbPos.Items.Add(ss);
                    }
                }
                cmbPos.SelectedIndex = 0;
            }            
            
            cmbNeg.DataSource = t;
            cmbNeg.DisplayMember = "tab_des";
            cmbNeg.ValueMember = "tab_cod";
            cmbNeg.SelectedValue = "";
            if(t.Rows.Count == 2)
                cmbNeg.SelectedValue = (string)t.Rows[1]["tab_cod"];
        }

        private void FillDati()
        {
            if (cmbNeg.DataSource != null)
            {
                SetDgv1();

                string sPos = cmbPos.Text.Substring(0, 1).Trim();
                string sNeg = cmbNeg.SelectedValue.ToString();
                string sUsr = cmbUsr.SelectedValue.ToString();

                if (sPos != "")
                    sPos = sPos.PadLeft(2, Convert.ToChar("0"));

                string s = "";
                DataRow x;
                DataRow[] j;
                decimal dInc = 0;

                s = "SELECT * FROM TabCauCassa";
                DataTable tCau = _clsFun.FillTabSql("Cau", s, false, _strConSql);
                DataTable tUsr = (DataTable)cmbUsr.DataSource;
                DataTable tNeg = (DataTable)cmbNeg.DataSource;

                s = "SELECT ";
                s += "GesNegVet.vet_neg, ";
                s += "GesNegVet.vet_usr, ";
                s += "GesNegVet.vet_pos, ";
                s += "GesNegVet.vet_ora, ";
                s += "GesNegVet.vet_sco, ";
                s += "GesNegVet.vet_imp, ";
                s += "GesNegVep.vep_cod, ";
                s += "GesNegVep.vep_tip, ";
                s += "GesNegVep.vep_imp ";
                s += "FROM GesNegVep ";
                s += "INNER JOIN GesNegVet ON ";
                s += "GesNegVep.vep_neg = GesNegVet.vet_neg AND ";
                s += "GesNegVep.vep_cau = GesNegVet.vet_cau AND ";
                s += "GesNegVep.vep_day = GesNegVet.vet_day AND ";
                s += "GesNegVep.vep_ora = GesNegVet.vet_ora AND ";
                s += "GesNegVep.vep_pos = GesNegVet.vet_pos AND ";
                s += "GesNegVep.vep_sco = GesNegVet.vet_sco ";
                s += "WHERE ";
                s += "(GesNegVet.vet_cau='SCO') AND ";
                s += "(GesNegVep.vep_tip='PAG' OR GesNegVep.vep_tip='RES') AND ";
                s += "(GesNegVep.vep_day >= " + _clsFun.DaySql(dtpIni.Value) + " AND GesNegVep.vep_day <= " + _clsFun.DaySql(dtpFin.Value) + ") ";
                if (sUsr != "")
                    s += " AND GesNegVet.vet_usr = '" + sUsr + "' ";
                if (sNeg != "")
                    s += " AND GesNegVet.vet_neg = '" + sNeg + "' ";
                if (sPos != "")
                    s += " AND GesNegVet.vet_pos = '" + sPos + "' ";
                s += "ORDER BY GesNegVet.vet_pos, GesNegVet.vet_sco";
                DataTable tVep = _clsFun.FillTabSql("Vep", s, false, _strConSqlStat);

                tVep.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "CauTip",
                    Caption = "Tipo",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (string)""
                });

                tVep.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "PagDes",
                    Caption = "Descrizione",
                    MaxLength = 50,
                    ReadOnly = false,
                    DefaultValue = (string)""
                });

                tVep.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "VepUsr",
                    Caption = "Descrizione utente",
                    MaxLength = 20,
                    ReadOnly = false,
                    DefaultValue = (string)""
                });

                tVep.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "VetNeg",
                    Caption = "Descrizione negozio",
                    MaxLength = 20,
                    ReadOnly = false,
                    DefaultValue = (string)""
                });

                tVep.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "TotSco",
                    Caption = "Importo scontrini",
                    ReadOnly = false,
                    DefaultValue = (decimal)0
                });

                tVep.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "TotVer",
                    Caption = "Importo versamenti",
                    ReadOnly = false,
                    DefaultValue = (decimal)0
                });

                tVep.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "TotPre",
                    Caption = "Importo prelievi",
                    ReadOnly = false,
                    DefaultValue = (decimal)0
                });

                tVep.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "TotDif",
                    Caption = "Importo saldo",
                    ReadOnly = false,
                    DefaultValue = (decimal)0
                });

                DataTable t = tVep.Clone();

                dInc = 0;

                string sKey = "";

                if (tVep.Rows.Count > 0)
                {
                    foreach (DataRow y in tVep.Rows)
                    {
                        if ((string)y["vep_tip"] == "RES")
                            y["vep_cod"] = "001";

                        //sKey = "vep_cod='" + y["vep_cod"] + "'";
                        //if (chkPos.Checked)
                        //    sKey = "vet_pos='" + y["vet_pos"] + "' AND vep_cod='" + y["vep_cod"] + "'";

                        sKey = "vep_cod='" + y["vep_cod"] + "'";
                        if (chkPos.Checked)
                            sKey += " AND vet_pos='" + y["vet_pos"] + "'";
                        if (chkUsr.Checked)
                            sKey += " AND vet_usr='" + y["vet_usr"] + "'";

                        j = t.Select(sKey);
                        if (j.Length == 0)
                        {
                            x = t.NewRow();
                            if (sPos != "" || chkPos.Checked)
                                x["vet_pos"] = y["vet_pos"];
                            x["vep_tip"] = y["vep_tip"];
                            x["vep_cod"] = y["vep_cod"];
                            x["vep_imp"] = 0;

                            j = tCau.Select("tab_cod='" + y["vep_cod"] + "'");
                            if (j.Length > 0)
                            {
                                x["PagDes"] = (string)j[0]["tab_des"];
                                x["CauTip"] = (string)j[0]["tab_key"];

                                if ((string)x["CauTip"] == "MOV")
                                    Console.WriteLine("zzzz");
                            }

                            //if (sUsr != "")
                            //{
                            if(chkUsr.Checked)
                            { 
                                x["vet_usr"] = y["vet_usr"];
                                j = tUsr.Select("tab_cod='" + y["vet_usr"] + "'");
                                if (j.Length > 0)
                                    x["VepUsr"] = (string)j[0]["tab_des"];
                            }

                            if (sNeg != "")
                            {
                                x["vet_neg"] = y["vet_neg"];
                                j = tNeg.Select("tab_cod='" + y["vet_neg"] + "'");
                                if (j.Length > 0)
                                    x["VetNeg"] = (string)j[0]["tab_des"];
                            }

                            t.Rows.Add(x);
                            j = t.Select(sKey);
                        }
                        if ((string)y["vep_tip"] == "RES")
                            j[0]["vep_imp"] = (decimal)j[0]["vep_imp"] - (decimal)y["vep_imp"];
                        else
                            j[0]["vep_imp"] = (decimal)j[0]["vep_imp"] + (decimal)y["vep_imp"];

                        j[0]["TotSco"] = (decimal)j[0]["vep_imp"];

                        if ((string)y["vep_tip"] == "RES")
                            dInc -= (decimal)y["vep_imp"];
                        else
                            dInc += (decimal)y["vep_imp"];
                    }
                }

                s = "SELECT * FROM GesNegMca WHERE ";
                s += "(GesNegMca.mca_day >= " + _clsFun.DaySql(dtpIni.Value) + " AND GesNegMca.mca_day <= " + _clsFun.DaySql(dtpFin.Value) + ") ";
                if (sUsr != "")
                    s += " AND GesNegMca.mca_usr = '" + sUsr + "' ";
                if (sNeg != "")
                    s += " AND GesNegMca.mca_neg = '" + sNeg + "' ";
                if (sPos != "")
                    s += " AND GesNegMca.mca_pos = '" + sPos + "' ";
                s += "ORDER BY GesNegMca.mca_pos";
                tVep = _clsFun.FillTabSql("GesNegMca", s, false, _strConSqlStat);

                /**/

                foreach (DataRow y in tVep.Rows)
                {
                    if (DBNull.Value.Equals(y["mca_neg"]) || ((string)y["mca_neg"]).Trim() == "")
                        y["mca_neg"] = "001";

                    string sCau = (string)y["mca_cau"];
                    string sCod = "";
                    string sDes = "";
                    string sTip = "";
                    string sSgn = "";
                    string sMcaUsr = "";
                    string sMcaPos = "";
                    string sMcaNeg = "";

                    j = tCau.Select("tab_tip='" + (string)y["mca_cau"] + "'");
                    //j = tCau.Select("tab_cod='" + y["mca_tpa"] + "'");
                    if (j.Length > 0)
                    {
                        sCod = (string)j[0]["tab_cod"];
                        sDes = (string)j[0]["tab_des"];
                        sTip = (string)j[0]["tab_tip"];
                        sSgn = (string)j[0]["tab_sgn"];

                        sMcaPos = (string)y["mca_pos"];
                        sMcaUsr = (string)y["mca_usr"];
                        sMcaUsr = (string)y["mca_usr"];
                        sMcaNeg = (string)y["mca_neg"];
                        if (sMcaNeg.Trim() == "")
                            sMcaNeg = "001";

                        sCod = (string)y["mca_tpa"];

                        //sKey = "vep_cod='" + y["vep_cod"] + "'";
                        //if (chkDiv.Checked)
                        //    sKey = "vet_pos='" + y["vet_pos"] + "' AND vep_cod='" + y["vep_cod"] + "'";
                        //j = t.Select(sKey);

                        sKey = "vep_cod='" + y["mca_tpa"] + "' AND vet_pos='" + y["mca_pos"] + "' AND vet_usr='" + y["mca_usr"] + "'";
                        //j = t.Select("vep_cod='" + sCod + "'");
                        j = t.Select(sKey);
                        if (j.Length == 0)
                        {
                            x = t.NewRow();
                            if (sPos != "")
                                x["vet_pos"] = sMcaPos;
                            x["vep_cod"] = sCod;
                            x["vep_tip"] = "MOV"; // y["mca_cau"];
                            x["vet_neg"] = sMcaNeg;
                            x["vet_usr"] = sMcaUsr;
                            x["vet_pos"] = sMcaPos;
                            x["vep_imp"] = 0;
                            x["CauTip"] = sTip;

                            x["PagDes"] = sDes;

                            if (sMcaUsr != "")
                            {
                                j = tUsr.Select("tab_cod='" + sMcaUsr + "'");
                                if (j.Length > 0)
                                    x["VepUsr"] = (string)j[0]["tab_des"];
                                else
                                    Console.WriteLine("zzzzz");
                            }

                            if (sMcaNeg != "")
                            {
                                j = tNeg.Select("tab_cod='" + sMcaNeg + "'");
                                if (j.Length > 0)
                                    x["VetNeg"] = (string)j[0]["tab_des"];
                                else
                                    Console.WriteLine("zzzzz");

                            }
                            else
                                Console.WriteLine("zzzzz");

                            t.Rows.Add(x);
                            //j = t.Select("vep_cod='" + sCod + "'");
                            j = t.Select(sKey);
                        }

                        if (sSgn == "-")
                        {

                            j[0]["vep_imp"] = (decimal)j[0]["vep_imp"] - (decimal)y["mca_imp"];

                            if ((string)y["mca_cau"] == "PRE")
                                j[0]["TotPre"] = (decimal)j[0]["TotPre"] - (decimal)y["mca_imp"];
                            else if ((string)y["mca_cau"] == "VER")
                                j[0]["TotVer"] = (decimal)j[0]["TotVer"] + (decimal)y["mca_imp"];
                        }
                        else
                        {
                            j[0]["vep_imp"] = (decimal)j[0]["vep_imp"] + (decimal)y["mca_imp"];

                            if ((string)y["mca_cau"] == "PRE")
                                j[0]["TotPre"] = (decimal)j[0]["TotPre"] - (decimal)y["mca_imp"];
                            else if ((string)y["mca_cau"] == "VER")
                                j[0]["TotVer"] = (decimal)j[0]["TotVer"] + (decimal)y["mca_imp"];
                        }

                        //if((string)x[""])

                    }
                }
                /**/

                decimal d = 0;

                foreach (DataRow y in t.Rows)
                {
                    //if ((string)y["CauTip"] == "CON")
                    //    d += (decimal)y["vep_imp"];
                    //else if ((string)y["CauTip"] == "VER")
                    //    d += (decimal)y["vep_imp"];
                    //else if ((string)y["CauTip"] == "PRE")
                    //    d += (decimal)y["vep_imp"];

                    y["TotDif"] = ((decimal)y["TotSco"] + (decimal)y["TotVer"] + (decimal)y["TotPre"]) * -1;

                    d += (decimal)y["TotDif"];
                }

                lblTot.Text = dInc.ToString("###,##0.00");
                lblDif.Text = d.ToString("###,##0.00");

                dgv1.DataSource = t;
            }
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Dgv2Csv(dgv1);
        }

        private void Dgv2Csv(DataGridView dgv)
        {
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv.DataSource, "", "", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            string s = "";
            //decimal d = 0;

            foreach (DataRowView r in v)
            {
                t.ImportRow(r.Row);
            }

            string sTit = DateTime.Today.ToString("dd/MM/yyyy") + " - Estrazione cassa dal " + dtpIni.Value.ToString("dd/MM/yyyy") + " al " + dtpFin.Value.ToString("dd/MM/yyyy");
            string sFil = "";
            string sFld = "";

            int ii = 0;

            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                sFld += dgv.Columns[i].DataPropertyName + "," + dgv.Columns[i].Name + "," + dgv.Columns[i].Width.ToString() + "," + dgv.Columns[i].ValueType.ToString() + ";";
                ii = i;
            }

            s = new string(';',ii-1);

            string sFoo = s + "Totale;" + lblTot.Text.Replace(",", ".");

            (new clsExcel()).exportToCsv1(sFld, t, sFil, sTit, sFoo);
        }

        private void pdfToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string sPar = "";

            if (chkUsr.Checked)
                sPar = "S;";
            else
                sPar = "N;";

            if (chkPos.Checked)
                sPar += "S;";
            else
                sPar += "N;";

            if (cmbNeg.Text.Substring(0, 1).Trim() != "")
                sPar += cmbNeg.Text + ";";
            else
                sPar += ";";
                
            if (cmbUsr.Text.Substring(0, 1).Trim() != "")
                sPar += cmbUsr.Text + ";";
            else
                sPar += ";";

            if (cmbPos.Text.Substring(0,1).Trim() != "")
                sPar += cmbPos.Text + ";";
            else
                sPar += ";";

            DataTable t = (DataTable)dgv1.DataSource;

            clsGenPdfSta1 cls = new clsGenPdfSta1();
            cls._dayIni = dtpIni.Value;
            cls._dayFin = dtpFin.Value;
            cls.PrnPdfStaPos(t, lblTot.Text, lblDif.Text, sPar);
        }

        private void pdfPerCassiereToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s = "";
            string sPar = "";

            DataRow[] j;

            DataTable t = (DataTable)dgv1.DataSource;

            string sUsr = "";
            string sUsd = "";

            foreach (DataRow y in t.Rows)
            {
                s = "V_" + (string)y["vet_pos"] + (string)y["vet_usr"];

                if (!sUsr.Contains(s))
                {
                    sUsr += s + ";";
                    sUsd += (string)y["VepUsr"] + " " + "Cassa " + (string)y["vet_pos"] + ";";
                }
            }

            DataTable tTmp = new clsGenTabTmp().TabTmpPosCassieri("TabPos", sUsr);

            foreach(DataRow y in t.Rows)
            {
                j = tTmp.Select("TmpTre='P1' AND TmpCod='" + y["vep_cod"] + "'");
                if(j.Length == 0)
                {
                    DataRow x = tTmp.NewRow();
                    x["TmpTre"] = "P1";
                    x["TmpCod"] = y["vep_cod"];
                    x["TmpDes"] = y["PagDes"];
                    tTmp.Rows.Add(x);

                    j = tTmp.Select("TmpTre='P1' AND TmpCod='" + y["vep_cod"] + "'");
                }

                s =  "V_" + (string)y["vet_pos"] + (string)y["vet_usr"];

                j[0][s] = y["vep_imp"];

            }

            DataRow xx = tTmp.NewRow();
            xx["TmpTre"] = "P2";
            xx["TmpCod"] = "";
            xx["TmpDes"] = "Totali";
            tTmp.Rows.Add(xx);

            string[] a = sUsr.Split(';');

            foreach (DataRow y in tTmp.Rows)
            {

                foreach (string ss in a)
                {
                    if (ss != "")
                    {
                        if ((string)y["TmpTre"] == "P1")
                        {
                            //s = "V_" + ss;
                            y["TmpTot"] = (decimal)y["TmpTot"] + (decimal)y[ss];

                            j = tTmp.Select("TmpTre='P2'");
                            j[0][ss] = (decimal)j[0][ss] + (decimal)y[ss];
                        }
                    }
                }

            }

            dgv2.DataSource = tTmp;

            sPar = sUsr + "-" + sUsd;

            string sNeg = cmbNeg.SelectedValue.ToString();

            frmGesStatIVA f = new frmGesStatIVA();
            DataView vIva = f.FillDatiIva(sNeg, dtpIni.Value, dtpFin.Value);

            clsGenPdfSta1 cls = new clsGenPdfSta1();
            cls._dayIni = dtpIni.Value;
            cls._dayFin = dtpFin.Value;
            cls.PrnPdfStaPos3(tTmp, lblTot.Text, lblDif.Text, sPar, vIva);
        }

        //private void chkDiv_CheckedChanged(object sender, EventArgs e)
        //{
        //    if(chkDiv.Checked)
        //        chkDiv.Text = "SUDDIVISO";
        //    else
        //        chkDiv.Text = "RIEPILOGATIVO";
        //}

        private void cmbNeg_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillDati();
        }

        private void cmbUsr_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillDati();
        }

        private void cmbPos_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillDati();
        }

        private void cmbPos_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillDati();
        }
    }
}
