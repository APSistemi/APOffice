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
    public partial class frmGesStatMovCassa2 : Form
    {
        private const string TABSTAVET = "GesNegVet";
        private const string TABSTAVEP = "GesNegVep";

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strConSqlStat = "";

        public frmGesStatMovCassa2()
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
            cTbc.Width = 150;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

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
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_pos";
            cTbc.Name = "Cassa";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

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

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_imp";
            cTbc.Name = "Importi";
            cTbc.Width = 70;
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
            s = "SELECT * FROM TabUtenti WHERE tab_ann=0 ORDER BY tab_cod";
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
        }

        private void FillDati()
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

            DataTable t = tVep.Clone();

            dInc = 0; 

            if (tVep.Rows.Count > 0)
            {
                string sKey = "";

                foreach (DataRow y in tVep.Rows)
                {
                    if ((string)y["vep_tip"] == "RES")
                        y["vep_cod"] = "001";

                    sKey = "vet_pos='" + y["vet_pos"] + "' AND ";
                    sKey +="vet_pos='" + y["vet_pos"] + "' AND ";
                    sKey += "vep_cod='" + y["vep_cod"] + "'";

                    j = t.Select(sKey);
                    if (j.Length == 0)
                    {
                        x = t.NewRow();
                        //if (sPos != "" || chkDiv.Checked)
                        x["vet_pos"] = y["vet_pos"];
                        x["vep_tip"] = y["vep_tip"];
                        x["vep_cod"] = y["vep_cod"];
                        x["vep_imp"] = 0;

                        j = tCau.Select("tab_cod='" + y["vep_cod"] + "'");
                        if (j.Length > 0)
                        {
                            x["PagDes"] = (string)j[0]["tab_des"];
                            x["CauTip"] = (string)j[0]["tab_key"];
                        }

                        //if (sUsr != "")
                        //{
                        x["vet_usr"] = y["vet_usr"];
                        j = tUsr.Select("tab_cod='" + y["vet_usr"] + "'");
                        if (j.Length > 0)
                            x["VepUsr"] = (string)j[0]["tab_des"];
                        //}

                        //if (sNeg != "")
                        //{
                        x["vet_neg"] = y["vet_neg"];
                        j = tNeg.Select("tab_cod='" + y["vet_neg"] + "'");
                        if (j.Length > 0)
                            x["VetNeg"] = (string)j[0]["tab_des"];
                        //}

                        t.Rows.Add(x);
                        j = t.Select(sKey);
                    }
                    if ((string)y["vep_tip"] == "RES")
                        j[0]["vep_imp"] = (decimal)j[0]["vep_imp"] - (decimal)y["vep_imp"];
                    else
                        j[0]["vep_imp"] = (decimal)j[0]["vep_imp"] + (decimal)y["vep_imp"];

                    if((string)y["vep_tip"] == "RES")
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
            tVep = _clsFun.FillTabSql("", s, false, _strConSqlStat);

            foreach(DataRow y in tVep.Rows)
            {
                string sCod = "";
                string sDes = "";
                string sTip = "";
                string sSgn = "";
                string sMcaUsr = "";
                string sMcaPos = "";
                string sMcaNeg = "";

                j = tCau.Select("tab_tip='" + y["mca_cau"] + "'");
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

                    j = t.Select("vep_cod='" + sCod + "'");
                    if (j.Length == 0)
                    {
                        x = t.NewRow();
                        x["vet_pos"] = sMcaPos;
                        x["vep_cod"] = sCod;
                        x["vep_tip"] = "MOV"; // y["mca_cau"];
                        x["vep_imp"] = 0;
                        x["CauTip"] = sTip;
                        x["PagDes"] = sDes;

                        x["vet_usr"] = sMcaUsr;
                        j = tUsr.Select("tab_cod='" + sMcaUsr + "'");
                        if (j.Length > 0)
                            x["VepUsr"] = (string)j[0]["tab_des"];

                        x["vet_neg"] = sMcaNeg;
                        j = tNeg.Select("tab_cod='" + sMcaNeg + "'");
                        if (j.Length > 0)
                            x["VetNeg"] = (string)j[0]["tab_des"];

                        t.Rows.Add(x);
                        j = t.Select("vep_cod='" + sCod + "'");
                    }

                    if(sSgn == "-")
                        j[0]["vep_imp"] = (decimal)j[0]["vep_imp"] - (decimal)y["mca_imp"];
                    else
                        j[0]["vep_imp"] = (decimal)j[0]["vep_imp"] + (decimal)y["mca_imp"];
                }
            }

            decimal d = 0;

            foreach(DataRow y in t.Rows)
            {
                if ((string)y["CauTip"] == "CON")
                    d += (decimal)y["vep_imp"];
                else if ((string)y["CauTip"] == "VER")
                    d += (decimal)y["vep_imp"];
                else if ((string)y["CauTip"] == "PRE")
                    d += (decimal)y["vep_imp"];
            }

            lblTot.Text = dInc.ToString("###,##0.00");
            lblDif.Text = d.ToString("###,##0.00");

            dgv1.DataSource = t;
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
            string s = "";
            string sPar = "N;";
            if (chkDiv.Checked)
                sPar = "S;";

            sPar += cmbNeg.SelectedValue.ToString() + ";";
            sPar += cmbUsr.SelectedValue.ToString() + ";";
            sPar += cmbPos.Text.Substring(0,1).Trim() + ";";

            DataTable t = (DataTable)dgv1.DataSource;

            clsGenPdfSta1 cls = new clsGenPdfSta1();
            cls._dayIni = dtpIni.Value;
            cls._dayFin = dtpFin.Value;
            s = cls.PrnPdfStaPos2(t, lblTot.Text, lblDif.Text, sPar);
            if(s != "")
                MessageBox.Show(s, "CONTROLLO ACCESSO DATI", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void chkDiv_CheckedChanged(object sender, EventArgs e)
        {
            if(chkDiv.Checked)
                chkDiv.Text = "SUDDIVISO";
            else
                chkDiv.Text = "RIEPILOGATIVO";
        }
    }
}
