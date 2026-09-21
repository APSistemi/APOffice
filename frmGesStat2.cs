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
    public partial class frmGesStat2 : Form
    {
        private const string TABSTAVET = "GesNegVet";
        private const string TABSTAVEP = "GesNegVep";
        private const string TABSTAVEN = "GesNegVen";
        private const string TABMOVFAT = "GesFatTestate";

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strConSqlSta = "";
        private string _strStatVisProf = "";

        public string _strArtCod = "";
        public Boolean _bolPrnAuto = false;

        public frmGesStat2()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
        }
        private void frmGesStat2_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
            _strStatVisProf = _clsFun.ParGet(clsDefine.enuParametri.ParStatVisProf, _strConSql);

            SetDgv1();
            ApplyModernUi();
            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, this.Name);

            DataTable t = new clsGenTabTmp().TabTmpStatistiche("TabSta");
            dgv1.DataSource = new DataView(t,"","tmp_day,tmp_odx",DataViewRowState.CurrentRows);
        }

        private void frmGesStat2_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                clsUiIcons.SaveFormBounds(this);
                clsUiIcons.SaveGridColumnWidths(dgv1, this.Name);
            }
            catch { }
        }

        private void ApplyModernUi()
        {
            try
            {
                // Top Menu Strip
                if (menuStrip1 != null)
                {
                    menuStrip1.Renderer = clsUiIcons.GetModernMenuRenderer();
                    if (esciToolStripMenuItem != null)
                        esciToolStripMenuItem.Image = clsUiIcons.GetIcon("exit", 16);
                    if (utilityToolStripMenuItem != null)
                        utilityToolStripMenuItem.Image = clsUiIcons.GetIcon("tools_gear", 16);
                    if (aggiornamentoCostiToolStripMenuItem != null)
                        aggiornamentoCostiToolStripMenuItem.Image = clsUiIcons.GetIcon("refresh", 16);
                }

                // Form background
                this.BackColor = Color.FromArgb(243, 244, 246);

                // Date Filters & Labels
                if (label1 != null)
                {
                    label1.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
                    label1.ForeColor = Color.FromArgb(51, 65, 85);
                }
                if (label2 != null)
                {
                    label2.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
                    label2.ForeColor = Color.FromArgb(51, 65, 85);
                }
                if (dtpIni != null) dtpIni.Font = new Font("Segoe UI", 9.0f);
                if (dtpFin != null) dtpFin.Font = new Font("Segoe UI", 9.0f);

                // Bottone Estrai
                if (btnEstrai != null)
                {
                    clsUiIcons.StyleStatButton(btnEstrai, "Estrai Dati", "extract", 20,
                        Color.FromArgb(37, 99, 235), Color.FromArgb(29, 78, 216), Color.FromArgb(30, 58, 138),
                        Color.White, Color.White);
                }

                // DataGridView
                if (dgv1 != null)
                {
                    clsUiIcons.StyleDataGridView(dgv1);
                    dgv1.AllowUserToResizeColumns = true;
                }

                // 14 Stat Action Buttons with Wow Gradients & Contextual Icons:

                // 1. Dettaglio per articoli (btnArt) - Blue
                clsUiIcons.StyleStatButton(btnArt, "Dettaglio per articoli", "articles", 22,
                    Color.FromArgb(238, 246, 255), Color.FromArgb(198, 220, 252), Color.FromArgb(96, 165, 250),
                    Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));

                // 2. Dettaglio per reparto (button2) - Cyan
                clsUiIcons.StyleStatButton(button2, "Dettaglio per reparto", "department", 22,
                    Color.FromArgb(236, 254, 255), Color.FromArgb(165, 243, 252), Color.FromArgb(56, 189, 248),
                    Color.FromArgb(22, 78, 99), Color.FromArgb(2, 132, 199));

                // 3. Dettaglio per merceologia (button10) - Teal
                clsUiIcons.StyleStatButton(button10, "Dettaglio per merceologia", "categories", 22,
                    Color.FromArgb(240, 253, 250), Color.FromArgb(153, 246, 228), Color.FromArgb(45, 212, 191),
                    Color.FromArgb(19, 78, 74), Color.FromArgb(13, 148, 136));

                // 4. Estrazioni settimanali (button12) - Sky Blue
                clsUiIcons.StyleStatButton(button12, "Estrazioni settimanali", "calendar_week", 22,
                    Color.FromArgb(240, 249, 255), Color.FromArgb(186, 230, 253), Color.FromArgb(56, 189, 248),
                    Color.FromArgb(12, 74, 110), Color.FromArgb(2, 132, 199));

                // 5. Dettaglio per IVA (button9) - Amber / Gold
                clsUiIcons.StyleStatButton(button9, "Dettaglio per IVA", "discount", 22,
                    Color.FromArgb(255, 253, 245), Color.FromArgb(254, 230, 138), Color.FromArgb(251, 191, 36),
                    Color.FromArgb(120, 53, 15), Color.FromArgb(217, 119, 6));

                // 6. Dettaglio per scontrino (button1) - Emerald Green
                clsUiIcons.StyleStatButton(button1, "Dettaglio per scontrino", "receipt", 22,
                    Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208), Color.FromArgb(52, 211, 153),
                    Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));

                // 7. Premi fidelity (button3) - Violet Purple
                clsUiIcons.StyleStatButton(button3, "Premi fidelity", "gift_award", 22,
                    Color.FromArgb(245, 243, 255), Color.FromArgb(221, 214, 254), Color.FromArgb(167, 139, 250),
                    Color.FromArgb(76, 29, 149), Color.FromArgb(124, 58, 237));

                // 8. Movimenti punti fidelity (button7) - Violet Deep
                clsUiIcons.StyleStatButton(button7, "Movimenti punti fidelity", "fidelity_card", 22,
                    Color.FromArgb(245, 243, 255), Color.FromArgb(221, 214, 254), Color.FromArgb(167, 139, 250),
                    Color.FromArgb(76, 29, 149), Color.FromArgb(124, 58, 237));

                // 9. Movimenti cassa (button4) - Forest Green
                clsUiIcons.StyleStatButton(button4, "Movimenti cassa", "pos_balance", 22,
                    Color.FromArgb(240, 253, 244), Color.FromArgb(187, 247, 208), Color.FromArgb(74, 222, 128),
                    Color.FromArgb(20, 83, 45), Color.FromArgb(22, 163, 74));

                // 10. Movimenti cassa 2 (button8) - Forest Green
                clsUiIcons.StyleStatButton(button8, "Movimenti cassa 2", "pos_balance", 22,
                    Color.FromArgb(240, 253, 244), Color.FromArgb(187, 247, 208), Color.FromArgb(74, 222, 128),
                    Color.FromArgb(20, 83, 45), Color.FromArgb(22, 163, 74));

                // 11. Rotazione per articoli (button5) - Coral / Orange
                clsUiIcons.StyleStatButton(button5, "Rotazione per articoli", "refresh", 22,
                    Color.FromArgb(255, 247, 237), Color.FromArgb(254, 215, 170), Color.FromArgb(251, 146, 60),
                    Color.FromArgb(124, 45, 18), Color.FromArgb(234, 88, 12));

                // 12. Movimentazione per articolo (button6) - Orange Warm
                clsUiIcons.StyleStatButton(button6, "Movimentazione articolo", "inventory_warehouse", 22,
                    Color.FromArgb(255, 247, 237), Color.FromArgb(254, 215, 170), Color.FromArgb(251, 146, 60),
                    Color.FromArgb(124, 45, 18), Color.FromArgb(234, 88, 12));

                // 13. Celiachia (button13) - Rose / Crimson
                clsUiIcons.StyleStatButton(button13, "Celiachia", "celiachia", 22,
                    Color.FromArgb(255, 241, 242), Color.FromArgb(254, 205, 211), Color.FromArgb(251, 113, 133),
                    Color.FromArgb(136, 19, 55), Color.FromArgb(225, 29, 72));

                // 14. Storico (button11) - Steel Slate
                clsUiIcons.StyleStatButton(button11, "Storico", "history_clock", 22,
                    Color.FromArgb(248, 250, 252), Color.FromArgb(203, 213, 225), Color.FromArgb(148, 163, 184),
                    Color.FromArgb(15, 23, 42), Color.FromArgb(51, 65, 85));

                // Style bottom summary labels
                StyleTotalSummary(label6, lblTotMov, Color.FromArgb(30, 58, 138));
                StyleTotalSummary(label5, lblTotAcq, Color.FromArgb(120, 53, 15));
                StyleTotalSummary(label4, lblTotVen, Color.FromArgb(6, 78, 59));
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmGesStat2.ApplyModernUi", ex.Message);
            }
        }

        private void StyleTotalSummary(Label lblTitle, Label lblVal, Color accent)
        {
            if (lblTitle != null)
            {
                lblTitle.Font = new Font("Segoe UI", 9.0f, FontStyle.Bold);
                lblTitle.ForeColor = Color.FromArgb(71, 85, 105);
            }
            if (lblVal != null)
            {
                lblVal.Font = new Font("Segoe UI", 10.0f, FontStyle.Bold);
                lblVal.BackColor = Color.White;
                lblVal.ForeColor = accent;
            }
        }

        private void frmGesStat2_KeyDown(object sender, KeyEventArgs e)
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
            cTbc.DataPropertyName = "tmp_tre";
            cTbc.Name = "tmp_tre";
            cTbc.Width = 0;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            cTbc.Visible = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_mag";
            cTbc.Name = "Magazzino";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_neg";
            cTbc.Name = "Negozio";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_ndo";
            cTbc.Name = "Documento";
            cTbc.Width = 120;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_day";
            cTbc.Name = "Data";
            cTbc.Width = 120;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_imp";
            cTbc.Name = "Importo";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
        }

        private void FillDati()
        {
            DataTable t = new clsGenTabTmp().TabTmpStatistiche("TabSta");
            dgv1.DataSource = new DataView(t, "", "tmp_day,tmp_odx", DataViewRowState.CurrentRows);

            FillVenduto();
            FillFatture();
            FillMovimenti();
            FillTotali();
        }

        private void FillVenduto()
        {
            DataRow[] j;

            string s = "SELECT * FROM TabNegozi";
            DataTable tNeg = _clsFun.FillTabSql("TabNegozi", s, false, _strConSql);

            DataTable t = ((DataView)dgv1.DataSource).Table;
            DataRow x;

            s = "";
            s = "SELECT ";
            s += "vet_neg, ";
            s += "vet_day, ";
            //s += "SUM(vet_imp) AS vet_imp ";
            s += "SUM(CASE WHEN vet_cau = 'SCO' THEN vet_imp ELSE 0 END) AS vet_imp ";

            s += "FROM GesNegVet ";
            s += "WHERE ";

            if (_strStatVisProf == "N")
                s += "GesNegVet.vet_cau<>'PRO' AND ";

            s += "GesNegVet.vet_cau <> 'MOV' AND ";
            s += "GesNegVet.vet_day >= " + _clsFun.DaySql(dtpIni.Value) + " AND ";
            s += "GesNegVet.vet_day <= " + _clsFun.DaySql(dtpFin.Value) + " ";
            s += "GROUP BY vet_neg, vet_day";
            DataTable tVet = _clsFun.FillTabSql("TabVet", s, false, _strConSqlSta);

            foreach(DataRow y in tVet.Rows)
            {
                x = t.NewRow();
                x["tmp_odx"] = "1";
                x["tmp_day"] = y["vet_day"];
                x["tmp_neg"] = y["vet_neg"];
                x["tmp_tre"] = "POS";
                x["tmp_des"] = "Casse";
                x["tmp_imp"] = y["vet_imp"];

                x["tmp_mag"] = "01";
                j = tNeg.Select("tab_cod='" + y["vet_neg"] + "'");
                if(j.Length > 0 && ((string)j[0]["tab_mag"]).Trim() != "")
                    x["tmp_mag"] = j[0]["tab_mag"];
                t.Rows.Add(x);
            }
        }

        private void FillFatture()
        {
            DataTable t = ((DataView)dgv1.DataSource).Table;
            DataRow x;
            DataRow[] j;
            string s = "";

            s = "SELECT * FROM AnaClienti";
            DataTable tCli = _clsFun.FillTabSql("TabCli", s, false, _strConSql);

            s = "SELECT * FROM AnaFornitori";
            DataTable tFor = _clsFun.FillTabSql("TabFor", s, false, _strConSql);

            s = "SELECT ";
            s += "GesFatTestate.fat_tpd, ";
            s += "GesFatTestate.fat_yfa, ";
            s += "GesFatTestate.fat_nfa, ";
            s += "GesFatTestate.fat_ndo, ";
            s += "GesFatTestate.fat_ddo, ";
            s += "GesFatTestate.fat_neg, ";
            s += "GesFatTestate.fat_cfo, ";
            s += "SUM(GesMovimenti.mov_imp) AS mov_imp ";
            s += "FROM GesFatTestate ";
            s += "LEFT OUTER JOIN GesMovimenti ON GesFatTestate.fat_yfa = GesMovimenti.mov_yfa AND GesFatTestate.fat_nfa = GesMovimenti.mov_nfa ";
            s += "WHERE ";
            s += "(GesFatTestate.fat_ann = 0) AND ";
            s += "GesFatTestate.fat_ddo >= " + _clsFun.DaySql(dtpIni.Value) + " AND ";
            s += "GesFatTestate.fat_ddo <= " + _clsFun.DaySql(dtpFin.Value) + " ";
            s += "GROUP BY GesFatTestate.fat_tpd, GesFatTestate.fat_yfa, GesFatTestate.fat_nfa, GesFatTestate.fat_ndo, GesFatTestate.fat_ddo, GesFatTestate.fat_neg, GesFatTestate.fat_cfo";
            DataTable tFat = _clsFun.FillTabSql(TABMOVFAT, s, false, _strConSql);

            foreach (DataRow y in tFat.Rows)
            {
                x = t.NewRow();
                x["tmp_odx"] = "4";
                x["tmp_key"] = (string)y["fat_tpd"] + "|" + (string)y["fat_yfa"] + "|" + y["fat_nfa"];
                x["tmp_ndo"] = y["fat_ndo"];
                x["tmp_day"] = y["fat_ddo"];
                x["tmp_neg"] = y["fat_neg"];
                x["tmp_des"] = "Fatture";
                if ((string)y["fat_tpd"] == "FA")
                {
                    x["tmp_tre"] = "FAC";
                    j = tFor.Select("for_cod='" + y["fat_cfo"] + "'");
                    if (j.Length > 0)
                        x["tmp_des"] += " acquisto " + j[0]["for_des"]; 
                }
                else
                {
                    x["tmp_tre"] = "FVE";
                    j = tCli.Select("cli_cod='" + y["fat_cfo"] + "'");
                    if (j.Length > 0)
                        x["tmp_des"] += " vendita " + j[0]["cli_des"]; 
                }
                x["tmp_imp"] = y["mov_imp"]; 

                t.Rows.Add(x);
            }
        }

        private void FillMovimenti()
        {
            DataTable t = ((DataView)dgv1.DataSource).Table;
            DataRow x;
            DataRow[] j;
            string s = "";

            s = "SELECT * FROM TabMovCausali";
            DataTable tCau = _clsFun.FillTabSql("TabCau", s, false, _strConSql);

            s = "SELECT ";
            s += "GesMovTestate.mot_ddo, ";
            s += "GesMovTestate.mot_cau, ";
            s += "GesMovTestate.mot_mag, ";
            s += "GesMovTestate.mot_neg, ";
            s += "SUM(GesMovimenti.mov_imp) AS mov_imp ";
            s += "FROM GesMovTestate ";
            s += "LEFT OUTER JOIN GesMovimenti ON GesMovTestate.mot_ymo = GesMovimenti.mov_ymo AND GesMovTestate.mot_nmo = GesMovimenti.mov_nmo ";
            s += "WHERE ";
            s += "(GesMovTestate.mot_ddo >= " + _clsFun.DaySql(dtpIni.Value) + ") AND ";
            s += "(GesMovTestate.mot_ddo <= " + _clsFun.DaySql(dtpFin.Value) + ") AND ";
            s += "(GesMovTestate.mot_ann = 0) ";
            s += "GROUP BY GesMovTestate.mot_ddo, GesMovTestate.mot_cau, GesMovTestate.mot_mag, GesMovTestate.mot_neg";
            DataTable tMot = _clsFun.FillTabSql("TabMot", s, false, _strConSql);

            Boolean bNeg = false;

            foreach (DataRow y in tMot.Rows)
            {
                bNeg = false;

                x = t.NewRow();
                x["tmp_odx"] = "5";
                x["tmp_day"] = y["mot_ddo"];
                x["tmp_tre"] = "MMO";
                x["tmp_des"] = "Movimenti ";

                if (DBNull.Value.Equals(y["mot_mag"]) || ((string)y["mot_mag"]).Trim() == "")
                    y["mot_mag"] = "01";
                x["tmp_mag"] = y["mot_mag"];

                j = tCau.Select("tab_cod='" + y["mot_cau"] + "'");
                if (j.Length > 0)
                {
                    x["tmp_des"] += (string)j[0]["tab_des"];

                    if ((string)j[0]["tab_sgm"] == "+")
                        x["tmp_tre"] = "MPI";
                    if ((string)j[0]["tab_sgm"] == "-")
                        x["tmp_tre"] = "MME";
                }
                if ((string)j[0]["tab_sgm"] == "-")
                    x["tmp_imp"] = (decimal)y["mov_imp"] *-1;
                else
                    x["tmp_imp"] = y["mov_imp"];

                t.Rows.Add(x);
            }
        }

        private void FillTotali()
        {
            DataRow x;
            DataRow[] j;

            DataTable t = ((DataView)dgv1.DataSource).Table;

            DataTable tTot = t.Clone();

            string sTre = "";

            decimal dMov = 0;

            foreach(DataRow y in t.Rows)
            {

                sTre = ((string)y["tmp_tre"]).Substring(0, 2) + "T";

                j = tTot.Select("tmp_tre='" + sTre + "'");
                if(j.Length == 0)
                {
                    x = tTot.NewRow();
                    x["tmp_odx"] = 0;
                    x["tmp_tre"] = sTre;
                    x["tmp_des"] = "Totale " + y["tmp_des"];
                    x["tmp_imp"] = 0;
                    tTot.Rows.Add(x);
                    j = tTot.Select("tmp_tre='" + sTre + "'");
                }

                j[0]["tmp_imp"] = (decimal)j[0]["tmp_imp"] + (decimal)y["tmp_imp"];

            }

            foreach (DataRow y in tTot.Rows)
            {
                t.ImportRow(y);

                if ((string)y["tmp_tre"] == "POT")
                    lblTotVen.Text = ((decimal)y["tmp_imp"]).ToString("##,##0.00");
                else if ((string)y["tmp_tre"] == "FAT")
                    lblTotAcq.Text = ((decimal)y["tmp_imp"]).ToString("##,##0.00");
                else if (((string)y["tmp_tre"]).Substring(0, 1) == "M" && ((string)y["tmp_tre"]).Substring(2, 1) == "T")
                    dMov += (decimal)y["tmp_imp"];
            }

            lblTotMov.Text = (dMov).ToString("##,##0.00");

        }

        private void dgv1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dgv1.Columns[e.ColumnIndex].Name == "tmp_tre")
            {
                if (e.Value != null)
                {
                    string sVal = (string)e.Value;
                    if (sVal.Length >= 3 && sVal.Substring(2, 1) == "T")
                    {
                        dgv1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(254, 240, 138);
                        dgv1.Rows[e.RowIndex].DefaultCellStyle.Font = new Font("Segoe UI", 9.0f, FontStyle.Bold);
                        dgv1.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.FromArgb(113, 63, 18);
                    }
                    else if (sVal == "POS")
                        dgv1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(240, 253, 250);
                    else if (sVal == "FVE")
                        dgv1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(238, 242, 255);
                    else if (sVal == "FAC")
                        dgv1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(254, 249, 195);
                    else if (sVal == "MOV" || sVal == "MME" || sVal == "MPI")
                        dgv1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(255, 241, 242);
                }
            }
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                string sTre = (string)r.Row["tmp_tre"];

                if (sTre == "POS")
                {
                    DateTime dDay = (DateTime)r.Row["tmp_day"];
                    string sNeg = (String)r.Row["tmp_neg"];

                    //frmGesStat f = new frmGesStat();
                    //f._bolAuto = true;
                    //f._dayDay = dDay;
                    //f.ShowDialog();

                    frmGesStatScontrini f = new frmGesStatScontrini();
                    f._strNeg = sNeg;
                    f._dayDay = dDay;
                    f.ShowDialog();
                }
                else if (sTre == "FAC")
                {
                    string s = (string)r.Row["tmp_key"];

                    string[] a = s.Split('|');

                    string sTpd = a[0];
                    string sYfa = a[1];
                    string sNfa = a[2];

                    s = "SELECT * FROM TabDocTpd WHERE tab_cod='" + sTpd + "'";
                    DataTable tTpd = _clsFun.FillTabSql("Tpd", s, false, _strConSql);
                    if (tTpd.Rows.Count == 0)
                        MessageBox.Show("tipo documento '" + sTpd + "' mancante!", "COMPILARE TABELLA TIPO DOCUMENTO FATTURA", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    else
                    {
                        frmGesDocTouch f = new frmGesDocTouch();
                        f._strMovFat = "F";
                        f._strCauTpd = sTpd;
                        f._strMftYea = sYfa;
                        f._strMftNum = sNfa;
                        f.ShowDialog();
                    }
                }

            }

        }

        private void btnArt_Click(object sender, EventArgs e)
        {
            frmGesStatArticoli f = new frmGesStatArticoli();
            f._dayDti = dtpIni.Value;
            f._dayDtf = dtpFin.Value;
            f._strStatVisProf = _strStatVisProf;
            f.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            frmGesStatPremi f = new frmGesStatPremi();
            f._dayStaIni = dtpIni.Value;
            f._dayStaFin = dtpFin.Value;
            f.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            frmGesStatReparto f = new frmGesStatReparto();
            f._dayDti = dtpIni.Value;
            f._dayDtf = dtpFin.Value;
            f._strStatVisProf = _strStatVisProf;
            f.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmGesStatVenDettaglio f = new frmGesStatVenDettaglio();
            f._dayDti = dtpIni.Value;
            f._dayDtf = dtpFin.Value;
            f.ShowDialog();
        }

        private void aggiornamentoCostiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new frmGesStatCostiAggiorna().ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            new frmGesStatRotazione().ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            new frmGesStatMovArticoli().ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            frmGesStatMovCassa f = new frmGesStatMovCassa();
            f._dayDti = dtpIni.Value;
            f._dayDtf = dtpFin.Value;
            f.ShowDialog();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            new frmGesStatMovCassa2().ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            frmGesStatPunti f = new frmGesStatPunti();
            f._dayDti = dtpIni.Value;
            f._dayDtf = dtpFin.Value;
            f.ShowDialog();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            frmGesStatIVA f = new frmGesStatIVA();
            f._dayDti = dtpIni.Value;
            f._dayDtf = dtpFin.Value;
            f.ShowDialog();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            frmGesStatMerceologia f = new frmGesStatMerceologia();
            f._dayDti = dtpIni.Value;
            f._dayDtf = dtpFin.Value;
            f._strStatVisProf = _strStatVisProf;
            f.ShowDialog();

        }

        private void button11_Click(object sender, EventArgs e)
        {
            frmGesStatStorico f = new frmGesStatStorico();
            f.ShowDialog();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            frmGesStatSettimanali f = new frmGesStatSettimanali(); 
            f._strStatVisProf = _strStatVisProf;
            f.ShowDialog();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            frmGesStatCeliachia f = new frmGesStatCeliachia();
            f._dayDti = dtpIni.Value;
            f._dayDtf = dtpFin.Value; 
            f.ShowDialog();
        }
    }
}
