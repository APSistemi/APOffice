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
    public partial class frmAnaTessera2 : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private const string TABTESANA = "AnaTessere";
        private const string TABTESLEG = "AnaTessLegami";
        private const string TABSTAVET = "GesNegVet";
        private const string TABTABCAU = "TabCauCassa";

        private string _strConSql = "";
        private string _strConSqlSta = "";
        private string _strFidParam = "";

        public string _strCod = "";
        public string _strDes = "";
        public string _strGru = "";
        public DataTable _tabTmp = new DataTable();

        public Boolean _bolCodNew = false;
        public Boolean _bolCodSostituito = false;

        public frmAnaTessera2()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmAnaTessere_Load(object sender, EventArgs e)
        {
            ApplyModernUi();
            this.Text += " Multi";

            _strConSql = _clsFun.ConSql(""); 
            _strConSqlSta = _clsFun.ConSql("3");

            string s = _clsFun.ParGet(clsDefine.enuParametri.Par027Fidelity, _strConSql);
            if(s.Length > 1)
            {
                string[] a = s.Split(',');
                if(a.Length > 1)
                    _strFidParam = a[1];
            }

            SetDgv1();
            SetDgv2();
            FillTabs();
            FillData();
            FillGru();
            FillMov();

            pnlAnag.Enabled = false;
            if(txtTesCod.Text != "")
                pnlAnag.Enabled = true;

            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmAnaTessera2_dgv1");
            clsUiIcons.RestoreGridColumnWidths(dgv2, "frmAnaTessera2_dgv2");
        }

        private void frmAnaTessera2_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmAnaTessera2_dgv1");
            clsUiIcons.SaveGridColumnWidths(dgv2, "frmAnaTessera2_dgv2");
        }

        private void ApplyModernUi()
        {
            try
            {
                if (menuStrip1 != null)
                {
                    menuStrip1.Renderer = clsUiIcons.GetModernMenuRenderer();
                    if (esciToolStripMenuItem != null)
                        esciToolStripMenuItem.Image = clsUiIcons.GetIcon("exit", 16);
                    if (stampaTesseraToolStripMenuItem != null)
                        stampaTesseraToolStripMenuItem.Image = clsUiIcons.GetIcon("print", 16);
                }

                this.BackColor = Color.FromArgb(243, 244, 246);

                if (btnNewTes != null)
                {
                    clsUiIcons.StyleStatButton(btnNewTes, "Nuova", "plus", 14,
                        Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208),
                        Color.FromArgb(52, 211, 153), Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));
                }

                if (btnSos != null)
                {
                    clsUiIcons.StyleStatButton(btnSos, "Sostituzione", "sync", 14,
                        Color.FromArgb(254, 252, 232), Color.FromArgb(254, 240, 138),
                        Color.FromArgb(250, 204, 21), Color.FromArgb(113, 63, 18), Color.FromArgb(202, 138, 4));
                }

                if (btnCollega != null)
                {
                    clsUiIcons.StyleStatButton(btnCollega, "Collega", "link", 14,
                        Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254),
                        Color.FromArgb(96, 165, 250), Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));
                }

                if (dgv1 != null)
                    clsUiIcons.StyleDataGridView(dgv1);
                if (dgv2 != null)
                    clsUiIcons.StyleDataGridView(dgv2);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaTessera2.ApplyModernUi", ex.Message);
            }
        }

        private void frmAnaTessera_KeyDown(object sender, KeyEventArgs e)
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
            Boolean b = Salva();
            _strDes = txtTesDes.Text;
            //_strGru = cmbTesGru.SelectedValue.ToString();
            
            if (!b)
            {
                if (MessageBox.Show("Continui l'uscita?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                      == DialogResult.Yes)
                    this.Close();
            }
            else this.Close();
        }

        private void txtTesCod_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Return)
                CtrlTes(txtTesCod.Text);
        }
    
        private void txtTesCod_Validated(object sender, EventArgs e)
        {
            CtrlTes(txtTesCod.Text);
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
            cTbc.DataPropertyName = "vep_day";
            cTbc.Name = "Data";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Format = "dd/MM/yy";
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_cau";
            cTbc.Name = "Causale";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_ora";
            cTbc.Name = "Ora";
            cTbc.Width = 35;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.DefaultCellStyle.Format = "##.##";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_pos";
            cTbc.Name = "Cassa";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);
            //dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_sco";
            cTbc.Name = "Scontrino";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "TipDes";
            cTbc.Name = "Descrizione";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(DateTime);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VepImp";
            cTbc.Name = "Importo";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0.00";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VepPun";
            cTbc.Name = "Punti";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VepPua";
            cTbc.Name = "Punti articoli";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VepPup";
            cTbc.Name = "Punti premi";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VepArt";
            cTbc.Name = "Articolo";
            cTbc.Width = 170;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VepFid";
            cTbc.Name = "Tessera";
            cTbc.Width = 170;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
        }

        private void SetDgv2()
        {
            string p = "TabFidGruppi";
            string s = "SELECT * FROM TabFidGruppi WHERE tab_ann=0";
            DataTable tGru = _clsFun.FillTabSql(p, s, false, _strConSql);

            dgv2.AutoGenerateColumns = false;
            //dgv2.VirtualMode = true;
            //dgv2.Dock = DockStyle.Fill;
            dgv2.AllowUserToAddRows = false;
            dgv2.ReadOnly = false;
            dgv2.AllowUserToDeleteRows = false;
            //dgv2.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            DataGridViewComboBoxColumn cCmb;
            DataGridViewButtonColumn cBtn;

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "let_gru";
            cCmb.Name = "Gruppo";
            cCmb.Width = 200;
            cCmb.DataSource = tGru;
            cCmb.ValueMember = "tab_cod";
            cCmb.DisplayMember = "tab_des";
            cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            dgv2.Columns.Add((DataGridViewColumn)cCmb);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "let_cod";
            //cTbc.Name = "Codice tesserato";
            //cTbc.Width = 55;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = true;
            //dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "let_tes";
            cTbc.Name = "Tessera";
            cTbc.Width = 120;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv2.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "tes_des";
            //cTbc.Name = "Nome";
            //cTbc.Width = 100;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = false;
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dgv2.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.DataPropertyName = "let_ann";
            cCbc.Name = "Annullato";
            cCbc.Width = 60;
            cCbc.ValueType = typeof(string);
            cCbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv2.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "LetMdy";
            cTbc.Name = "Mdy";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.Visible = false;
            dgv2.Columns.Add(cTbc);

            cBtn = new DataGridViewButtonColumn();
            cBtn.UseColumnTextForButtonValue = true;
            cBtn.FlatStyle = FlatStyle.System;
            cBtn.Text = "Rettifice";
            cBtn.Width = 100;
            cBtn.ValueType = typeof(string);
            cBtn.ReadOnly = false;
            cBtn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cBtn.ToolTipText = "Aggiunta e diminuzione punti";
            dgv2.Columns.Add(cBtn);
        }

        private void FillTabs()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabProfessioni";
            s = "SELECT * FROM " + p + " WHERE tab_ann=0";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbTesPro.DataSource = t;
            cmbTesPro.DisplayMember = "tab_des";
            cmbTesPro.ValueMember = "tab_cod";
            cmbTesPro.SelectedValue = "";

            p = "TabFidGruppi";
            s = "SELECT * FROM TabFidGruppi WHERE tab_ann=0";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbTesGru.DataSource = t;
            cmbTesGru.DisplayMember = "tab_des";
            cmbTesGru.ValueMember = "tab_cod";
            cmbTesGru.SelectedValue = "";

            p = "TabFidCampagne";
            s = "SELECT * FROM TabFidCampagne WHERE tab_ann=0 ORDER BY tab_dti DESC";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.Add(x);
            cmbFidCam.DataSource = t;
            cmbFidCam.DisplayMember = "tab_des";
            cmbFidCam.ValueMember = "tab_cod";
            //cmbFidCam.SelectedValue = "";
        }

        private void FillData()
        {
            if (_strCod == _clsDef.CODNEW)
            {
                //_strCod = "";
                txtTesCod.Text = _strCod;
                txtTesCod.Select();
            }
            else if (_tabTmp.Rows.Count > 0)
            {
                txtTesCod.ReadOnly = true;
                _strCod = (string)_tabTmp.Rows[0]["tmp_tes"];

                string s = "SELECT * FROM AnaTessere WHERE tes_cod = '" + _strCod + "'";
                DataTable t = _clsFun.FillTabSql(TABTESANA, s, true, _strConSql);

                if (t.Rows.Count > 0)
                {
                    txtTesCod.Text = (string)t.Rows[0]["tes_cod"];
                    if (!_bolCodSostituito)
                    {
                        txtTesDes.Text = (string)t.Rows[0]["tes_des"];
                        txtTesInd.Text = (string)t.Rows[0]["tes_ind"];
                        txtTesLoc.Text = (string)t.Rows[0]["tes_loc"];
                        txtTesCap.Text = (string)t.Rows[0]["tes_cap"];
                        txtTesPrv.Text = (string)t.Rows[0]["tes_prv"];
                        txtTesTel.Text = (string)t.Rows[0]["tes_tel"];
                        txtTesCel.Text = (string)t.Rows[0]["tes_cel"];
                        txtTesMai.Text = (string)t.Rows[0]["tes_mai"];

                        if ((string)t.Rows[0]["tes_sex"] == "M")
                            rdbTesSxM.Checked = true;

                        dtpTesDna.Value = (DateTime)t.Rows[0]["tes_dna"];

                        cmbTesPro.SelectedValue = (string)t.Rows[0]["tes_pro"];
                        //cmbTesGru.SelectedValue = (string)t.Rows[0]["tes_gru"];

                        chkTesAnn.Checked = (Boolean)t.Rows[0]["tes_ann"];
                    }
                    txtTesDes.Select();
                }
            }
        }

        private void FillGru()
        {
            string s = "";

            s = "SELECT ";
            s += "AnaTessLegami.let_cod, ";
            s += "AnaTessLegami.let_tes, ";
            s += "AnaTessLegami.let_gru, ";
            s += "AnaTessLegami.let_ann, ";
            s += "TabFidGruppi.tab_des AS GruDes, ";
            s += "TabFidGruppi.tab_cam AS GruCam ";
            s += "FROM AnaTessLegami LEFT OUTER JOIN TabFidGruppi ON TabFidGruppi.tab_cod = AnaTessLegami.let_gru ";
            s += "WHERE (AnaTessLegami.let_cod = '" + txtTesCod.Text + "')";
            DataTable t = _clsFun.FillTabSql(TABTESLEG, s, false, _strConSql);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "LetMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            dgv2.DataSource = t;
        }
 
        private void CtrlTes(string strTes)
        {
            DataRow[] j;
            string s = "";

            //string s = strTes.Substring(0, strTes.Length - 1) + new clsCtrlCodici().FindMod10Digit(strTes.Substring(0, strTes.Length - 1));
            //if (s != strTes)
            //{
            //    MessageBox.Show("Codice non corretto! (" + s + ")");
            //    txtTesCod.Select();
            //}
            //else
            //{
            s = "SELECT * FROM AnaTessere WHERE tes_cod='" + strTes + "'";
            DataTable t = _clsFun.FillTabSql(TABTESANA, s, false, _strConSql);
            if (t.Rows.Count > 0)
            {
                MessageBox.Show("Codice tessera già presente!", "CONTROLLO TESSERA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtTesCod.Select();
            }
            else
            {
                if (txtTesCod.Text != "")
                    pnlAnag.Enabled = true;
            }
            //}
        }

        private void FillMov()
        {
            if (txtTesCod.Text != "")
            {
                DataRow x;
                DataRow[] j = ((DataTable)cmbFidCam.DataSource).Select("tab_cod='" + cmbFidCam.SelectedValue.ToString() + "'");
                if (j.Length > 0)
                {
                    string sCam = (string)j[0]["tab_cod"];
                    DateTime dDti = (DateTime)j[0]["tab_dti"];
                    DateTime dDtf = (DateTime)j[0]["tab_dtf"];
                    string sGru = (string)j[0]["tab_gru"];

                    string s = "SELECT * FROM TabFidGruppi WHERE tab_ann=0";
                    DataTable tGru = _clsFun.FillTabSql("TabFidGruppi", s, false, _strConSql);

                    decimal dPun = 0;
                    decimal dImp = 0;

                    string sTesWhe = "";

                    DataTable t = new DataTable();
                    if (dgv1.DataSource != null)
                        ((DataTable)dgv1.DataSource).Clear(); ;

                    t = (DataTable)dgv2.DataSource;
                    foreach (DataRow y in t.Rows)
                    {
                        if (DBNull.Value.Equals(y["GruCam"]) || ((string)y["GruCam"]) == "")
                        {
                            j = tGru.Select("tab_cod='" + y["let_gru"] + "'");
                            if (j.Length > 0)
                                y["GruCam"] = (string)j[0]["tab_cam"];
                        }
                        if((string)y["let_gru"] == sGru)
                            sTesWhe += " GesNegVep.vep_fid = '" + (string)y["let_tes"] + "' OR ";
                    }

                    if (sTesWhe != "")
                    {
                        sTesWhe = "(" + sTesWhe.Substring(0, sTesWhe.Length - 3) + ")";

                        s = "SELECT * FROM TabCauCassa";
                        DataTable tCau = _clsFun.FillTabSql(TABTABCAU, s, false, _strConSql);

                        s = "SELECT ";
                        s += "GesNegVep.vep_cau, ";
                        s += "GesNegVep.vep_neg, ";
                        s += "GesNegVep.vep_day, ";
                        s += "GesNegVep.vep_ora, ";
                        s += "GesNegVep.vep_pos, ";
                        s += "GesNegVep.vep_sco, ";
                        s += "GesNegVep.vep_cod, ";
                        s += "GesNegVep.vep_tri, ";
                        s += "GesNegVep.vep_tip, ";
                        s += "GesNegVep.vep_off, ";
                        s += "GesNegVep.vep_imp, ";
                        s += "GesNegVep.vep_val, ";
                        //s += "GesNegVep.vep_ppo, ";
                        //s += "GesNegVep.vep_ord, ";
                        s += "GesNegVep.vep_ean, ";
                        s += "GesNegVep.vep_fid, ";
                        s += "GesNegVet.vet_imp, ";
                        s += "GesNegVet.vet_pun ";
                        s += "FROM GesNegVep ";
                        s += "LEFT OUTER JOIN GesNegVet ON ";
                        s += "GesNegVep.vep_sco = GesNegVet.vet_sco AND ";
                        s += "GesNegVep.vep_pos = GesNegVet.vet_pos AND ";
                        s += "GesNegVep.vep_ora = GesNegVet.vet_ora AND ";
                        s += "GesNegVep.vep_day = GesNegVet.vet_day AND ";
                        s += "GesNegVep.vep_neg = GesNegVet.vet_neg AND ";
                        s += "GesNegVep.vep_cau = GesNegVet.vet_cau ";
                        s += "WHERE ";
                        s += "GesNegVep.vep_day >= " + _clsFun.DaySql(dDti) + " AND ";
                        s += "GesNegVep.vep_day <= " + _clsFun.DaySql(dDtf) + " AND ";
                        //s += "GesNegVet.vet_fid = '" + txtTesCod.Text + "' ";
                        s += sTesWhe;
                        s += "ORDER BY GesNegVet.vet_day DESC, GesNegVet.vet_ora DESC";
                        t = _clsFun.FillTabSql(TABSTAVET, s, false, _strConSqlSta);

                        DataTable tVep = t.Clone();
                        tVep.Columns.Add(new DataColumn()
                        {
                            DataType = Type.GetType("System.String"),
                            ColumnName = "TipDes",
                            Caption = "Descrizione",
                            MaxLength = 50,
                            ReadOnly = false,
                            DefaultValue = (String)""
                        });
                        tVep.Columns.Add(new DataColumn()
                        {
                            DataType = Type.GetType("System.Decimal"),
                            ColumnName = "VepPun",
                            Caption = "Punti",
                            ReadOnly = false,
                            DefaultValue = (Decimal)0
                        });
                        tVep.Columns.Add(new DataColumn()
                        {
                            DataType = Type.GetType("System.Decimal"),
                            ColumnName = "VepPua",
                            Caption = "Punti",
                            ReadOnly = false,
                            DefaultValue = (Decimal)0
                        });
                        tVep.Columns.Add(new DataColumn()
                        {
                            DataType = Type.GetType("System.Decimal"),
                            ColumnName = "VepPup",
                            Caption = "Punti premi",
                            ReadOnly = false,
                            DefaultValue = (Decimal)0
                        });
                        tVep.Columns.Add(new DataColumn()
                        {
                            DataType = Type.GetType("System.Decimal"),
                            ColumnName = "VepImp",
                            Caption = "Importo",
                            ReadOnly = false,
                            DefaultValue = (Decimal)0
                        });
                        tVep.Columns.Add(new DataColumn()
                        {
                            DataType = Type.GetType("System.String"),
                            ColumnName = "VepArt",
                            Caption = "Articolo",
                            MaxLength = 100,
                            ReadOnly = false,
                            DefaultValue = (String)""
                        });
                        tVep.Columns.Add(new DataColumn()
                        {
                            DataType = Type.GetType("System.String"),
                            ColumnName = "VepFid",
                            Caption = "Tessera",
                            MaxLength = 100,
                            ReadOnly = false,
                            DefaultValue = (String)""
                        });

                        DataColumn[] keys = new DataColumn[7];
                        keys[0] = tVep.Columns["vep_cau"];
                        keys[1] = tVep.Columns["vep_day"];
                        keys[2] = tVep.Columns["vep_ora"];
                        keys[3] = tVep.Columns["vep_pos"];
                        keys[4] = tVep.Columns["vep_sco"];
                        keys[5] = tVep.Columns["vep_tip"];
                        keys[6] = tVep.Columns["vep_ean"];
                        tVep.PrimaryKey = keys;

                        foreach (DataRow y in t.Rows)
                        {
                            if (((string)y["vep_tip"]).Substring(0, 2) == "PU")
                            {
                                //if ((string)y["vep_tip"] == "PUM")
                                string ss = (string)y["vep_ean"];
                                Console.WriteLine("xxxxxxxx");

                                string sSql = "";
                                sSql = "vep_cau='" + y["vep_cau"] + "' AND ";
                                sSql += "vep_day=" + _clsFun.DayMdb((DateTime)y["vep_day"]) + " AND ";
                                sSql += "vep_ora='" + y["vep_ora"] + "' AND ";
                                sSql += "vep_pos='" + y["vep_pos"] + "' AND ";
                                sSql += "vep_sco='" + y["vep_sco"] + "' AND ";
                                //sSql += "vep_tip='" + y["vep_tip"] + "' AND ";
                                sSql += "vep_fid='" + y["vep_fid"] + "'";
                                j = tVep.Select(sSql);
                                if (j.Length == 0)
                                {
                                    x = tVep.NewRow();
                                    x["vep_neg"] = y["vep_neg"];
                                    x["vep_day"] = y["vep_day"];
                                    x["vep_cau"] = y["vep_cau"];
                                    x["vep_ora"] = y["vep_ora"];
                                    x["vep_pos"] = y["vep_pos"];
                                    x["vep_sco"] = y["vep_sco"];
                                    x["vep_tip"] = y["vep_tip"];
                                    x["vep_ean"] = y["vep_ean"];
                                    x["vep_fid"] = y["vep_fid"];
                                    x["VepPun"] = 0; // y["vet_pun"];
                                    x["VepPua"] = 0;
                                    x["VepPup"] = 0;
                                    x["VepImp"] = y["vet_imp"];
                                    x["VepFid"] = y["vep_fid"];

                                    dImp += (decimal)y["vet_imp"];
                                    //dPun += (decimal)y["vet_pun"];



                                    j = tCau.Select("tab_cod='" + (string)y["vep_cod"] + "'");
                                    if (j.Length > 0)
                                        x["TipDes"] = (string)j[0]["tab_des"];

                                    if ((string)y["vep_cau"] != "PUN" && (string)y["vep_cau"] != "PUM" && (string)y["vep_ean"] != "" && (string)y["vep_ean"] != _clsDef.PUNTITESSERAX)
                                    {
                                        DataTable tArt = _clsQry.SeekEanArt((string)y["vep_ean"]);
                                        if (tArt.Rows.Count > 0)
                                            x["VepArt"] = ((string)tArt.Rows[0]["tmp_ard"]).Trim() + "-" + (string)tArt.Rows[0]["tmp_art"];
                                    }

                                    tVep.Rows.Add(x);
                                }

                                //if ((string)y["vep_tip"] == "RES")
                                //    Console.WriteLine("aaaa");

                                j = tVep.Select(sSql);

                                if ((string)y["vep_tip"] == "PUN")
                                {
                                    j[0]["VepPun"] = (decimal)j[0]["VepPun"] + (decimal)y["vep_imp"];
                                    dPun += (decimal)j[0]["VepPun"];
                                    //dImp += (decimal)y["vet_imp"];

                                }
                                else if ((string)y["vep_tip"] == "PUM")
                                {
                                    if (((string)y["vep_ean"]).Trim() != "" && (string)y["vep_ean"] != _clsDef.PUNTITESSERAX)
                                    {
                                        j[0]["VepPup"] = (decimal)j[0]["VepPup"] + (decimal)y["vep_imp"];
                                        //dPun += (decimal)y["vep_imp"];
                                        //dImp += (decimal)y["vet_imp"];

                                        if(((string)j[0]["VepArt"]).Trim() != "")
                                            j[0]["VepArt"] = "Più Articoli ";
                                        else
                                        {
                                            if ((string)y["vep_cau"] != "PUN" && (string)y["vep_cau"] != "PUM" && (string)y["vep_ean"] != "" && (string)y["vep_ean"] != _clsDef.PUNTITESSERAX)
                                            {
                                                DataTable tArt = _clsQry.SeekEanArt((string)y["vep_ean"]);
                                                if (tArt.Rows.Count > 0)
                                                    j[0]["VepArt"] = ((string)tArt.Rows[0]["tmp_ard"]).Trim() + "-" + (string)tArt.Rows[0]["tmp_art"];
                                            }
                                        }
                                    }
                                    else
                                    {

                                        //if ((string)y["vep_cau"] == "PUN")
                                        //{
                                        //    j[0]["VepPun"] = (decimal)j[0]["VepPun"] + (decimal)y["vep_imp"];
                                        //    dPun -= (decimal)j[0]["VepPun"];
                                        //}
                                        //else
                                        //{
                                        //    j[0]["VepPun"] = (decimal)j[0]["VepPun"] + (decimal)y["vep_imp"];
                                        //    dPun += (decimal)j[0]["VepPun"];
                                        //    //dImp += (decimal)y["vet_imp"];
                                        //}

                                        DataRow[] jj = tCau.Select("tab_cod='" + (string)y["vep_cod"] + "'");
                                        if (jj.Length > 0)
                                        {
                                            if((string)jj[0]["tab_sgn"] == "-")
                                            {
                                                dPun -= (decimal)y["vep_imp"];
                                                j[0]["VepPun"] = (decimal)j[0]["VepPun"] - (decimal)y["vep_imp"];
                                            }
                                            else
                                            {
                                                dPun += (decimal)y["vep_imp"];
                                                j[0]["VepPun"] = (decimal)j[0]["VepPun"] + (decimal)y["vep_imp"];
                                            }
                                        }

                                    }
                                }
                                else
                                    j[0]["VepPua"] = (decimal)j[0]["VepPua"] + (decimal)y["vep_imp"];



                                //if ((string)y["vep_tip"] == "PAG" || (string)y["vep_tip"] == "SCT" || (string)y["vep_tip"] == "RES")
                                //{
                                //    j[0]["VepImp"] = (decimal)j[0]["VepImp"] + (decimal)y["vep_imp"];
                                //    if ((string)y["vep_tip"] == "PAG")
                                //        dImp += (decimal)y["vep_imp"];
                                //    else
                                //        dImp -= (decimal)y["vep_imp"];
                                //}
                                //else if (((string)y["vep_cau"]).Trim() != "PUN" && ((string)y["vep_ean"]).Trim() != "")
                                //{
                                //    j[0]["VepPua"] = (decimal)j[0]["VepPua"] + (decimal)y["vep_imp"];
                                //    //dPun += (decimal)y["vep_imp"]; NON CONTEGGIATI PERCHE' GIA' COMPRESI NEL TOTALE SCONTRINO
                                //}
                                //else if ((string)y["vep_tip"] == "PUM")
                                //{
                                //    j[0]["VepPun"] = (decimal)j[0]["VepPun"] + (decimal)y["vep_imp"];
                                //    if ((decimal)y["vep_imp"] > 0)
                                //        dPun -= (decimal)y["vep_imp"];
                                //    else
                                //        dPun += (decimal)y["vep_imp"];
                                //}
                                //else if ((string)y["vep_tip"] == "PUN")
                                //{
                                //    j[0]["VepPun"] = (decimal)j[0]["VepPun"] + (decimal)y["vep_imp"];
                                //    dPun += (decimal)y["vep_imp"];
                                //}
                            }
                        }

                        dgv1.DataSource = tVep;
                    }

                    lblTotPun.Text = dPun.ToString("##,###,##0");
                    lblTotImp.Text = dImp.ToString("##,###,##0.00");
                }
            }
        }

        private void btnSos_Click(object sender, EventArgs e)
        {
            frmUtyTessSostituzione f = new frmUtyTessSostituzione();
            f._strOldTes = txtTesCod.Text;
            f._strOldDes = txtTesDes.Text;
            f._strOldPun = lblTotPun.Text;
            f.ShowDialog();
            if (f._bolFatto)
            {
                _strCod = f._strNewTes;
                _bolCodSostituito = true;

                //chkTesAnn.Checked = true;

                _tabTmp.Clear();

                DataRow k = _tabTmp.NewRow();
                k["tmp_tes"] = f._strNewTes;
                k["tmp_ted"] = f._strNewDes;
                _tabTmp.Rows.Add(k);

                FillData();
                FillMov();
                FillGru();
            }
        }

        //private void btnRett_Click(object sender, EventArgs e)
        //{
        //    CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
        //    if (cm.Position >= 0)
        //    {
        //        DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
        //        DataRow x = r.Row;

        //        string sTes = (string)x["let_tes"];

        //        RettifichePunti(sTes);
        //        FillMov();
        //    }
        //}

        private void RettifichePunti( string strTes, string strGru)
        {
            string s = "SELECT * FROM TabFidGruppi WHERE tab_cod='" + strGru + "'";
            DataTable tGru = _clsFun.FillTabSql("TabFidGruppi", s, false, _strConSql);
            if (tGru.Rows.Count > 0)
            {
                string sNeg = (string)tGru.Rows[0]["tab_neg"];

                frmUtyTessRettifiche f = new frmUtyTessRettifiche();
                f._strTesNeg = sNeg;
                f._strTesCod = strTes;
                f._strTesCam = cmbFidCam.SelectedValue.ToString();
                f.ShowDialog();
            }
        }
        
        private void txtCollega_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                Collega();
        }

        private void btnCollega_Click(object sender, EventArgs e)
        {
            Collega();
        }

        private void Collega()
        {
            string s = "";
            Boolean b = true;
            DataRow[] j;
            DataRow x;
            string sGru = "";

            if (txtCollega.Text == "")
                MessageBox.Show("Codice tessera non inserito!", "CONTROLLO TESSERA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else if (cmbTesGru.SelectedValue == null || cmbTesGru.SelectedValue.ToString() == "")
                MessageBox.Show("Gruppo non definito!", "CONTROLLO TESSERA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else if (!_clsFun.Numerico(txtCollega.Text))
                MessageBox.Show("Codice tessera non corretto!", "CONTROLLO TESSERA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                s = txtCollega.Text.Substring(0, txtCollega.Text.Length - 1) + new clsCtrlCodici().FindMod10Digit(txtCollega.Text.Substring(0, txtCollega.Text.Length - 1));
                if (s != txtCollega.Text)
                {
                    b = false;
                    MessageBox.Show("Barcode non corretto sul digit! (" + s + ")");
                }

                if (b && _strFidParam != "")
                {
                    string sMsg = "";

                    sGru = cmbTesGru.SelectedValue.ToString();

                    string[] a = _strFidParam.Split('|');

                    foreach(string ss in a)
                    {
                        string[] aa = ss.Split('-');

                        string sGrp = aa[0];
                        string sPrf = aa[1];
                        int iLen = Convert.ToInt16(aa[2]);

                        Console.WriteLine("zzzz");

                        if (sGru == sGrp)
                        {
                            if (txtCollega.Text.Length != iLen || txtCollega.Text.Substring(0, sPrf.Length) != sPrf)
                            {
                                MessageBox.Show("Codice tessera non corretto!", "CONTROLLO TESSERA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                b = false;
                            }
                            else
                                b = true;
                            Console.WriteLine("zzzz");
                            break;
                        }
                    }
                }
                else if (b && txtCollega.Text.Length != 13)
                {
                    MessageBox.Show("Codice tessera non corretto!", "CONTROLLO TESSERA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    b = false;
                }

                if (b)
                {
                    s = "SELECT * FROM AnaTessLegami WHERE let_tes='" + txtCollega.Text + "'";
                    DataTable t = _clsFun.FillTabSql(TABTESLEG, s, false, _strConSql);
                    if (t.Rows.Count > 0)
                    {
                        MessageBox.Show("Codice tessera già presente sul codice del tesserato " + (string)t.Rows[0]["let_cod"] + "!", "CONTROLLO TESSERA", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        if (MessageBox.Show("Continui con lo spostamento della tessera in " + txtTesDes.Text, "SPOSTAMENTO TESSERA", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                        {
                            s = "UPDATE AnaTessLegami SET let_cod='" + txtTesCod.Text + "' WHERE let_tes='" + txtCollega.Text + "'";
                            _clsFun.SqlWrite(s, _strConSql);
                            FillGru();
                        }

                    }
                    else
                    {
                        t = (DataTable)dgv2.DataSource;
                        if (t.Rows.Count > 0)
                        {
                            //sGru = (string)t.Rows[0]["let_gru"];
                            j = t.Select("let_tes='" + txtCollega.Text + "'");
                            if (j.Length > 0)
                            {
                                b = false;
                                MessageBox.Show("Tessera da collegare già presente sul codice del tesserato " + (string)j[0]["let_cod"] + "!", "CONTROLLO TESSERA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        if (b)
                        {
                            x = t.NewRow();
                            x["let_cod"] = txtTesCod.Text;
                            x["let_tes"] = txtCollega.Text;
                            x["let_ann"] = false;
                            x["let_gru"] = cmbTesGru.SelectedValue.ToString();
                            x["GruCam"] = "";
                            x["LetMdy"] = "S";
                            t.Rows.Add(x);
                        }
                    }
                }
            }

        }

        private void cmbFidCam_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillMov();
        }

        private Boolean Salva()
        {
            Boolean b = true;
            string s = "";
            DataRow x;
            DataRow[] j;

            DataTable tLet = (DataTable)dgv2.DataSource;

            if (txtTesDes.Text == "")
            {
                s = "Nome mancante!" + _clsDef.CRLF;
                b = false;
            }
            else if(tLet.Rows.Count == 0)
            {
                s = "Tessere non presenti!" + _clsDef.CRLF;
                b = false;
            }
            else
            {
                foreach (DataRow y in tLet.Rows)
                {
                    if (DBNull.Value.Equals(y["let_gru"]) || ((string)y["let_gru"]).Trim() == "")
                    {
                        s += "Gruppo non definito sulla tessera!" + _clsDef.CRLF;
                        b = false;
                    }
                }
            }

            if (!b)
                MessageBox.Show(s, "CONTROLLO DATI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                //if (txtTesCod.Text == _clsDef.CODNEW)
                //{
                //    txtTesCod.Text = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaClienti, 5, _strConSql);
                //    _strCod = txtTesCod.Text;
                //}

                if (txtTesCod.Text == _clsDef.CODNEW)
                    txtTesCod.Text = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaTessere, 6, _strConSql);

                _strCod = txtTesCod.Text;

                s = "SELECT * FROM AnaTessere WHERE tes_cod = '" + txtTesCod.Text + "'";
                DataTable t = _clsFun.FillTabSql(TABTESANA, s, true, _strConSql);

                x = t.NewRow();

                x["tes_cod"] = txtTesCod.Text;
                x["tes_des"] = txtTesDes.Text;
                x["tes_ind"] = txtTesInd.Text;
                x["tes_loc"] = txtTesLoc.Text;
                x["tes_cap"] = txtTesCap.Text;
                x["tes_prv"] = txtTesPrv.Text;
                x["tes_tel"] = txtTesTel.Text;
                x["tes_cel"] = txtTesCel.Text;
                x["tes_mai"] = txtTesMai.Text;

                if (rdbTesSxM.Checked)
                    x["tes_sex"] = "M";
                else
                    x["tes_sex"] = "F";

                x["tes_dna"] = dtpTesDna.Value;

                x["tes_pro"] = cmbTesPro.SelectedValue;
                //x["tes_gru"] = cmbTesGru.SelectedValue;
                x["tes_gru"] = _clsDef.COD03X;

                x["tes_ann"] = chkTesAnn.Checked;

                if (t.Rows.Count == 0)
                    s = _clsFun.SqlInsertRow(TABTESANA, t, x);
                else
                {
                    ArrayList aWhe = new ArrayList();
                    aWhe.Add("tes_cod");
                    ArrayList aExl = new ArrayList();
                    s = _clsFun.SqlUpdRow(TABTESANA, t, t.Rows[0], x, aWhe, aExl);
                }

                if (s != "")
                {
                    _clsFun.SqlWrite(s, _strConSql);
                    _clsFun.FileLog(TABTESANA, txtTesCod.Text, s);
                }

                s = "SELECT * FROM AnaTessLegami WHERE let_cod = '" + txtTesCod.Text + "'";
                t = _clsFun.FillTabSql(TABTESLEG, s, false, _strConSql);

                //DataTable tLet = (DataTable)dgv2.DataSource;

                foreach (DataRow y in tLet.Rows)
                {
                    if ((string)y["LetMdy"] == "S")
                    {
                        //x = t.NewRow();
                        //x["let_gru"] = y["let_gru"];
                        //x["let_cod"] = y["let_cod"];
                        //x["let_tes"] = y["let_tes"];
                        //x["let_ann"] = y["let_ann"];
                        //x["let_idx"] = "";

                        y["let_cod"] = txtTesCod.Text;

                        s = "let_cod='" + y["let_cod"] + "' AND let_tes='" + y["let_tes"] + "'";
                        j = t.Select(s);
                        if (j.Length > 0)
                        {
                            //y["let_idx"] = y["let_idx"];
                            s = _clsFun.SqlUpdRowIdx(TABTESLEG, t, j[0], y, null);
                        }
                        else
                        {
                            //x["inv_num"] = lblIntNum.Text;
                            s = _clsFun.SqlInsertRow(TABTESLEG, t, y);
                        }
                        if (s != "")
                        {
                            _clsFun.SqlWrite(s, _strConSql);
                            _clsFun.FileLog("Legami tessere", (string)y["let_cod"], s);
                        }
                    }
                }
            }
            return (b);
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                string sNeg = (string)r.Row["vep_neg"];
                string sCau = (string)r.Row["vep_cau"];
                DateTime dDay = (DateTime)r.Row["vep_day"];
                string sOra = (string)r.Row["vep_ora"];
                string sPos = (string)r.Row["vep_pos"];
                string sSco = (string)r.Row["vep_sco"];

                frmGesStatScoDettaglio f = new frmGesStatScoDettaglio();
                f._strNeg = sNeg;
                f._strCau = sCau;
                f._dayDay = dDay;
                f._strOra = sOra;
                f._strPos = sPos;
                f._strSco = sSco;
                f.ShowDialog();
            }
        }

        private void dgv2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;
                x["LetMdy"] = "S";
            }
        }

        private void dgv2_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dgv2.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void dgv2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            dgv2.Rows[e.RowIndex].Cells["Mdy"].Value = "S";
        }

        private void dgv2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 4)
            {
                Boolean b = Salva();

                if (b)
                {
                    CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {
                        DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
                        DataRow x = r.Row;

                        string sTes = (string)x["let_tes"];
                        string sGru = (string)x["let_gru"];

                        RettifichePunti(sTes, sGru);
                        FillMov();

                    }
                }
            }
            else if (e.RowIndex >= 0)
                dgv2.Rows[e.RowIndex].Cells["Mdy"].Value = "S";
        }

        private void stampaTesseraToolStripMenuItem_Click(object sender, EventArgs e)
        {

            CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                string sDes = txtTesDes.Text;
                string sTes = (string)x["let_tes"];
                string sGru = (string)x["GruDes"];
                new clsGenPdfArtEcr().PrnPdfTessere(sGru, sTes, sDes);
            }
        }

        private void btnNewTes_Click(object sender, EventArgs e)
        {
            if(cmbTesGru.SelectedValue == null || cmbTesGru.SelectedValue.ToString() == "")
                MessageBox.Show("Gruppo non definito!", "CONTROLLO CODICE", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                string s = "";
                string sGru = cmbTesGru.SelectedValue.ToString();

                string[] a = _strFidParam.Split('|');

                if (a.Length < 1)
                    MessageBox.Show("Paramentri inserimento aut5omatico tessera non definiti!", "CONTROLLO CODICE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    foreach (string ss in a)
                    {
                        string[] aa = ss.Split('-');
                        string sGrp = aa[0];
                        string sPrf = aa[1];
                        int iLen = Convert.ToInt16(aa[2]);

                        if (sGru == sGrp)
                        {
                            string sPref = (sPrf + "0000000000").Substring(0, 7);

                            s = "SELECT * FROM AnaTessLegami WHERE SUBSTRING(let_tes,1," + sPrf.Length.ToString() + ") = '" + sPrf + "' ORDER BY let_tes DESC";
                            DataTable t = _clsFun.FillTabSql(TABTESANA, s, true, _strConSql);

                            string sCod = "00000";

                            if (t.Rows.Count > 0)
                            {
                                s = (string)t.Rows[0]["let_tes"];
                                sPref = s.Substring(0, 7);
                                sCod = s.Substring(7, 5);
                            }

                            sCod = (Convert.ToInt32(sCod) + 1).ToString("00000");

                            string sTes = sPref + sCod;


                            sTes = sTes + new clsCtrlCodici().FindMod10Digit(sTes);

                            txtCollega.Text = sTes;



                            break;
                        }
                    }
                }
            }

        }

    }
}
