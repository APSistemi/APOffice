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
    public partial class frmGesOffTransazione : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private const string TABOFFTRA = "GesOffTransazione";
        private const string TABOFFTRD = "GesOffTranDate";

        private string _strConSql = "";

        DataSet _dasGen = new DataSet();

        public frmGesOffTransazione()
        {
            InitializeComponent(); new clsGesGraph().SetGraph(this, 0);;
            _strConSql = _clsFun.ConSql("");
        }

        private void frmGesOffTransazione_Load(object sender, EventArgs e)
        {
            ApplyModernUi();
            FillTab();
            SetDgv1();
            SetDgv2();
            FillDati();
            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmGesOffTransazione_dgv1");
            clsUiIcons.RestoreGridColumnWidths(dgv2, "frmGesOffTransazione_dgv2");
        }

        private void frmGesOffTransazione_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmGesOffTransazione_dgv1");
            clsUiIcons.SaveGridColumnWidths(dgv2, "frmGesOffTransazione_dgv2");
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
                }

                this.BackColor = Color.FromArgb(243, 244, 246);

                if (panel1 != null)
                {
                    panel1.BackColor = Color.FromArgb(248, 250, 252);
                }

                // 1. Tasto "Nuova offerta"
                if (btnNew != null)
                {
                    clsUiIcons.StyleStatButton(btnNew, "Nuova offerta", "offers_promo", 20,
                        Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254),
                        Color.FromArgb(96, 165, 250), Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));
                }

                // 2. Tasto "Invio a cassa"
                if (btnInv != null)
                {
                    clsUiIcons.StyleStatButton(btnInv, "Invio a cassa", "send", 20,
                        Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208),
                        Color.FromArgb(52, 211, 153), Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));
                }

                // 3. Tasto "Nuova" data attiva
                if (btnDayNew != null)
                {
                    clsUiIcons.StyleStatButton(btnDayNew, "Nuova", "calendar_week", 16,
                        Color.FromArgb(240, 249, 255), Color.FromArgb(186, 230, 253),
                        Color.FromArgb(56, 189, 248), Color.FromArgb(12, 74, 110), Color.FromArgb(2, 132, 199));
                }

                // 4. Tasto "MIX"
                if (btnOtrMix != null)
                {
                    clsUiIcons.StyleButton(btnOtrMix, clsUiIcons.GetIcon("sync", 12), Color.FromArgb(241, 245, 249), Color.FromArgb(30, 41, 59));
                }

                if (dgv1 != null)
                {
                    clsUiIcons.StyleDataGridView(dgv1);
                    dgv1.CellFormatting += dgv1_CellFormatting;
                }

                if (dgv2 != null)
                {
                    clsUiIcons.StyleDataGridView(dgv2);
                    dgv2.CellFormatting += dgv2_CellFormatting;
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmGesOffTransazione.ApplyModernUi", ex.Message);
            }
        }

        private void dgv1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            try
            {
                if (dgv1.Columns.Contains("Ann.") && dgv1.Rows[e.RowIndex].Cells["Ann."].Value is bool isAnn && isAnn)
                {
                    e.CellStyle.ForeColor = Color.FromArgb(148, 163, 184);
                }

                string colName = dgv1.Columns[e.ColumnIndex].Name;
                if ((colName == "Valore/punti" || colName == "Passo" || colName == "Soglia minima valore" || colName == "Soglia minima" || colName == "Soglia massima") && e.Value != null)
                {
                    if (double.TryParse(e.Value.ToString(), out double valDbl))
                    {
                        if (valDbl != 0)
                        {
                            e.Value = valDbl.ToString("N2");
                            e.FormattingApplied = true;
                        }
                    }
                }
            }
            catch { }
        }

        private void dgv2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            try
            {
                if (dgv2.Columns.Contains("Ann.") && dgv2.Rows[e.RowIndex].Cells["Ann."].Value is bool isAnn && isAnn)
                {
                    e.CellStyle.ForeColor = Color.FromArgb(148, 163, 184);
                }
            }
            catch { }
        }

        private void frmGesOffTransazione_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.DoEvents();
            dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);

            Boolean b = Salva();
            if (!b)
            {
                if (MessageBox.Show("Continui l'uscita?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                      == DialogResult.Yes)
                    Esci();
            }
            else 
                Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void SetDgv1()
        {
            DataTable tCam = (DataTable)cmbOtrCam.DataSource;
            DataTable tNeg = (DataTable)cmbOtrNeg.DataSource;
            DataTable tTip = (DataTable)cmbOtrTip.DataSource;
            DataTable tGru = (DataTable)cmbOtrGru.DataSource;

            dgv1.AutoGenerateColumns = false;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            dgv1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv1.MultiSelect = false;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            DataGridViewComboBoxColumn cCmb;
            DataGridViewButtonColumn cBtn;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "otr_cod";
            cTbc.Name = "Codice";
            cTbc.HeaderText = "Codice";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "otr_des";
            cTbc.Name = "Descrizione";
            cTbc.HeaderText = "Descrizione offerta";
            cTbc.Width = 180;
            cTbc.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            cTbc.FillWeight = 140;
            cTbc.MinimumWidth = 140;
            cTbc.MaxInputLength = 30;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            cTbc.ReadOnly = false;
            cTbc.ToolTipText = "Dicitura su scontrino";
            dgv1.Columns.Add(cTbc);

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "otr_gru";
            cCmb.Name = "Gruppo";
            cCmb.HeaderText = "Gruppo";
            cCmb.Width = 110;
            cCmb.DataSource = tGru;
            cCmb.ValueMember = "tab_cod";
            cCmb.DisplayMember = "tab_des";
            dgv1.Columns.Add((DataGridViewColumn)cCmb);

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "otr_neg";
            cCmb.Name = "Negozio";
            cCmb.HeaderText = "Negozio";
            cCmb.Width = 110;
            cCmb.DataSource = tNeg;
            cCmb.ValueMember = "tab_cod";
            cCmb.DisplayMember = "tab_des";
            dgv1.Columns.Add((DataGridViewColumn)cCmb);

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "otr_cam";
            cCmb.Name = "Campagna";
            cCmb.HeaderText = "Campagna";
            cCmb.Width = 110;
            cCmb.DataSource = tCam;
            cCmb.ValueMember = "tab_cod";
            cCmb.DisplayMember = "tab_des";
            dgv1.Columns.Add((DataGridViewColumn)cCmb);

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "otr_tip";
            cCmb.Name = "Tipo";
            cCmb.HeaderText = "Tipo";
            cCmb.Width = 100;
            cCmb.DataSource = tTip;
            cCmb.ValueMember = "tab_cod";
            cCmb.DisplayMember = "tab_des";
            dgv1.Columns.Add((DataGridViewColumn)cCmb);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "otr_val";
            cTbc.Name = "Valore/punti";
            cTbc.HeaderText = "Val/Punti";
            cTbc.Width = 65;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#,##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "otr_pas";
            cTbc.Name = "Passo";
            cTbc.HeaderText = "Passo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#,##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "otr_sgl";
            cTbc.Name = "Soglia minima valore";
            cTbc.HeaderText = "Soglia val.";
            cTbc.Width = 65;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.ToolTipText = "Valore da togliere per il calcolo dell'ammontare su cui conteggiare i punti";
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#,##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "otr_smi";
            cTbc.Name = "Soglia minima";
            cTbc.HeaderText = "Soglia min";
            cTbc.Width = 65;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.ToolTipText = "Valore minimo scontrino da cui conteggiare i punti/sconto";
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#,##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "otr_smx";
            cTbc.Name = "Soglia massima";
            cTbc.HeaderText = "Soglia max";
            cTbc.Width = 65;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#,##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "otr_mix";
            cTbc.Name = "Codice mix";
            cTbc.HeaderText = "Cod mix";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.ToolTipText = "Codice offerta mix-match per Ditron";
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "###";

            cBtn = new DataGridViewButtonColumn();
            cBtn.UseColumnTextForButtonValue = true;
            cBtn.FlatStyle = FlatStyle.System;
            cBtn.Text = "Mix";
            cBtn.HeaderText = "Mix";
            cBtn.Width = 40;
            cBtn.ValueType = typeof(string);
            cBtn.ReadOnly = false;
            cBtn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cBtn.ToolTipText = "Gen variazione";
            dgv1.Columns.Add(cBtn);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "otr_ann";
            cCbc.Name = "Ann.";
            cCbc.HeaderText = "Ann.";
            cCbc.Width = 40;
            dgv1.Columns.Add(cCbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "otr_vof";
            cCbc.Name = "Offerta";
            cCbc.HeaderText = "In Off.";
            cCbc.Width = 45;
            cCbc.ToolTipText = "Calcolo sconto % con prodotti in offerta";
            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.ValueType = typeof(Boolean);
            cTbc.DataPropertyName = "OtrMdy";
            cTbc.Name = "Modificato";
            cTbc.Width = 1;
            cTbc.Visible = false;
            dgv1.Columns.Add(cTbc);
        }

        private void SetDgv2()
        {
            dgv2.AutoGenerateColumns = false;
            dgv2.AllowUserToAddRows = false;
            dgv2.ReadOnly = false;
            dgv2.AllowUserToDeleteRows = false;
            dgv2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv2.MultiSelect = false;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "otd_dti";
            cTbc.Name = "D.inizio";
            cTbc.HeaderText = "D.inizio";
            cTbc.Width = 90;
            cTbc.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            cTbc.FillWeight = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv2.Columns.Add(cTbc);
            dgv2.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yyyy";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "otd_dtf";
            cTbc.Name = "D.fine";
            cTbc.HeaderText = "D.fine";
            cTbc.Width = 90;
            cTbc.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            cTbc.FillWeight = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv2.Columns.Add(cTbc);
            dgv2.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yyyy";

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "otd_ann";
            cCbc.Name = "Ann.";
            cCbc.HeaderText = "Ann.";
            cCbc.Width = 40;
            dgv2.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.ValueType = typeof(Boolean);
            cTbc.DataPropertyName = "OtdMdy";
            cTbc.Name = "Modificato";
            cTbc.Width = 15;
            cTbc.Visible = false;
            dgv2.Columns.Add(cTbc);
        }

        private void FillTab()
        {
            string s = "";
            DataTable t = new DataTable();
            DataRow x;

            t = _clsFun.FillTabSql("TabFidCampagne", "SELECT * FROM TabFidCampagne", false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbOtrCam.DataSource = t;
            cmbOtrCam.DisplayMember = "tab_des";
            cmbOtrCam.ValueMember = "tab_cod";

            t = _clsFun.FillTabSql("TabNegozi", "SELECT * FROM TabNegozi WHERE tab_ann=0", false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbOtrNeg.DataSource = t;
            cmbOtrNeg.DisplayMember = "tab_des";
            cmbOtrNeg.ValueMember = "tab_cod";

            t = _clsFun.FillTabSql("TabOffTransazione", "SELECT * FROM TabOffTransazione", false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbOtrTip.DataSource = t;
            cmbOtrTip.DisplayMember = "tab_des";
            cmbOtrTip.ValueMember = "tab_cod";

            t = _clsFun.FillTabSql("TabFidGruppi", "SELECT * FROM TabFidGruppi", false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbOtrGru.DataSource = t;
            cmbOtrGru.DisplayMember = "tab_des";
            cmbOtrGru.ValueMember = "tab_cod";

            t = new DataTable("TabTca");
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "TabCod",
                Caption = "Codice",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "TabDes",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            x = t.NewRow();
            x["TabCod"] = "T";
            x["TabDes"] = "Transazione";
            t.Rows.Add(x);
            x = t.NewRow();
            x["TabCod"] = "A";
            x["TabDes"] = "Articoli";
            t.Rows.Add(x);
            cmbOtrTca.DataSource = t;
            cmbOtrTca.DisplayMember = "TabDes";
            cmbOtrTca.ValueMember = "TabCod";
            cmbOtrTca.SelectedValue = "T";
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            TranNew();
        }

        private void dgv1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;
                x["OtrMdy"] = "S";
            }
        }

        private void FillDati()
        {
            string s = "SELECT * FROM GesOffTranDate ORDER BY otd_cod, otd_dti";
            DataTable t = _clsFun.FillTabSql(TABOFFTRD, s, false, _strConSql);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "OtdMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            _dasGen.Tables.Add(t);

            s = "SELECT * FROM GesOffTransazione ORDER BY otr_cod";
            t = _clsFun.FillTabSql(TABOFFTRA, s, false, _strConSql);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "OtrMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            dgv1.DataSource = t;
        }

        private void TranNew()
        {
            DataTable t = (DataTable)dgv1.DataSource; 
            DataRow y = t.NewRow();
            string sCod = "001";
            if (t.Rows.Count > 0)
            {
                DataView v = new DataView(t, "", "otr_cod DESC", DataViewRowState.CurrentRows);
                sCod = (Convert.ToInt32(v[0]["otr_cod"]) + 1).ToString("000");
            }
            y["otr_cod"] = sCod;
            y["otr_des"] = "";
            y["otr_gru"] = "";
            y["otr_cam"] = "";
            y["otr_neg"] = "";
            y["otr_dti"] = DateTime.Today;
            y["otr_dtf"] = DateTime.Today;
            y["otr_tip"] = "";
            y["otr_tca"] = "T";
            y["otr_val"] = 0;
            y["otr_pas"] = 0;
            y["otr_sgl"] = 0;
            y["otr_smi"] = 0;
            y["otr_mix"] = 0;
            y["otr_smx"] = 0;
            y["otr_vof"] = false;
            y["otr_ann"] = false;
            y["otr_rep"] = "";
            y["otr_ren"] = "";

            t.Rows.Add(y);
        }

        private Boolean Salva()
        {
            dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);

            Boolean b = true;
            DataRow[] j;

            string s = "SELECT * FROM GesOffTransazione ORDER BY otr_cod";
            DataTable tOtr = _clsFun.FillTabSql(TABOFFTRA, s, false, _strConSql);

            DataTable t = (DataTable)dgv1.DataSource;

            ArrayList aWhe = new ArrayList();
            aWhe.Add("otr_cod");
            ArrayList aExl = new ArrayList();

            foreach(DataRow y in t.Rows)
            {
                if((string)y["OtrMdy"] == "S")
                {
                    j = tOtr.Select("otr_cod='" + y["otr_cod"] + "'");
                    if (j.Length == 0)
                        s = _clsFun.SqlInsertRow(TABOFFTRA, tOtr, y);
                    else
                        s = _clsFun.SqlUpdRow(TABOFFTRA, tOtr, j[0], y, aWhe, aExl);
                    if(s != "")
                        _clsFun.SqlWrite(s, _strConSql);
                }
            }

            aWhe = new ArrayList();
            aWhe.Add("otd_cod");
            aWhe.Add("otd_dti");
            aExl = new ArrayList();

            s = "SELECT * FROM GesOffTranDate ORDER BY otd_cod";
            DataTable tOtd = _clsFun.FillTabSql(TABOFFTRA, s, false, _strConSql);

            t = _dasGen.Tables[TABOFFTRD];

            foreach (DataRow y in t.Rows)
            {
                if ((string)y["OtdMdy"] == "S")
                {
                    if (DBNull.Value.Equals(y["otd_ann"]))
                        y["otd_ann"] = false;

                    Console.WriteLine("xxxxx");

                    if (!DBNull.Value.Equals(y["otd_dti"]) && !DBNull.Value.Equals(y["otd_dtf"]))
                    {
                        s = "";
                        j = tOtd.Select("otd_cod='" + y["otd_cod"] + "' AND otd_dti=" + _clsFun.DayMdb((DateTime)y["otd_dti"]));
                        if (j.Length == 0)
                        {
                            if(!(Boolean)y["otd_ann"])
                                s = _clsFun.SqlInsertRow(TABOFFTRD, tOtd, y);
                        }
                        else
                        {
                            if ((Boolean)y["otd_ann"])
                                s = "DELETE FROM GesOffTranDate WHERE otd_idx=" + Convert.ToString(y["otd_idx"]);
                            else
                                s = _clsFun.SqlUpdRow(TABOFFTRD, tOtd, j[0], y, aWhe, aExl);
                        }
                        if (s != "")
                            _clsFun.SqlWrite(s, _strConSql);
                    }
                }
            }

            s = "SELECT * FROM GesOffTranDate ORDER BY otd_cod";
            tOtd = _clsFun.FillTabSql(TABOFFTRA, s, false, _strConSql);

            foreach(DataRow y in tOtd.Rows)
            {
                j = t.Select("otd_cod='" + y["otd_cod"] + "' AND otd_dti=" + _clsFun.DayMdb((DateTime)y["otd_dti"]));
                if(j.Length == 0)
                {
                    s = "DELETE FROM GesOffTranDate WHERE otd_idx=" + Convert.ToString(y["otd_idx"]);
                    _clsFun.SqlWrite(s, _strConSql);
                }
            }

            return b;
        }

        private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 0)
            {
                if (dgv1.Columns[e.ColumnIndex].Name == "Ann." && dgv1.CurrentCell is DataGridViewCheckBoxCell)
                {
                    dgv1.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
                }
            }
        }
        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 13)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    ArrayList aMix = new clsQuery().MixMatch();
                    x["otr_mix"] = aMix[0];
                    x["OtrMdy"] = "S";
                }
            }
        }

        private void dgv1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.CellStyle.BackColor = Color.Aquamarine;

            if (e.Control is TextBox)
            {
                var txt = e.Control as TextBox;
                if (txt != null)
                {
                    e.Control.KeyPress += new KeyPressEventHandler(txt_KeyPress);
                }
            }
        }

        private void dgv1_CurrentCellChanged(object sender, EventArgs e)
        {
            FillDettaglio();
        }

        private void txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((sender as TextBox).Text != null && (sender as TextBox).Text.Trim() != "")
            {
                if (e.KeyChar == '.')
                    e.KeyChar = ',';
            }
        }

        private void btnInv_Click(object sender, EventArgs e)
        {
            Otr2Pos();
        }

        private void Otr2Pos()
        {
            Salva();
            DataTable t = (DataTable)dgv1.DataSource;
            frmGesVarPos f = new frmGesVarPos();
            f._tabOtr = t;
            f.ShowDialog();
        }

        private void FillDettaglio()    //Boolean bolCls)
        {
            if (dgv1.DataSource != null && ((DataTable)dgv1.DataSource).Rows.Count > 0)
            {
                try
                {
                    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {
                        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                        DataRow x = r.Row;

                        txtOtrCod.Text = (string)x["otr_cod"];
                        txtOtrDes.Text = (string)x["otr_des"];

                        //dtpOtrDti.Value = (DateTime)x["otr_dti"];
                        //dtpOtrDtf.Value = (DateTime)x["otr_dtf"];

                        chkOtrAnn.Checked = (Boolean)x["otr_ann"];
                        chkOtrVof.Checked = (Boolean)x["otr_vof"];

                        cmbOtrNeg.SelectedValue = (string)x["otr_neg"];
                        cmbOtrCam.SelectedValue = (string)x["otr_cam"];
                        cmbOtrGru.SelectedValue = (string)x["otr_gru"];
                        cmbOtrTip.SelectedValue = (string)x["otr_tip"];
                        if (DBNull.Value.Equals(x["otr_tca"]) || (string)x["otr_tca"] != "A")
                            cmbOtrTca.SelectedValue = "T";      //Transazione
                        else
                            cmbOtrTca.SelectedValue = (string)x["otr_tca"];

                        txtOtrVal.Text = Convert.ToString(x["otr_val"]);
                        txtOtrPas.Text = Convert.ToString(x["otr_pas"]);
                        txtOtrSgl.Text = Convert.ToString(x["otr_sgl"]);
                        txtOtrSmx.Text = Convert.ToString(x["otr_smx"]);
                        txtOtrSmi.Text = Convert.ToString(x["otr_smi"]);
                        txtOtrMix.Text = Convert.ToString(x["otr_mix"]);
                        txtOtrRep.Text = Convert.ToString(x["otr_rep"]);
                        txtOtrRen.Text = Convert.ToString(x["otr_ren"]);

                        //DataTable t = _dasGen.Tables[TABOFFTRD];

                        if (_dasGen.Tables.Count > 0)
                        {
                            DataTable t = _dasGen.Tables[TABOFFTRD];
                            DataView v = new DataView(t, "otd_cod='" + txtOtrCod.Text + "'", "otd_dti", DataViewRowState.CurrentRows);
                            dgv2.DataSource = v;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        private void btnOtrMix_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource != null && ((DataTable)dgv1.DataSource).Rows.Count > 0)
            {
                try
                {
                    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {
                        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                        DataRow x = r.Row;

                        ArrayList aMix = new clsQuery().MixMatch();
                        x["otr_mix"] = aMix[0];
                        x["OtrMdy"] = "S";
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        private void txtOtrDes_Validated(object sender, EventArgs e)
        {
            TextBox txt = (TextBox)sender;
            AggRiga(txt.Name, txt.Text);
        }

        private void txtOtrVal_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                TextBox txt = (TextBox)sender;
                AggRiga(txt.Name, txt.Text);
            }
        }

        private void txtOtrDes_Leave(object sender, EventArgs e)
        {
            TextBox txt = (TextBox)sender;
            AggRiga(txt.Name, txt.Text);
        }

        private void txtOtr_TextChanged(object sender, EventArgs e)
        {
            TextBox txt = (TextBox)sender;
            AggRiga(txt.Name, txt.Text);
        }

        private void chkOtrVof_Click(object sender, EventArgs e)
        {
            CheckBox chk = (CheckBox)sender;
            AggRiga(chk.Name, chk.Checked.ToString());
        }

        private void cmbOtrNeg_SelectionChangeCommitted(object sender, EventArgs e)
        {
            ComboBox cmb = (ComboBox)sender;
            AggRiga(cmb.Name, cmb.SelectedValue.ToString());
        }

        private void dtpOtrDti_Leave(object sender, EventArgs e)
        {
            DateTimePicker dtp = (DateTimePicker)sender;
            Console.WriteLine("zzzz");
            AggRiga(dtp.Name, dtp.Value.ToString("yyyyMMdd"));
        }

        private void AggRiga(string strFld, string strVal)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                if (strFld == "txtOtrDes")
                    x["otr_des"] = strVal;
                else if (strFld == "txtOtrRep")
                    x["otr_rep"] = strVal;
                else if (strFld == "txtOtrRen")
                    x["otr_ren"] = strVal;
                else if (strFld == "txtOtrPas" && _clsFun.Numerico(strVal, "0123456789"))
                    x["otr_pas"] = Convert.ToDecimal(strVal);
                else if (strFld == "txtOtrVal" && _clsFun.Numerico(strVal, "0123456789"))
                    x["otr_val"] = Convert.ToDecimal(strVal);
                else if (strFld == "txtOtrSgl" && _clsFun.Numerico(strVal, "0123456789"))
                    x["otr_sgl"] = Convert.ToDecimal(strVal);
                else if (strFld == "txtOtrSmx" && _clsFun.Numerico(strVal, "0123456789"))
                    x["otr_smx"] = Convert.ToDecimal(strVal);
                else if (strFld == "txtOtrSmi" && _clsFun.Numerico(strVal, "0123456789"))
                    x["otr_smi"] = Convert.ToDecimal(strVal);
                else if (strFld == "txtOtrMix" && _clsFun.Numerico(strVal, "0123456789"))
                    x["otr_mix"] = Convert.ToDecimal(strVal);
                else if (strFld == "chkOtrVof" && strVal == "True")
                    x["otr_vof"] = true;
                else if (strFld == "chkOtrVof" && strVal == "False")
                    x["otr_vof"] = false;
                else if (strFld == "chkOtrAnn" && strVal == "True")
                    x["otr_ann"] = true;
                else if (strFld == "chkOtrAnn" && strVal == "False")
                    x["otr_ann"] = false;
                else if (strFld == "cmbOtrNeg")
                    x["otr_neg"] = strVal;
                else if (strFld == "cmbOtrGru")
                    x["otr_gru"] = strVal;
                else if (strFld == "cmbOtrCam")
                    x["otr_cam"] = strVal;
                else if (strFld == "cmbOtrTip")
                    x["otr_tip"] = strVal;
                else if (strFld == "cmbOtrTca")
                    x["otr_tca"] = strVal;
                else if (strFld == "dtpOtrDti")
                    x["otr_dti"] = _clsFun.Str2Day(strVal);
                else if (strFld == "dtpOtrDtf")
                    x["otr_dtf"] = _clsFun.Str2Day(strVal);

                x["OtrMdy"] = "S";
            }
        }

        private void btnDayNew_Click(object sender, EventArgs e)
        {
            DataView v = (DataView)dgv2.DataSource;
            DataRowView r = v.AddNew();
            r["otd_cod"] = txtOtrCod.Text;
        }

        private void dgv2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0 || e.ColumnIndex == 1)
            {
                DateTime d = DateTime.Today;
                //if(!DBNull.Value.Equals(dgv2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value))
                //    d = (DateTime)dgv2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;

                if (!DBNull.Value.Equals(dgv2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value))
                    d = (DateTime)dgv2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;

                frmUtyDay f = new frmUtyDay();
                f.ShowDialog();
                //e.RowIndex = f._strDay;

                if (f._strDay != "")
                {
                    d = Convert.ToDateTime(f._strDay);

                    dgv2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = d;
                    dgv2.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
                    SendKeys.Send("{ENTER}");
                }
            }
            else if (e.ColumnIndex == 2)
            {
                dgv2.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
                SendKeys.Send("{ENTER}");
            }
        }
    }
}
