using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmAnaCliente : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private const string TABANA = "AnaClienti";
        private string _strConSql = "";

        public string _strCod = "";
        public string _strDes = "";
        public DataTable _tabTmp = new DataTable();

        public Boolean _bolCodNew = false;

        public frmAnaCliente()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmAnaCliente_Load(object sender, EventArgs e)
        {
            ApplyModernUi();
            _strConSql = _clsFun.ConSql("");
            _clsFun.SqlWrite("DELETE FROM AnaCliDestinazioni WHERE cld_cli='NEW' OR cld_cli=''", _strConSql);
            SetDgv1();
            FillTabs();
            FillData();

            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmAnaCliente_dgv1");
        }

        private void frmAnaCliente_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmAnaCliente_dgv1");
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
                    if (documentiToolStripMenuItem != null)
                        documentiToolStripMenuItem.Image = clsUiIcons.GetIcon("document_edit", 16);
                }

                this.BackColor = Color.FromArgb(243, 244, 246);

                if (btnNewDest != null)
                {
                    clsUiIcons.StyleStatButton(btnNewDest, "Nuova", "plus", 14,
                        Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254),
                        Color.FromArgb(96, 165, 250), Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));
                }

                if (dgv1 != null)
                {
                    clsUiIcons.StyleDataGridView(dgv1);
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaCliente.ApplyModernUi", ex.Message);
            }
        }

        private void frmAnaCliente_KeyDown(object sender, KeyEventArgs e)
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
            dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);

            Boolean b = true;
            b = Salva();
            _strDes = txtCliDes.Text;

            SalvaCliDes();

            if (!b)
            {
                if (MessageBox.Show("Continui l'uscita?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                      == DialogResult.Yes)
                    this.Close();
            }
            else
            {
                this.Close();
            }
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
            dgv1.SelectionMode = DataGridViewSelectionMode.CellSelect;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "cld_cod";
            cTbc.Name = "Riga";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "cld_des";
            cTbc.Name = "Indirizzo";
            cTbc.Width = 600;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 100;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "cld_ann";
            cCbc.Name = "Ann.";
            cCbc.Width = 25;
            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "CldMdy";
            cTbc.Name = "Modificato";
            cTbc.Width = 10;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.Visible = false;
            dgv1.Columns.Add(cTbc);
        }

        private void FillTabs()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabListini";
            s = "SELECT * FROM TabListini WHERE tab_ann=0 ORDER BY tab_cod";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbCliLic.DataSource = t;
            cmbCliLic.DisplayMember = "tab_des";
            cmbCliLic.ValueMember = "tab_cod";
            cmbCliLic.SelectedValue = "";

            cmbCliLir.DataSource = t.Copy();
            cmbCliLir.DisplayMember = "tab_des";
            cmbCliLir.ValueMember = "tab_cod";
            cmbCliLir.SelectedValue = "";

            p = "TabIva";
            s = "SELECT * FROM TabIva WHERE tab_ann=0 ORDER BY tab_cod";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbCliIva.DataSource = t;
            cmbCliIva.DisplayMember = "tab_des";
            cmbCliIva.ValueMember = "tab_cod";
            cmbCliIva.SelectedValue = "";

            p = "TabClienteTipo";
            s = "SELECT * FROM TabClienteTipo WHERE tab_ann=0 ORDER BY tab_cod";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbCliTip.DataSource = t;
            cmbCliTip.DisplayMember = "tab_des";
            cmbCliTip.ValueMember = "tab_cod";
            cmbCliTip.SelectedValue = "";

            p = "TabPagamenti";
            s = "SELECT * FROM TabPagamenti ORDER BY tab_cod";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbCliTpa.DataSource = t;
            cmbCliTpa.DisplayMember = "tab_des";
            cmbCliTpa.ValueMember = "tab_cod";
            cmbCliTpa.SelectedValue = "";
        }

        private void FillData()
        {
            string s = "";
            DataTable t = new DataTable();

            if (_strCod == _clsDef.CODNEW)
            {
                txtCliCod.Text = _clsDef.CODNEW;
                _strCod = _clsDef.CODNEW;
                //dtpcliDti.Value = DateTime.Today;
                //dtpcliDtm.Value = DateTime.Today;
            }
            else
            {
                if (_tabTmp.Rows.Count > 0)
                {
                    _strCod = (string)_tabTmp.Rows[0]["tmp_cli"];
                }

                txtCliCod.Text = _strCod;
                s = "SELECT * FROM AnaClienti WHERE cli_cod = '" + _strCod + "'";
                t = _clsFun.FillTabSql(TABANA, s, true, _strConSql);

                if (t.Rows.Count > 0)
                {
                    txtCliCod.Text = (string)t.Rows[0]["cli_cod"];
                    txtCliDes.Text = (string)t.Rows[0]["cli_des"];
                    txtCliDe2.Text = (string)t.Rows[0]["cli_de2"];

                    txtCliInd.Text = (string)t.Rows[0]["cli_ind"];
                    txtCliNcv.Text = (string)t.Rows[0]["cli_ncv"];
                    txtCliLoc.Text = (string)t.Rows[0]["cli_loc"];
                    txtCliCap.Text = (string)t.Rows[0]["cli_cap"];
                    txtCliPrv.Text = (string)t.Rows[0]["cli_prv"];
                    txtCliNaz.Text = (string)t.Rows[0]["cli_naz"];
                    txtCliSdi.Text = (string)t.Rows[0]["cli_sdi"];
                    txtCliCcp.Text = (string)t.Rows[0]["cli_ccp"];

                    //txtCliIn2.Text = (string)t.Rows[0]["cli_in2"];
                    //txtCliLo2.Text = (string)t.Rows[0]["cli_lo2"];
                    //txtCliCa2.Text = (string)t.Rows[0]["cli_ca2"];
                    //txtCliPr2.Text = (string)t.Rows[0]["cli_pr2"];

                    txtCliPiv.Text = (string)t.Rows[0]["cli_piv"];
                    txtCliCfi.Text = (string)t.Rows[0]["cli_cfi"];

                    txtCliTel.Text = (string)t.Rows[0]["cli_tel"];
                    txtCliCel.Text = (string)t.Rows[0]["cli_cel"];
                    txtCliFax.Text = (string)t.Rows[0]["cli_fax"];
                    txtCliMai.Text = (string)t.Rows[0]["cli_mai"];
                    txtCliPec.Text = (string)t.Rows[0]["cli_pec"];

                    if (!DBNull.Value.Equals(t.Rows[0]["cel_cod"]))
                        txtCelCod.Text = (string)t.Rows[0]["cel_cod"];
                    if (!DBNull.Value.Equals(t.Rows[0]["cel_pin"]))
                        txtCelPin.Text = (string)t.Rows[0]["cel_pin"];

                    t.Rows[0]["cli_iva"] = "";
                    if (!DBNull.Value.Equals(t.Rows[0]["cli_iva"]))
                        cmbCliIva.SelectedValue = (string)t.Rows[0]["cli_iva"];

                    //t.Rows[0]["cli_tip"] = "";
                    if (!DBNull.Value.Equals(t.Rows[0]["cli_tip"]))
                        cmbCliTip.SelectedValue = (string)t.Rows[0]["cli_tip"];

                    txtCliNot.Text = (string)t.Rows[0]["cli_not"];

                    cmbCliLic.SelectedValue = (string)t.Rows[0]["cli_lic"];
                    cmbCliLir.SelectedValue = (string)t.Rows[0]["cli_lir"];
                    cmbCliTpa.SelectedValue = (string)t.Rows[0]["cli_tpa"];

                    s = new string(' ', 10);
                    if (!DBNull.Value.Equals(t.Rows[0]["cli_etc"]))
                        s = ((string)t.Rows[0]["cli_etc"]) + new string(' ', 10).Substring(0, 10);

                    if(s.Substring(0,1) == "S")
                        chkCliEnm.Checked = true;   //Nome sull'etichetta
                    if (s.Substring(1, 1) == "S")
                        chkCliEin.Checked = true;   //Indirizzo sull'etichetta
                    if (s.Substring(2, 1) == "S")
                        chkCliEpv.Checked = true;   //Prezzo vendita sull'etichetta
                    if (s.Substring(3, 1) == "S")
                        chkCliLot.Checked = true;   //Lotto fornitore sull'etichetta
                }
            }

            if (_strCod == _clsDef.CODNEW)
                s = "SELECT * FROM AnaCliDestinazioni WHERE 1=0";
            else
                s = "SELECT * FROM AnaCliDestinazioni WHERE cld_cli='" + _strCod + "' ORDER BY cld_cod";
            t = _clsFun.FillTabSql("AnaCliDestinazioni", s, false, _strConSql);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "CldMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            dgv1.DataSource = t;
        }

        private Boolean Salva()
        {
            Boolean b = true;
            string s = "";
            DataRow y;

            if (txtCliDes.Text == "")
            {
                MessageBox.Show("Nome mancante!");
                return false;
            }

            if (txtCliCod.Text == _clsDef.CODNEW)
            {
                txtCliCod.Text = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaClienti, 5, _strConSql);
                _strCod = txtCliCod.Text;
            }

            if (txtCliPiv.Text != "")
            {
                s = "SELECT * FROM AnaClienti WHERE cli_cod <> '" + txtCliCod.Text + "' AND cli_piv='" + txtCliPiv.Text + "'";
                DataTable tt = _clsFun.FillTabSql("AnaClienti", s, false, _strConSql);
                if (tt.Rows.Count > 0)
                { 
                    MessageBox.Show("Cliente con partita IVA su " + tt.Rows[0]["cli_des"] + "!", "CONTROLLO PARTITA IVA", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //txtCliPiv.Text = "";
                }
            }

            s = "SELECT * FROM AnaClienti WHERE cli_cod = '" + txtCliCod.Text + "'";

            DataTable t = _clsFun.FillTabSql(TABANA, s, true, _strConSql);

            y = t.NewRow();
            y["cli_cod"] = txtCliCod.Text;

            y["cli_des"] = txtCliDes.Text.ToUpper();
            y["cli_de2"] = txtCliDe2.Text.ToUpper();

            y["cli_ind"] = txtCliInd.Text.ToUpper();

            y["cli_ncv"] = txtCliNcv.Text.ToUpper();

            y["cli_loc"] = txtCliLoc.Text.ToUpper();
            y["cli_cap"] = txtCliCap.Text;
            y["cli_prv"] = txtCliPrv.Text.ToUpper();
            y["cli_naz"] = txtCliNaz.Text.ToUpper();
            y["cli_sdi"] = txtCliSdi.Text.ToUpper();
            y["cli_ccp"] = txtCliCcp.Text.ToUpper();

            //y["cli_in2"] = txtCliInd.Text.ToUpper();
            //y["cli_lo2"] = txtCliLoc.Text.ToUpper(); 
            //y["cli_ca2"] = txtCliCap.Text;
            //y["cli_pr2"] = txtCliPrv.Text.ToUpper(); 

            y["cli_piv"] = txtCliPiv.Text;
            y["cli_cfi"] = txtCliCfi.Text;
            y["cli_tel"] = txtCliTel.Text;
            y["cli_cel"] = txtCliCel.Text;
            y["cli_fax"] = txtCliFax.Text;
            y["cli_mai"] = txtCliMai.Text.ToLower();
            y["cli_pec"] = txtCliPec.Text.ToLower();
            y["cli_iva"] = cmbCliIva.SelectedValue.ToString();
            y["cli_tip"] = cmbCliTip.SelectedValue.ToString();

            y["cli_not"] = txtCliNot.Text;

            y["cli_lic"] = cmbCliLic.SelectedValue.ToString();
            y["cli_lir"] = cmbCliLir.SelectedValue.ToString();
            y["cli_tpa"] = cmbCliTpa.SelectedValue.ToString();

            y["cel_cod"] = txtCelCod.Text;
            y["cel_pin"] = txtCelPin.Text;

            //y["cli_enm"] = chkCliEnm.Checked;
            //y["cli_epv"] = chkCliEpv.Checked;

            s = "";
            if (chkCliEnm.Checked)
                s += "S";
            else
                s += "N";
            if (chkCliEin.Checked)
                s += "S";
            else
                s += "N";
            if (chkCliEpv.Checked)
                s += "S";
            else
                s += "N";
            if (chkCliLot.Checked)
                s += "S";
            else
                s += "N";

            y["cli_etc"] = s;

            if (t.Rows.Count == 0)
                s = _clsFun.SqlInsertRow(TABANA, t, y);
            else
            {
                ArrayList aWhe = new ArrayList();
                aWhe.Add("cli_cod");
                ArrayList aExl = new ArrayList();
                s = _clsFun.SqlUpdRow(TABANA, t, t.Rows[0], y, aWhe, aExl);
            }

            if (s != "")
            { 
                _clsFun.SqlWrite(s, _strConSql);
                _clsFun.FileLog("AnaCliente", txtCliCod.Text, s);
            }

            return (b);
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void btnNewDest_Click(object sender, EventArgs e)
        {
            if (txtCliCod.Text == _clsDef.CODNEW)
            {
                if (!Salva())
                    return;
            }
            else
            {
                Salva();
            }

            if (txtCliCod.Text == _clsDef.CODNEW || _strCod == _clsDef.CODNEW)
                return;

            DataTable t = ((DataTable)dgv1.DataSource);

            string sCod = "01";

            if (t.Rows.Count > 0)
            {
                DataView v = new DataView(t, "", "cld_cod DESC", DataViewRowState.CurrentRows);
                sCod = (Convert.ToInt16(v[0]["cld_cod"]) + 1).ToString("00");
            }

            DataRow x = t.NewRow();
            x["cld_cli"] = _strCod;
            x["cld_cod"] = sCod;
            x["cld_des"] = "";
            x["cld_ann"] = false;
            x["CldMdy"] = "S";
            t.Rows.Add(x);

        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dgv1.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
        }

        private void dgv1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                if (e.ColumnIndex == 1)
                {
                    x["CldMdy"] = "S";
                }
            }
        }
        private void dgv1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgv1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private Boolean SalvaCliDes()
        {
            if (_strCod == _clsDef.CODNEW || txtCliCod.Text == _clsDef.CODNEW)
                return true;

            string s = "SELECT * FROM AnaCliDestinazioni WHERE cld_cli='" + _strCod + "'ORDER BY cld_cod";
            string p = "AnaCliDestinazioni";
            DataTable t = _clsFun.FillTabSql("AnaCliDestinazioni", s, false, _strConSql);
            DataRow[] j;
            Boolean b = true;
            DataRow x;

            ArrayList aWhe = new ArrayList();
            aWhe.Add("cld_cli");
            aWhe.Add("cld_cod");
            ArrayList aExl = new ArrayList();

            foreach (DataRow y in ((DataTable)dgv1.DataSource).Rows)
            {
                if ((string)y["CldMdy"] == "S")
                {
                    //Boolean b = true;
                    x = t.NewRow();
                    x["cld_cli"] = _strCod;
                    x["cld_cod"] = y["cld_cod"];
                    x["cld_des"] = y["cld_des"];
                    x["cld_ann"] = y["cld_ann"];

                    s = "cld_cod='" + y["cld_cod"] + "'";
                    j = t.Select(s);
                    if (j.Length > 0)
                        s = _clsFun.SqlUpdRow(p, t, j[0], x, aWhe, aExl);
                    else
                        s = _clsFun.SqlInsertRow(p, t, x);

                    if (s != "")
                        _clsFun.SqlWrite(s, _strConSql);
                }
            }

            return b;
        }

        //private void dgv1_CellLeave(object sender, DataGridViewCellEventArgs e)
        //{
        //    Console.WriteLine("AAAAAAAAAAA");
        //}

        //private void dgv1_Leave(object sender, EventArgs e)
        //{
        //    Console.WriteLine("AAAAAAAAAAA");
        //}

        private void dgv1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Console.WriteLine("AAAAAAAAAAA");

            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                if (e.ColumnIndex == 1)
                {
                    x["CldMdy"] = "S";
                }
            }
        }

        private void txtCliPiv_Validated(object sender, EventArgs e)
        {
            if (!new clsCtrlCodici().ControllaPartitaIva(txtCliPiv.Text))
                MessageBox.Show("Partita IVA non corretta!", "CONTROLLO CODIFICA", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void txtCliCfi_Validated(object sender, EventArgs e)
        {
            if (!new clsCtrlCodici().VerificaCodiceFiscale(txtCliCfi.Text))
                MessageBox.Show("Codice fiscale non corretto!", "CONTROLLO CODIFICA", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void documentiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmAnaCfoDocumenti f = new frmAnaCfoDocumenti();
            f._strCod = txtCliCod.Text;
            f._strDes = txtCliDes.Text;
            f._strCfo = _clsDef.TIPCLI;
            f.ShowDialog();
        }

    }
}
