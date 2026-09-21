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
using System.IO;

namespace APOffice
{
    public partial class frmGesOfferta : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public string _strOftYea = "";
        public string _strOftCod = "";
        public string _strOftSta = "";

        private const string TABGESOFT = "GesOffTestate";
        private const string TABGESOFA = "GesOffArticoli";
        private const string TABTABOFF = "TabOfferte";
        private const string TABTABNEG = "TabNegozi";
        private const string TABTABSTA = "TabStatoOfferte";

        private string _strConSql = "";

        public DataSet _dasGen = new DataSet();

        public frmGesOfferta()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmGesOfferta_Load(object sender, EventArgs e)
        {
            this.Show();

            lblOftYea.Text = _strOftYea;
            lblOftCod.Text = _strOftCod;
            _strConSql = _clsFun.ConSql("");
            ApplyModernUi();
            FillTab();
            OffTipo("INI");
            SetDgv1();
            FillDati();
            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmGesOfferta_dgv1");
            txtSeek.Select();
        }

        private void frmGesOfferta_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmGesOfferta_dgv1");
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
                    if (stampaToolStripMenuItem != null)
                        stampaToolStripMenuItem.Image = clsUiIcons.GetIcon("excel", 16);
                    if (importDaTerminalinoToolStripMenuItem != null)
                        importDaTerminalinoToolStripMenuItem.Image = clsUiIcons.GetIcon("terminal", 16);
                    if (invioANegoziToolStripMenuItem != null)
                        invioANegoziToolStripMenuItem.Image = clsUiIcons.GetIcon("send", 16);
                }

                this.BackColor = Color.FromArgb(243, 244, 246);

                if (pnlOft != null)
                {
                    pnlOft.BackColor = Color.FromArgb(248, 250, 252);
                }

                if (groupBox1 != null)
                {
                    groupBox1.BackColor = Color.FromArgb(248, 250, 252);
                    groupBox1.ForeColor = Color.FromArgb(30, 41, 59);
                }

                if (groupBox2 != null)
                {
                    groupBox2.BackColor = Color.FromArgb(248, 250, 252);
                    groupBox2.ForeColor = Color.FromArgb(30, 41, 59);
                }

                if (pnlMxN != null)
                {
                    pnlMxN.BackColor = Color.FromArgb(241, 245, 249);
                }

                // Stile Pulsante Cerca Articolo "..."
                if (btnSeek != null)
                {
                    clsUiIcons.StyleButton(btnSeek, clsUiIcons.GetIcon("search", 14), Color.FromArgb(241, 245, 249), Color.FromArgb(30, 41, 59));
                }

                // Stile Pulsante "Genera etichette"
                if (btnEtiPrn != null)
                {
                    clsUiIcons.StyleStatButton(btnEtiPrn, "Genera etichette", "labels", 16,
                        Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208),
                        Color.FromArgb(52, 211, 153), Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));
                }

                // Stile Pulsante "Attribuzione mix-match"
                if (btnMix != null)
                {
                    clsUiIcons.StyleStatButton(btnMix, "Attribuzione mix-match", "sync", 18,
                        Color.FromArgb(245, 243, 255), Color.FromArgb(221, 214, 254),
                        Color.FromArgb(167, 139, 250), Color.FromArgb(76, 29, 149), Color.FromArgb(124, 58, 237));
                }

                // Badge estetici per Anno, Codice, Numero
                if (lblOftYea != null)
                {
                    lblOftYea.BackColor = Color.FromArgb(239, 246, 255);
                    lblOftYea.ForeColor = Color.FromArgb(30, 58, 138);
                }
                if (lblOftCod != null)
                {
                    lblOftCod.BackColor = Color.FromArgb(239, 246, 255);
                    lblOftCod.ForeColor = Color.FromArgb(30, 58, 138);
                }
                if (lblOftNum != null)
                {
                    lblOftNum.BackColor = Color.FromArgb(241, 245, 249);
                    lblOftNum.ForeColor = Color.FromArgb(51, 65, 85);
                }

                if (dgv1 != null)
                {
                    clsUiIcons.StyleDataGridView(dgv1);
                    dgv1.CellFormatting += dgv1_CellFormatting;
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmGesOfferta.ApplyModernUi", ex.Message);
            }
        }

        private void dgv1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            try
            {
                // Se la riga è annullata (ofa_ann = true), visualizza con colore spento
                if (dgv1.Columns.Contains("Ann.") && dgv1.Rows[e.RowIndex].Cells["Ann."].Value is bool isAnn && isAnn)
                {
                    e.CellStyle.ForeColor = Color.FromArgb(148, 163, 184);
                }

                string colName = dgv1.Columns[e.ColumnIndex].Name;

                // Formattazione Costo, Prezzo, Valore Offerta
                if ((colName == "Costo" || colName == "Prezzo" || colName == "Offerta") && e.Value != null)
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

        private void frmGesOfferta_KeyDown(object sender, KeyEventArgs e)
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
            FillMixMatch();
            Boolean b = Salva();
            if (!b)
            {
                if (MessageBox.Show("Continui l'uscita?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                        == DialogResult.Yes)
                    this.Close();
            }
            else this.Close();
           
        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            dgv1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv1.MultiSelect = false;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ofa_rig";
            cTbc.Name = "Riga";
            cTbc.HeaderText = "Riga";
            cTbc.Width = 45;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ofa_art";
            cTbc.Name = "Articolo";
            cTbc.HeaderText = "Articolo";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "OfaArd";
            cTbc.Name = "Descrizione";
            cTbc.HeaderText = "Descrizione articolo";
            cTbc.Width = 220;
            cTbc.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            cTbc.FillWeight = 150;
            cTbc.MinimumWidth = 150;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "OfaTpd";
            cTbc.Name = "Offerta";
            cTbc.HeaderText = "Tipo offerta";
            cTbc.Width = 95;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ofa_cos";
            cTbc.Name = "Costo";
            cTbc.HeaderText = "Costo";
            cTbc.Width = 65;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ofa_prv";
            cTbc.Name = "Prezzo";
            cTbc.HeaderText = "Prezzo";
            cTbc.Width = 65;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ofa_val";
            cTbc.Name = "Offerta";
            cTbc.HeaderText = "Valore off.";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ofa_xem";
            cTbc.Name = "M";
            cTbc.HeaderText = "M";
            cTbc.Width = 35;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.DefaultCellStyle.Format = "0";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ofa_xen";
            cTbc.Name = "N";
            cTbc.HeaderText = "N";
            cTbc.Width = 35;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.DefaultCellStyle.Format = "0";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ofa_mix";
            cTbc.Name = "Mix Match";
            cTbc.HeaderText = "Mix Match";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.DefaultCellStyle.Format = "0";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ofa_pun";
            cTbc.Name = "Punti fidelity";
            cTbc.HeaderText = "Punti fide";
            cTbc.Width = 65;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "0";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "ofa_ann";
            cCbc.Name = "Ann.";
            cCbc.HeaderText = "Ann.";
            cCbc.Width = 45;
            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "OfaEti";
            cTbc.Name = "Etichetta";
            cTbc.HeaderText = "Etic.";
            cTbc.Width = 45;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            cTbc.ToolTipText = "Etichetta offerta";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ofa_not";
            cTbc.Name = "Note";
            cTbc.HeaderText = "Note";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "OfaMdy";
            cTbc.Name = "Modificato";
            cTbc.Width = 1;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void btnSeek_Click(object sender, EventArgs e)
        {
            frmSeekArt f = new frmSeekArt();
            f.ShowDialog();
            if(f._tabArt != null && f._tabArt.Rows.Count > 0)
            {
                if (!DBNull.Value.Equals(f._tabArt.Rows[0]["tmp_eqp"]) && ((string)f._tabArt.Rows[0]["tmp_eqp"]).Trim() != "")
                    FillArtEqp((string)f._tabArt.Rows[0]["tmp_eqp"]);
                else
                    OffArtNew(f._tabArt, 0);
            }
        }

        private void txtSeek_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                if (txtSeek.Text != "")
                {
                    DataTable t = _clsQry.ArtSeek(txtSeek.Text, "SEEK");
                    if (t != null && t.Rows.Count > 0)
                    {
                        //OffArtNew(t, 0);
                        if (!DBNull.Value.Equals(t.Rows[0]["tmp_eqp"]) && ((string)t.Rows[0]["tmp_eqp"]).Trim() != "")
                            FillArtEqp((string)t.Rows[0]["tmp_eqp"]);
                        else
                            OffArtNew(t, 0);

                        txtOfaVal.Select();
                    }
                }
            }
        }

        private void FillTab()
        {
            string s = "SELECT * FROM TabOfferte";
            DataTable t = _clsFun.FillTabSql(TABTABOFF, s, false, _strConSql);
            if (_dasGen.Tables.IndexOf(t.TableName) >= 0)
                _dasGen.Tables.Remove(t.TableName);
            _dasGen.Tables.Add(t);

            s = "SELECT * FROM TabStato WHERE tab_off=1";
            t = _clsFun.FillTabSql(TABTABSTA, s, false, _strConSql);
            cmbOftSta.DataSource = t;
            cmbOftSta.DisplayMember = "tab_des";
            cmbOftSta.ValueMember = "tab_cod";
            cmbOftSta.SelectedValue = "N";

            s = "SELECT * FROM TabNegozi WHERE tab_ann=0";
            t = _clsFun.FillTabSql(TABTABNEG, s, false, _strConSql);
            DataRow x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = " Non definita";
            t.Rows.Add(x);
            cmbOftNeg.DataSource = t;
            cmbOftNeg.DisplayMember = "tab_des";
            cmbOftNeg.ValueMember = "tab_cod";
            cmbOftNeg.SelectedValue = "";

            s = "SELECT * FROM TabFidCampagne";
            t = _clsFun.FillTabSql("TabCam", s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = " Non definita";
            t.Rows.Add(x);
            cmbOftCam.DataSource = t;
            cmbOftCam.DisplayMember = "tab_des";
            cmbOftCam.ValueMember = "tab_cod";
            cmbOftCam.SelectedValue = "";
        }

        private void FillDati()
        {
            string s = "";
            string sMsg = "";

            s = "SELECT * FROM GesOffTestate ";
             s += "WHERE ";
            if (_strOftCod == "")
                s += "oft_cod='99999'";
            else
                s += "oft_yea='" + _strOftYea + "' AND oft_cod='" + _strOftCod + "'";
            DataTable t = _clsFun.FillTabSql(TABGESOFT, s, false, _strConSql);
            if(t.Rows.Count > 0)
            {
                lblOftYea.Text = (String)t.Rows[0]["oft_yea"];
                lblOftCod.Text = (String)t.Rows[0]["oft_cod"];
                lblOftNum.Text = (String)t.Rows[0]["oft_num"];
                dtpOftDti.Value = (DateTime)t.Rows[0]["oft_dti"];
                dtpOftDtf.Value = (DateTime)t.Rows[0]["oft_dtf"];
                txtOftDes.Text = (String)t.Rows[0]["oft_des"];
                cmbOftCam.SelectedValue = (String)t.Rows[0]["oft_cam"];
                if (_strOftSta != "")
                    t.Rows[0]["oft_sta"] = _strOftSta;
                cmbOftSta.SelectedValue = (String)t.Rows[0]["oft_sta"];
                cmbOftNeg.SelectedValue = (String)t.Rows[0]["oft_neg"];
            }

            s = "";
            s = "SELECT ";
            s += "ofa_yea, ";
            s += "ofa_cod, ";
            s += "ofa_rig, ";
            s += "ofa_art, ";
            s += "ofa_tip, ";
            s += "ofa_cos, ";
            s += "ofa_prv, ";
            s += "ofa_val, ";
            s += "ofa_xem, ";
            s += "ofa_xen, ";
            s += "ofa_mix, ";
            s += "ofa_pun, ";
            s += "ofa_ann, ";
            s += "ofa_not, ";
            s += "TabOfferte.tab_des AS OfaTpd, ";
            s += "AnaArticoli.art_des AS OfaArd, ";
            s += "AnaArticoli.art_sta AS OfaSta ";
            s += "FROM GesOffArticoli ";
            s += "LEFT OUTER JOIN AnaArticoli ON GesOffArticoli.ofa_art = AnaArticoli.art_cod ";
            s += "LEFT OUTER JOIN TabOfferte ON GesOffArticoli.ofa_tip = TabOfferte.tab_cod ";
            s += "WHERE ";
            s += "ofa_yea='" + _strOftYea + "' AND ";
            if (_strOftCod == "")
                s += "ofa_cod='99999'";
            else
                s += "ofa_cod='" + _strOftCod + "'";

            t = _clsFun.FillTabSql(TABGESOFA, s, false, _strConSql);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "OfaMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "OfaEti",
                Caption = "Etichetta",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                DataTable tTmp = _clsQry.ArtSeek((string)y["ofa_art"], "");
                if (tTmp.Rows.Count > 0)
                {
                    y["ofa_prv"] = (decimal)tTmp.Rows[0]["tmp_prv"];

                    if ( (string)y["ofa_tip"] == _clsDef.OFAPRZ && (decimal)y["ofa_prv"] <= (decimal)y["ofa_val"])
                        sMsg += (string)y["ofa_art"] + " " + (string)y["OfaArd"] + " con prezzo <= offerta" + _clsDef.CRLF;
                }
            }

            dgv1.DataSource = new DataView(t, "", "", DataViewRowState.CurrentRows);

            if (sMsg != "")
                MessageBox.Show(sMsg);
        }

        private void OffArtNew(DataTable tabArt, decimal decPro)
        {
            if (tabArt.Rows.Count > 0)
            {
                string sRig = "0001";
                decimal dPrv = 0;
                decimal dOff = 0;
                decimal dXem = 0;
                decimal dXen = 0;

                DataTable t = ((DataView)dgv1.DataSource).Table;

                if(t.Rows.Count > 0)
                {
                    DataView v = new DataView(t, "", "ofa_rig DESC", DataViewRowState.CurrentRows);
                    sRig = (Convert.ToInt16( v[0]["ofa_rig"])+1).ToString("0000");
                    if ((string)t.Rows[0]["ofa_tip"] == OffTipo("COD"))
                    {
                        dPrv = (decimal)v[0]["ofa_prv"];
                        dOff = (decimal)v[0]["ofa_val"];
                        dXem = (decimal)v[0]["ofa_xem"];
                        dXen = (decimal)v[0]["ofa_xen"];
                    }
                }

                DataRow[] j = t.Select("ofa_art='" + tabArt.Rows[0]["tmp_art"] + "'");
                if (j.Length > 0)
                    MessageBox.Show((string)j[0]["ofa_art"] + " " + (string)j[0]["OfaArd"], "ARTICOLO GIA' PRESENTE");
                else
                {
                    
                    DataTable tTmp = _clsQry.OfaSeek((string)tabArt.Rows[0]["tmp_art"], dtpOftDti.Value, dtpOftDtf.Value);
                    if (tTmp.Rows.Count > 0)
                        MessageBox.Show("Articolo già presente in offerta n. " + (string)tTmp.Rows[0]["ofa_yea"] +"-"+ (string)tTmp.Rows[0]["ofa_cod"] + "!");
                    else
                    {
                        if ((decimal)tabArt.Rows[0]["tmp_prv"] == 0)
                            tabArt = _clsQry.ArtSeek((string)tabArt.Rows[0]["tmp_art"], "");

                        DataRow x = t.NewRow();
                        x["ofa_rig"] = sRig;
                        x["ofa_art"] = tabArt.Rows[0]["tmp_art"];
                        x["ofa_cos"] = tabArt.Rows[0]["tmp_cos"];
                        x["ofa_prv"] = tabArt.Rows[0]["tmp_prv"];
                        x["OfaArd"] = tabArt.Rows[0]["tmp_ard"];
                        x["OfaSta"] = tabArt.Rows[0]["tmp_sta"];
                        x["ofa_tip"] = OffTipo("COD");
                        x["OfaTpd"] = OffTipo("DES");
                        x["ofa_val"] = decPro;
                        x["ofa_xem"] = 0;
                        x["ofa_xen"] = 0;
                        x["ofa_mix"] = 0;
                        x["ofa_pun"] = 0;
                        x["ofa_ann"] = false;
                        x["OfaMdy"] = "S";

                        ((DataView)dgv1.DataSource).Table.Rows.Add(x);

                        dgv1.CurrentCell = dgv1.Rows[dgv1.RowCount - 1].Cells[1];
                        dgv1.Rows[dgv1.RowCount - 1].Selected = true;

                        OffRigaAggiorna();

                        if ((decimal)x["ofa_prv"] != dPrv)
                            txtOfaVal.Text = "0";
                    }
                }
                txtOfaVal.Select();
            }
        }

        private string OffTipo(string strTip)
        {
            string sTip = "";
            if (strTip == "COD")
            {
                if (rdb001.Checked)
                    sTip = _clsDef.OFAPRZ;
                if (rdb002.Checked)
                    sTip = _clsDef.OFASCO;
                if (rdb003.Checked)
                    sTip = _clsDef.OFAMXN;
            }
            else if (strTip == "DES")
            {
                if (rdb001.Checked)
                    sTip = rdb001.Text;
                if (rdb002.Checked)
                    sTip = rdb002.Text;
                if (rdb003.Checked)
                    sTip = rdb003.Text;
            }
            else if (strTip == "INI")
            {
                if (_dasGen.Tables[TABTABOFF].Rows.Count > 1)
                {
                    rdb001.Text = (string)_dasGen.Tables[TABTABOFF].Rows[0]["tab_des"];
                    rdb002.Text = (string)_dasGen.Tables[TABTABOFF].Rows[1]["tab_des"];
                    rdb003.Text = (string)_dasGen.Tables[TABTABOFF].Rows[2]["tab_des"];
                }
            }

            return sTip;
        }

        private void rdbOff_Click(object sender, EventArgs e)
        {
            OffRigaAggiorna();
            txtOfaVal.Select();
        }

        private void txtOfaVal_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                TextBox txt = (TextBox)sender;

                OffRigaAggiorna();

                if (txt.Name == "txtOfaVal")
                {
                    if (rdb003.Checked)
                        txtOfaXem.Select();
                    else
                        txtSeek.Select();
                }
                else if (txt.Name == "txtOfaXem")
                    txtOfaXen.Select();
                else if (txt.Name == "txtOfaXen")
                    txtSeek.Select();
            }
        }

        private void OffRigaAggiorna()
        {
            string sRdb = "";
            string sRdd = "";

            pnlMxN.Enabled = false;

            if (rdb001.Checked)
            {
                sRdb = rdb001.Name.Substring(3, 3);
                sRdd = rdb001.Text;
                txtOfaXem.Text = "0";
                txtOfaXen.Text = "0";
            }
            else if (rdb002.Checked)
            {
                sRdb = rdb002.Name.Substring(3, 3);
                sRdd = rdb002.Text;
                txtOfaXem.Text = "0";
                txtOfaXen.Text = "0";
            }
            else if (rdb003.Checked)
            {
                sRdb = rdb003.Name.Substring(3, 3);
                sRdd = rdb003.Text;
                pnlMxN.Enabled = true;
            }

            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                x["ofa_tip"] = sRdb;
                x["OfaTpd"] = sRdd;
                x["ofa_val"] = _clsFun.Txt2Dec(txtOfaVal.Text);
                x["ofa_xem"] = _clsFun.Txt2Dec(txtOfaXem.Text);
                x["ofa_xen"] = _clsFun.Txt2Dec(txtOfaXen.Text);
                //x["ofa_ann"] = x["ofa_ann"];
                x["OfaMdy"] = "S";

            }
        }

        private Boolean OfaCtrl()
        {
            string sMsg = "";
            string sPun = "";

            if (((DataView)dgv1.DataSource).Count == 0)
                sMsg = "Non sono presenti righe valide" + _clsDef.CRLF;
            else if (((DataView)dgv1.DataSource).Count > 0)
                {
                if (txtOftDes.Text == "")
                    txtOftDes.Text = "Off/Art dal " + dtpOftDti.Value.ToString("dd/MM/yyyy") + " al " + dtpOftDtf.Value.ToString("dd/MM/yyyy");
                //sMsg = "Descrizione promozione non definita";
                else if (dtpOftDti.Value > dtpOftDtf.Value)
                    sMsg = "Data inizio maggiore della data fine";

                if (sMsg == "")
                {
                    DataTable t = ((DataView)dgv1.DataSource).ToTable();
                    foreach (DataRow y in t.Rows)
                    {
                        if (!(Boolean)y["ofa_ann"])
                        {
                            if ((string)y["OfaSta"] != _clsDef.STAPRE && ( DBNull.Value.Equals(y["ofa_val"]) || Convert.ToDecimal(y["ofa_val"]) <= 0))
                                sMsg += "Articolo " + (string)y["OfaArd"] + " con valore non definito!" + _clsDef.CRLF;
                            if (Convert.ToString(y["ofa_tip"]) == "")
                                sMsg += "Articolo " + (string)y["OfaArd"] + " con tipo offerta non definito!" + _clsDef.CRLF;
                            if (Convert.ToString(y["ofa_tip"]) == _clsDef.OFAMXN && ((decimal)y["ofa_xem"] <= 0 || (decimal)y["ofa_xen"] <= 0))
                                sMsg += "Articolo " + (string)y["OfaArd"] + " con MxN non corretto!" + _clsDef.CRLF;
                            if (Convert.ToDecimal(y["ofa_pun"]) > 0 && cmbOftCam.SelectedValue.ToString() == "")
                                sPun += "Articolo " + (string)y["OfaArd"] + " con punti ma definizione campagna mancante!" + _clsDef.CRLF;
                        }
                    }
                }
            }
            if(sMsg != "")
                MessageBox.Show(sMsg, "DATI MANCANTI");
            if(sPun != "")
                MessageBox.Show("Sono presenti articoli che assegnano punti alla spesa ma non è stata definita la CAMPAGNA, i punti verranno assegnati a tutti i clienti fidelizzati.", "ARTICOLI CON PUNTI", MessageBoxButtons.OK, MessageBoxIcon.Information);

            return sMsg == "";
        }

        private Boolean Salva()
        {
            Boolean b = OfaCtrl();

            if (cmbOftSta.SelectedValue.ToString() == _clsDef.STACAN)
                b = true;

            b = true;       //Salviamo tutto Seck 20181115
            if (b)
            {
                clsVariazioni clsVar = new clsVariazioni();

                ArrayList aWhe = new ArrayList();
                ArrayList aExl = new ArrayList();

                if (lblOftCod.Text == _clsDef.CODNEW)
                {
                    lblOftCod.Text = _clsFun.NewNum(DateTime.Today.Year.ToString("0000"), clsDefine.enuNumeratori.NumGesOfferte, 3, _strConSql);
                }

                string s = "SELECT * FROM " + TABGESOFT + " WHERE oft_yea='" + lblOftYea.Text + "' AND  oft_cod='" + lblOftCod.Text + "'";
                DataTable t = _clsFun.FillTabSql(TABGESOFT, s, false, _strConSql);

                DataRow y = t.NewRow();
                y["oft_yea"] = lblOftYea.Text;
                y["oft_cod"] = lblOftCod.Text;
                y["oft_num"] = lblOftNum.Text;
                y["oft_des"] = txtOftDes.Text;
                y["oft_dti"] = dtpOftDti.Value;
                y["oft_dtf"] = dtpOftDtf.Value;
                y["oft_sta"] = cmbOftSta.SelectedValue.ToString();
                y["oft_cam"] = cmbOftCam.SelectedValue.ToString();
                y["oft_neg"] = cmbOftNeg.SelectedValue.ToString();

                if (t.Rows.Count == 0)
                    s = _clsFun.SqlInsertRow(TABGESOFT, t, y);
                else
                {
                    aWhe = new ArrayList();
                    aExl = new ArrayList();
                    aWhe.Add("oft_yea");
                    aWhe.Add("oft_cod");
                    s = _clsFun.SqlUpdRow(TABGESOFT, t, t.Rows[0], y, aWhe, aExl);
                }
                if (s != "")
                {
                    _clsFun.SqlWrite(s, _strConSql);
                    _clsFun.FileLog("GesOfferta", lblOftYea.Text +"-"+ lblOftCod.Text, s);
                }

                t.Rows.Add(y);

                if (t.Rows.Count == 2)
                    t.Rows[0].Delete();

                if (_dasGen.Tables.Contains(t.TableName))
                    _dasGen.Tables.Remove(t.TableName);
                _dasGen.Tables.Add(t);
                _dasGen.Tables[TABGESOFT].Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "OftStd",
                    Caption = "Mdy",
                    MaxLength = 50,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                _dasGen.Tables[TABGESOFT].Rows[0]["OftStd"] = cmbOftSta.Text.Trim();

                s = "SELECT * FROM " + TABGESOFA + " WHERE ";
                s += "ofa_yea = '" + lblOftYea.Text + "' AND ";
                s += "ofa_cod = '" + lblOftCod.Text + "' ";
                s += "ORDER BY ofa_art";
                t = _clsFun.FillTabSql(TABGESOFA, s, false, _strConSql);
                DataRow[] j;

                aWhe = new ArrayList();
                aWhe.Add("ofa_yea");
                aWhe.Add("ofa_cod");
                aWhe.Add("ofa_art");
                aExl = new ArrayList();

                foreach (DataRow x in (((DataView)dgv1.DataSource).ToTable()).Rows)
                {
                    if ((string)x["OfaMdy"] == "S")
                    {
                        b = true;
                        x["ofa_yea"] = lblOftYea.Text;
                        x["ofa_cod"] = lblOftCod.Text;

                        s = "ofa_art='" + x["ofa_art"] + "'";
                        j = t.Select(s);
                        if (j.Length > 0)
                            s = _clsFun.SqlUpdRow(TABGESOFA, t, j[0], x, aWhe, aExl);
                        else
                            s = _clsFun.SqlInsertRow(TABGESOFA, t, x);

                        if (s != "")
                        {
                            _clsFun.SqlWrite(s, _strConSql);
                            _clsFun.FileLog("GesOfferta", (string)x["ofa_art"], s);

                            if(cmbOftSta.SelectedValue.ToString() == _clsDef.STAATT)
                                clsVar.Variazioni((string)x["ofa_art"], "Forza", "Offerta " + lblOftCod.Text, _clsDef.VARPOS);
                        }
                    }
                    if ((string)x["OfaEti"] == "S")
                    {
                        clsVar.Variazioni((string)x["ofa_art"], "OFF" + lblOftCod.Text, "Offerta " + lblOftCod.Text, _clsDef.VARETI);
                    }
                }
            }

            return b;
        }

        private void dgv1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;
                x["OfaMdy"] = "S";
            }
        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv1.Columns[e.ColumnIndex].Name == "Ann." && dgv1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv1.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
            }
        }

        private void dgv1_CurrentCellChanged(object sender, EventArgs e)
        {
            string s = "";

            try
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    txtOfaVal.Text = "0";
                    if (_clsFun.Numerico(x["ofa_val"]))
                        txtOfaVal.Text = _clsFun.Dec2Txt(x["ofa_val"], 2);

                    s = (string)x["ofa_tip"];

                    if (s == _clsDef.OFAPRZ)
                        rdb001.Checked = true;
                    else if (s == _clsDef.OFASCO)
                        rdb002.Checked = true;
                    else if (s == _clsDef.OFAMXN)
                        rdb003.Checked = true;

                    txtOfaXem.Text = Convert.ToString(x["ofa_xem"]);
                    txtOfaXen.Text = Convert.ToString(x["ofa_xen"]);

                    lblOfaMix.Text = (string)x["ofa_mix"];

                    //txtMovSco.Text = "0";

                    //txtMovImp.Text = "0";
                    //if (_clsFun.Numerico(x["mov_imp"]))
                    //    txtMovImp.Text = _clsFun.Dec2Txt(x["mov_imp"], 2);

                    //chkOfaAnn.Checked = false;
                    //if (!DBNull.Value.Equals(x["mov_ann"]))
                    //    chkMovAnn.Checked = Convert.ToBoolean(x["mov_ann"]);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private void dgv1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgv1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = (string)x["ofa_art"];
                    f.ShowDialog();

                    DataTable t = f._tabArt;
                    if(t.Rows.Count > 0)
                    {
                        x["OfaSta"] = (string)t.Rows[0]["art_sta"];
                    }

                }
            }
            else if (e.ColumnIndex == 12)
            {
                string s = dgv1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();

                if (s == "")
                    dgv1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "S";
                else
                    dgv1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
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

        private void txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((sender as TextBox).Text != null && (sender as TextBox).Text.Trim() != "")
            {
                if (e.KeyChar == '.')
                    e.KeyChar = ',';
            }
        }

        private void btnMix_Click(object sender, EventArgs e)
        {
            FillMixMatch();
        }

        private void FillMixMatch()
        {
            Salva();

            string s= "";
            string sKey = "";
            int iMix = 0;

            DataView v = (DataView)dgv1.DataSource;

            v.Sort = "ofa_xem,ofa_xen,ofa_val";

            ArrayList aMix = _clsQry.MixMatch();

            foreach(DataRowView y in v)
            {
                if (((string)y["ofa_tip"] == _clsDef.OFAMXN || (string)y["ofa_tip"] == _clsDef.OFASCO) && (decimal)y["ofa_mix"] == 0)
                {
                    s = ((decimal)y["ofa_xem"]).ToString("00");
                    s += ((decimal)y["ofa_xen"]).ToString("00");
                    s += ((decimal)y["ofa_val"]).ToString("000.00");

                    if (s != sKey)
                    {
                        sKey = s;
                        iMix = (int)aMix[0];
                        aMix.RemoveAt(0);
                    }

                    y["ofa_mix"] = iMix;
                    y["OfaMdy"] = "S";
                }
                //else if ((string)y["ofa_tip"] == _clsDef.OFASCO && (decimal)y["ofa_mix"] == 0)
                //{
                //    s = ((decimal)y["ofa_xem"]).ToString("00");
                //    s += ((decimal)y["ofa_xen"]).ToString("00");
                //    s += ((decimal)y["ofa_val"]).ToString("000.00");

                //    if (s != sKey)
                //    {
                //        sKey = s;
                //        iMix = (int)aMix[0];
                //        aMix.RemoveAt(0);
                //    }

                //    y["ofa_mix"] = iMix;
                //    y["OfaMdy"] = "S";
                //}
            }
        }

        private void btnEtiPrn_Click(object sender, EventArgs e)
        {
            Salva();
            FillEtichette();
        }

        private void FillEtichette()
        {
            clsVariazioni cls = new clsVariazioni();
            DataTable t = ((DataView)dgv1.DataSource).Table;

            foreach(DataRow y in t.Rows)
            {
                if(!(Boolean)y["ofa_ann"])
                    cls.Variazioni((string)y["ofa_art"], "OFF" + lblOftCod.Text, "Offerta " + lblOftCod.Text, _clsDef.VARETI);
            }
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {

        }

        private void txtOffSeek_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                string s = txtOffSeek.Text.Trim();

                string sFld = "Descrizione";
                if (_clsFun.Numerico(s))
                    sFld = "Articolo";

                if (s != "" && dgv1.RowCount > 0)
                {
                    Boolean b = false;
                    int iCur = dgv1.CurrentCell.RowIndex;
                    int i = 0;

                    foreach (DataGridViewRow row in dgv1.Rows)
                    {
                        i++;

                        if (i > iCur + 1)
                        {
                            //if (row.Cells["Cliente/Fornitore"].Value.ToString().ToLower().Contains(s.ToLower()))
                            if (row.Cells[sFld].Value.ToString().ToLower().Contains(s.ToLower()))
                            {
                                dgv1.Rows[row.Index].Selected = true;
                                dgv1.FirstDisplayedScrollingRowIndex = row.Index;
                                dgv1.CurrentCell = dgv1.Rows[row.Index].Cells[0];
                                b = true;
                                break;
                            }
                        }
                    }
                    if (!b)
                        dgv1.CurrentCell = dgv1.Rows[0].Cells[0];
                }
            }
        }

        private void importDaTerminalinoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Importazione da terminalino, confermi?", "IMPORT ARTICOLI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                ImpTerm();
        }

        private void ImpTerm()
        {
            string s = "";
            string sTer = "";
            DataTable tOrd = ((DataView)dgv1.DataSource).Table;
            if (tOrd.Rows.Count > 0)
            {
                if ((string)tOrd.Rows[tOrd.Rows.Count - 1]["ofa_art"] == "")
                    tOrd.Rows[tOrd.Rows.Count - 1].Delete();
            }

            //string sDivTip = "";
            //s = "SELECT * FROM TabForImport WHERE tab_cod='" + _strOftFor + "'";
            //DataTable t = _clsFun.FillTabSql(TABTABFIM, s, false, _strConSql);
            //if (t.Rows.Count == 0)
            //    MessageBox.Show("Tabella importazione fornitore non configurata!");
            //else if (DBNull.Value.Equals(t.Rows[0]["tab_tip"]) || ((string)t.Rows[0]["tab_tip"]).Trim() == "")
            //    MessageBox.Show("Tipo divulgazione non configurato in Tabella importazione fornitore!");
            //else
            //{

            Boolean b = true;

            if(b)
            {

                int i = tOrd.Rows.Count;

                frmGesImpTerm f = new frmGesImpTerm();
                f._strDivTip = "";
                f._strTip = "ADV";
                f.ShowDialog();
                if (f._tabImp != null)
                {
                    sTer = f._strTer;

                    foreach (DataRow y in f._tabImp.Rows)
                    {

                        DataTable t = _clsQry.ArtSeek((string)y["ord_art"], "SEEK");
                        if (t != null && t.Rows.Count > 0)
                        {
                            OffArtNew(t, (decimal)y["ord_qta"]);
                            txtOfaVal.Select();
                        }

                    }
                }
            }
        }

        private void FillArtEqp(string strEqp)
        {
            frmAnaArtEquivalenze f = new frmAnaArtEquivalenze();
            f._strEqp = strEqp;
            f._decPrv = 0;
            f._strTip = "Selezione";
            f.ShowDialog();

            if(f._tabTab != null && f._tabTab.Rows.Count > 0)
            {
                DataTable t = new clsGenTabTmp().TabTmpArt("ArtTmp");
                DataRow x = t.NewRow();
                t.Rows.Add(x);

                foreach (DataRow y in f._tabTab.Rows)
                {
                    t.Rows[0]["tmp_art"] = y["art_cod"];
                    OffArtNew(t, (decimal)y["art_prv"]);
                }

            }

        }

        private void stampaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrnOfferta();
        }

        private void OldPrnOfferta()
        {
            DataTable t = ((DataView)dgv1.DataSource).ToTable();

            string sTit = "Offerta " + txtOftDes.Text.Trim() + " dal " + dtpOftDti.Value.ToString("dd/MM/yyyy") + " dal " + dtpOftDtf.Value.ToString("dd/MM/yyyy");
            string sFil = "";

            string s = "";
            s += "ofa_art, Codice, 50, StringLiteral;";
            s += "OfaArd, Descrizione, 180, StringLiteral;";
            s += "OfaTpd, Tipo, 70, StringLiteral;";
            s += "ofa_cos, Costo, 50, Decimal2;";
            s += "ofa_prv, Prezzo, 50, Decimal2;";
            s += "ofa_val, Offerta, 50, Decimal2;";
            //s += "ofa_ann, Ann, 50, string;";

            for (int i = 0; i <= 100; i++)
            {
                sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_Offerta.xls";
                if (File.Exists(sFil))
                {
                    try
                    {
                        File.Delete(sFil);
                        break;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
                else
                    break;
            }

            DataSet ds = new DataSet();

            DataTable tTit = (new clsExcel()).FillTitStr(s, "TabTit");

            ds.Tables.Add(tTit);
            ds.Tables.Add(t);

            //(new clsExcel()).exportToExcel(ds, sFil, sTit, "Totale " + d.ToString("#,###,##0.00"), true);
            (new clsExcel()).exportToExcel(ds, sFil, sTit, "", true);

            ds.Tables.Remove(t);
            ds.Clear();
            ds.Dispose();
        }

        private void PrnOfferta()
        {
            DataTable t = ((DataView)dgv1.DataSource).ToTable();

            string sTit = "Offerta " + txtOftDes.Text.Trim() + " dal " + dtpOftDti.Value.ToString("dd/MM/yyyy") + " dal " + dtpOftDtf.Value.ToString("dd/MM/yyyy");
            string sFil = "";
            string sFoo = "";

            string sFld = "";
            sFld += "ofa_art, Codice, 50, StringLiteral;";
            sFld += "OfaArd, Descrizione, 180, StringLiteral;";
            sFld += "OfaTpd, Tipo, 70, StringLiteral;";
            sFld += "ofa_cos, Costo, 50, Decimal;";
            sFld += "ofa_prv, Prezzo, 50, Decimal;";
            sFld += "ofa_val, Offerta, 50, Decimal;";
            //s += "ofa_ann, Ann, 50, string;";

            //string sFld = "";
            //sFld += "InvL1c, Liv 1, 30, StringLiteral;";
            //sFld += "InvL1d, Liv 1 des, 70, StringLiteral;";
            //sFld += "InvL2c, Liv 2, 30, StringLiteral;";
            //sFld += "InvL2d, Liv 2 des, 90, StringLiteral;";
            //sFld += "InvL3c, Liv 3, 30, StringLiteral;";
            //sFld += "InvL3d, Liv 3 des, 90, StringLiteral;";
            //sFld += "inv_art, Articolo, 50, StringLiteral;";
            //sFld += "InvArd, Descrizione, 180, StringLiteral;";
            //sFld += "InvUmi, UM, 20, StringLiteral;";
            //sFld += "InvIva, IVA, 20, StringLiteral;";

            for (int i = 0; i <= 100; i++)
            {
                sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_Offerta.xls";
                if (File.Exists(sFil))
                {
                    try
                    {
                        File.Delete(sFil);
                        break;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
                else
                    break;
            }

            sFil = "Offerta.xls";

            (new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);

            //DataSet ds = new DataSet();
            //DataTable tTit = (new clsExcel()).FillTitStr(s, "TabTit");
            //ds.Tables.Add(tTit);
            //ds.Tables.Add(t);
            ////(new clsExcel()).exportToExcel(ds, sFil, sTit, "Totale " + d.ToString("#,###,##0.00"), true);
            //(new clsExcel()).exportToExcel(ds, sFil, sTit, "", true);
            //ds.Tables.Remove(t);
            //ds.Clear();
            //ds.Dispose();

            //string sTit = "INVENTARIO AL " + dtpIntDay.Value.ToShortDateString() + " - " + txtIntDes.Text;
            //if (rdbPrnQtaGia.Checked)
            //    sTit += " su giacenza";
            ////string sFil = "";
            //string sFoo = ",,,,,,,'Totali',,,,," + d.ToString("#0.00").Replace(",", ".");

            //string sFld = "";
            //sFld += "InvL1c, Liv 1, 30, StringLiteral;";
            //sFld += "InvL1d, Liv 1 des, 70, StringLiteral;";
            //sFld += "InvL2c, Liv 2, 30, StringLiteral;";
            //sFld += "InvL2d, Liv 2 des, 90, StringLiteral;";
            //sFld += "InvL3c, Liv 3, 30, StringLiteral;";
            //sFld += "InvL3d, Liv 3 des, 90, StringLiteral;";
            //sFld += "inv_art, Articolo, 50, StringLiteral;";
            //sFld += "InvArd, Descrizione, 180, StringLiteral;";
            //sFld += "InvUmi, UM, 20, StringLiteral;";
            //sFld += "InvIva, IVA, 20, StringLiteral;";

            //if (rdbPrnQtaGia.Checked)
            //    sFld += "inv_gia, Q.tà, 40, Decimal;";
            //else
            //    sFld += "inv_qta, Q.tà, 40, Decimal;";

            //sFld += "inv_cos, Prezzo, 40, Decimal;";
            //sFld += "InvImp, Importo, 50, Decimal;";

            //sFil = "Inventario.xls";

            //(new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);
        }

        private void chkOffAnn_CheckedChanged(object sender, EventArgs e)
        {
            string sFilt = "ofa_ann<>1";
            if (chkOffAnn.Checked)
                sFilt = "";

            ((DataView)dgv1.DataSource).RowFilter = sFilt;
        }

        private void invioANegoziToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string  sPar016PathDivNegozi = _clsFun.ParGet(clsDefine.enuParametri.Par016PathDivNegozi, _strConSql);

            if (sPar016PathDivNegozi != "")
            {
                Boolean b = Salva();

                if (MessageBox.Show("Procedi con l'invio?", "DIVULGAZIONE AI NEGOZI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    //ArrayList a = new ArrayList();
                    //a.Add(lblOftYea.Text);
                    //a.Add(lblOftCod.Text);

                    string s = lblOftYea.Text + "," + lblOftCod.Text;

                    new clsVariazioni().DivNegOfferta(sPar016PathDivNegozi, s, progressBar1);
                }
            }
        }
 
    }
}
